import time
import os
import json
import re
import urllib3
import requests
import pandas as pd
import datetime
import base64
from bs4 import BeautifulSoup
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
GITHUB_TOKEN = os.environ.get("MY_GITHUB_TOKEN") 
SCRAPER_API_KEY = os.environ.get("SCRAPER_API_KEY")
GITHUB_USER = "speed4458-arch"
REPO_NAME = "saju-lotto-db"
# ==========================================

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}

def clean_text(text):
    if pd.isna(text) or not text: return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def load_fixed_stores(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [1단계] 회원님의 고정 명당(엑셀) 데이터 불러오기...")
    candidates = ["fixed_stores.xlsx", "fixed_stores.csv"]
    target_path = None
    for fname in candidates:
        if os.path.exists(os.path.join(folder_path, fname)):
            target_path = os.path.join(folder_path, fname); break
            
    if not target_path: return pd.DataFrame()
    try:
        if target_path.endswith(".xlsx"): df = pd.read_excel(target_path)
        else:
            try: df = pd.read_csv(target_path, encoding='utf-8-sig')
            except: df = pd.read_csv(target_path, encoding='cp949')
                
        df.columns = [str(c).strip() for c in df.columns]
        rename_map = {}
        for c in df.columns:
            if "지역" in c: rename_map[c] = "지역"
            elif any(k in c for k in ["상호", "판매점", "이름"]): rename_map[c] = "상호명"
            elif any(k in c for k in ["주소", "소재지", "위치"]): rename_map[c] = "상세주소"
            
        df = df.rename(columns=rename_map)
        df["상호명"] = df["상호명"].apply(clean_text)
        df["상세주소"] = df["상세주소"].apply(clean_text)
        print(f"✅ 명당 리스트 {len(df)}곳 로드 완료!")
        return df[['지역', '상호명', '상세주소']]
    except: return pd.DataFrame()

def get_lotto_from_daum(drw):
    try:
        url = f"https://search.daum.net/search?w=tot&q=로또+{drw}회"
        res = requests.get(url, headers=BROWSER_HEADERS, timeout=10)
        soup = BeautifulSoup(res.text, "html.parser")
        
        balls = [int(s.text.strip()) for s in soup.select("span") if s.text.strip().isdigit() and 'ball' in str(s.get('class', ''))]
        
        prize1 = 0
        prize_match = re.search(r'1등\s*당첨금[^\d]*([\d,]+)\s*원', soup.get_text(" "))
        if prize_match:
            prize1 = int(prize_match.group(1).replace(",", ""))
            
        if len(balls) >= 7:
            return {
                "회차": drw, "추첨일": datetime.datetime.now().strftime("%Y-%m-%d"), 
                "번호1": balls[0], "번호2": balls[1], "번호3": balls[2], "번호4": balls[3], "번호5": balls[4], "번호6": balls[5],
                "보너스": balls[6], "1등당첨금": prize1
            }
    except: pass
    return None

def get_latest_draw_and_update_numbers(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [2단계] 역대 당첨번호 최신화...")
    base_date = datetime.date(2026, 10, 24)
    base_drw = 1247
    now = datetime.datetime.now()
    estimated_drw = base_drw + ((now.date() - base_date).days // 7)
    if now.weekday() == 5 and now.hour < 21: estimated_drw -= 1
        
    latest_drw = estimated_drw
    num_file = os.path.join(folder_path, "lotto_numbers_raw.csv")
    history_excel = os.path.join(folder_path, "lotto_history.xlsx")
    parsed_dict = {}
    
    if os.path.exists(history_excel):
        try:
            df_hist = pd.read_excel(history_excel, header=None)
            for idx, row in df_hist.iterrows():
                try:
                    drw = int(row[1])
                    parsed_dict[drw] = {
                        "회차": drw, "추첨일": (base_date - datetime.timedelta(weeks=(base_drw - drw))).strftime("%Y-%m-%d"),
                        "번호1": int(row[2]), "번호2": int(row[3]), "번호3": int(row[4]), 
                        "번호4": int(row[5]), "번호5": int(row[6]), "번호6": int(row[7]), "보너스": int(row[8]), 
                        "1등당첨금": int(str(row[11]).replace(',', '').replace('원', '').strip() if str(row[11]).replace(',', '').replace('원', '').strip().isdigit() else 0)
                    }
                except: continue
        except: pass

    missing_draws = sorted([d for d in range(1, latest_drw + 1) if d not in parsed_dict], reverse=True)
    if missing_draws:
        for drw in missing_draws:
            daum_data = get_lotto_from_daum(drw)
            if daum_data:
                parsed_dict[drw] = daum_data
            time.sleep(0.5)

    df_nums = pd.DataFrame(list(parsed_dict.values())).sort_values(by="회차", ascending=False)
    df_nums.to_csv(num_file, index=False, encoding='utf-8-sig')
    return df_nums

def get_winning_stores(drw):
    print(f"\n🚀 [추가기능] {drw}회차 1등 당첨 판매점 크롤링 시도...")
    stores = []
    try:
        # 동행복권 구형 주소 대신 해외 IP 차단이 없는 공개 당첨판매점 페이지 활용 (ScraperAPI 소모 없음)
        url = f"https://lotto.inforjin.com/stores?drw={drw}"
        res = requests.get(url, headers=BROWSER_HEADERS, timeout=20)
        html = res.text

        if "1등 당첨 판매점" in html:
            sec_html = html.split("1등 당첨 판매점", 1)[1]
            if "2등 당첨 판매점" in sec_html:
                sec_html = sec_html.split("2등 당첨 판매점", 1)[0]
        else:
            sec_html = html

        soup = BeautifulSoup(sec_html, "html.parser")
        lines = [clean_text(line) for line in soup.get_text("\n").splitlines() if clean_text(line)]

        i = 0
        while i < len(lines):
            line = lines[i]
            # 케이스 1: '자동/수동/반자동'이 단독 줄로 분리된 경우
            if line in ("자동", "수동", "반자동") and i >= 1 and i + 1 < len(lines):
                name = re.sub(r'^\d+[\.\s]*', '', lines[i - 1]).strip()
                addr = lines[i + 1].strip()
                if name and name not in ("상호명", "판매점", "구분", "1등 당첨 판매점"):
                    stores.append({"상호명": name, "구분": line, "소재지": addr})
                i += 2
                continue

            # 케이스 2: '상호명 자동' 형태가 한 줄에 붙어있는 경우
            m = re.match(r'^(?:\d+[\.\s]*)?(.+?)\s+(자동|수동|반자동)(?:\.?\s*(.*))?$', line)
            if m:
                name = m.group(1).strip()
                gubun = m.group(2).strip()
                addr = m.group(3).strip() if m.group(3) else (lines[i + 1].strip() if i + 1 < len(lines) else "")
                if name and name not in ("상호명", "판매점", "구분"):
                    stores.append({"상호명": name, "구분": gubun, "소재지": addr})
                i += 2 if not m.group(3) else 1
                continue

            i += 1

        print(f"✅ {drw}회차 1등 당첨 판매점 {len(stores)}곳 수집 완료!")
        for s in stores[:5]:
            print(f"   - {s['상호명']} ({s['구분']}) | {s['소재지']}")
    except Exception as e:
        print(f"❌ 당첨 판매점 수집 에러 발생: {e}")
    return stores

def generate_banners(latest_stores, df_fixed):
    banners = []
    fixed_names = df_fixed['상호명'].tolist()
    
    for store in latest_stores:
        win_name = store.get("상호명", "")
        win_name_clean = win_name.replace(" ", "")
        
        for fixed_name in fixed_names:
            fixed_name_clean = fixed_name.replace(" ", "")
            if fixed_name_clean and (fixed_name_clean in win_name_clean or win_name_clean in fixed_name_clean):
                msg = f"🎉 경축! 우리 명당 [{win_name}]에서 1등 당첨자가 나왔습니다! 🎉"
                if msg not in banners:
                    banners.append(msg)
                break
                
    return banners

def export_to_github(df_nums, df_fixed, latest_stores, banners, folder_path):
    if df_nums.empty or df_fixed.empty: return
    history = [{"drwNo": int(r["회차"]), "date": str(r["추첨일"]), "nums": [int(r[f"번호{i}"]) for i in range(1, 7)], "bns": int(r["보너스"]), "prize1": int(r["1등당첨금"])} for _, r in df_nums.iterrows()]
    stores = [{"region": str(r["지역"]), "name": str(r["상호명"]), "address": str(r["상세주소"])} for _, r in df_fixed.iterrows()]
    
    payload = {
        "updatedAt": time.strftime("%Y-%m-%d %H:%M:%S"),
        "totalStoresCount": len(stores),
        "latestDraw": history[0] if history else {},
        "latestWinningStores": latest_stores, 
        "fixedStores": stores,
        "weeklyBanner": banners,
        "history": history
    }
    
    out_file = os.path.join(folder_path, "lotto_app_db.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        
    print("\n🚀 [3단계] 깃허브 서버로 앱 DB 자동 업로드 중...")
    if not GITHUB_TOKEN or "여기에" in GITHUB_TOKEN: return

    api_url = f"https://api.github.com/repos/{GITHUB_USER}/{REPO_NAME}/contents/lotto_app_db.json"
    headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json", "User-Agent": GITHUB_USER}
    
    res = requests.get(api_url, headers=headers)
    sha = res.json().get("sha", "") if res.status_code == 200 else ""

    with open(out_file, "rb") as f: content = f.read()
    data = {"message": f"Auto-update: DB sync {payload['updatedAt']}", "content": base64.b64encode(content).decode("utf-8")}
    if sha: data["sha"] = sha

    requests.put(api_url, headers=headers, json=data)
    print("🎉 배포 완료!")

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    df_fixed = load_fixed_stores(folder)
    
    if not df_fixed.empty:
        df_nums = get_latest_draw_and_update_numbers(folder)
        
        # 1100회차 테스트 (테스트 확인 후 int(df_nums.iloc[0]["회차"]) 로 복구)
        latest_drw = 1100
        latest_stores = get_winning_stores(latest_drw)
        banners = generate_banners(latest_stores, df_fixed)
        
        if banners:
            print(f"\n🎯 [매칭 성공] 총 {len(banners)}개의 배너 생성!")
            for b in banners: print(" -", b)
            
        export_to_github(df_nums, df_fixed, latest_stores, banners, folder)
