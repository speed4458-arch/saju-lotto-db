import os
import re
import json
import time
import base64
import urllib3
import datetime
import requests
import warnings
import pandas as pd
from bs4 import BeautifulSoup

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# GitHub 환경변수 설정
# ==========================================
GITHUB_TOKEN = os.environ.get("MY_GITHUB_TOKEN")
GITHUB_USER = "speed4458-arch"
REPO_NAME = "saju-lotto-db"
# ==========================================

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}

def clean_text(text):
    if pd.isna(text) or not text: return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def calculate_latest_draw():
    """1회차(2002-12-07) 기준 최신 회차 자동 계산"""
    start_date = datetime.date(2002, 12, 7)
    today = datetime.date.today()
    weeks = (today - start_date).days // 7
    calc_drw = 1 + weeks
    now = datetime.datetime.now()
    # 토요일 저녁 추첨 전이면 이전 회차로 유지
    if now.weekday() == 5 and (now.hour < 20 or (now.hour == 20 and now.minute < 45)):
        calc_drw -= 1
    return calc_drw

# ==========================================
# [1단계] 고정 명당 데이터 불러오기
# ==========================================
def load_fixed_stores(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [1단계] 고정 명당 데이터 불러오기...")
    candidates = ["fixed_stores.xlsx", "fixed_stores.csv"]
    target_path = None
    for fname in candidates:
        p = os.path.join(folder_path, fname)
        if os.path.exists(p):
            target_path = p
            break
            
    if not target_path: 
        print("⚠️ 명당 엑셀/CSV 파일이 존재하지 않습니다.")
        return pd.DataFrame()
        
    try:
        if target_path.endswith(".xlsx"): 
            df = pd.read_excel(target_path)
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
    except Exception as e: 
        print(f"❌ 명당 파일 로드 에러: {e}")
        return pd.DataFrame()

# ==========================================
# [2단계] 당첨번호 수집 (해외 IP 차단 없는 포털 파싱)
# ==========================================
def get_lotto_numbers(drw):
    try:
        url = f"https://search.daum.net/search?w=tot&q=로또+{drw}회"
        res = requests.get(url, headers=BROWSER_HEADERS, timeout=12)
        if res.status_code != 200:
            return None
        
        soup = BeautifulSoup(res.text, "html.parser")
        
        # 공 번호 추출
        balls = []
        for s in soup.select("span"):
            c_str = " ".join(s.get("class", []))
            txt = s.get_text().strip()
            if "ball" in c_str and txt.isdigit():
                balls.append(int(txt))
                if len(balls) == 7:
                    break
                    
        # 1등 당첨금 추출
        prize1 = 0
        prize_match = re.search(r'1등\s*(?:당첨금)?\s*[\(]?\s*([0-9,]+)\s*원', soup.get_text(" "))
        if prize_match:
            prize1 = int(prize_match.group(1).replace(",", ""))

        # 추첨일 추출
        draw_date = datetime.date.today().strftime("%Y-%m-%d")
        date_match = re.search(r'(\d{4})[\.\-](\d{2})[\.\-](\d{2})', soup.get_text(" "))
        if date_match:
            draw_date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"

        if len(balls) >= 7:
            return {
                "회차": drw,
                "추첨일": draw_date,
                "번호1": balls[0], "번호2": balls[1], "번호3": balls[2],
                "번호4": balls[3], "번호5": balls[4], "번호6": balls[5],
                "보너스": balls[6],
                "1등당첨금": prize1
            }
    except Exception as e:
        print(f"⚠️ {drw}회차 번호 수집 오류: {e}")
    return None

def get_latest_draw_and_update_numbers(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [2단계] 역대 당첨번호 최신화...")
    latest_drw = calculate_latest_draw()
    num_file = os.path.join(folder_path, "lotto_numbers_raw.csv")
    history_excel = os.path.join(folder_path, "lotto_history.xlsx")
    parsed_dict = {}

    if os.path.exists(num_file):
        try:
            df_old = pd.read_csv(num_file)
            for _, r in df_old.iterrows():
                drw = int(r["회차"])
                parsed_dict[drw] = r.to_dict()
        except: pass

    if os.path.exists(history_excel):
        try:
            df_hist = pd.read_excel(history_excel, header=None)
            for _, row in df_hist.iterrows():
                try:
                    drw = int(row[1])
                    if drw not in parsed_dict:
                        parsed_dict[drw] = {
                            "회차": drw, "추첨일": str(row[2])[:10],
                            "번호1": int(row[2]), "번호2": int(row[3]), "번호3": int(row[4]),
                            "번호4": int(row[5]), "번호5": int(row[6]), "번호6": int(row[7]), "보너스": int(row[8]),
                            "1등당첨금": int(re.sub(r'[^\d]', '', str(row[11]))) if pd.notna(row[11]) else 0
                        }
                except: continue
        except: pass

    missing_draws = sorted([d for d in range(1, latest_drw + 1) if d not in parsed_dict], reverse=True)
    if missing_draws:
        print(f"🔄 누락된 {len(missing_draws)}개 회차 번호 수집 중...")
        for drw in missing_draws:
            data = get_lotto_numbers(drw)
            if data:
                parsed_dict[drw] = data
            time.sleep(0.4)

    df_nums = pd.DataFrame(list(parsed_dict.values())).sort_values(by="회차", ascending=False)
    df_nums.to_csv(num_file, index=False, encoding='utf-8-sig')
    return df_nums

# ==========================================
# [3단계] 1등 당첨 판매점 크롤링 (기사 본문 정밀 파싱)
# ==========================================
def get_winning_stores(drw):
    print(f"\n🚀 [판매점 수집] {drw}회차 1등 당첨 판매점 조회 시도...")
    stores = []
    search_url = f"https://search.daum.net/search?w=news&q=로또+{drw}회+1등+당첨+판매점"
    
    try:
        res = requests.get(search_url, headers=BROWSER_HEADERS, timeout=12)
        soup = BeautifulSoup(res.text, "html.parser")
        
        article_links = []
        for a in soup.select("a.item-title, div.item-title a, a.tit_main"):
            href = a.get("href", "")
            if href.startswith("http") and href not in article_links:
                article_links.append(href)
                
        print(f"🔎 관련 뉴스 기사 {len(article_links)}건 분석 중...")
        
        for link in article_links[:6]:
            try:
                art_res = requests.get(link, headers=BROWSER_HEADERS, timeout=10)
                art_soup = BeautifulSoup(art_res.text, "html.parser")
                text = art_soup.get_text("\n")
                
                # 패턴 1: "1. 상호명, 자동, 주소"
                p1 = re.findall(r'(?:\d+[\.\s]*)?([가-힣0-9A-Za-z\s\(\)]+?)\s*[,·/]\s*(자동|수동|반자동)\s*[,·/]\s*([가-힣0-9\s\-]+?[구군시동읍면길로]\s*[0-9\-]*[가-힣0-9\s]*)', text)
                for name, gubun, addr in p1:
                    name = clean_text(re.sub(r'^\d+[\.\s]*', '', name))
                    addr = clean_text(addr)
                    if len(name) >= 2 and len(addr) >= 5 and "판매점" not in name[:2]:
                        if not any(s['상호명'] == name for s in stores):
                            stores.append({"상호명": name, "구분": gubun, "소재지": addr})

                # 패턴 2: "상호명(자동·주소)"
                p2 = re.findall(r'(?:\d+[\.\s]*)?([가-힣0-9A-Za-z\s]+?)\((자동|수동|반자동)[·\s]+([가-힣0-9\s\-]+?[구군시동읍면길로]\s*[0-9\-]*[가-힣0-9\s\(\)]*)\)', text)
                for name, gubun, addr in p2:
                    name = clean_text(re.sub(r'^\d+[\.\s]*', '', name))
                    addr = clean_text(addr)
                    if len(name) >= 2 and len(addr) >= 5:
                        if not any(s['상호명'] == name for s in stores):
                            stores.append({"상호명": name, "구분": gubun, "소재지": addr})

                # 패턴 3: "△상호명(주소)"
                p3 = re.findall(r'[△▲]\s*([가-힣0-9A-Za-z\s]+?)\(([가-힣0-9\s\-]+?[구군시동읍면길로]\s*[0-9\-]*[가-힣0-9\s]*)\)', text)
                for name, addr in p3:
                    name = clean_text(name)
                    addr = clean_text(addr)
                    if len(name) >= 2 and len(addr) >= 5:
                        if not any(s['상호명'] == name for s in stores):
                            stores.append({"상호명": name, "구분": "자동", "소재지": addr})

                if len(stores) >= 5:
                    break
            except Exception:
                continue

    except Exception as e:
        print(f"⚠️ 기사 수집 중 예외: {e}")

    print(f"✅ {drw}회차 1등 당첨 판매점 {len(stores)}곳 수집 완료!")
    for s in stores[:5]:
        print(f"   - {s['상호명']} ({s['구분']}) | {s['소재지']}")
        
    return stores

def generate_banners(latest_stores, df_fixed):
    banners = []
    if df_fixed.empty or not latest_stores: return banners
    fixed_names = df_fixed['상호명'].tolist()
    
    for store in latest_stores:
        win_name = store.get("상호명", "")
        win_name_clean = win_name.replace(" ", "")
        
        for fixed_name in fixed_names:
            fixed_name_clean = fixed_name.replace(" ", "")
            if fixed_name_clean and (fixed_name_clean in win_name_clean or win_name_clean in fixed_name_clean):
                msg = f"경축! 우리 명당 [{win_name}]에서 1등 당첨자가 나왔습니다!"
                if msg not in banners:
                    banners.append(msg)
                break
    return banners

# ==========================================
# [4단계] GitHub 저장소에 배포
# ==========================================
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
        
    print("\n🚀 [4단계] 깃허브 서버로 앱 DB 자동 업로드 중...")
    if not GITHUB_TOKEN:
        print("⚠️ MY_GITHUB_TOKEN이 설정되지 않아 로컬 파일만 저장 후 종료합니다.")
        return

    api_url = f"https://api.github.com/repos/{GITHUB_USER}/{REPO_NAME}/contents/lotto_app_db.json"
    headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json", "User-Agent": GITHUB_USER}
    
    res = requests.get(api_url, headers=headers)
    sha = res.json().get("sha", "") if res.status_code == 200 else ""

    with open(out_file, "rb") as f: content = f.read()
    data = {"message": f"Auto-update: DB sync {payload['updatedAt']}", "content": base64.b64encode(content).decode("utf-8")}
    if sha: data["sha"] = sha

    requests.put(api_url, headers=headers, json=data)
    print("🎉 배포 완료!")

# ==========================================
# 메인 실행부
# ==========================================
if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    df_fixed = load_fixed_stores(folder)
    
    if not df_fixed.empty:
        df_nums = get_latest_draw_and_update_numbers(folder)
        latest_drw = int(df_nums.iloc[0]["회차"]) if not df_nums.empty else calculate_latest_draw()
        
        latest_stores = get_winning_stores(latest_drw)
        banners = generate_banners(latest_stores, df_fixed)
        
        if banners:
            print(f"\n🎯 [매칭 성공] 총 {len(banners)}개의 배너 생성!")
            for b in banners: print(" -", b)
            
        export_to_github(df_nums, df_fixed, latest_stores, banners, folder)
