import time
import os
import json
import re
import urllib3
from curl_cffi import requests
import pandas as pd
import datetime
import base64
from bs4 import BeautifulSoup
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# 🔴 [필수 입력] 깃허브 연동 설정
# 환경변수 대신 여기에 직접 토큰을 문자열로 붙여넣으세요 (예: "ghp_xxxxxxxxxxxxxxx")
GITHUB_TOKEN = os.environ.get("MY_GITHUB_TOKEN")
GITHUB_USER = "speed4458-arch"
REPO_NAME = "saju-lotto-db"
# ==========================================

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Referer": "https://dhlottery.co.kr/gameResult.do?method=byWin",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "X-Requested-With": "XMLHttpRequest",
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
            target_path = os.path.join(folder_path, fname)
            break
            
    if not target_path:
        print("❌ 폴더에 fixed_stores.xlsx 파일이 없습니다!")
        return pd.DataFrame()

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
    except Exception as e:
        print(f"❌ 명당 엑셀 읽기 오류: {e}")
        return pd.DataFrame()

def get_lotto_from_naver(drw):
    print(f"⚠️ 동행복권 차단됨. 포털 사이트 우회 크롤링으로 {drw}회차 수집 시도...")
    try:
        url = f"https://search.daum.net/search?w=tot&q=로또+{drw}회"
        res = requests.get(url, impersonate="chrome120", timeout=5)
        soup = BeautifulSoup(res.text, "html.parser")
        
        balls = []
        for span in soup.select("span"):
            if span.text.strip().isdigit() and 'ball' in str(span.get('class', '')):
                balls.append(int(span.text.strip()))
                
        if len(balls) >= 7:
            return {
                "회차": drw, "추첨일": datetime.datetime.now().strftime("%Y-%m-%d"), 
                "번호1": balls[0], "번호2": balls[1], "번호3": balls[2],
                "번호4": balls[3], "번호5": balls[4], "번호6": balls[5],
                "보너스": balls[6], "1등당첨금": 0 
            }
    except Exception as e:
        print(f"다음 백업 오류: {e}")
    return None

def get_latest_draw_and_update_numbers(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [2단계] 역대 당첨번호 엑셀 로드 및 최신화...")
    
    base_date = datetime.date(2026, 10, 24)
    base_drw = 1247
    now = datetime.datetime.now()
    
    days_diff = (now.date() - base_date).days
    estimated_drw = base_drw + (days_diff // 7)
    if now.weekday() == 5 and now.hour < 21:
        estimated_drw -= 1
        
    latest_drw = estimated_drw
    print(f"📌 달력 계산으로 도출해낸 현재 최신 회차는 [ {latest_drw}회 ] 입니다.")

    num_file = os.path.join(folder_path, "lotto_numbers_raw.csv")
    history_excel = os.path.join(folder_path, "lotto_history.xlsx")
    
    parsed_dict = {}
    
    if os.path.exists(history_excel):
        print("📁 공식 엑셀(lotto_history.xlsx)을 발견했습니다! 데이터를 흡수합니다.")
        try:
            df_hist = pd.read_excel(history_excel, header=None)
            for idx, row in df_hist.iterrows():
                try:
                    drw = int(row[1])
                    weeks_diff = base_drw - drw
                    draw_date = base_date - datetime.timedelta(weeks=weeks_diff)
                    
                    prize_str = str(row[11]).replace(',', '').replace('원', '').strip()
                    prize1 = int(prize_str) if prize_str.isdigit() else 0
                    
                    parsed_dict[drw] = {
                        "회차": drw, "추첨일": draw_date.strftime("%Y-%m-%d"),
                        "번호1": int(row[2]), "번호2": int(row[3]), "번호3": int(row[4]), 
                        "번호4": int(row[5]), "번호5": int(row[6]), "번호6": int(row[7]),
                        "보너스": int(row[8]), "1등당첨금": prize1
                    }
                except: continue
        except Exception as e:
            pass

    collected_draws = set(parsed_dict.keys())
    missing_draws = [d for d in range(1, latest_drw + 1) if d not in collected_draws]
    missing_draws.sort(reverse=True)
    
    if missing_draws:
        print(f"🎯 엑셀에 없는 새로운 {len(missing_draws)}개 회차 수집을 시작합니다...")
        session = requests.Session()
        for drw in missing_draws:
            print(f"👉 {drw}회차 업데이트 중...", end="\r")
            success = False
            try:
                url = f"https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={drw}"
                res = session.get(url, headers=BROWSER_HEADERS, impersonate="chrome120", timeout=5)
                
                if res.status_code == 200 and res.text.strip():
                    d_json = res.json()
                    if d_json.get("returnValue") == "success":
                        parsed_dict[drw] = {
                            "회차": int(d_json.get("drwNo", drw)), "추첨일": str(d_json.get("drwNoDate", "")),
                            "번호1": int(d_json.get("drwtNo1", 0)), "번호2": int(d_json.get("drwtNo2", 0)),
                            "번호3": int(d_json.get("drwtNo3", 0)), "번호4": int(d_json.get("drwtNo4", 0)),
                            "번호5": int(d_json.get("drwtNo5", 0)), "번호6": int(d_json.get("drwtNo6", 0)),
                            "보너스": int(d_json.get("bnusNo", 0)), "1등당첨금": int(d_json.get("firstWinamnt", 0))
                        }
                        success = True
            except: pass
                
            if not success:
                naver_data = get_lotto_from_naver(drw)
                if naver_data:
                    parsed_dict[drw] = naver_data
                    print(f"✅ 다음(Daum) 백업 경로로 {drw}회차 수집 성공!")
                else:
                    print(f"\n❌ {drw}회차 수집 완전 실패 (동행복권, 다음 모두 차단됨)")
            time.sleep(0.5)
        print("\n✅ 최신 회차 동기화 완료!")
    else:
        print("✨ 누락된 회차 없이 모든 데이터가 완벽한 최신 상태입니다!")

    parsed_rows = list(parsed_dict.values())
    df_nums = pd.DataFrame(parsed_rows).sort_values(by="회차", ascending=False)
    df_nums.to_csv(num_file, index=False, encoding='utf-8-sig')
    return df_nums

def export_to_github(df_nums, df_fixed, folder_path):
    if df_nums.empty or df_fixed.empty: return
    
    history = [{
        "drwNo": int(r["회차"]), "date": str(r["추첨일"]),
        "nums": [int(r[f"번호{i}"]) for i in range(1, 7)],
        "bns": int(r["보너스"]), "prize1": int(r["1등당첨금"])
    } for _, r in df_nums.iterrows()]
    
    stores = [{"region": str(r["지역"]), "name": str(r["상호명"]), "address": str(r["상세주소"])} for _, r in df_fixed.iterrows()]
    
    payload = {
        "updatedAt": time.strftime("%Y-%m-%d %H:%M:%S"),
        "totalStoresCount": len(stores),
        "latestDraw": history[0] if history else {},
        "fixedStores": stores,
        "weeklyBanner": [], # 앱 오류 방지를 위해 빈 리스트로 구조 유지
        "history": history
    }
    
    out_file = os.path.join(folder_path, "lotto_app_db.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        
    print("\n" + "=" * 65)
    print("🚀 [3단계] 깃허브 서버로 앱 DB 자동 업로드 중...")
    
    if not GITHUB_TOKEN or "여기에" in GITHUB_TOKEN:
        print("⚠️ 깃허브 토큰이 입력되지 않아 로컬 저장만 완료되었습니다. 앱에 반영되지 않습니다.")
        return

    file_path = "lotto_app_db.json"
    api_url = f"https://api.github.com/repos/{GITHUB_USER}/{REPO_NAME}/contents/{file_path}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": GITHUB_USER
    }

    sha = ""
    res = requests.get(api_url, headers=headers)
    if res.status_code == 200:
        sha = res.json().get("sha", "")

    with open(out_file, "rb") as f:
        content = f.read()
    b64_content = base64.b64encode(content).decode("utf-8")

    data = {
        "message": f"Auto-update: DB sync {payload['updatedAt']}",
        "content": b64_content,
    }
    if sha:
        data["sha"] = sha

    put_res = requests.put(api_url, headers=headers, json=data)
    
    if put_res.status_code in [200, 201]:
        print("🎉 [완전 자동화 성공]: 깃허브 서버에 최신 데이터가 성공적으로 배포되었습니다!")
        print("💡 앱을 새로고침 하시면 1248회가 즉시 반영됩니다.")
    else:
        print(f"❌ 깃허브 업로드 실패 (상태코드 {put_res.status_code}): {put_res.text}")
    print("=" * 65)

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    df_fixed = load_fixed_stores(folder)
    if not df_fixed.empty:
        df_nums = get_latest_draw_and_update_numbers(folder)
        export_to_github(df_nums, df_fixed, folder)
