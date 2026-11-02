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
GITHUB_TOKEN = os.environ.get("MY_GITHUB_TOKEN")
SCRAPER_API_KEY = os.environ.get("SCRAPER_API_KEY")  # 해외 서버(GitHub Actions) 실행 시 필수
GITHUB_USER = "speed4458-arch"
REPO_NAME = "saju-lotto-db"
# ==========================================

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://dhlottery.co.kr/",
}

def clean_text(text):
    if pd.isna(text) or not text: return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def send_request(url, method="GET", data=None):
    """
    해외 IP 차단 방지용 요청 함수:
    SCRAPER_API_KEY가 등록되어 있으면 프록시를 통해 우회하고, 없으면 일반 요청을 수행합니다.
    """
    try:
        if SCRAPER_API_KEY and "여기에" not in SCRAPER_API_KEY:
            proxy_url = "http://api.scraperapi.com"
            payload = {'api_key': SCRAPER_API_KEY, 'url': url}
            if method == "POST":
                payload['keep_headers'] = 'true'
                return requests.post(proxy_url, params=payload, data=data, headers=BROWSER_HEADERS, timeout=30)
            return requests.get(proxy_url, params=payload, headers=BROWSER_HEADERS, timeout=30)
        else:
            if method == "POST":
                return requests.post(url, data=data, headers=BROWSER_HEADERS, timeout=15)
            return requests.get(url, headers=BROWSER_HEADERS, timeout=15)
    except Exception as e:
        print(f"⚠️ 요청 중 네트워크 오류 발생 ({url}): {e}")
        return None

def get_lotto_number_official(drw):
    """동행복권 공식 JSON API를 통해 당첨번호 수집"""
    url = f"https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={drw}"
    res = send_request(url, method="GET")
    if res and res.status_code == 200:
        try:
            data = res.json()
            if data.get("returnValue") == "success":
                return {
                    "회차": drw,
                    "추첨일": data.get("drwNoDate"),
                    "번호1": data.get("drwtNo1"),
                    "번호2": data.get("drwtNo2"),
                    "번호3": data.get("drwtNo3"),
                    "번호4": data.get("drwtNo4"),
                    "번호5": data.get("drwtNo5"),
                    "번호6": data.get("drwtNo6"),
                    "보너스": data.get("bnusNo"),
                    "1등당첨금": data.get("firstWinamnt", 0)
                }
        except Exception as e:
            print(f"❌ {drw}회차 번호 JSON 파싱 실패: {e}")
    return None

def get_winning_stores(drw):
    """동행복권 공식 사이트에서 POST 방식으로 1등 당첨 판매점 크롤링"""
    print(f"\n🚀 [판매점 수집] {drw}회차 1등 당첨 판매점 조회 시도...")
    url = "https://dhlottery.co.kr/store.do?method=topStore&pageGubun=L645"
    post_data = {
        "method": "topStore",
        "nowPage": "1",
        "rankNo": "",
        "gameNo": "5133",
        "drwNo": str(drw),
        "schKey": "all",
        "schVal": ""
    }

    res = send_request(url, method="POST", data=post_data)
    stores = []
    if not res or res.status_code != 200:
        print(f"❌ 동행복권 판매점 페이지 접근 실패 (응답 코드: {getattr(res, 'status_code', 'None')})")
        return stores

    try:
        soup = BeautifulSoup(res.text, "html.parser")
        # 동행복권 당첨 판매점 페이지의 테이블 목록 (첫 번째 테이블이 1등 배출점)
        tables = soup.select("table.tbl_data_col")
        if not tables:
            print("⚠️ 당첨 판매점 테이블을 찾을 수 없습니다. (IP 차단 또는 구조 변경)")
            return stores

        first_rank_table = tables[0]
        rows = first_rank_table.select("tbody tr")

        for row in rows:
            cols = row.select("td")
            if len(cols) >= 4:
                # 0: 순번, 1: 상호명, 2: 구분(자동/수동/반자동), 3: 소재지
                name = clean_text(cols[1].text)
                gubun = clean_text(cols[2].text)
                addr = clean_text(cols[3].text)

                if name and "조회 결과가 없습니다" not in name:
                    stores.append({
                        "상호명": name,
                        "구분": gubun,
                        "소재지": addr
                    })

        print(f"✅ {drw}회차 1등 당첨 판매점 {len(stores)}곳 수집 완료!")
        for s in stores[:3]:
            print(f"   - {s['상호명']} ({s['구분']}) | {s['소재지']}")
    except Exception as e:
        print(f"❌ 파싱 중 에러 발생: {e}")

    return stores

def calculate_latest_draw():
    """1회차(2002-12-07 토요일) 기준 최신 회차 자동 계산"""
    start_date = datetime.date(2002, 12, 7)
    today = datetime.date.today()
    weeks = (today - start_date).days // 7
    calc_drw = 1 + weeks
    now = datetime.datetime.now()
    # 토요일 오후 8시 45분 이전이면 아직 이번 주 추첨 전
    if now.weekday() == 5 and (now.hour < 20 or (now.hour == 20 and now.minute < 45)):
        calc_drw -= 1
    return calc_drw

def get_latest_draw_and_update_numbers(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [2단계] 역대 당첨번호 최신화...")
    latest_drw = calculate_latest_draw()
    num_file = os.path.join(folder_path, "lotto_numbers_raw.csv")
    history_excel = os.path.join(folder_path, "lotto_history.xlsx")
    parsed_dict = {}

    # 기존 CSV 우선 로드 (중복 요청 방지)
    if os.path.exists(num_file):
        try:
            df_old = pd.read_csv(num_file)
            for _, r in df_old.iterrows():
                drw = int(r["회차"])
                parsed_dict[drw] = r.to_dict()
        except: pass

    # 엑셀 파일 보완
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

    # 누락된 회차 공식 API로 채우기
    missing_draws = sorted([d for d in range(1, latest_drw + 1) if d not in parsed_dict], reverse=True)
    if missing_draws:
        print(f"🔄 누락된 {len(missing_draws)}개 회차 번호 수집 중...")
        for drw in missing_draws:
            data = get_lotto_number_official(drw)
            if data:
                parsed_dict[drw] = data
            time.sleep(0.3)  # 과도한 요청 방지

    df_nums = pd.DataFrame(list(parsed_dict.values())).sort_values(by="회차", ascending=False)
    df_nums.to_csv(num_file, index=False, encoding='utf-8-sig')
    return df_nums

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
        print(f"❌ 명당 파일 로드 에러: {e}")
        return pd.DataFrame()

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
    if not GITHUB_TOKEN or "여기에" in GITHUB_TOKEN: 
        print("⚠️ GITHUB_TOKEN이 설정되지 않아 업로드를 건너뜁니다.")
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

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    df_fixed = load_fixed_stores(folder)
    
    if not df_fixed.empty:
        df_nums = get_latest_draw_and_update_numbers(folder)
        
        # 최신 회차 설정 (최신 당첨 번호 데이터의 1번째 행 회차 기준)
        latest_drw = int(df_nums.iloc[0]["회차"]) if not df_nums.empty else 1100
        latest_stores = get_winning_stores(latest_drw)
        banners = generate_banners(latest_stores, df_fixed)
        
        if banners:
            print(f"\n🎯 [매칭 성공] 총 {len(banners)}개의 배너 생성!")
            for b in banners: print(" -", b)
            
        export_to_github(df_nums, df_fixed, latest_stores, banners, folder)
