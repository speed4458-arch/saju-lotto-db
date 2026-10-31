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

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# 🔴 [필수 입력] 깃허브 연동 설정
GITHUB_TOKEN = os.environ.get("MY_GITHUB_TOKEN") # 👈 이렇게 금고를 열어보라고 지시합니다!
GITHUB_USER = "speed4458-arch"
REPO_NAME = "saju-lotto-db" # 깃허브 저장소 이름
# ==========================================

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
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

def get_latest_draw_and_update_numbers(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [2단계] 역대 당첨번호 엑셀 로드 및 최신화...")
    
    base_date = datetime.date(2026, 10, 24)
    base_drw = 1247
    # 깃허브 서버(영국) 시간을 한국 시간(KST, UTC+9)으로 강제 변환
    kst = datetime.timezone(datetime.timedelta(hours=9))
    now = datetime.datetime.now(kst)
    
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
        print("📁 공식 엑셀(lotto_history.xlsx)을 발견했습니다! 데이터를 0.1초 만에 흡수합니다.")
        try:
            df_hist = pd.read_excel(history_excel, header=None)
            count = 0
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
                    count += 1
                except: continue
            print(f"✅ 엑셀에서 총 {count}개의 역대 당첨 기록 로드 완료!")
        except Exception as e:
            print(f"❌ 엑셀 읽기 오류: {e}")

    collected_draws = set(parsed_dict.keys())
    missing_draws = [d for d in range(1, latest_drw + 1) if d not in collected_draws]
    missing_draws.sort(reverse=True)
    
    if missing_draws:
        print(f"🎯 엑셀에 없는 새로운 {len(missing_draws)}개 회차만 인터넷에서 가져옵니다...")
        session = requests.Session()
        for drw in missing_draws:
            try:
                print(f"👉 {drw}회차 업데이트 중...", end="\r")
                url = f"https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={drw}"
                d_json = session.get(url, headers=BROWSER_HEADERS, verify=False, timeout=5).json()
                if d_json.get("returnValue") == "success":
                    parsed_dict[drw] = {
                        "회차": int(d_json.get("drwNo", drw)), "추첨일": str(d_json.get("drwNoDate", "")),
                        "번호1": int(d_json.get("drwtNo1", 0)), "번호2": int(d_json.get("drwtNo2", 0)),
                        "번호3": int(d_json.get("drwtNo3", 0)), "번호4": int(d_json.get("drwtNo4", 0)),
                        "번호5": int(d_json.get("drwtNo5", 0)), "번호6": int(d_json.get("drwtNo6", 0)),
                        "보너스": int(d_json.get("bnusNo", 0)), "1등당첨금": int(d_json.get("firstWinamnt", 0))
                    }
                time.sleep(0.3)
            except: continue
        print("\n✅ 최신 회차 동기화 완료!")
    else:
        print("✨ 누락된 회차 없이 모든 데이터가 완벽한 최신 상태입니다! (인터넷 접속 패스)")

    parsed_rows = list(parsed_dict.values())
    df_nums = pd.DataFrame(parsed_rows).sort_values(by="회차", ascending=False)
    df_nums.to_csv(num_file, index=False, encoding='utf-8-sig')
    return df_nums, latest_drw

def get_weekly_hit_stores(latest_drw, df_fixed):
    print("\n" + "=" * 65)
    print(f"🏆 [3단계] 이번 주({latest_drw}회) 우리 명당 배출 여부 확인...")
    session = requests.Session()
    url = "https://www.dhlottery.co.kr/store.do?method=topStore&pageGubun=L645"
    req_headers = BROWSER_HEADERS.copy()
    req_headers["Content-Type"] = "application/x-www-form-urlencoded"
    
    latest_stores = []
    try:
        res = session.post(url, data={"method": "topStore", "pageGubun": "L645", "drwNo": latest_drw, "nowPage": 1}, headers=req_headers, verify=False, timeout=5)
        res.encoding = 'euc-kr'
        soup = BeautifulSoup(res.text, 'html.parser')
        tables = soup.find_all('table', class_='tbl_data')
        
        if len(tables) >= 1 and tables[0].find('tbody'):
            for tr in tables[0].find('tbody').find_all('tr'):
                tds = tr.find_all('td')
                if len(tds) >= 4 and "없습니다" not in tds[0].text:
                    latest_stores.append({'상호명': clean_text(tds[1].text), '상세주소': clean_text(tds[3].text), '등수': 1})
    except: pass

    banner_hits = []
    for store in latest_stores:
        match = df_fixed[df_fixed['상호명'].str.contains(store['상호명'][:3], na=False)]
        if not match.empty:
            banner_hits.append({"rank": store['등수'], "name": store['상호명'], "address": store['상세주소']})
            
    if banner_hits: print(f"🎉 대박! 이번 주 우리 명당에서 1등이 나왔습니다! ({len(banner_hits)}건)")
    else: print(f"🥲 이번 주 당첨 매칭 내역 없음.")
    return banner_hits

def export_to_github(df_nums, df_fixed, banner_hits, folder_path):
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
        "weeklyBanner": banner_hits,
        "history": history
    }
    
    # 1. 로컬에 백업 저장
    out_file = os.path.join(folder_path, "lotto_app_db.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        
    print("\n" + "=" * 65)
    print("🚀 [4단계] 깃허브 서버로 앱 DB 자동 업로드 중...")
    
    if not GITHUB_TOKEN or "여기에" in GITHUB_TOKEN:
        print("⚠️ 깃허브 토큰이 입력되지 않아 로컬 저장만 완료되었습니다.")
        return

    # 2. 깃허브 API 연동
    file_path = "lotto_app_db.json"
    api_url = f"https://api.github.com/repos/{GITHUB_USER}/{REPO_NAME}/contents/{file_path}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 기존 파일 SHA 값 가져오기 (덮어쓰기에 필수)
    sha = ""
    res = requests.get(api_url, headers=headers)
    if res.status_code == 200:
        sha = res.json().get("sha", "")

    # 파일 인코딩 및 업로드
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
        print("💡 이제 앱을 켜면 전 세계 어디서든 이 최신 데이터를 다운받습니다.")
    else:
        print(f"❌ 깃허브 업로드 실패: {put_res.json()}")
    print("=" * 65)

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    df_fixed = load_fixed_stores(folder)
    if not df_fixed.empty:
        df_nums, latest_drw = get_latest_draw_and_update_numbers(folder)
        banner_hits = get_weekly_hit_stores(latest_drw, df_fixed)
        export_to_github(df_nums, df_fixed, banner_hits, folder)
