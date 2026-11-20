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
    """1회차(2002-12-07) 기준 최신 회차 자동 계산 (한국 시간 KST 고정)"""
    KST = datetime.timezone(datetime.timedelta(hours=9))
    now = datetime.datetime.now(KST)
    today = now.date()
    start_date = datetime.date(2002, 12, 7)
    weeks = (today - start_date).days // 7
    calc_drw = 1 + weeks
    # 한국 시간 토요일 저녁 20시 45분 전이면 이전 회차로 유지
    if now.weekday() == 5 and (now.hour < 20 or (now.hour == 20 and now.minute < 45)):
        calc_drw -= 1
    print(f"🕒 현재 한국시간: {now.strftime('%Y-%m-%d %H:%M:%S')} -> 목표 회차: {calc_drw}회")
    return calc_drw

# ==========================================
# [1단계] 고정 명당 데이터 불러오기
# ==========================================
# ==========================================
# [1단계] 명당 및 일반 판매점 데이터 불러오기
# ==========================================
def load_fixed_stores(folder_path):
    print("\n" + "=" * 65)
    print("🚀 [1단계] 고정 명당 데이터 불러오기...")
    return _load_excel_data(folder_path, ["fixed_stores.xlsx", "fixed_stores.csv"])

def load_general_stores(folder_path):
    print("🚀 [1-2단계] 일반 판매점 데이터 불러오기...")
    return _load_excel_data(folder_path, ["general_stores.xlsx", "general_stores.csv"])

def _load_excel_data(folder_path, candidates):
    target_path = None
    for fname in candidates:
        p = os.path.join(folder_path, fname)
        if os.path.exists(p):
            target_path = p
            break
            
    if not target_path: 
        print(f"⚠️ {candidates[0]} 파일이 존재하지 않습니다.")
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
        print(f"✅ 리스트 {len(df)}곳 로드 완료!")
        return df[['지역', '상호명', '상세주소']]
    except Exception as e: 
        print(f"❌ 파일 로드 에러: {e}")
        return pd.DataFrame()

# ==========================================
# [2단계] 당첨번호 수집 (해외 IP 차단 없는 포털 파싱)
# ==========================================
def get_lotto_numbers(drw):
    KST = datetime.timezone(datetime.timedelta(hours=9))
    urls = [
        ("다음", f"https://search.daum.net/search?w=tot&q=로또+{drw}회"),
        ("네이버", f"https://search.naver.com/search.naver?query=로또+{drw}회"),
    ]
    for portal_name, url in urls:
        try:
            res = requests.get(url, headers=BROWSER_HEADERS, timeout=12)
            if res.status_code != 200:
                continue
            
            soup = BeautifulSoup(res.text, "html.parser")
            
            # 공 번호 추출
            balls = []
            for s in soup.select("span"):
                c_str = " ".join(s.get("class", []))
                txt = s.get_text().strip()
                if "ball" in c_str and txt.isdigit():
                    num = int(txt)
                    if 1 <= num <= 45:
                        balls.append(num)
                        if len(balls) == 7:
                            break
                        
            if len(balls) >= 7:
                # 1등 당첨금 추출
                prize1 = 0
                prize_match = re.search(r'1등\s*(?:당첨금)?\s*[\(]?\s*([0-9,]+)\s*원', soup.get_text(" "))
                if prize_match:
                    prize1 = int(prize_match.group(1).replace(",", ""))

                # 추첨일 추출
                draw_date = datetime.datetime.now(KST).strftime("%Y-%m-%d")
                date_match = re.search(r'(\d{4})[\.\-](\d{2})[\.\-](\d{2})', soup.get_text(" "))
                if date_match:
                    draw_date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"

                print(f"✅ [{portal_name}] {drw}회차 번호 수집 성공: {balls[:6]} + {balls[6]}")
                return {
                    "회차": drw,
                    "추첨일": draw_date,
                    "번호1": balls[0], "번호2": balls[1], "번호3": balls[2],
                    "번호4": balls[3], "번호5": balls[4], "번호6": balls[5],
                    "보너스": balls[6],
                    "1등당첨금": prize1
                }
        except Exception as e:
            print(f"⚠️ [{portal_name}] {drw}회차 번호 수집 오류: {e}")
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
    article_links = []

    search_urls = [
        f"https://search.naver.com/search.naver?where=news&query=로또+{drw}회+1등+당첨+판매점",
        f"https://search.daum.net/search?w=news&q=로또+{drw}회+1등+당첨+판매점",
    ]

    for s_url in search_urls:
        try:
            res = requests.get(s_url, headers=BROWSER_HEADERS, timeout=12)
            soup = BeautifulSoup(res.text, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a["href"]
                # 네이버 뉴스, 다음 뉴스, 일반 언론사 기사 링크 폭넓게 수집
                if any(domain in href for domain in ["n.news.naver.com/mnews/article", "v.daum.net/v/"]) or "news_tit" in " ".join(a.get("class", [])):
                    if href.startswith("http") and href not in article_links:
                        article_links.append(href)
        except Exception as e:
            print(f"⚠️ 검색 페이지 조회 실패 ({s_url}): {e}")

    print(f"🔎 관련 뉴스 기사 {len(article_links)}건 발견 -> 본문 분석 시작...")

    def _add_store(name, gubun, addr):
        name = clean_text(re.sub(r'^[\d\.\)\-\s△▲•·]+', '', name))
        addr = clean_text(addr)
        if len(name) >= 2 and len(name) <= 30 and len(addr) >= 5:
            if not any(bad in name for bad in ["판매점", "당첨", "번호", "회차", "복권", "구분", "소재지"]):
                if not any(s["상호명"] == name for s in stores):
                    stores.append({"상호명": name, "구분": gubun, "소재지": addr})

    for link in article_links[:10]:
        try:
            art_res = requests.get(link, headers=BROWSER_HEADERS, timeout=10)
            art_res.encoding = art_res.apparent_encoding or "utf-8"
            art_soup = BeautifulSoup(art_res.text, "html.parser")

            # 1) HTML 표(<table>) 파싱
            for tr in art_soup.select("tr"):
                cols = [clean_text(td.get_text(" ")) for td in tr.select("td, th")]
                if len(cols) >= 3:
                    row_str = " ".join(cols)
                    if any(g in row_str for g in ["자동", "수동", "반자동"]):
                        gubun = next((c for c in cols if c in ["자동", "수동", "반자동"]), "자동")
                        non_gubun = [c for c in cols if c != gubun and not c.isdigit()]
                        if len(non_gubun) >= 2:
                            _add_store(non_gubun[0], gubun, non_gubun[-1])

            # 2) 본문 텍스트 패턴 파싱
            text = art_soup.get_text("\n")

            # 패턴 A: "상호명(자동) - 주소" 또는 "상호명(자동·주소)" 또는 "상호명(자동, 주소)"
            p_a = re.findall(r'([가-힣0-9A-Za-z\s]{2,25})\(\s*(자동|수동|반자동)\s*\)\s*[-–:,·/\s]+([가-힣0-9\s\-]{5,50}[구군시동읍면길로][^\n]{0,30})', text)
            for name, gubun, addr in p_a:
                _add_store(name, gubun, addr)

            # 패턴 B: "상호명, 자동, 주소" / "상호명 / 자동 / 주소" / "상호명 · 자동 · 주소"
            p_b = re.findall(r'([가-힣0-9A-Za-z\s\(\)]{2,25})\s*[,·/\-]\s*(자동|수동|반자동)\s*[,·/\-]\s*([가-힣0-9\s\-]{5,50}[구군시동읍면길로][^\n]{0,30})', text)
            for name, gubun, addr in p_b:
                _add_store(name, gubun, addr)

            # 패턴 C: "상호명(자동·주소)"
            p_c = re.findall(r'([가-힣0-9A-Za-z\s]{2,25})\((자동|수동|반자동)[·,\s]+([가-힣0-9\s\-]{5,50}[구군시동읍면길로][^)\n]{0,30})\)', text)
            for name, gubun, addr in p_c:
                _add_store(name, gubun, addr)

            # 패턴 D: "△상호명(주소)" 또는 "▲상호명(주소)"
            p_d = re.findall(r'[△▲•]\s*([가-힣0-9A-Za-z\s]{2,25})\(([가-힣0-9\s\-]{5,50}[구군시동읍면길로][^)\n]{0,30})\)', text)
            for name, addr in p_d:
                _add_store(name, "자동", addr)

            if len(stores) >= 5:
                break
        except Exception:
            continue

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
# ==========================================
# [4단계] GitHub 저장소에 배포
# ==========================================
def export_to_github(df_nums, df_fixed, df_general, latest_stores, banners, folder_path):
    if df_nums.empty or df_fixed.empty: return
    history = [{"drwNo": int(r["회차"]), "date": str(r["추첨일"]), "nums": [int(r[f"번호{i}"]) for i in range(1, 7)], "bns": int(r["보너스"]), "prize1": int(r["1등당첨금"])} for _, r in df_nums.iterrows()]
    stores = [{"region": str(r["지역"]), "name": str(r["상호명"]), "address": str(r["상세주소"])} for _, r in df_fixed.iterrows()]
    
    # 💡 일반 판매점 데이터 JSON 변환
    g_stores = [{"region": str(r["지역"]), "name": str(r["상호명"]), "address": str(r["상세주소"])} for _, r in df_general.iterrows()]
    
    KST = datetime.timezone(datetime.timedelta(hours=9))
    payload = {
        "updatedAt": datetime.datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"),
        "totalStoresCount": len(stores),
        "latestDraw": history[0] if history else {},
        "latestWinningStores": latest_stores, 
        "fixedStores": stores,
        "generalStores": g_stores, # 💡 JSON DB에 추가!
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
# ==========================================
# 메인 실행부
# ==========================================
if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    df_fixed = load_fixed_stores(folder)
    df_general = load_general_stores(folder) # 💡 일반 판매점 로드 실행
    
    if not df_fixed.empty:
        df_nums = get_latest_draw_and_update_numbers(folder)
        latest_drw = int(df_nums.iloc[0]["회차"]) if not df_nums.empty else calculate_latest_draw()
        
        latest_stores = get_winning_stores(latest_drw)
        banners = generate_banners(latest_stores, df_fixed)
        
        if banners:
            print(f"\n🎯 [매칭 성공] 총 {len(banners)}개의 배너 생성!")
            for b in banners: print(" -", b)
            
        # 💡 df_general 변수 같이 넘겨주기
        export_to_github(df_nums, df_fixed, df_general, latest_stores, banners, folder)
