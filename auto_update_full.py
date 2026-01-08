"""
부태리 신고가 완전 자동 업데이트 스크립트
데이터 갱신 + HTML 생성 통합 버전

매일 새벽 6시 실행용 (Windows 작업 스케줄러)
"""

import sqlite3
import os
import sys
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from collections import Counter
from html import escape
import logging
import time
import random
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 로깅 설정
log_dir = os.path.join(os.path.dirname(__file__), 'logs')
os.makedirs(log_dir, exist_ok=True)

log_file = os.path.join(log_dir, f'auto_update_full_{datetime.now().strftime("%Y%m%d")}.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'db', 'monitoring.db')
OUTPUT_DIR = os.path.join(os.path.dirname(BASE_DIR), 'newtrade')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# API 설정
API_KEY = "발급받은 API 키를 여기에 입력하세요"  # ⚠️ 실제 API 키로 교체 필요
API_TIMEOUT = (5, 15)

# API 키 확인
if API_KEY == "발급받은 API 키를 여기에 입력하세요":
    print("\n" + "="*80)
    print("[경고] API 키가 설정되지 않았습니다!")
    print("="*80)
    print("\n1. auto_update_full.py 파일을 텍스트 에디터로 엽니다.")
    print("2. 47번째 줄의 API_KEY = \"...\" 부분을 찾습니다.")
    print("3. 따옴표 안에 실제 API 키를 입력합니다.")
    print("4. 저장 후 다시 실행합니다.")
    print("\nAPI 키 발급: https://www.data.go.kr/")
    print("검색: '국토교통부 아파트매매 실거래 상세 자료'\n")
    sys.exit(1)

# 모니터링 리스트 설정
MONITORING_CONFIGS = [
    {'list_name': '서울 수도권', 'output_file': '서울&수도권 신고가.html'},
    {'list_name': '부산', 'output_file': '부산신고가.html'},
    {'list_name': '대구', 'output_file': '대구신고가.html'}
]

def build_session():
    """SSL 문제 해결된 세션"""
    s = requests.Session()
    retry = Retry(
        total=5,
        connect=3,
        read=3,
        backoff_factor=0.6,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=40)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.headers.update({"User-Agent": "RealEstateMonitor/1.0"})
    s.verify = False
    return s

def jitter_sleep(max_ms=300):
    """API 호출 간 지연"""
    time.sleep(random.uniform(0, max_ms / 1000.0))

def fetch_trade_data_from_api(session, sigungu_code, deal_ymd, api_type='existing'):
    """공공 API에서 실거래 데이터 조회"""
    try:
        if api_type == 'existing':
            url = "http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTrade"
        else:  # new
            url = "http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTradeDev"

        params = {
            'serviceKey': API_KEY,
            'LAWD_CD': sigungu_code,
            'DEAL_YMD': deal_ymd,
            'numOfRows': '999'
        }

        jitter_sleep()
        response = session.get(url, params=params, timeout=API_TIMEOUT)

        if response.status_code != 200:
            logging.warning(f"API 오류: {response.status_code}")
            return []

        root = ET.fromstring(response.content)
        items = root.findall('.//item')

        trades = []
        for item in items:
            try:
                apt_name = item.findtext('아파트', '').strip()
                area_str = item.findtext('전용면적', '0').strip()
                price_str = item.findtext('거래금액', '0').strip().replace(',', '')
                floor_str = item.findtext('층', '0').strip()
                dong = item.findtext('법정동', '').strip()

                year = item.findtext('년', '').strip()
                month = item.findtext('월', '').strip()
                day = item.findtext('일', '').strip()

                if not all([apt_name, year, month, day]):
                    continue

                trade_date = datetime.strptime(f"{year}-{month.zfill(2)}-{day.zfill(2)}", '%Y-%m-%d')

                trades.append({
                    'apt_name': apt_name,
                    'area': float(area_str),
                    'price': int(price_str),
                    'floor': int(floor_str) if floor_str.isdigit() else 0,
                    'dong': dong,
                    'date': trade_date
                })
            except Exception as e:
                logging.debug(f"항목 파싱 오류: {str(e)}")
                continue

        return trades

    except Exception as e:
        logging.error(f"API 조회 오류: {str(e)}")
        return []

def update_apartment_data(db_conn, list_name):
    """특정 모니터링 리스트의 아파트 데이터 갱신"""
    logging.info(f"\n{'='*60}")
    logging.info(f"'{list_name}' 데이터 갱신 시작")
    logging.info(f"{'='*60}")

    try:
        cursor = db_conn.cursor()

        # 리스트 ID 조회
        cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (list_name,))
        result = cursor.fetchone()

        if not result:
            logging.warning(f"모니터링 리스트 '{list_name}'을 찾을 수 없습니다.")
            return False

        list_id = result[0]

        # 아파트 목록 조회
        cursor.execute("""
            SELECT id, apt_name, area, sigungu_code, dong, last_max_price,
                   max_price_date, max_price_floor, max_price_dong,
                   prev_max_price, prev_max_date, prev_max_floor
            FROM apartments
            WHERE list_id = ?
        """, (list_id,))

        apartments = cursor.fetchall()
        total = len(apartments)
        logging.info(f"총 {total}개 아파트 데이터 갱신 시작")

        session = build_session()
        updated_count = 0
        new_high_count = 0

        for idx, apt_row in enumerate(apartments, 1):
            (apt_id, apt_name, area, sigungu_code, dong,
             last_max_price, max_date, max_floor, max_dong,
             prev_max_price, prev_date, prev_floor) = apt_row

            # 진행 상황을 더 자주 표시 (10개마다)
            if idx % 10 == 0:
                print(f"진행: {idx}/{total} ({idx/total*100:.1f}%) - {apt_name}", flush=True)
                logging.info(f"진행: {idx}/{total} ({idx/total*100:.1f}%)")

            try:
                target_area = float(str(area).replace('㎡', '').strip())

                # 최근 4개월 데이터 조회
                all_trades = []
                current_date = datetime.now()

                for month_offset in range(4):
                    search_date = current_date - timedelta(days=30 * month_offset)
                    deal_ymd = search_date.strftime("%Y%m")

                    # 기존 아파트 API
                    trades_existing = fetch_trade_data_from_api(session, sigungu_code, deal_ymd, 'existing')
                    # 신규 아파트 API
                    trades_new = fetch_trade_data_from_api(session, sigungu_code, deal_ymd, 'new')

                    all_trades.extend(trades_existing)
                    all_trades.extend(trades_new)

                # 해당 아파트, 동, 면적 필터링
                filtered_trades = []
                for trade in all_trades:
                    if (trade['apt_name'] == apt_name and
                        trade['dong'] == dong and
                        abs(trade['area'] - target_area) <= 0.5):
                        filtered_trades.append(trade)

                if not filtered_trades:
                    continue

                # 최고가 찾기
                max_trade = max(filtered_trades, key=lambda x: x['price'])
                new_max_price = max_trade['price']
                new_max_date = max_trade['date'].strftime('%Y-%m-%d')
                new_max_floor = max_trade['floor']
                new_max_dong = max_trade['dong']

                # 신고가 갱신 체크
                if new_max_price > (last_max_price or 0):
                    # 신고가 발생!
                    cursor.execute("""
                        UPDATE apartments
                        SET prev_max_price = ?,
                            prev_max_date = ?,
                            prev_max_floor = ?,
                            last_max_price = ?,
                            max_price_date = ?,
                            max_price_floor = ?,
                            max_price_dong = ?,
                            updated_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                    """, (last_max_price, max_date, max_floor,
                          new_max_price, new_max_date, new_max_floor, new_max_dong,
                          apt_id))

                    db_conn.commit()
                    updated_count += 1
                    new_high_count += 1

                    logging.info(f"✨ 신고가: {apt_name} {area} - {last_max_price:,}만원 → {new_max_price:,}만원")

            except Exception as e:
                logging.error(f"'{apt_name}' 처리 중 오류: {str(e)}")
                continue

        logging.info(f"\n{'='*60}")
        logging.info(f"'{list_name}' 갱신 완료")
        logging.info(f"  총 아파트: {total}개")
        logging.info(f"  갱신: {updated_count}개")
        logging.info(f"  신고가: {new_high_count}개")
        logging.info(f"{'='*60}\n")

        return True

    except Exception as e:
        logging.error(f"'{list_name}' 데이터 갱신 중 오류: {str(e)}")
        import traceback
        logging.error(traceback.format_exc())
        return False

def get_new_high_data(db_conn, list_name):
    """신고가 데이터 조회"""
    try:
        cursor = db_conn.cursor()

        cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (list_name,))
        result = cursor.fetchone()

        if not result:
            return []

        list_id = result[0]

        cursor.execute("""
            SELECT
                apt_name, area, sido, sigungu, dong as location_dong, build_year,
                max_price_date as date, last_max_price as new_price,
                max_price_floor as floor, max_price_dong as dong,
                prev_max_price as old_price, prev_max_date as old_date,
                prev_max_floor as old_floor
            FROM apartments
            WHERE list_id = ?
            AND prev_max_price > 0
            AND last_max_price > prev_max_price
            ORDER BY last_max_price DESC
        """, (list_id,))

        results = cursor.fetchall()

        apt_list = []
        for row in results:
            apt_list.append({
                'apt_name': row[0], 'area': row[1], 'sido': row[2],
                'sigungu': row[3], 'location_dong': row[4], 'build_year': row[5],
                'date': row[6], 'new_price': row[7], 'floor': row[8],
                'dong': row[9], 'old_price': row[10] or 0,
                'old_date': row[11], 'old_floor': row[12]
            })

        return apt_list

    except Exception as e:
        logging.error(f"신고가 데이터 조회 오류: {str(e)}")
        return []

def generate_html(apt_list, list_name):
    """HTML 생성 (간소화 버전)"""
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    current_year = datetime.now().year

    apt_list = sorted(apt_list, key=lambda x: x.get('new_price', 0), reverse=True)
    total = len(apt_list)

    young_count = sum(1 for apt in apt_list
                     if apt.get('build_year') == '분양' or
                     (apt.get('build_year', '').isdigit() and current_year - int(apt['build_year']) <= 10))

    cards_html = []
    for apt in apt_list:
        name = escape(str(apt.get('apt_name', '')))
        area = escape(str(apt.get('area', '')))
        old_price = apt.get('old_price', 0) or 0
        new_price = apt.get('new_price', 0) or 0
        date = escape(str(apt.get('date', '')))
        floor = escape(str(apt.get('floor', '')))
        build_year = escape(str(apt.get('build_year', '')))
        sido = escape(str(apt.get('sido', '')))
        sigungu = escape(str(apt.get('sigungu', '')))
        location_dong = escape(str(apt.get('location_dong', '')))

        inc = new_price - old_price if old_price else 0
        pct = f"{(inc/old_price*100):.1f}%" if old_price else "-"

        year_badge = f"<span class='year-badge'>{build_year}년</span>" if build_year else ""

        card = f"""
        <section class="card">
          <h3>{name} {year_badge} <span class="tag">전용면적 {area}㎡</span></h3>
          <div class="card-sub">📍 {sido} {sigungu} {location_dong}</div>
          <div class="grid">
            <div class="k">이번 실거래</div>
            <div><span class="highlight">{new_price:,}만원</span> ({date}{f' | {floor}층' if floor else ''})</div>
            <div class="k">직전 최고가</div>
            <div>{old_price:,}만원</div>
            <div class="k">변화</div>
            <div class="rise">{inc:,}만원 상승 ({pct})</div>
          </div>
        </section>
        """
        cards_html.append(card)

    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8"/>
<title>{escape(list_name)} 신고가 - {escape(now)}</title>
<style>
:root {{ --bg:#F2F2F7; --fg:#111; --card:#fff; --primary:#007AFF; --accent:#ff3b30; }}
body{{ margin:0; font-family:"Malgun Gothic",sans-serif; background:var(--bg); }}
.page{{ max-width:960px; margin:24px auto }}
.header{{ background:linear-gradient(135deg,var(--primary),#63A4FF); color:#fff;
         border-radius:16px; padding:20px; }}
.header h1{{ margin:0; font-size:22px }}
.cards{{ display:grid; gap:16px; margin-top:16px }}
.card{{ background:var(--card); border-radius:14px; padding:18px;
       box-shadow:0 2px 12px rgba(0,0,0,.04); }}
.card h3{{ margin:0 0 8px; font-size:17px; display:flex; align-items:center; gap:8px; flex-wrap:wrap }}
.card-sub{{ color:#6b7280; font-size:13px; margin-bottom:12px }}
.grid{{ display:grid; grid-template-columns:auto 1fr; gap:8px 12px; font-size:14px }}
.k{{ color:#6b7280; font-weight:600 }}
.highlight{{ color:var(--accent); font-weight:700; font-size:16px }}
.rise{{ color:#10b981; font-weight:600 }}
.tag{{ background:#e0f2fe; color:#0369a1; padding:4px 10px;
      border-radius:999px; font-size:12px; font-weight:600 }}
.year-badge{{ background:#f1f5f9; color:#334155; padding:4px 10px;
             border-radius:999px; font-size:12px; font-weight:600 }}
footer{{ margin:24px 0; color:#6b7280; font-size:12px; text-align:center }}
</style>
</head>
<body>
<div class="page">
  <header class="header">
    <h1>🔔 {escape(list_name)} 신고가</h1>
    <p style="margin:8px 0 0;opacity:.95">자동 업데이트: {escape(now)} · 총 {total}건 (10년이하 {young_count}건)</p>
  </header>
  <section class="cards">
    {"".join(cards_html)}
  </section>
  <footer>© 부태리 실거래가 모니터링 시스템 (자동생성)</footer>
</div>
</body>
</html>"""

    return html_content

def main():
    """메인 실행 함수"""
    start_time = datetime.now()

    print("\n" + "="*80)
    print("[시작] 부태리 신고가 완전 자동 업데이트")
    print(f"시작 시간: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    print("\n예상 소요 시간: 40~75분")
    print("  - 서울 수도권: 30~60분")
    print("  - 부산: 5~10분")
    print("  - 대구: 2~5분")
    print("\n[주의] 창을 닫지 마세요. 백그라운드에서 실행 중입니다...")
    print("="*80 + "\n")

    logging.info("\n" + "="*80)
    logging.info(f"부태리 신고가 완전 자동 업데이트 시작: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("="*80)

    if not os.path.exists(DB_PATH):
        logging.error(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")
        return

    db_conn = sqlite3.connect(DB_PATH, check_same_thread=False)

    total_success = 0

    try:
        # 1단계: 각 리스트별 데이터 갱신
        for config in MONITORING_CONFIGS:
            list_name = config['list_name']
            if update_apartment_data(db_conn, list_name):
                total_success += 1

        # 2단계: HTML 파일 생성
        logging.info("\n" + "="*60)
        logging.info("HTML 파일 생성 시작")
        logging.info("="*60)

        for config in MONITORING_CONFIGS:
            list_name = config['list_name']
            output_file = config['output_file']
            output_path = os.path.join(OUTPUT_DIR, output_file)

            try:
                apt_list = get_new_high_data(db_conn, list_name)

                if apt_list:
                    html_content = generate_html(apt_list, list_name)

                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(html_content)

                    logging.info(f"✓ '{list_name}' HTML 저장: {output_file} ({len(apt_list)}건)")
                else:
                    logging.warning(f"'{list_name}' 신고가 데이터 없음")

            except Exception as e:
                logging.error(f"'{list_name}' HTML 생성 오류: {str(e)}")

    finally:
        db_conn.close()

    end_time = datetime.now()
    elapsed = end_time - start_time

    logging.info("\n" + "="*80)
    logging.info(f"전체 작업 완료!")
    logging.info(f"  소요 시간: {elapsed}")
    logging.info(f"  성공: {total_success}/{len(MONITORING_CONFIGS)}")
    logging.info("="*80)

if __name__ == "__main__":
    main()
