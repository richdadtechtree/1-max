"""
부태리 신고가 자동 HTML 생성 스크립트
매일 새벽 6시 실행용 (Windows 작업 스케줄러)

사용법:
1. Windows 작업 스케줄러에 등록
2. 매일 06:00에 자동 실행
3. 서울&수도권, 부산, 대구 신고가 HTML 자동 생성
"""

import sqlite3
import os
import sys
from datetime import datetime, timedelta
from collections import Counter
from html import escape
import logging

# 로깅 설정
log_dir = os.path.join(os.path.dirname(__file__), 'logs')
os.makedirs(log_dir, exist_ok=True)

log_file = os.path.join(log_dir, f'auto_update_{datetime.now().strftime("%Y%m%d")}.log')
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

# 출력 디렉토리 생성
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 모니터링 리스트와 출력 파일 매핑
MONITORING_CONFIGS = [
    {
        'list_name': '서울 수도권',
        'output_file': '서울&수도권 신고가.html'
    },
    {
        'list_name': '부산',
        'output_file': '부산신고가.html'
    },
    {
        'list_name': '대구',
        'output_file': '대구신고가.html'
    }
]


def get_new_high_data(db_conn, list_name):
    """특정 모니터링 리스트의 신고가 데이터 조회"""
    try:
        cursor = db_conn.cursor()

        # 리스트 ID 조회
        cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (list_name,))
        result = cursor.fetchone()

        if not result:
            logging.warning(f"모니터링 리스트 '{list_name}'을 찾을 수 없습니다.")
            return []

        list_id = result[0]

        # 신고가 데이터 조회 (핑크색 단지: prev_max_price > 0 AND last_max_price > prev_max_price)
        cursor.execute("""
            SELECT
                apt_name,
                area,
                sido,
                sigungu,
                dong as location_dong,
                build_year,
                max_price_date as date,
                last_max_price as new_price,
                max_price_floor as floor,
                max_price_dong as dong,
                prev_max_price as old_price,
                prev_max_date as old_date,
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
            apt_dict = {
                'apt_name': row[0],
                'area': row[1],
                'sido': row[2],
                'sigungu': row[3],
                'location_dong': row[4],
                'build_year': row[5],
                'date': row[6],
                'new_price': row[7],
                'floor': row[8],
                'dong': row[9],
                'old_price': row[10] or 0,
                'old_date': row[11],
                'old_floor': row[12]
            }
            apt_list.append(apt_dict)

        logging.info(f"'{list_name}' 신고가 {len(apt_list)}건 조회 완료")
        return apt_list

    except Exception as e:
        logging.error(f"'{list_name}' 신고가 데이터 조회 중 오류: {str(e)}")
        import traceback
        logging.error(traceback.format_exc())
        return []


def generate_html(apt_list, list_name):
    """신고가 HTML 생성 (기존 코드 재사용)"""
    from html import escape

    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    current_year = datetime.now().year

    # 신고가 높은 순으로 정렬
    apt_list = sorted(apt_list, key=lambda x: x.get('new_price', 0), reverse=True)
    total = len(apt_list)

    # 지역별 건수 집계 및 연식별 카운트
    region_counter = Counter()
    young_count = 0
    very_young_count = 0
    old_count = 0

    for apt in apt_list:
        sido = apt.get('sido', '')
        sigungu = apt.get('sigungu', '')
        location_dong = apt.get('location_dong', '')
        if sido and sigungu:
            sigungu_clean = sigungu.split('(')[0] if '(' in sigungu else sigungu
            # 성남시 특정 동 매핑
            if sido == "경기도" and "성남시" in sigungu:
                if location_dong in ['구미동', '금곡동', '대장동', '백현동', '분당동', '서현동', '수내동', '야탑동', '운중동', '정자동', '판교동', '삼평동', '동막동', '궁내동', '율동', '매송동']:
                    sigungu_clean = "성남시 분당구"
                elif location_dong in ['고등동', '금토동', '단대동', '복정동', '신흥동', '양지동', '오야동', '태평동', '신촌동', '수진동', '창곡동', '시흥동', '둔전동']:
                    sigungu_clean = "성남시 수정구"
                elif location_dong in ['갈현동', '도촌동', '상대원동', '성남동', '은행동', '중앙동', '하대원동', '금광동', '여수동']:
                    sigungu_clean = "성남시 중원구"
            # 용인시 특정 동 매핑
            elif sido == "경기도" and "용인시" in sigungu:
                if location_dong == "보정동":
                    sigungu_clean = "용인시 기흥구"
                elif location_dong in ["김량장동", "고림동"] or location_dong.startswith("이동읍"):
                    sigungu_clean = "용인시 처인구"
            # 고양시 특정 동 매핑
            elif sido == "경기도" and "고양시" in sigungu:
                if location_dong == "지축동":
                    sigungu_clean = "고양시 덕양구"
                elif location_dong == "풍동":
                    sigungu_clean = "고양시 일산동구"
            # 수원시 특정 동 매핑
            elif sido == "경기도" and "수원시" in sigungu:
                if location_dong == "인계동":
                    sigungu_clean = "수원시 팔달구"
                elif location_dong in ["이의동", "하동"]:
                    sigungu_clean = "수원시 영통구"
            region_key = f"{sido} {sigungu_clean}"
            region_counter[region_key] += 1

        # build_year 처리
        build_year = apt.get('build_year', '')
        if build_year == '분양':
            young_count += 1
            very_young_count += 1
            apt['is_young'] = True
            apt['is_very_young'] = True
            apt['is_old'] = False
        elif build_year:
            try:
                year = int(build_year)
                years_old = current_year - year

                if year < 2000:
                    old_count += 1
                    apt['is_old'] = True
                else:
                    apt['is_old'] = False

                if years_old <= 5:
                    young_count += 1
                    very_young_count += 1
                    apt['is_young'] = True
                    apt['is_very_young'] = True
                elif years_old <= 10:
                    young_count += 1
                    apt['is_young'] = True
                    apt['is_very_young'] = False
                else:
                    apt['is_young'] = False
                    apt['is_very_young'] = False
            except:
                apt['is_young'] = False
                apt['is_very_young'] = False
                apt['is_old'] = False
        else:
            apt['is_young'] = False
            apt['is_very_young'] = False
            apt['is_old'] = False

    sorted_regions = sorted(region_counter.items(), key=lambda x: x[1], reverse=True)

    # 카드 HTML 생성
    cards_html = []
    for apt in apt_list:
        name = escape(str(apt.get('apt_name', '')))
        area = escape(str(apt.get('area', '')))
        old_price = apt.get('old_price', 0) or 0
        new_price = apt.get('new_price', 0) or 0
        date = escape(str(apt.get('date', '')))
        floor = escape(str(apt.get('floor', '')))
        dong = escape(str(apt.get('dong', '-')))
        build_year = escape(str(apt.get('build_year', '')))

        old_date = escape(str(apt.get('old_date', '')))
        old_floor = escape(str(apt.get('old_floor', '')))

        sido = escape(str(apt.get('sido', '')))
        sigungu = escape(str(apt.get('sigungu', '')))
        location_dong = escape(str(apt.get('location_dong', '')))
        location = f"{sido} {sigungu} {location_dong}"

        region_for_filter = sigungu.split('(')[0] if '(' in sigungu else sigungu
        if sido == "경기도" and "용인시" in sigungu:
            if location_dong == "보정동":
                region_for_filter = "용인시 기흥구"
            elif location_dong in ["김량장동", "고림동"] or location_dong.startswith("이동읍"):
                region_for_filter = "용인시 처인구"
        elif sido == "경기도" and "고양시" in sigungu:
            if location_dong == "지축동":
                region_for_filter = "고양시 덕양구"
            elif location_dong == "풍동":
                region_for_filter = "고양시 일산동구"
        elif sido == "경기도" and "수원시" in sigungu:
            if location_dong == "인계동":
                region_for_filter = "수원시 팔달구"
            elif location_dong in ["이의동", "하동"]:
                region_for_filter = "수원시 영통구"

        inc = new_price - old_price if old_price else 0
        pct = f"{(inc/old_price*100):.1f}%" if old_price else "-"

        if old_price:
            old_str = f"{old_price:,}만원"
            if old_date or old_floor:
                old_str += " ("
                if old_date:
                    old_str += old_date
                if old_floor:
                    old_str += f"{' | ' if old_date else ''}{old_floor}층"
                old_str += ")"
        else:
            old_str = "-"

        new_str = f"{new_price:,}만원"
        inc_str = f"{inc:,}만원" if old_price else "-"

        # 연식 표시
        year_badge = ""
        data_year = ""
        if build_year == '분양':
            year_badge = "<span class='year-badge new'>분양</span>"
            data_year = str(current_year)
        elif build_year:
            try:
                year_int = int(build_year)
                if year_int < 2000:
                    year_badge = f"<span class='year-badge old'>{build_year}년</span>"
                elif current_year - year_int <= 5:
                    year_badge = f"<span class='year-badge very-young'>{build_year}년</span>"
                elif current_year - year_int <= 10:
                    year_badge = f"<span class='year-badge young'>{build_year}년</span>"
                else:
                    year_badge = f"<span class='year-badge'>{build_year}년</span>"
                data_year = build_year
            except:
                year_badge = f"<span class='year-badge'>{build_year}년</span>"
                data_year = build_year

        is_young = "1" if apt.get('is_young', False) else "0"
        is_very_young = "1" if apt.get('is_very_young', False) else "0"
        is_old = "1" if apt.get('is_old', False) else "0"

        card = f"""
        <section class="card" data-region="{sido} {region_for_filter}"
                 data-build-year="{data_year}" data-young="{is_young}"
                 data-very-young="{is_very_young}" data-old="{is_old}" data-area="{area}">
          <h3>{name} {year_badge} <span class="tag">전용면적 {area}㎡</span></h3>
          <div class="card-sub">📍 {location}</div>
          <div class="grid">
            <div class="k">이번 실거래</div>
            <div><span class="highlight">{new_str}</span> ({date}{f' | {floor}층' if floor else ''})</div>
            <div class="k">직전 최고가</div>
            <div>{old_str}</div>
            <div class="k">변화</div>
            <div class="rise">{inc_str} 상승 ({pct})</div>
          </div>
        </section>
        """
        cards_html.append(card)

    cards = "\n".join(cards_html)

    html_content = f"""<!DOCTYPE html>
    <html lang="ko">
    <head>
    <meta charset="utf-8"/>
    <meta content="width=device-width, initial-scale=1" name="viewport"/>
    <title>{escape(list_name)} 신고가 알림 - {escape(now)}</title>
    <style>
      :root {{
        --bg:#F2F2F7; --fg:#111; --sub:#6b7280; --card:#fff; --bd:#e5e7eb;
        --primary:#007AFF; --primary-dark:#0051D2; --accent:#ff3b30;
      }}
      *{{ box-sizing:border-box }}
      body{{
        margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Noto Sans KR","Apple SD Gothic Neo","Malgun Gothic",sans-serif;
        color:var(--fg); background:var(--bg);
      }}
      .page{{ max-width:960px; width:92vw; margin:24px auto 80px }}
      .header{{
        background:linear-gradient(135deg,var(--primary),#63A4FF);
        color:#fff; border-radius:16px; padding:20px 20px 16px;
        box-shadow:0 6px 18px rgba(0,0,0,.08);
      }}
      .header h1{{ margin:0 0 8px; font-size:22px; font-weight:800; letter-spacing:-.2px }}
      .header .sub{{ margin:0; opacity:.95; font-size:14px }}

      .info-box{{
        background:#fff; border-radius:12px; padding:16px; margin:12px 0 0;
        box-shadow:0 2px 8px rgba(0,0,0,.04); border:1px solid var(--bd);
      }}
      .info-title{{ font-weight:700; font-size:14px; margin-bottom:8px; color:var(--primary) }}
      .info-items{{ display:flex; flex-wrap:wrap; gap:8px }}
      .info-item{{
        background:#f8fafc; padding:6px 12px; border-radius:8px;
        font-size:13px; border:1px solid #e2e8f0;
      }}
      .info-item .num{{ font-weight:600; color:var(--primary) }}

      .cards{{ display:grid; gap:16px; margin-top:16px }}
      .card{{
        background:var(--card); border:1px solid var(--bd);
        border-radius:14px; padding:18px; box-shadow:0 2px 12px rgba(0,0,0,.04);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
      }}
      .card:hover{{ transform: translateY(-2px); box-shadow:0 4px 20px rgba(0,0,0,.08) }}
      .card h3{{
        margin:0 0 8px; font-size:17px; font-weight:700;
        display:flex; align-items:center; gap:8px; flex-wrap:wrap;
      }}
      .card-sub{{ color:var(--sub); font-size:13px; margin-bottom:12px }}
      .grid{{
        display:grid; grid-template-columns: auto 1fr; gap:8px 12px;
        font-size:14px; line-height:1.6;
      }}
      .k{{ color:var(--sub); font-weight:600 }}
      .highlight{{ color:var(--accent); font-weight:700; font-size:16px }}
      .rise{{ color:#10b981; font-weight:600 }}

      .tag{{
        background:#e0f2fe; color:#0369a1; padding:4px 10px;
        border-radius:999px; font-size:12px; font-weight:600;
      }}

      .year-badge{{
        display:inline-flex; padding:4px 10px; border-radius:999px;
        font-size:12px; font-weight:600; background:#f1f5f9; color:#334155;
      }}
      .year-badge.new{{ background:#dbeafe; color:#1e40af }}
      .year-badge.very-young{{ background:#dcfce7; color:#15803d }}
      .year-badge.young{{ background:#fef3c7; color:#a16207 }}
      .year-badge.old{{ background:#fee2e2; color:#b91c1c }}

      footer{{ margin:24px 0 0; color:#6b7280; font-size:12px; text-align:center }}
    </style>
    </head>
    <body>
      <div class="page">
        <header class="header">
          <h1>🔔 {escape(list_name)} 신고가 알림</h1>
          <p class="sub">자동 업데이트: {escape(now)} · 총 {total}건</p>

          <div class="info-box">
            <div class="info-title">📊 요약 정보</div>
            <div class="info-items">
              <div class="info-item">전체 <span class="num">{total}</span>건</div>
              <div class="info-item">10년이하 <span class="num">{young_count}</span>건</div>
              <div class="info-item">5년이하 <span class="num">{very_young_count}</span>건</div>
              <div class="info-item">2000년미만 <span class="num">{old_count}</span>건</div>
            </div>
          </div>
        </header>

        <section class="cards">
          {cards}
        </section>

        <footer>
          © 부태리의 실거래가 모니터링 시스템 (자동생성: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
        </footer>
      </div>
    </body>
    </html>
    """

    return html_content


def update_html_files():
    """모든 모니터링 리스트의 HTML 파일 업데이트"""
    logging.info("=" * 60)
    logging.info(f"신고가 HTML 자동 업데이트 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("=" * 60)

    try:
        # DB 연결
        if not os.path.exists(DB_PATH):
            logging.error(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")
            return

        db_conn = sqlite3.connect(DB_PATH, check_same_thread=False)

        success_count = 0
        total_count = len(MONITORING_CONFIGS)

        for config in MONITORING_CONFIGS:
            list_name = config['list_name']
            output_file = config['output_file']
            output_path = os.path.join(OUTPUT_DIR, output_file)

            try:
                logging.info(f"\n처리 중: '{list_name}' -> {output_file}")

                # 신고가 데이터 조회
                apt_list = get_new_high_data(db_conn, list_name)

                if not apt_list:
                    logging.warning(f"'{list_name}' 신고가 데이터 없음")
                    continue

                # HTML 생성
                html_content = generate_html(apt_list, list_name)

                # 파일 저장
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)

                logging.info(f"✓ '{list_name}' HTML 저장 완료: {output_path}")
                logging.info(f"  신고가 건수: {len(apt_list)}건")
                success_count += 1

            except Exception as e:
                logging.error(f"✗ '{list_name}' 처리 중 오류: {str(e)}", exc_info=True)

        db_conn.close()

        logging.info("\n" + "=" * 60)
        logging.info(f"업데이트 완료: {success_count}/{total_count} 성공")
        logging.info("=" * 60)

    except Exception as e:
        logging.error(f"업데이트 중 치명적 오류: {str(e)}", exc_info=True)


if __name__ == "__main__":
    update_html_files()
