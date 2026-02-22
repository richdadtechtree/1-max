
##gui에서 '신고가 다시 찾기' 눌러서 한번 조회한 단지는 또 눌러도 다시 조회하지 않고 새롭게 등록된 단지만 6년치 신고가 데이터를 찾게 수정해 줘.


import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import os
os.environ['PYTHONHTTPSVERIFY'] = '0'
import json
import sqlite3  # DB 지원 추가
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import logging
import time
import concurrent.futures
import threading
import schedule
import tkinter.font as tkFont
from plyer import notification  # 윈도우 알림
from PIL import Image, ImageTk, ImageGrab  # 캡쳐
import io
import random
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from tkinter import simpledialog



import webbrowser
import html  # html.escape 사용
# (선택) 월 이동 정확도 개선 시 사용
# from dateutil.relativedelta import relativedelta

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# --------------------------- 공통 네트워킹 유틸 ---------------------------
def build_session():
    """SSL 문제 해결된 세션"""
    import ssl
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
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
    s.verify = False  # SSL 검증 비활성화
    return s

# (connect, read) 타임아웃
API_TIMEOUT = (5, 15)

def jitter_sleep(max_ms=300):
    """0~max_ms ms 사이 임의 지연 (429 완화용)"""
    time.sleep(random.uniform(0, max_ms / 1000.0))

# --------------------------- DB 초기화 및 관리 ---------------------------
def init_database(db_path):
    """데이터베이스 초기화 및 스키마 생성"""
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON")
    cursor = conn.cursor()

    # 모니터링 리스트 테이블
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS monitoring_lists (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            is_active BOOLEAN DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # 아파트 정보 테이블 (확장된 스키마)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS apartments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            list_id INTEGER NOT NULL,
            apt_name TEXT NOT NULL,
            area TEXT,
            sido TEXT,
            sigungu TEXT,
            dong TEXT,
            sigungu_code TEXT,
            jibun_addr TEXT,
            build_year TEXT,
            prev_max_price INTEGER DEFAULT 0,
            prev_max_date TEXT,
            prev_max_floor TEXT,
            prev_max_dong TEXT,
            last_max_price INTEGER DEFAULT 0,
            max_price_date TEXT,
            max_price_floor TEXT,
            max_price_dong TEXT,
            last_update TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY (list_id) REFERENCES monitoring_lists(id) ON DELETE CASCADE
        )
    """)

    # 거래 데이터 테이블
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS trade_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            apt_id INTEGER NOT NULL,
            price REAL NOT NULL,
            trade_date DATE NOT NULL,
            area REAL,
            floor INTEGER,
            dong TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (apt_id) REFERENCES apartments(id) ON DELETE CASCADE
        )
    """)

    # 신고가 히스토리 테이블
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS notifications_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            apt_id INTEGER,
            apt_name TEXT NOT NULL,
            price REAL NOT NULL,
            prev_max_price REAL,
            trade_date DATE,
            area REAL,
            floor INTEGER,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (apt_id) REFERENCES apartments(id) ON DELETE CASCADE
        )
    """)

    # 평형별 순위 이력 테이블
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS area_ranking_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            list_id INTEGER NOT NULL,
            area_type TEXT NOT NULL,
            apt_name TEXT NOT NULL,
            area TEXT,
            ranking INTEGER NOT NULL,
            price INTEGER NOT NULL,
            sido TEXT,
            sigungu TEXT,
            dong TEXT,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (list_id) REFERENCES monitoring_lists(id) ON DELETE CASCADE
        )
    """)

    # 기존 테이블에 누락된 컬럼 추가 (마이그레이션)
    try:
        # monitoring_lists 테이블 컬럼 확인
        cursor.execute("PRAGMA table_info(monitoring_lists)")
        ml_columns = [col[1] for col in cursor.fetchall()]
        if 'is_active' not in ml_columns:
            cursor.execute("ALTER TABLE monitoring_lists ADD COLUMN is_active BOOLEAN DEFAULT 0")
        if 'updated_at' not in ml_columns:
            cursor.execute("ALTER TABLE monitoring_lists ADD COLUMN updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        # trade_data 테이블 컬럼 확인
        cursor.execute("PRAGMA table_info(trade_data)")
        td_columns = [col[1] for col in cursor.fetchall()]
        if 'created_at' not in td_columns:
            cursor.execute("ALTER TABLE trade_data ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        # notifications_history 테이블 컬럼 확인 및 추가
        cursor.execute("PRAGMA table_info(notifications_history)")
        nh_columns = [col[1] for col in cursor.fetchall()]

        if 'dong' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN dong TEXT")
            logging.info("notifications_history 테이블에 dong 컬럼 추가됨")

        if 'sido' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN sido TEXT")
            logging.info("notifications_history 테이블에 sido 컬럼 추가됨")

        if 'sigungu' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN sigungu TEXT")
            logging.info("notifications_history 테이블에 sigungu 컬럼 추가됨")

        if 'location_dong' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN location_dong TEXT")
            logging.info("notifications_history 테이블에 location_dong 컬럼 추가됨")

        if 'build_year' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN build_year TEXT")
            logging.info("notifications_history 테이블에 build_year 컬럼 추가됨")

        if 'prev_max_date' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN prev_max_date TEXT")
            logging.info("notifications_history 테이블에 prev_max_date 컬럼 추가됨")

        if 'prev_max_floor' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN prev_max_floor TEXT")
            logging.info("notifications_history 테이블에 prev_max_floor 컬럼 추가됨")

        if 'prev_max_dong' not in nh_columns:
            cursor.execute("ALTER TABLE notifications_history ADD COLUMN prev_max_dong TEXT")
            logging.info("notifications_history 테이블에 prev_max_dong 컬럼 추가됨")

        conn.commit()
    except Exception as e:
        logging.warning(f"컬럼 추가 중 경고: {e}")

    # 인덱스 생성 (성능 향상)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_apt_name ON apartments(apt_name)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_apt_sido ON apartments(sido)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_apt_sigungu ON apartments(sigungu)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_apt_sigungu_code ON apartments(sigungu_code)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_apt_list ON apartments(list_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_trade_apt ON trade_data(apt_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_trade_date ON trade_data(trade_date)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_noti_timestamp ON notifications_history(timestamp)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_noti_apt ON notifications_history(apt_id)")

    conn.commit()
    return conn

# ------------------------------------------------------------------------

class RealEstateMonitorApp:
    def __init__(self):
        self.palette = {
            'bg':            '#1E1E2F',
            'surface':       '#2A2A3D',
            'text_primary':  '#E0E0E8',
            'text_secondary':'#A0A0A8',
            'accent':        '#3AA6FF',
        }
        self.root = tk.Tk()
        self.root.title("부태리의 실거래가 모니터링")
        self.root.geometry("1500x800")  # 700 -> 850으로 변경
        
        # 공통 HTTP 세션
        self.http = build_session()
        
        # ★★★ API 캐싱 시스템 추가 ★★★
        self.api_cache = {}  # 캐시 저장소
        self.cache_ttl = timedelta(hours=24)  # 캐시 유효시간 24시간
        self.cache_hit_count = 0  # 캐시 히트 통계
        self.api_call_count = 0  # API 호출 통계

        # ▼▼ 추가: '선택시 해당 지역 전체 단지 모니터링' 체크 상태
        self.bulk_monitor_enabled = tk.BooleanVar(value=False)        
        
        # 기본 설정 (초기값)
        self.download_path = "C:\\Download"
        self.lawdong_path = "C:/law-dong/law-dong.txt"
        self.monitored_apts_file = os.path.join(self.download_path, "monitored_apts.json")
        self.db_path = os.path.join(self.download_path, "monitoring.db")  # 기본 DB 경로

        # 알림 설정
        self.auto_update_enabled = tk.BooleanVar(value=False)
        self.update_time = tk.StringVar(value="09:10")
        self.setup_fonts()

        # 설정 파일에서 경로 먼저 불러오기 (DB 초기화 전)
        self.load_settings()

        # DB 초기화 (설정에서 불러온 경로 사용)
        self.db_conn = init_database(self.db_path)

        # 백업 관련 설정
        self.backup_dir = os.path.join(self.download_path, "backups")
        self.auto_backup_interval = 24  # 24시간마다 자동 백업
        self.last_auto_backup = None
        os.makedirs(self.backup_dir, exist_ok=True)

        # ---- 여러 리스트 관리로 전환 ----
        self.monitored_lists = self.load_monitored_apts()  # {"lists": {...}, "active_list": "기본"} 형태로 로드/마이그레이션
        if "lists" not in self.monitored_lists:
            self.monitored_lists = {"lists": {"기본": []}, "active_list": "기본"}
        
        self.active_list = tk.StringVar(value=self.monitored_lists.get("active_list", "기본"))
        if self.active_list.get() not in self.monitored_lists["lists"]:
            self.monitored_lists["lists"][self.active_list.get()] = []        
        
        # 폴더 생성
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)

    
        # 신고가 히스토리 파일 경로 추가
        self.notifications_history_file = os.path.join(self.download_path, "notifications_history.json")
        
        # 신고가 히스토리 로드
        self.notifications_history = self.load_notifications_history()
        
        # 폰트 설정
        self.setup_fonts()
        
        # 법정동 코드 관련 변수 초기화
        self.region_codes = {}
        self.sido_list = []
        self.sigungu_dict = {}
        self.dong_dict = {}

        # 마지막 검색 지역 저장 변수
        self.last_search_region = ""
        
        # 법정동 파일 로드
        self.load_lawdong_file()
    
        self.sort_column = "last_max_price"  # 기본 정렬 열: 현재 최고가
        self.sort_reverse = True  # 기본 정렬 방향: 내림차순 (높은 값부터)
        
        # API 키 설정 (환경변수로 빼는 것을 권장)
        self.service_key = "Vs5lXsSo6iEI8no3pP%2FT0udWF9s7Cc8oP1SIWnEI5F4h6dKq92fLvnKmxkoWGJxSeW2%2FSOLQECGxOJzWcjJEXQ%3D%3D"
        
        # 모니터링 중인 아파트 목록


        # GUI 설정
        self.setup_gui()

        # 마지막 검색 지역 GUI에 표시
        if self.last_search_region:
            self.last_search_label.config(text=f"이전 검색: {self.last_search_region}")

        # 스케줄러 설정
        self.setup_scheduler()
        
        # 종료 시 처리
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    @property
    def monitored_apts(self):
        # 현재 활성 리스트의 실제 목록을 반환
        return self.monitored_lists["lists"].setdefault(self.active_list.get(), [])
    
    @monitored_apts.setter
    def monitored_apts(self, new_list):
        # 현재 활성 리스트에 새 목록을 설정
        self.monitored_lists["lists"][self.active_list.get()] = new_list
    
    def get_cached_api_data(self, sigungu_code, deal_ymd, api_type='existing'):
        """캐시된 API 데이터 반환 또는 새로 조회"""
        cache_key = (sigungu_code, deal_ymd, api_type)
        
        # 캐시 확인
        if cache_key in self.api_cache:
            cached_data, timestamp = self.api_cache[cache_key]
            if datetime.now() - timestamp < self.cache_ttl:
                self.cache_hit_count += 1
                logging.info(f"캐시 히트: {cache_key} (총 히트: {self.cache_hit_count})")
                return cached_data
            else:
                del self.api_cache[cache_key]
        
        # API 호출
        self.api_call_count += 1
        logging.info(f"API 호출: {cache_key} (총 호출: {self.api_call_count})")
        
        if api_type == 'existing':
            url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                   f"?serviceKey={self.service_key}"
                   f"&LAWD_CD={sigungu_code}"
                   f"&DEAL_YMD={deal_ymd}"
                   f"&numOfRows=1000")
        else:
            url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcSilvTrade/getRTMSDataSvcSilvTrade"
                   f"?serviceKey={self.service_key}"
                   f"&LAWD_CD={sigungu_code}"
                   f"&DEAL_YMD={deal_ymd}"
                   f"&numOfRows=1000")
        
        try:
            response = self.http.get(url, timeout=API_TIMEOUT)
            jitter_sleep(200)
            
            if response.status_code == 200:
                root = ET.fromstring(response.text)
                items = root.findall('.//item')
                
                parsed_data = []
                for item in items:
                    try:
                        data = {
                            'apt_name': item.findtext('aptNm', '').strip(),
                            'dong': item.findtext('umdNm', '').strip(),
                            'area': float(item.findtext('excluUseAr', '0')),
                            'year': int(item.findtext('dealYear')),
                            'month': int(item.findtext('dealMonth')),
                            'day': int(item.findtext('dealDay', '1')),
                            'price': int(item.findtext('dealAmount').replace(',', '')),
                            'floor': int(item.findtext('floor', '0')),
                            'jibun': item.findtext('jibun', '').strip(),
                            'kaptdong': item.findtext('kaptdong', ''),
                            'build_year': item.findtext('buildYear', '').strip()
                        }
                        parsed_data.append(data)
                    except (ValueError, TypeError):
                        continue
                
                # 캐시 저장
                self.api_cache[cache_key] = (parsed_data, datetime.now())
                
                # 캐시 크기 관리
                if len(self.api_cache) > 100:
                    self.cleanup_old_cache()
                
                return parsed_data
            return []
        except Exception as e:
            logging.error(f"API 호출 중 오류: {str(e)}")
            return []


    def _build_search_index_text(self, apt: dict) -> str:
        """
        한 아파트(dict)를 검색 가능한 문자열로 합쳐줌.
        - 단어 포함(부분일치) 위주 검색, 공백으로 여러 키워드 AND 매칭
        """
        fields = [
            apt.get('apt_name', ''),
            str(apt.get('build_year', '')),
            f"{apt.get('area','')}㎡",
            f"{apt.get('sido','')} {apt.get('sigungu','')} {apt.get('dong','')}",
            str(apt.get('max_price_dong', '')),
            str(apt.get('max_price_floor', '')),
            str(apt.get('prev_max_price','')),
            str(apt.get('prev_max_date','')),
            str(apt.get('last_max_price','')),
            str(apt.get('max_price_date','')),
            str(apt.get('last_update','')),
        ]
        return " ".join(map(str, fields)).lower()
    
    def _filter_apts_by_query(self, apts: list, query: str) -> list:
        """
        공백으로 분리된 여러 키워드를 모두 포함(AND)하는 항목만 반환.
        """
        q = (query or '').strip().lower()
        if not q:
            return apts
        tokens = [t for t in q.split() if t]
        if not tokens:
            return apts
    
        filtered = []
        for apt in apts:
            hay = self._build_search_index_text(apt)
            if all(tok in hay for tok in tokens):
                filtered.append(apt)
        return filtered
    
    def calculate_regional_rankings(self):
        """전체 모니터링 단지 기준 지역구별 타입별 순위 계산"""
        rankings = {}
        
        # 모든 모니터링 단지를 순회하며 지역/타입별로 분류
        for apt in self.monitored_apts:
            sido = apt.get('sido', '')
            sigungu = apt.get('sigungu', '').split('(')[0] if apt.get('sigungu') else ''
            
            if not sido or not sigungu:
                continue
                
            region_key = f"{sido} {sigungu}"
            
            if region_key not in rankings:
                rankings[region_key] = {'84': [], '59': []}
            
            # 면적과 가격 정보 추출
            try:
                area_str = str(apt.get('area', '')).replace('㎡', '').strip()
                area = float(area_str)
                price = apt.get('last_max_price', 0)
                apt_name = apt.get('apt_name', '')
                
                if not apt_name or price <= 0:
                    continue
                
                # 84타입 (82-86㎡)
                if 82 <= area <= 86:
                    rankings[region_key]['84'].append((apt_name, price))
                # 59타입 (57-61㎡)
                elif 57 <= area <= 61:
                    rankings[region_key]['59'].append((apt_name, price))
            except:
                continue
        
        # 각 지역/타입별로 가격순 정렬 후 TOP 5 순위 부여
        final_rankings = {}
        for region_key in rankings:
            final_rankings[region_key] = {'84': {}, '59': {}}
            
            # 84타입 TOP 5
            sorted_84 = sorted(rankings[region_key]['84'], key=lambda x: x[1], reverse=True)[:5]
            for rank, (name, _) in enumerate(sorted_84, 1):
                final_rankings[region_key]['84'][name] = rank
            
            # 59타입 TOP 5
            sorted_59 = sorted(rankings[region_key]['59'], key=lambda x: x[1], reverse=True)[:5]
            for rank, (name, _) in enumerate(sorted_59, 1):
                final_rankings[region_key]['59'][name] = rank
        
        return final_rankings
    
    def backfill_build_year(self, apt_info, months=24, fallback_text=None):
        """
        apt_info에 build_year가 비어 있으면 resolve_build_year로 보정한다.
        - months: 캐시 조회 기간(기본 24개월)
        - fallback_text: 목록 문자열(있으면 '(준공: ####년)' / '분양중' 파싱)
        반환: 보정된 연식 문자열('2012' / '분양' / '') 
        """
        by = (apt_info.get('build_year') or '').strip()
        if by:
            return by
    
        try:
            by = self.resolve_build_year(
                sigungu_code=apt_info.get('sigungu_code', ''),
                dong=apt_info.get('dong', ''),
                apt_name=apt_info.get('apt_name', ''),
                months=months,
                fallback_text=fallback_text
            )
            if by:
                apt_info['build_year'] = by
                return by
        except Exception as e:
            logging.error(f"build_year 보정 중 오류: {e}")
        return ''



    
    def resolve_build_year(self, sigungu_code, dong, apt_name, months=12, fallback_text=None):
        """
        최근 months개월의 캐시에서 해당 단지의 buildYear를 우선 탐색.
        없으면 fallback_text(목록 문자열)에서 (준공: ####년) 또는 (분양중) 파싱.
        반환: '2012' / '분양' / ''(미상)
        """
        import re
        now = datetime.now()
    
        # 1) 캐시된 '기축/신축' API에서 buildYear 찾기
        for m in range(months):
            deal_ymd = (now - timedelta(days=30*m)).strftime("%Y%m")
            # 기축
            try:
                for it in self.get_cached_api_data(sigungu_code, deal_ymd, 'existing'):
                    if it.get('apt_name','').strip() == apt_name and it.get('dong','').strip() == dong:
                        by = (it.get('build_year') or '').strip()
                        if by and by.isdigit():
                            return by
            except Exception:
                pass
            # 신축(분양권)
            try:
                for it in self.get_cached_api_data(sigungu_code, deal_ymd, 'new'):
                    if it.get('apt_name','').strip() == apt_name and it.get('dong','').strip() == dong:
                        by = (it.get('build_year') or '').strip()
                        # 신축 API는 보통 buildYear가 없지만, 있으면 사용. 없으면 '분양' 처리
                        if by and by.isdigit():
                            return by
                        # 신축 표기만 있고 연식이 없으면 '분양'으로 간주
                        if not by:
                            return '분양'
            except Exception:
                pass
    
        # 2) 목록 문자열에서 후순위 파싱
        if fallback_text:
            # (준공: 2012년)
            m = re.search(r'준공:\s*(\d{4})년', fallback_text)
            if m:
                return m.group(1)
            # (분양중)
            if '분양중' in fallback_text:
                return '분양'
    
        return ''



    
    def cleanup_old_cache(self):
        """오래된 캐시 항목 정리"""
        current_time = datetime.now()
        expired_keys = []
        for key, (data, timestamp) in self.api_cache.items():
            if current_time - timestamp >= self.cache_ttl:
                expired_keys.append(key)
        for key in expired_keys:
            del self.api_cache[key]
        logging.info(f"캐시 정리: {len(expired_keys)}개 항목 제거")            

    def filter_apt_data(self, all_data, apt_name, dong, target_area):
        """전체 데이터에서 특정 아파트 정보만 필터링"""
        filtered = []
        for item in all_data:
            if (item['apt_name'] == apt_name and 
                item['dong'] == dong and 
                abs(item['area'] - target_area) <= 1):
                
                building_dong = item['kaptdong']
                if not building_dong and item['jibun'] and '동' in item['jibun']:
                    dong_parts = item['jibun'].split('동')
                    if len(dong_parts) > 0 and dong_parts[0].isdigit():
                        building_dong = dong_parts[0] + '동'
                building_dong = building_dong or '-'
                
                trade = {
                    'date': datetime(item['year'], item['month'], item['day']),
                    'price': item['price'],
                    'floor': item['floor'],
                    'area': item['area'],
                    'dong': building_dong
                }
                filtered.append(trade)
        return filtered
    
    def adjust_column_widths(self):
        """Treeview의 모든 열을 내용에 맞춰 자동으로 너비 조정"""
        font = tkFont.Font(font=self.font_normal)
        for col in self.apt_tree["columns"]:
            header_text = self.apt_tree.heading(col, option="text")
            max_width = font.measure(header_text)
            for iid in self.apt_tree.get_children():
                cell_text = self.apt_tree.set(iid, col)
                max_width = max(max_width, font.measure(cell_text))
            self.apt_tree.column(col, width=max_width + 20)

    def load_notifications_history(self):
        """JSON 파일에서 신고가 히스토리 로드 (r7 방식)"""
        if os.path.exists(self.notifications_history_file):
            try:
                with open(self.notifications_history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                    # timestamp 변환
                    for notification in data:
                        if 'timestamp' in notification and isinstance(notification['timestamp'], str):
                            try:
                                notification['timestamp'] = datetime.strptime(notification['timestamp'], '%Y-%m-%d %H:%M:%S')
                            except:
                                notification['timestamp'] = datetime.now()

                    # 최근 30일 데이터만 필터링
                    cutoff_date = datetime.now() - timedelta(days=30)
                    filtered_data = [n for n in data if n.get('timestamp', datetime.now()) > cutoff_date]

                    logging.info(f"[JSON 로드] 신고가 히스토리 로드 완료 ({len(filtered_data)}개 그룹)")
                    return filtered_data
            except Exception as e:
                logging.error(f"신고가 히스토리 로드 중 오류: {str(e)}")
                import traceback
                logging.error(traceback.format_exc())
        return []

    def migrate_notifications_to_db(self):
        """JSON 신고가 히스토리를 DB로 마이그레이션"""
        try:
            with open(self.notifications_history_file, 'r', encoding='utf-8') as f:
                json_data = json.load(f)

            cursor = self.db_conn.cursor()

            for notification in json_data:
                timestamp = notification.get('timestamp')
                if isinstance(timestamp, str):
                    try:
                        timestamp = datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S')
                    except:
                        timestamp = datetime.now()
                elif not isinstance(timestamp, datetime):
                    timestamp = datetime.now()

                # apt_list에서 정보 추출
                for apt in notification.get('apt_list', []):
                    # apt_id 찾기 (apt_name으로 검색)
                    apt_name = apt.get('apt_name', '')
                    if not apt_name:
                        continue

                    cursor.execute("SELECT id FROM apartments WHERE apt_name = ? LIMIT 1", (apt_name,))
                    result = cursor.fetchone()
                    apt_id = result[0] if result else None

                    # apt_id가 없으면 건너뛰기 (NOT NULL 제약조건)
                    if not apt_id:
                        logging.warning(f"신고가 히스토리 마이그레이션: {apt_name}의 아파트 ID를 찾을 수 없어 건너뜁니다.")
                        continue

                    # price가 None이면 건너뛰기 (NOT NULL 제약조건)
                    price = apt.get('new_max_price') or apt.get('new_price')
                    if not price:
                        logging.warning(f"신고가 히스토리 마이그레이션: {apt_name}의 가격 정보가 없어 건너뜁니다.")
                        continue

                    cursor.execute("""
                        INSERT INTO notifications_history
                        (apt_id, apt_name, price, prev_max_price, trade_date, area, floor, timestamp)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        apt_id,
                        apt_name,
                        price,
                        apt.get('prev_max_price'),
                        apt.get('date'),
                        apt.get('area'),
                        apt.get('floor'),
                        timestamp.strftime('%Y-%m-%d %H:%M:%S')
                    ))

            self.db_conn.commit()
            logging.info(f"신고가 히스토리 JSON -> DB 마이그레이션 완료")

            # 백업 후 JSON 파일 이름 변경
            backup_name = self.notifications_history_file.replace('.json', '_migrated_backup.json')
            # 백업 파일이 이미 있으면 삭제
            if os.path.exists(backup_name):
                os.remove(backup_name)
                logging.info(f"기존 백업 파일 삭제: {backup_name}")
            os.rename(self.notifications_history_file, backup_name)
            logging.info(f"기존 JSON 파일을 백업했습니다: {backup_name}")

        except Exception as e:
            logging.error(f"신고가 히스토리 마이그레이션 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            self.db_conn.rollback()

    def save_notifications_history(self):
        """신고가 히스토리를 JSON 파일로 저장 (r7 방식 - 안정적)"""
        try:
            # JSON 파일로 저장
            save_dir = os.path.dirname(self.notifications_history_file)
            os.makedirs(save_dir, exist_ok=True)

            save_data = []
            cutoff_date = datetime.now() - timedelta(days=30)

            for notification in self.notifications_history:
                # 30일 이전 데이터는 제외
                if notification.get('timestamp', datetime.now()) <= cutoff_date:
                    continue

                notification_copy = notification.copy()
                if 'timestamp' in notification_copy and isinstance(notification_copy['timestamp'], datetime):
                    notification_copy['timestamp'] = notification_copy['timestamp'].strftime('%Y-%m-%d %H:%M:%S')
                save_data.append(notification_copy)

            with open(self.notifications_history_file, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)

            logging.info(f"[JSON 저장] 신고가 히스토리 저장 완료 ({len(save_data)}개 그룹)")

        except Exception as e:
            logging.error(f"신고가 히스토리 저장 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())

    def save_last_search_region(self):
        """마지막 검색 지역만 빠르게 저장"""
        try:
            settings_file = os.path.join(os.getcwd(), 'monitor_settings.json')
            settings_data = {}

            # 기존 설정 로드
            if os.path.exists(settings_file):
                try:
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        settings_data = json.load(f)
                except:
                    pass

            # 마지막 검색 지역 업데이트
            settings_data['last_search_region'] = self.last_search_region

            # 저장
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"마지막 검색 지역 저장 중 오류: {str(e)}")

    def save_last_active_list(self):
        """마지막 활성 리스트만 빠르게 저장"""
        try:
            settings_file = os.path.join(os.getcwd(), 'monitor_settings.json')
            settings_data = {}

            # 기존 설정 로드
            if os.path.exists(settings_file):
                try:
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        settings_data = json.load(f)
                except:
                    pass

            # 마지막 활성 리스트 업데이트
            settings_data['active_list'] = self.active_list.get()

            # 저장
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)

            logging.info(f"마지막 활성 리스트 저장: {self.active_list.get()}")
        except Exception as e:
            logging.error(f"마지막 활성 리스트 저장 중 오류: {str(e)}")

    def load_settings(self):
        """설정 파일 불러오기"""
        try:
            settings_file = os.path.join(os.getcwd(), 'monitor_settings.json')
            if os.path.exists(settings_file):
                try:
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        settings_data = json.load(f)
                        if 'download_path' in settings_data:
                            self.download_path = settings_data['download_path']
                            self.monitored_apts_file = os.path.join(self.download_path, "monitored_apts.json")
                        if 'lawdong_path' in settings_data:
                            self.lawdong_path = settings_data['lawdong_path']
                        if 'db_path' in settings_data:
                            self.db_path = settings_data['db_path']
                            logging.info(f"설정에서 DB 경로 로드: {self.db_path}")
                        if 'auto_update' in settings_data:
                            self.auto_update_enabled.set(settings_data['auto_update'])
                        if 'update_time' in settings_data:
                            self.update_time.set(settings_data['update_time'])
                        # active_list 로드
                        if 'active_list' in settings_data:
                            self.loaded_active_list = settings_data['active_list']
                            logging.info(f"설정에서 활성 리스트 로드: {self.loaded_active_list}")
                        # 마지막 백업 디렉토리 로드
                        if 'last_backup_dir' in settings_data:
                            self.last_backup_dir = settings_data['last_backup_dir']
                            logging.info(f"설정에서 마지막 백업 디렉토리 로드: {self.last_backup_dir}")
                        # 마지막 검색 지역 로드
                        if 'last_search_region' in settings_data:
                            self.last_search_region = settings_data['last_search_region']
                            logging.info(f"설정에서 마지막 검색 지역 로드: {self.last_search_region}")
                except Exception as e:
                    logging.error(f"설정 파일 불러오기 중 오류: {str(e)}")
        except Exception as e:
            logging.error(f"설정 파일 경로 확인 중 오류: {str(e)}")
                
    def setup_fonts(self):
        """폰트 설정 (2pt씩 확대)"""
        self.font_normal = ('Malgun Gothic', 11)
        self.font_large  = ('Malgun Gothic', 13)
        self.font_title  = ('Malgun Gothic', 16, 'bold')
        self.font_button = ('Malgun Gothic', 11)
        
    def load_lawdong_file(self):
        """법정동 코드 파일 로드"""
        try:
            if not os.path.exists(self.lawdong_path):
                messagebox.showerror("오류", "법정동 코드 파일이 존재하지 않습니다.")
                return False
            for encoding in ['cp949', 'euc-kr', 'utf-8']:
                try:
                    with open(self.lawdong_path, 'r', encoding=encoding) as file:
                        law_dong_data = []
                        for line in file:
                            parts = line.strip().split('\t')
                            if len(parts) < 2:
                                continue
                            code = parts[0].strip()
                            name = parts[1].strip()
                            if any('폐지' in part for part in parts):
                                continue
                            sido_code = code[:2]
                            sigungu_code = code[2:5]
                            dong_code = code[5:]
                            law_dong_data.append({
                                'code': code,
                                'name': name,
                                'sido_code': sido_code,
                                'sigungu_code': sigungu_code, 
                                'dong_code': dong_code
                            })
                        self.sido_list = []
                        self.sigungu_dict = {}
                        self.dong_dict = {}
                        self.region_codes = {}
                        self.sigungu_to_full_info = {}
                        self.special_sigungu_names = {}
                        self.gu_info = {}  # 구 정보 저장 (구이름 -> 구코드)
                        # 시도
                        sido_data = [item for item in law_dong_data if item['code'].endswith('00000000')]
                        for sido in sido_data:
                            sido_name = sido['name']
                            self.sido_list.append(sido_name)
                            self.sigungu_dict[sido_name] = []
                        # 시군구 및 구
                        sigungu_data = [item for item in law_dong_data 
                                       if item['dong_code'] == '00000' and not item['code'].endswith('00000000')]
                        sigungu_name_count = {}
                        temp_sigungu_list = []
                        for item in sigungu_data:
                            names = item['name'].split()
                            if len(names) >= 2:
                                if len(names) >= 3 and names[2].endswith('구'):
                                    si_name = names[1]
                                    sigungu_name_count[si_name] = sigungu_name_count.get(si_name, 0) + 1
                                    if si_name not in temp_sigungu_list:
                                        temp_sigungu_list.append(si_name)
                                else:
                                    sigungu_name = names[1]
                                    sigungu_name_count[sigungu_name] = sigungu_name_count.get(sigungu_name, 0) + 1
                                    temp_sigungu_list.append(sigungu_name)
                        duplicate_sigungu_names = {name for name, count in sigungu_name_count.items() if count > 1}
                        processed_si = set()
                        for item in sigungu_data:
                            names = item['name'].split()
                            if len(names) >= 2:
                                sido_name = names[0]
                                if len(names) >= 3 and names[2].endswith('구'):
                                    si_name = names[1]
                                    gu_name = names[2]
                                    gu_code = f"{item['sido_code']}{item['sigungu_code']}"
                                    self.gu_info[f"{sido_name}_{si_name}_{gu_name}"] = gu_code
                                    if (sido_name, si_name) not in processed_si:
                                        processed_si.add((sido_name, si_name))
                                        display_name = si_name
                                        if si_name in duplicate_sigungu_names:
                                            # 시도명에서 광역시/특별시/도 제거하고 사용
                                            sido_abbr = sido_name.replace('광역시', '').replace('특별시', '').replace('특별자치시', '').replace('특별자치도', '').replace('도', '')
                                            display_name = f"{si_name}({sido_abbr})"
                                        if sido_name in self.sigungu_dict:
                                            self.sigungu_dict[sido_name].append(display_name)
                                            self.dong_dict[display_name] = []
                                            si_code = gu_code[:5]
                                            self.sigungu_to_full_info[display_name] = (sido_name, si_name, si_code)
                                else:
                                    sigungu_name = names[1]
                                    if sido_name in self.sido_list:
                                        sigungu_full_code = f"{item['sido_code']}{item['sigungu_code']}"
                                        display_name = sigungu_name
                                        if sigungu_name in duplicate_sigungu_names:
                                            # 시도명에서 광역시/특별시/도 제거하고 사용
                                            sido_abbr = sido_name.replace('광역시', '').replace('특별시', '').replace('특별자치시', '').replace('특별자치도', '').replace('도', '')
                                            display_name = f"{sigungu_name}({sido_abbr})"
                                            self.special_sigungu_names[display_name] = (sido_name, sigungu_name)
                                        self.sigungu_to_full_info[display_name] = (sido_name, sigungu_name, sigungu_full_code)
                                        if display_name not in self.sigungu_dict[sido_name]:
                                            self.sigungu_dict[sido_name].append(display_name)
                                            self.dong_dict[display_name] = []
                        # 읍면동
                        for item in law_dong_data:
                            if item['dong_code'] != '00000' and not item['code'].endswith('00000'):
                                names = item['name'].split()
                                if len(names) >= 4 and names[2].endswith('구'):
                                    sido_name = names[0]
                                    si_name = names[1]
                                    gu_name = names[2]
                                    dong_name = names[3]
                                    si_display_name = None
                                    for display_name, (s_name, sg_name, _) in self.sigungu_to_full_info.items():
                                        if s_name == sido_name and sg_name == si_name:
                                            si_display_name = display_name
                                            break
                                    if si_display_name:
                                        if gu_name not in self.dong_dict[si_display_name]:
                                            self.dong_dict[si_display_name].append(gu_name)
                                        gu_key = f"{si_display_name}_{gu_name}"
                                        if gu_key not in self.dong_dict:
                                            self.dong_dict[gu_key] = []
                                        if dong_name not in self.dong_dict[gu_key]:
                                            self.dong_dict[gu_key].append(dong_name)
                                        sigungu_code_5digits = f"{item['sido_code']}{item['sigungu_code']}"
                                        self.region_codes[f"{si_display_name}_{gu_name}_{dong_name}"] = (item['code'], sigungu_code_5digits)
                                elif len(names) >= 3:
                                    sido_name = names[0]
                                    sigungu_name = names[1]
                                    dong_name = names[2]
                                    display_name = None
                                    for d_name, (s_name, sg_name, _) in self.sigungu_to_full_info.items():
                                        if s_name == sido_name and sg_name == sigungu_name:
                                            display_name = d_name
                                            break
                                    if display_name and display_name in self.dong_dict:
                                        if dong_name not in self.dong_dict[display_name]:
                                            self.dong_dict[display_name].append(dong_name)
                                            sigungu_code_5digits = f"{item['sido_code']}{item['sigungu_code']}"
                                            self.region_codes[(sido_name, display_name, dong_name)] = (item['code'], sigungu_code_5digits)
                        self.sido_list = sorted(set(self.sido_list))
                        for sido in self.sido_list:
                            self.sigungu_dict[sido] = sorted(set(self.sigungu_dict[sido]))
                        for key in self.dong_dict:
                            self.dong_dict[key] = sorted(set(self.dong_dict[key]))

                        # 디버깅: 중복 시군구 확인
                        logging.info(f"[법정동코드 로드 완료]")
                        for sigungu_key in sorted(self.dong_dict.keys()):
                            if '동구' in sigungu_key or '서구' in sigungu_key or '중구' in sigungu_key:
                                logging.info(f"  {sigungu_key}: {len(self.dong_dict[sigungu_key])}개 동")

                        return True
                except UnicodeDecodeError:
                    continue
            messagebox.showerror("오류", "법정동 코드 파일을 읽을 수 없습니다.")
            return False
        except Exception as e:
            messagebox.showerror("오류", f"법정동 코드 파일 로드 중 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def migrate_json_to_db(self):
        """기존 JSON 데이터를 DB로 마이그레이션"""
        try:
            # JSON 파일이 존재하는지 확인
            if not os.path.exists(self.monitored_apts_file):
                logging.info("마이그레이션할 JSON 파일이 없습니다.")
                return

            # JSON 데이터 로드
            with open(self.monitored_apts_file, 'r', encoding='utf-8') as f:
                json_data = json.load(f)

            cursor = self.db_conn.cursor()

            # 구형 리스트 형식 처리
            if isinstance(json_data, list):
                json_data = {"lists": {"기본": json_data}, "active_list": "기본"}

            # 각 리스트 마이그레이션
            for list_name, apt_list in json_data.get("lists", {}).items():
                # 리스트가 이미 DB에 있는지 확인
                cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (list_name,))
                result = cursor.fetchone()

                if result:
                    list_id = result[0]
                    # ⚠️ 기존 데이터를 삭제하지 않음 (데이터 손실 방지)
                    # 대신 중복 체크를 하면서 아파트를 추가합니다
                    logging.info(f"리스트 '{list_name}' 발견 - 기존 데이터 유지하며 병합")
                else:
                    # 새 리스트 생성
                    is_active = 1 if list_name == json_data.get("active_list", "기본") else 0
                    cursor.execute(
                        "INSERT INTO monitoring_lists (name, is_active) VALUES (?, ?)",
                        (list_name, is_active)
                    )
                    list_id = cursor.lastrowid

                # 아파트 데이터 마이그레이션 (중복 체크)
                for apt in apt_list:
                    # 중복 체크: 같은 list_id, apt_name, sigungu_code를 가진 아파트가 있는지 확인
                    cursor.execute("""
                        SELECT id FROM apartments
                        WHERE list_id = ? AND apt_name = ? AND sigungu_code = ?
                    """, (list_id, apt.get('apt_name', ''), apt.get('sigungu_code', '')))

                    existing_apt = cursor.fetchone()

                    if existing_apt:
                        # 이미 존재하면 건너뜀
                        apt_id = existing_apt[0]
                        continue

                    # 아파트 정보 삽입
                    cursor.execute("""
                        INSERT INTO apartments
                        (list_id, apt_name, area, sido, sigungu, dong, sigungu_code, jibun_addr,
                         build_year, prev_max_price, prev_max_date, prev_max_floor, prev_max_dong,
                         last_max_price, max_price_date, max_price_floor, max_price_dong,
                         last_update, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        list_id,
                        apt.get('apt_name', ''),
                        apt.get('area', ''),
                        apt.get('sido', ''),
                        apt.get('sigungu', ''),
                        apt.get('dong', ''),
                        apt.get('sigungu_code', ''),
                        apt.get('jibun_addr', ''),
                        apt.get('build_year', ''),
                        apt.get('prev_max_price', 0),
                        apt.get('prev_max_date', ''),
                        apt.get('prev_max_floor', ''),
                        apt.get('prev_max_dong', ''),
                        apt.get('last_max_price', 0),
                        apt.get('max_price_date', ''),
                        apt.get('max_price_floor', ''),
                        apt.get('max_price_dong', ''),
                        apt.get('last_update', ''),
                        apt.get('created_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                    ))
                    apt_id = cursor.lastrowid

                    # 거래 데이터 삽입
                    if 'trade_data' in apt:
                        for trade in apt['trade_data']:
                            trade_date = trade.get('date')
                            if isinstance(trade_date, str):
                                try:
                                    trade_date = datetime.strptime(trade_date, '%Y-%m-%d').date()
                                except:
                                    trade_date = None
                            elif isinstance(trade_date, datetime):
                                trade_date = trade_date.date()

                            cursor.execute("""
                                INSERT INTO trade_data
                                (apt_id, price, trade_date, area, floor, dong)
                                VALUES (?, ?, ?, ?, ?, ?)
                            """, (
                                apt_id,
                                trade.get('price'),
                                trade_date,
                                trade.get('area'),
                                trade.get('floor'),
                                trade.get('dong')
                            ))

            self.db_conn.commit()
            logging.info(f"JSON -> DB 마이그레이션 완료 (리스트 수: {len(json_data.get('lists', {}))})")

            # 백업 후 JSON 파일 이름 변경
            backup_name = self.monitored_apts_file.replace('.json', '_migrated_backup.json')
            # 백업 파일이 이미 있으면 삭제
            if os.path.exists(backup_name):
                os.remove(backup_name)
                logging.info(f"기존 백업 파일 삭제: {backup_name}")
            os.rename(self.monitored_apts_file, backup_name)
            logging.info(f"기존 JSON 파일을 백업했습니다: {backup_name}")

        except Exception as e:
            logging.error(f"JSON -> DB 마이그레이션 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            self.db_conn.rollback()

    def load_monitored_apts(self):
        """DB에서 모니터링 리스트 로드"""
        try:
            # JSON 파일 자동 마이그레이션 제거
            # (사용자가 명시적으로 '백업 불러오기'를 사용해야 함)
            # if os.path.exists(self.monitored_apts_file):
            #     self.migrate_json_to_db()

            cursor = self.db_conn.cursor()

            # 활성 리스트 확인 (설정 파일에서 가져오거나 기본값)
            active_list = getattr(self, 'loaded_active_list', None) or "기본"

            # 모든 리스트 로드
            cursor.execute("SELECT id, name FROM monitoring_lists")
            lists_data = {"lists": {}, "active_list": active_list}

            # DB 스키마 버전 체크 (area 컬럼이 있는지 확인)
            cursor.execute("PRAGMA table_info(apartments)")
            columns = [col[1] for col in cursor.fetchall()]
            has_area = 'area' in columns

            for list_id, list_name in cursor.execute("SELECT id, name FROM monitoring_lists").fetchall():
                apt_list = []

                # 스키마에 따라 다른 쿼리 실행
                if has_area:
                    # 원래 프로그램의 스키마 (area, sido 등 사용)
                    cursor.execute("""
                        SELECT id, apt_name, area, sido, sigungu, dong, sigungu_code,
                               jibun_addr, build_year, prev_max_price, prev_max_date,
                               prev_max_floor, prev_max_dong, last_max_price, max_price_date,
                               max_price_floor, max_price_dong, last_update
                        FROM apartments
                        WHERE list_id = ?
                    """, (list_id,))

                    for apt_row in cursor.fetchall():
                        (apt_id, apt_name, area, sido, sigungu, dong, sigungu_code,
                         jibun_addr, build_year, prev_max_price, prev_max_date,
                         prev_max_floor, prev_max_dong, last_max_price, max_price_date,
                         max_price_floor, max_price_dong, last_update) = apt_row

                        # 거래 데이터 로드
                        cursor.execute("""
                            SELECT price, trade_date, area, floor, dong
                            FROM trade_data
                            WHERE apt_id = ?
                            ORDER BY trade_date DESC
                        """, (apt_id,))

                        trade_data = []
                        for trade_row in cursor.fetchall():
                            price, trade_date, t_area, floor, t_dong = trade_row
                            trade_data.append({
                                'price': price,
                                'date': datetime.strptime(trade_date, '%Y-%m-%d') if isinstance(trade_date, str) else trade_date,
                                'area': t_area,
                                'floor': floor,
                                'dong': t_dong
                            })

                        apt_list.append({
                            'id': apt_id,
                            'apt_name': apt_name,
                            'area': area,
                            'sido': sido,
                            'sigungu': sigungu,
                            'dong': dong,
                            'sigungu_code': sigungu_code,
                            'jibun_addr': jibun_addr,
                            'build_year': build_year,
                            'prev_max_price': prev_max_price,
                            'prev_max_date': prev_max_date,
                            'prev_max_floor': prev_max_floor,
                            'prev_max_dong': prev_max_dong,
                            'last_max_price': last_max_price,
                            'max_price_date': max_price_date,
                            'max_price_floor': max_price_floor,
                            'max_price_dong': max_price_dong,
                            'last_update': last_update,
                            'trade_data': trade_data
                        })

                else:
                    # 새 스키마 (region_code 사용)
                    cursor.execute("""
                        SELECT id, apt_name, region_code, sigungu_code,
                               last_max_price, last_max_price_date, last_checked
                        FROM apartments
                        WHERE list_id = ?
                    """, (list_id,))

                    for apt_row in cursor.fetchall():
                        apt_id, apt_name, region_code, sigungu_code, last_max_price, last_max_price_date, last_checked = apt_row

                        # 새 스키마에서는 region_code, sigungu_code만 있으므로
                        # 다른 필드들은 거래 데이터에서 가져와야 합니다
                        area = ""
                        sido = region_code or ""  # region_code를 sido로 사용
                        sigungu = sigungu_code or ""  # sigungu_code를 sigungu로 사용
                        dong = ""
                        jibun_addr = ""
                        build_year = ""
                        prev_max_price = 0
                        prev_max_date = ""
                        prev_max_floor = ""
                        prev_max_dong = ""
                        max_price_date = last_max_price_date
                        max_price_floor = ""
                        max_price_dong = ""
                        last_update = last_checked

                        # 거래 데이터 로드 (가격 기준으로 정렬하여 최고가와 2등 가격을 찾음)
                        cursor.execute("""
                            SELECT price, trade_date, area, floor, dong
                            FROM trade_data
                            WHERE apt_id = ?
                            ORDER BY price DESC, trade_date DESC
                        """, (apt_id,))

                        trade_data = []
                        trades_rows = cursor.fetchall()

                        # 가격별로 그룹화하여 최고가와 이전 최고가 찾기
                        price_groups = {}
                        for trade_row in trades_rows:
                            price, trade_date, t_area, floor, t_dong = trade_row
                            if price not in price_groups:
                                price_groups[price] = {
                                    'price': price,
                                    'date': trade_date,
                                    'area': t_area,
                                    'floor': floor,
                                    'dong': t_dong
                                }

                            trade_data.append({
                                'price': price,
                                'date': datetime.strptime(trade_date, '%Y-%m-%d') if isinstance(trade_date, str) else trade_date,
                                'area': t_area,
                                'floor': floor,
                                'dong': t_dong
                            })

                        # 거래 데이터에서 정보 추출
                        if trades_rows:
                            # 최고가 거래 정보 (첫 번째 행)
                            max_trade = trades_rows[0]
                            _, _, max_area, max_floor, max_dong = max_trade

                            # area 설정 (최고가 거래의 면적 사용)
                            if max_area:
                                area = f"{max_area:.2f}"  # 소수점 2자리로 포맷

                            # 최고가 거래의 동, 층 정보
                            if max_floor:
                                max_price_floor = str(max_floor)
                            if max_dong and max_dong != "-동":
                                max_price_dong = max_dong
                                dong = max_dong  # 동 정보도 업데이트

                            # 이전 최고가 찾기 (최고가와 다른 가격 중 가장 높은 가격)
                            unique_prices = sorted(price_groups.keys(), reverse=True)
                            if len(unique_prices) > 1:
                                # 두 번째로 높은 가격 찾기
                                prev_price = unique_prices[1]
                                prev_trade = price_groups[prev_price]
                                prev_max_price = int(prev_price)
                                prev_max_date = prev_trade['date']
                                if prev_trade['floor']:
                                    prev_max_floor = str(prev_trade['floor'])
                                if prev_trade['dong'] and prev_trade['dong'] != "-동":
                                    prev_max_dong = prev_trade['dong']

                        # 시군구 코드로 지역명 매핑 (간단한 매핑)
                        sigungu_mapping = {
                            '11680': '강남구', '11740': '강동구', '11305': '강북구',
                            '11500': '강서구', '11620': '관악구', '11215': '광진구',
                            '11530': '구로구', '11545': '금천구', '11350': '노원구',
                            '11320': '도봉구', '11230': '동대문구', '11590': '동작구',
                            '11440': '마포구', '11410': '서대문구', '11650': '서초구',
                            '11200': '성동구', '11290': '성북구', '11710': '송파구',
                            '11470': '양천구', '11560': '영등포구', '11170': '용산구',
                            '11380': '은평구', '11110': '종로구', '11140': '중구',
                            '11260': '중랑구'
                        }
                        if sigungu_code in sigungu_mapping:
                            sigungu = sigungu_mapping[sigungu_code]
                            jibun_addr = f"서울특별시 {sigungu}"

                        apt_list.append({
                            'id': apt_id,
                            'apt_name': apt_name,
                            'area': area,
                            'sido': sido,
                            'sigungu': sigungu,
                            'dong': dong,
                            'sigungu_code': sigungu_code,
                            'jibun_addr': jibun_addr,
                            'build_year': build_year,
                            'prev_max_price': prev_max_price,
                            'prev_max_date': prev_max_date,
                            'prev_max_floor': prev_max_floor,
                            'prev_max_dong': prev_max_dong,
                            'last_max_price': last_max_price,
                            'max_price_date': max_price_date,
                            'max_price_floor': max_price_floor,
                            'max_price_dong': max_price_dong,
                            'last_update': last_update,
                            'trade_data': trade_data
                        })

                lists_data["lists"][list_name] = apt_list

            # 기본 리스트가 없으면 생성
            if not lists_data["lists"]:
                cursor.execute("INSERT INTO monitoring_lists (name) VALUES (?)", ("기본",))
                self.db_conn.commit()
                lists_data["lists"]["기본"] = []
                lists_data["active_list"] = "기본"

            logging.info(f"DB에서 모니터링 리스트 로드 완료 (리스트 수: {len(lists_data['lists'])})")
            return lists_data

        except Exception as e:
            logging.error(f"DB에서 모니터링 목록 로드 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            # 기본 골격 반환
            return {"lists": {"기본": []}, "active_list": "기본"}

    def merge_duplicate_apts(self, apt_list):
        """동일한 아파트(이름+동+면적)를 병합하여 거래 데이터 통합 및 최고가 재계산

        분양권 거래와 준공후 거래가 따로 등록된 경우 하나로 합침
        - 단지명의 공백 차이도 같은 단지로 인식 (예: '청라언덕역 서한포레스트' = '청라언덕역서한포레스트')
        """
        if not apt_list:
            return apt_list

        import re

        def normalize_name(name):
            """단지명 정규화: 공백 제거, 소문자 변환"""
            if not name:
                return ''
            # 모든 공백 제거
            return re.sub(r'\s+', '', str(name).strip())

        def normalize_area(area):
            """면적 정규화: 숫자만 추출하여 정수로 변환"""
            if not area:
                return '0'
            # 숫자와 소수점만 추출
            area_str = re.sub(r'[^\d.]', '', str(area))
            try:
                # 정수로 반올림 (84.97 -> 85, 84.12 -> 84)
                return str(int(round(float(area_str)))) if area_str else '0'
            except:
                return '0'

        merged = {}
        for apt in apt_list:
            # 고유 키: 정규화된 아파트명 + 동 + 정규화된 면적
            key = (
                normalize_name(apt.get('apt_name', '')),
                apt.get('dong', '').strip(),
                normalize_area(apt.get('area', ''))
            )

            if key not in merged:
                merged[key] = apt.copy()
                # trade_data가 없으면 빈 리스트로 초기화
                if 'trade_data' not in merged[key]:
                    merged[key]['trade_data'] = []
            else:
                # 중복 발견 - 거래 데이터 병합
                existing = merged[key]
                new_trades = apt.get('trade_data', [])

                # 기존 거래 데이터의 키 세트 생성 (중복 방지)
                existing_keys = set()
                for t in existing.get('trade_data', []):
                    if isinstance(t.get('date'), datetime):
                        t_key = (t['date'].year, t['date'].month, t.get('day', 1),
                                 t.get('price', 0), t.get('floor', 0))
                    else:
                        t_key = (str(t.get('date', '')), t.get('price', 0), t.get('floor', 0))
                    existing_keys.add(t_key)

                # 새 거래 데이터 추가 (중복 제외)
                for t in new_trades:
                    if isinstance(t.get('date'), datetime):
                        t_key = (t['date'].year, t['date'].month, t.get('day', 1),
                                 t.get('price', 0), t.get('floor', 0))
                    else:
                        t_key = (str(t.get('date', '')), t.get('price', 0), t.get('floor', 0))

                    if t_key not in existing_keys:
                        existing['trade_data'].append(t)
                        existing_keys.add(t_key)

                # 최고가 재계산
                all_trades = existing.get('trade_data', [])
                if all_trades:
                    max_trade = max(all_trades, key=lambda x: x.get('price', 0))
                    max_price = max_trade.get('price', 0)

                    # 기존 최고가보다 높으면 업데이트
                    if max_price > existing.get('last_max_price', 0):
                        if isinstance(max_trade.get('date'), datetime):
                            max_date = max_trade['date'].strftime('%Y-%m-%d')
                        else:
                            max_date = str(max_trade.get('date', ''))

                        existing['prev_max_price'] = max_price
                        existing['prev_max_date'] = max_date
                        existing['prev_max_floor'] = max_trade.get('floor', '')
                        existing['prev_max_dong'] = max_trade.get('dong', '-')
                        existing['last_max_price'] = max_price
                        existing['max_price_date'] = max_date
                        existing['max_price_floor'] = max_trade.get('floor', '')
                        existing['max_price_dong'] = max_trade.get('dong', '-')

                logging.info(f"[중복 병합] {apt.get('apt_name')} {apt.get('area')}㎡ - 거래 데이터 통합됨")

        result = list(merged.values())
        if len(result) < len(apt_list):
            logging.info(f"[중복 병합] {len(apt_list)}개 -> {len(result)}개로 병합됨 ({len(apt_list) - len(result)}개 중복 제거)")

        return result

    def save_monitored_apts(self):
        """모니터링 리스트를 DB에 저장 (동적 스키마 감지)"""
        try:
            # DB 연결 확인 및 재연결
            if not hasattr(self, 'db_conn') or self.db_conn is None:
                logging.warning("DB 연결이 없습니다. 재연결을 시도합니다.")
                self.db_conn = init_database(self.db_path)

            cursor = self.db_conn.cursor()

            # DB 스키마 버전 체크 (area 컬럼이 있는지 확인)
            cursor.execute("PRAGMA table_info(apartments)")
            columns = [col[1] for col in cursor.fetchall()]
            has_area = 'area' in columns

            # 각 리스트 저장
            for list_name, apt_list in self.monitored_lists["lists"].items():
                # ★ 저장 전 중복 아파트 병합 (분양권+준공후 거래 통합)
                apt_list = self.merge_duplicate_apts(apt_list)
                self.monitored_lists["lists"][list_name] = apt_list

                # 리스트가 존재하는지 확인
                cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (list_name,))
                result = cursor.fetchone()

                if result:
                    list_id = result[0]
                    # 기존 아파트 삭제 (새로 추가하기 위해)
                    cursor.execute("DELETE FROM apartments WHERE list_id = ?", (list_id,))
                else:
                    # 새 리스트 생성
                    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    # created_at, updated_at 컬럼이 있는지 확인
                    cursor.execute("PRAGMA table_info(monitoring_lists)")
                    list_columns = [col[1] for col in cursor.fetchall()]

                    if 'created_at' in list_columns and 'updated_at' in list_columns:
                        cursor.execute(
                            "INSERT INTO monitoring_lists (name, created_at, updated_at) VALUES (?, ?, ?)",
                            (list_name, now, now)
                        )
                    else:
                        cursor.execute("INSERT INTO monitoring_lists (name) VALUES (?)", (list_name,))

                    list_id = cursor.lastrowid

                # 아파트 데이터 저장 (스키마에 따라 다른 쿼리 사용)
                for apt in apt_list:
                    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                    if has_area:
                        # 원래 스키마 (area, sido, sigungu, dong 등 포함)
                        cursor.execute("""
                            INSERT INTO apartments
                            (list_id, apt_name, area, sido, sigungu, dong, sigungu_code,
                             jibun_addr, build_year, prev_max_price, prev_max_date,
                             prev_max_floor, prev_max_dong, last_max_price, max_price_date,
                             max_price_floor, max_price_dong, last_update, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            list_id,
                            apt.get('apt_name', ''),
                            apt.get('area', ''),
                            apt.get('sido', ''),
                            apt.get('sigungu', ''),
                            apt.get('dong', ''),
                            apt.get('sigungu_code', ''),
                            apt.get('jibun_addr', ''),
                            apt.get('build_year', ''),
                            apt.get('prev_max_price', 0),
                            apt.get('prev_max_date', ''),
                            apt.get('prev_max_floor', ''),
                            apt.get('prev_max_dong', ''),
                            apt.get('last_max_price', 0),
                            apt.get('max_price_date', ''),
                            apt.get('max_price_floor', ''),
                            apt.get('max_price_dong', ''),
                            apt.get('last_update', now),
                            now
                        ))
                    else:
                        # 새 스키마 (region_code, sigungu_code만 사용)
                        cursor.execute("""
                            INSERT INTO apartments
                            (list_id, apt_name, region_code, sigungu_code,
                             last_max_price, last_max_price_date, last_checked)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        """, (
                            list_id,
                            apt.get('apt_name', ''),
                            apt.get('region_code', ''),
                            apt.get('sigungu_code', ''),
                            apt.get('last_max_price', 0),
                            apt.get('max_price_date', ''),
                            now
                        ))

                    apt_id = cursor.lastrowid

                    # 거래 데이터 저장
                    if 'trade_data' in apt and apt['trade_data']:
                        for trade in apt['trade_data']:
                            trade_date = trade.get('date')
                            if isinstance(trade_date, datetime):
                                trade_date = trade_date.strftime('%Y-%m-%d')
                            elif isinstance(trade_date, str):
                                # 이미 문자열이면 그대로 사용
                                pass
                            else:
                                # date가 없으면 건너뛰기
                                continue

                            # 실제 스키마: apt_id, price, trade_date, area, floor, dong, created_at
                            cursor.execute("""
                                INSERT INTO trade_data
                                (apt_id, price, trade_date, area, floor, dong, created_at)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            """, (
                                apt_id,
                                trade.get('price'),
                                trade_date,
                                trade.get('area'),
                                trade.get('floor'),
                                trade.get('dong'),
                                now
                            ))

            self.db_conn.commit()
            logging.info(f"DB에 모니터링 리스트 저장 완료 (리스트 수: {len(self.monitored_lists['lists'])})")

            # 자동 백업 체크
            self.auto_backup_check()

        except Exception as e:
            logging.error(f"DB 저장 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            self.db_conn.rollback()

    def backup_monitored_lists(self, *, parent_window=None):
        """모니터링 리스트를 백업"""
        try:
            from tkinter import filedialog

            # 백업 파일명 생성
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"monitored_lists_backup_{timestamp}.json"

            # 마지막 사용 디렉토리 가져오기
            initial_dir = getattr(self, 'last_backup_dir', self.backup_dir)

            # 저장 위치 선택
            if parent_window:
                filepath = filedialog.asksaveasfilename(
                    parent=parent_window,
                    initialdir=initial_dir,
                    initialfile=default_filename,
                    title="모니터링 리스트 백업",
                    defaultextension=".json",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
            else:
                filepath = filedialog.asksaveasfilename(
                    initialdir=initial_dir,
                    initialfile=default_filename,
                    title="모니터링 리스트 백업",
                    defaultextension=".json",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )

            if not filepath:
                return False

            # 사용한 디렉토리 저장
            self.last_backup_dir = os.path.dirname(filepath)
            self.save_last_backup_dir()
            
            # 현재 데이터 준비
            backup_data = {
                "version": "2.0",
                "backup_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "lists": {},
                "active_list": self.active_list.get(),
                "statistics": {
                    "total_lists": len(self.monitored_lists["lists"]),
                    "total_apts": sum(len(lst) for lst in self.monitored_lists["lists"].values())
                }
            }
            
            # 데이터 복사 및 직렬화
            for list_name, lst in self.monitored_lists["lists"].items():
                lst_copy = []
                for apt in lst:
                    apt_copy = apt.copy()
                    if 'trade_data' in apt_copy:
                        for trade in apt_copy['trade_data']:
                            if 'date' in trade and isinstance(trade['date'], datetime):
                                trade['date'] = trade['date'].strftime('%Y-%m-%d')
                    lst_copy.append(apt_copy)
                backup_data["lists"][list_name] = lst_copy
            
            # 파일 저장
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("백업 완료", f"모니터링 리스트가 백업되었습니다.\n경로: {filepath}")
            logging.info(f"모니터링 리스트 백업 완료: {filepath}")
            return True
            
        except Exception as e:
            messagebox.showerror("백업 실패", f"백업 중 오류가 발생했습니다:\n{str(e)}")
            logging.error(f"백업 중 오류: {str(e)}")
            return False

    def manual_save_db(self, *, parent_window=None):
        """DB를 수동으로 저장"""
        try:
            # 현재 데이터를 DB에 저장
            self.save_monitored_apts()
            self.save_notifications_history()

            # 커밋 확인
            if hasattr(self, 'db_conn') and self.db_conn:
                self.db_conn.commit()

            messagebox.showinfo("저장 완료",
                              f"DB가 저장되었습니다.\n경로: {self.db_path}",
                              parent=parent_window)
            logging.info(f"DB 수동 저장 완료: {self.db_path}")
            return True

        except Exception as e:
            messagebox.showerror("저장 실패",
                               f"DB 저장 중 오류가 발생했습니다:\n{str(e)}",
                               parent=parent_window)
            logging.error(f"DB 수동 저장 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def backup_db_file(self, *, parent_window=None):
        """DB 파일을 백업"""
        try:
            from tkinter import filedialog
            import shutil

            # 먼저 현재 데이터 저장
            self.save_monitored_apts()
            self.save_notifications_history()
            if hasattr(self, 'db_conn') and self.db_conn:
                self.db_conn.commit()

            # 백업 파일명 생성
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            db_filename = os.path.basename(self.db_path)
            db_name, db_ext = os.path.splitext(db_filename)
            default_filename = f"{db_name}_backup_{timestamp}{db_ext}"

            # 마지막 사용 디렉토리 가져오기
            initial_dir = getattr(self, 'last_backup_dir', self.backup_dir)

            # 저장 위치 선택
            if parent_window:
                backup_path = filedialog.asksaveasfilename(
                    parent=parent_window,
                    initialdir=initial_dir,
                    initialfile=default_filename,
                    title="DB 백업 저장",
                    defaultextension=db_ext,
                    filetypes=[("Database files", "*.db"), ("All files", "*.*")]
                )
            else:
                backup_path = filedialog.asksaveasfilename(
                    initialdir=initial_dir,
                    initialfile=default_filename,
                    title="DB 백업 저장",
                    defaultextension=db_ext,
                    filetypes=[("Database files", "*.db"), ("All files", "*.*")]
                )

            if not backup_path:
                return False

            # 사용한 디렉토리 저장
            self.last_backup_dir = os.path.dirname(backup_path)
            self.save_last_backup_dir()

            # DB 파일 복사
            shutil.copy2(self.db_path, backup_path)

            messagebox.showinfo("백업 완료",
                              f"DB 파일이 백업되었습니다.\n경로: {backup_path}",
                              parent=parent_window)
            logging.info(f"DB 파일 백업 완료: {backup_path}")
            return True

        except Exception as e:
            messagebox.showerror("백업 실패",
                               f"DB 백업 중 오류가 발생했습니다:\n{str(e)}",
                               parent=parent_window)
            logging.error(f"DB 백업 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def restore_db_file(self, *, parent_window=None):
        """DB 파일을 복원"""
        try:
            from tkinter import filedialog
            import shutil

            # 마지막 사용 디렉토리 가져오기
            initial_dir = getattr(self, 'last_backup_dir', self.backup_dir)

            # 백업 파일 선택
            if parent_window:
                backup_path = filedialog.askopenfilename(
                    parent=parent_window,
                    initialdir=initial_dir,
                    title="복원할 DB 파일 선택",
                    filetypes=[("Database files", "*.db"), ("All files", "*.*")]
                )
            else:
                backup_path = filedialog.askopenfilename(
                    initialdir=initial_dir,
                    title="복원할 DB 파일 선택",
                    filetypes=[("Database files", "*.db"), ("All files", "*.*")]
                )

            if not backup_path:
                return False

            # 사용한 디렉토리 저장
            self.last_backup_dir = os.path.dirname(backup_path)
            self.save_last_backup_dir()

            # 확인 메시지
            if not messagebox.askyesno("DB 복원 확인",
                                      f"현재 DB를 다음 파일로 복원하시겠습니까?\n\n"
                                      f"복원 파일: {backup_path}\n"
                                      f"현재 DB: {self.db_path}\n\n"
                                      f"현재 DB는 백업됩니다.",
                                      parent=parent_window):
                return False

            # 현재 DB 백업 (자동)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            db_filename = os.path.basename(self.db_path)
            db_name, db_ext = os.path.splitext(db_filename)
            auto_backup_path = os.path.join(self.backup_dir, f"{db_name}_auto_backup_{timestamp}{db_ext}")

            # 기존 DB 연결 종료
            if hasattr(self, 'db_conn') and self.db_conn:
                try:
                    self.db_conn.close()
                except:
                    pass
                self.db_conn = None

            # 현재 DB를 자동 백업
            if os.path.exists(self.db_path):
                shutil.copy2(self.db_path, auto_backup_path)
                logging.info(f"현재 DB 자동 백업 완료: {auto_backup_path}")

            # 백업 DB로 복원
            shutil.copy2(backup_path, self.db_path)

            # DB 다시 연결
            self.db_conn = init_database(self.db_path)

            # 모니터링 리스트 다시 로드
            loaded_data = self.load_monitored_apts()
            if loaded_data:
                self.monitored_lists = loaded_data
                # active_list 설정
                if 'active_list' in loaded_data:
                    self.active_list.set(loaded_data['active_list'])
            else:
                self.monitored_lists = {"lists": {}}

            # 신고가 히스토리 다시 로드
            self.notifications_history = self.load_notifications_history()

            # UI 업데이트
            self.refresh_list_combobox_values()
            self.update_apt_tree()

            messagebox.showinfo("복원 완료",
                              f"DB가 복원되었습니다.\n\n"
                              f"복원된 DB: {backup_path}\n"
                              f"이전 DB 백업: {auto_backup_path}",
                              parent=parent_window)
            logging.info(f"DB 복원 완료: {backup_path}")
            return True

        except Exception as e:
            messagebox.showerror("복원 실패",
                               f"DB 복원 중 오류가 발생했습니다:\n{str(e)}",
                               parent=parent_window)
            logging.error(f"DB 복원 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())

            # 오류 발생 시 DB 재연결 시도
            try:
                if not hasattr(self, 'db_conn') or self.db_conn is None:
                    self.db_conn = init_database(self.db_path)
            except:
                pass

            return False

    def save_last_backup_dir(self):
        """마지막 사용한 백업 디렉토리를 설정 파일에 저장"""
        try:
            settings_file = os.path.join(os.getcwd(), 'monitor_settings.json')
            settings_data = {}

            # 기존 설정 읽기
            if os.path.exists(settings_file):
                try:
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        settings_data = json.load(f)
                except:
                    pass

            # 마지막 백업 디렉토리 저장
            settings_data['last_backup_dir'] = self.last_backup_dir

            # 파일에 쓰기
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)

            logging.info(f"마지막 백업 디렉토리 저장: {self.last_backup_dir}")
        except Exception as e:
            logging.error(f"마지막 백업 디렉토리 저장 중 오류: {str(e)}")

    def restore_monitored_lists(self, *, parent_window=None):
        """백업된 모니터링 리스트 복원"""
        try:
            from tkinter import filedialog

            # 마지막 사용 디렉토리 가져오기
            initial_dir = getattr(self, 'last_backup_dir', self.backup_dir)

            # 복원할 파일 선택
            if parent_window:
                filepath = filedialog.askopenfilename(
                    parent=parent_window,
                    initialdir=initial_dir,
                    title="백업 파일 선택",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
            else:
                filepath = filedialog.askopenfilename(
                    initialdir=initial_dir,
                    title="백업 파일 선택",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )

            if not filepath:
                return False

            # 사용한 디렉토리 저장
            self.last_backup_dir = os.path.dirname(filepath)
            self.save_last_backup_dir()
            
            # 현재 데이터 백업 (복원 실패시 롤백용)
            current_backup = {
                "lists": self.monitored_lists["lists"].copy(),
                "active_list": self.active_list.get()
            }
            
            try:
                # 백업 파일 읽기
                with open(filepath, 'r', encoding='utf-8') as f:
                    backup_data = json.load(f)
                
                # 데이터 유효성 검증
                if not isinstance(backup_data, dict):
                    raise ValueError("잘못된 백업 파일 형식입니다.")
                
                # 버전 확인
                version = backup_data.get("version", "1.0")
                
                # 리스트 복원
                if "lists" in backup_data:
                    restored_lists = backup_data["lists"]
                elif isinstance(backup_data, list):
                    # 구형 백업 (단일 리스트)
                    restored_lists = {"기본": backup_data}
                else:
                    raise ValueError("백업 파일에서 리스트를 찾을 수 없습니다.")
                
                # 날짜 문자열을 datetime으로 변환
                for list_name, lst in restored_lists.items():
                    for apt in lst:
                        if 'trade_data' in apt:
                            for trade in apt['trade_data']:
                                if 'date' in trade and isinstance(trade['date'], str):
                                    try:
                                        trade['date'] = datetime.strptime(trade['date'], '%Y-%m-%d')
                                    except:
                                        pass
                
                # 복원 확인
                stats = backup_data.get("statistics", {})
                total_lists = stats.get("total_lists", len(restored_lists))
                total_apts = stats.get("total_apts", sum(len(lst) for lst in restored_lists.values()))
                backup_time = backup_data.get("backup_timestamp", "알 수 없음")
                
                msg = f"백업 정보:\n"
                msg += f"- 백업 시간: {backup_time}\n"
                msg += f"- 리스트 수: {total_lists}개\n"
                msg += f"- 총 아파트 수: {total_apts}개\n\n"
                msg += "현재 데이터를 이 백업으로 교체하시겠습니까?"
                
                if not messagebox.askyesno("복원 확인", msg):
                    return False
                
                # 데이터 복원
                self.monitored_lists["lists"] = restored_lists
                
                # 활성 리스트 설정
                if "active_list" in backup_data and backup_data["active_list"] in restored_lists:
                    self.active_list.set(backup_data["active_list"])
                else:
                    # 첫 번째 리스트 선택
                    if restored_lists:
                        self.active_list.set(list(restored_lists.keys())[0])
                
                # UI 업데이트
                self.refresh_list_combobox_values()
                self.update_apt_tree()
                
                # 복원된 데이터 저장
                self.save_monitored_apts()
                
                messagebox.showinfo("복원 완료", 
                                  f"백업이 성공적으로 복원되었습니다.\n"
                                  f"복원된 리스트: {total_lists}개\n"
                                  f"복원된 아파트: {total_apts}개")
                logging.info(f"백업 복원 완료: {filepath}")
                return True
                
            except Exception as e:
                # 복원 실패시 롤백
                self.monitored_lists["lists"] = current_backup["lists"]
                self.active_list.set(current_backup["active_list"])
                raise e
                
        except Exception as e:
            messagebox.showerror("복원 실패", f"백업 복원 중 오류가 발생했습니다:\n{str(e)}")
            logging.error(f"백업 복원 중 오류: {str(e)}")
            return False
    
    def open_backup_folder(self):
        """백업 폴더 열기"""
        try:
            os.makedirs(self.backup_dir, exist_ok=True)
            if os.name == 'nt':  # Windows
                os.startfile(self.backup_dir)
            elif os.name == 'posix':  # macOS/Linux
                os.system(f'open "{self.backup_dir}"')
            else:
                messagebox.showinfo("경로", f"백업 폴더 경로:\n{self.backup_dir}")
        except Exception as e:
            messagebox.showerror("오류", f"백업 폴더를 열 수 없습니다:\n{str(e)}")
    
    def auto_backup_check(self):
        """자동 백업 체크 (save_monitored_apts 호출시) - DB 백업"""
        try:
            if self.last_auto_backup is None:
                self.last_auto_backup = datetime.now()
                return

            hours_passed = (datetime.now() - self.last_auto_backup).total_seconds() / 3600
            if hours_passed >= self.auto_backup_interval:
                # DB 파일 백업
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                auto_backup_file = os.path.join(self.backup_dir, f"auto_backup_{timestamp}.db")

                # SQLite DB 백업 (VACUUM INTO 사용)
                import shutil
                shutil.copy2(self.db_path, auto_backup_file)

                self.last_auto_backup = datetime.now()
                logging.info(f"자동 DB 백업 완료: {auto_backup_file}")

                # 오래된 자동 백업 삭제 (7일 이상)
                self.cleanup_old_auto_backups()
                
        except Exception as e:
            logging.error(f"자동 백업 중 오류: {str(e)}")
    
    def cleanup_old_auto_backups(self):
        """7일 이상된 자동 백업 파일 삭제"""
        try:
            cutoff_date = datetime.now() - timedelta(days=7)
            for filename in os.listdir(self.backup_dir):
                if filename.startswith("auto_backup_"):
                    filepath = os.path.join(self.backup_dir, filename)
                    file_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                    if file_time < cutoff_date:
                        os.remove(filepath)
                        logging.info(f"오래된 자동 백업 삭제: {filename}")
        except Exception as e:
            logging.error(f"자동 백업 정리 중 오류: {str(e)}")

    
    
    def setup_gui(self):
        """GUI 구성 (다크 테마)"""
        self.root.configure(bg=self.palette['bg'])
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure('TFrame', background=self.palette['surface'])
        style.configure('TLabelframe', background=self.palette['surface'], borderwidth=0, relief='flat')
        style.configure('TLabelframe.Label', background=self.palette['surface'], foreground=self.palette['text_primary'], font=self.font_large)
        style.configure('TLabel', background=self.palette['surface'], foreground=self.palette['text_primary'])
        style.configure('Title.TLabel', background=self.palette['surface'], foreground=self.palette['accent'], font=self.font_title)
        style.configure('Accent.TButton', background=self.palette['accent'], foreground='white', relief='flat', font=self.font_button)
        style.map('Accent.TButton', background=[('active', '#1A8CFF')])
        style.configure('TEntry', fieldbackground=self.palette['surface'], background=self.palette['surface'], foreground=self.palette['text_primary'])
        style.configure('TCombobox', fieldbackground=self.palette['surface'], background=self.palette['surface'], foreground=self.palette['text_primary'])
        style.map('TCombobox', fieldbackground=[('readonly', self.palette['surface'])], foreground=[('disabled', self.palette['text_secondary'])])
        style.configure('TCheckbutton', background=self.palette['bg'], foreground=self.palette['text_primary'])
        style.configure('Treeview', background=self.palette['surface'], fieldbackground=self.palette['surface'], foreground=self.palette['text_primary'], rowheight=26, font=self.font_normal)
        style.configure('Treeview.Heading', background=self.palette['bg'], foreground=self.palette['accent'], relief='flat', font=self.font_normal)
        style.map('Treeview.Heading', background=[('active', self.palette['bg']), ('pressed', self.palette['bg']), ('!active', self.palette['bg'])])
        main_frame = ttk.Frame(self.root, padding="10", style='TFrame')
        main_frame.pack(fill="both", expand=True)
        top_frame = ttk.Frame(main_frame, style='TFrame')
        top_frame.pack(fill="x")
        title_label = ttk.Label(top_frame, text="부태리의 실거래가 모니터", style='Title.TLabel')
        title_label.pack(side="left", pady=(0, 10))
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(side="right", pady=(0, 10))
        ttk.Button(button_frame, text="⚙️ 설정", style='Accent.TButton',
                  command=self.show_settings_dialog).pack(side="right", padx=(0, 5))
        ttk.Button(button_frame, text="💾 캐시", style='Accent.TButton',
                  command=self.show_cache_statistics).pack(side="right", padx=(0, 5))
        ttk.Button(button_frame, text="📥 이전 신고가 다운로드", style='Accent.TButton',
                  command=self.download_previous_high_prices).pack(side="right", padx=(0, 5))
        ttk.Button(button_frame, text="📊 이전 신고가", style='Accent.TButton',
                  command=self.show_previous_notifications).pack(side="right", padx=(0, 5))
        region_frame = ttk.LabelFrame(main_frame, text="지역 검색", padding=10)
        region_frame.pack(fill="x", pady=5)
        region_container = ttk.Frame(region_frame)
        region_container.pack(fill="x")
        ttk.Label(region_container, text="시/도:", style='TLabel').grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.sido_combobox = ttk.Combobox(region_container, values=self.sido_list, state="readonly", width=15, style='TCombobox')
        self.sido_combobox.set("시/도 선택")
        self.sido_combobox.grid(row=0, column=1, padx=(0, 10))
        self.sido_combobox.bind('<<ComboboxSelected>>', self.on_sido_selected)
        ttk.Label(region_container, text="시/군/구:", style='TLabel').grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.sigungu_combobox = ttk.Combobox(region_container, state="readonly", width=15, style='TCombobox')
        self.sigungu_combobox.set("시/군/구 선택")
        self.sigungu_combobox.grid(row=0, column=3, padx=(0, 10))
        self.sigungu_combobox.bind('<<ComboboxSelected>>', self.on_sigungu_selected)
        ttk.Label(region_container, text="읍/면/동:", style='TLabel').grid(row=0, column=4, sticky="w", padx=(0, 5))
        self.dong_combobox = ttk.Combobox(region_container, state="readonly", width=15, style='TCombobox')
        self.dong_combobox.set("읍/면/동 선택")
        self.dong_combobox.grid(row=0, column=5, padx=(0, 10))
        self.dong_combobox.bind('<<ComboboxSelected>>', self.on_dong_selected)
        search_button = ttk.Button(region_container, text="아파트 목록 조회", style='Accent.TButton', command=self.show_apt_list)
        search_button.grid(row=0, column=6, padx=(10, 0))
        
        # ▼▼ 추가: 체크박스 (선택시 해당 지역 전체 단지 모니터링)
        bulk_chk = ttk.Checkbutton(
            region_container,
            text="선택시 해당 지역 전체 단지 모니터링",
            variable=self.bulk_monitor_enabled
        )
        bulk_chk.grid(row=0, column=7, padx=(12, 0), sticky="w")
        
        for i in range(8):  # ← 7 → 8 로 변경 (컬럼 개수 증가)
            region_container.grid_columnconfigure(i, weight=1 if i == 7 else 0)

        # 마지막 검색 지역 표시 라벨 추가
        self.last_search_label = ttk.Label(
            region_frame,
            text="",
            style='TLabel',
            foreground=self.palette['text_secondary'],
            font=('맑은 고딕', 9)
        )
        self.last_search_label.pack(anchor="w", padx=5, pady=(5, 0))

        # === 리스트 선택/관리 바 ===
        list_bar = ttk.LabelFrame(main_frame, text="모니터링 리스트", padding=10)
        list_bar.pack(fill="x", pady=5)
        
        bar_left = ttk.Frame(list_bar)
        bar_left.pack(side="left")
        ttk.Label(bar_left, text="현재 리스트:", style='TLabel').pack(side="left", padx=(0, 6))
        
        # 콤보박스 (리스트 선택)
        self.list_combobox = ttk.Combobox(
            bar_left,
            values=sorted(self.monitored_lists["lists"].keys()),
            textvariable=self.active_list,
            state="readonly",
            width=20,
            style='TCombobox'
        )
        self.list_combobox.pack(side="left")
        self.list_combobox.bind('<<ComboboxSelected>>', self.on_active_list_changed)
        
        bar_right = ttk.Frame(list_bar)
        bar_right.pack(side="right")
        
        ttk.Button(bar_right, text="➕ 추가", style='Accent.TButton', command=self.create_new_list).pack(side="left", padx=4)
        ttk.Button(bar_right, text="✏️ 이름변경", style='Accent.TButton', command=self.rename_current_list).pack(side="left", padx=4)
        ttk.Button(bar_right, text="🗑 삭제", style='Accent.TButton', command=self.delete_current_list).pack(side="left", padx=4)
        ttk.Button(bar_right, text="↪ 선택 이동", style='Accent.TButton', command=self.move_selected_to_list).pack(side="left", padx=4)



            
        monitored_apt_frame = ttk.LabelFrame(main_frame, text="모니터링 중인 아파트", padding=10)
        monitored_apt_frame.pack(fill="both", expand=True, pady=5)

        # --- 검색 바 ---
        search_bar = ttk.Frame(monitored_apt_frame)
        search_bar.pack(fill="x", pady=(0, 6))
        
        self.search_var = tk.StringVar(value="")
        self.search_count_var = tk.StringVar(value="")
        
        def _on_search_change(*_):
            self.update_apt_tree()
        
        def _clear_search():
            self.search_var.set("")
            self.update_apt_tree()
        
        ttk.Label(search_bar, text="검색:").pack(side="left", padx=(0,6))
        search_entry = ttk.Entry(search_bar, textvariable=self.search_var, width=36)
        search_entry.pack(side="left")
        self.search_var.trace_add('write', _on_search_change)
        
        ttk.Button(search_bar, text="지우기", style='Accent.TButton', command=_clear_search)\
           .pack(side="left", padx=(6,6))
        
        # 결과 개수 표시
        ttk.Label(search_bar, textvariable=self.search_count_var)\
           .pack(side="right")
        
        # 단축키: Ctrl+F 로 검색창 포커스
        def _focus_search(event=None):
            try:
                search_entry.focus_set()
                search_entry.select_range(0, 'end')
                return "break"
            except:
                return
        
        self.root.bind_all("<Control-f>", _focus_search)





        
        list_frame = ttk.Frame(monitored_apt_frame)
        list_frame.pack(fill="both", expand=True)
        columns = ("apt_name", "build_year", "area", "location", "dong", "floor",
                   "prev_max_price", "prev_date", "last_max_price", "avg_py_price", "last_date", "last_update")
        
        self.apt_tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            height=10,            # (네가 앞서 10으로 바꿨다면 그대로 두기)
            style='Treeview'
        )
        
        self.apt_tree.heading("apt_name", text="단지명", anchor='center')
        self.apt_tree.heading("build_year", text="연식", anchor='center')
        self.apt_tree.heading("area", text="전용면적", anchor='center')
        self.apt_tree.heading("location", text="주소", anchor='center')
        self.apt_tree.heading("dong", text="동", anchor='center')
        self.apt_tree.heading("floor", text="층", anchor='center')
        self.apt_tree.heading("prev_max_price", text="이전 신고가", anchor='center')
        self.apt_tree.heading("prev_date", text="거래 날짜", anchor='center')
        self.apt_tree.heading("last_max_price", text="최근 신고가", anchor='center')
        self.apt_tree.heading("avg_py_price", text="평균평단가", anchor='center')   # ★ 추가
        self.apt_tree.heading("last_date", text="날짜", anchor='center')
        self.apt_tree.heading("last_update", text="갱신 날짜", anchor='center')
        
        self.apt_tree.column("apt_name", width=120, anchor='center')
        self.apt_tree.column("build_year", width=60, anchor='center')
        self.apt_tree.column("area", width=60, anchor='center')
        self.apt_tree.column("location", width=100, anchor='center')
        self.apt_tree.column("dong", width=50, anchor='center')
        self.apt_tree.column("floor", width=40, anchor='center')
        self.apt_tree.column("prev_max_price", width=100, anchor='center')
        self.apt_tree.column("prev_date", width=100, anchor='center')
        self.apt_tree.column("last_max_price", width=100, anchor='center')
        self.apt_tree.column("avg_py_price", width=90, anchor='center')            # ★ 추가
        self.apt_tree.column("last_date", width=100, anchor='center')
        self.apt_tree.column("last_update", width=140, anchor='center')

        # 헤더 클릭 정렬 바인딩 (모든 열)
        for col in self.apt_tree["columns"]:
            # lambda의 late-binding 방지: 기본값 c=col 캡처
            self.apt_tree.heading(col, command=lambda c=col: self.treeview_sort_column(
                c,
                # 같은 열을 또 누르면 방향 토글, 다른 열을 누르면 내림차순(또는 원하는 기본값) 시작
                (False if self.sort_column != c else not self.sort_reverse)
            ))

        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.apt_tree.yview)
        self.apt_tree.configure(yscrollcommand=scrollbar.set)
        self.apt_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        btns = ttk.Frame(monitored_apt_frame)
        btns.pack(fill="x", pady=5)
        ttk.Button(btns, text="선택 삭제", style='Accent.TButton', command=self.delete_selected_apt).pack(side="left", padx=5)
        ttk.Button(btns, text="모두 삭제", style='Accent.TButton', command=self.clear_all_apts).pack(side="left", padx=5)

        ttk.Button(btns, text="데이터 갱신", style='Accent.TButton', command=self.update_all_data).pack(side="right", padx=5)
        ttk.Button(btns, text="🔍 신고가 다시 찾기", style='Accent.TButton', command=self.recheck_max_prices).pack(side="right", padx=5)
        ttk.Button(btns, text="📈 가격대별 분위", style='Accent.TButton', command=self.export_price_distribution_html).pack(side="right", padx=5)
        ttk.Button(btns, text="📊 거래량 순위", style='Accent.TButton', command=self.export_trade_volume_ranking_html).pack(side="right", padx=5)
        ttk.Button(btns, text="🏆 59㎡ 순위", style='Accent.TButton', command=lambda: self.export_area_ranking_html('59')).pack(side="right", padx=5)
        ttk.Button(btns, text="🏆 84㎡ 순위", style='Accent.TButton', command=lambda: self.export_area_ranking_html('84')).pack(side="right", padx=5)
        
        auto_update_frame = ttk.LabelFrame(main_frame, text="자동 업데이트 설정", padding=10)
        auto_update_frame.pack(fill="x", pady=5)
        ttk.Checkbutton(auto_update_frame, text="자동 업데이트 사용", variable=self.auto_update_enabled, command=self.toggle_auto_update).pack(side="left", padx=5)
        ttk.Label(auto_update_frame, text="업데이트 시간:").pack(side="left", padx=5)
        time_entry = ttk.Entry(auto_update_frame, textvariable=self.update_time, width=8)
        time_entry.pack(side="left", padx=5)
        ttk.Label(auto_update_frame, text="(24시간 형식, 예: 09:10)").pack(side="left", padx=5)
        ttk.Button(auto_update_frame, text="적용", style='Accent.TButton', command=self.apply_update_time).pack(side="left", padx=5)
        self.status_var = tk.StringVar(value="준비 완료")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w", background=self.palette['surface'], foreground=self.palette['text_secondary'])
        status_bar.pack(side="bottom", fill="x")
        self.update_apt_tree()
        self.root.after_idle(self.adjust_column_widths)

        # 트리뷰에 더블클릭 이벤트 바인딩 추가
        self.apt_tree.bind('<Double-Button-1>', self.edit_apt_data)
        
        # 우클릭 메뉴 추가
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="데이터 수정", command=self.edit_selected_apt)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="삭제", command=self.delete_selected_apt)
        
        self.apt_tree.bind('<Button-3>', self.show_context_menu)


    

    def refresh_list_combobox_values(self):
        self.list_combobox['values'] = sorted(self.monitored_lists["lists"].keys())
        # active_list 값이 사라진 경우 보정
        if self.active_list.get() not in self.monitored_lists["lists"]:
            # 리스트가 있으면 하나 선택, 없으면 빈 문자열
            if self.monitored_lists["lists"]:
                any_name = sorted(self.monitored_lists["lists"].keys())[0]
                self.active_list.set(any_name)
            else:
                self.active_list.set("")

    def show_context_menu(self, event):
        """우클릭 메뉴 표시"""
        try:
            # 클릭한 위치의 아이템 선택
            item = self.apt_tree.identify('item', event.x, event.y)
            if item:
                self.apt_tree.selection_set(item)
                self.context_menu.post(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def edit_selected_apt(self):
        """선택된 아파트 데이터 수정"""
        selection = self.apt_tree.selection()
        if not selection:
            messagebox.showinfo("알림", "수정할 아파트를 선택해주세요.")
            return
        self.edit_apt_data(None)
    
    def edit_apt_data(self, event):
        """아파트 데이터 수정 다이얼로그"""
        selection = self.apt_tree.selection()
        if not selection:
            return
        
        # 선택된 아이템 정보 가져오기
        item_id = selection[0]
        item_values = self.apt_tree.item(item_id, "values")
        
        apt_name = item_values[0]
        area_str = item_values[2]
        
        # 면적 값 추출
        import re
        area_value = re.search(r'(\d+(?:\.\d+)?)', area_str)
        if area_value:
            area = area_value.group(1)
        else:
            area = area_str.replace('㎡', '').strip()
        
        # monitored_apts에서 해당 아파트 찾기
        apt_data = None
        apt_index = None
        for idx, apt in enumerate(self.monitored_apts):
            if apt.get('apt_name') == apt_name:
                apt_area = str(apt.get('area', '')).replace('㎡', '').strip()
                try:
                    if float(apt_area) == float(area):
                        apt_data = apt
                        apt_index = idx
                        break
                except:
                    if apt_area == area:
                        apt_data = apt
                        apt_index = idx
                        break
        
        if not apt_data:
            messagebox.showerror("오류", "아파트 정보를 찾을 수 없습니다.")
            return
        
        # 편집 다이얼로그 창
        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"데이터 수정 - {apt_name}")
        edit_window.geometry("500x600")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        # 창을 화면 중앙에 배치
        edit_window.update_idletasks()
        width = edit_window.winfo_width()
        height = edit_window.winfo_height()
        x = (edit_window.winfo_screenwidth() // 2) - (width // 2)
        y = (edit_window.winfo_screenheight() // 2) - (height // 2)
        edit_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # 스타일 설정
        frame = ttk.Frame(edit_window, padding="10")
        frame.pack(fill="both", expand=True)
        
        # 제목
        title_label = ttk.Label(frame, text=f"{apt_name} ({area}㎡)", 
                               font=self.font_large)
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # 입력 필드들
        ttk.Label(frame, text="현재 최고가 정보", font=self.font_large).grid(
            row=1, column=0, columnspan=2, pady=(10, 5), sticky="w")
        
        ttk.Label(frame, text="최고가 (만원):").grid(row=2, column=0, sticky="w", pady=5)
        current_price_var = tk.StringVar(value=str(apt_data.get('last_max_price', 0)))
        current_price_entry = ttk.Entry(frame, textvariable=current_price_var, width=20)
        current_price_entry.grid(row=2, column=1, sticky="w", pady=5)
        
        ttk.Label(frame, text="거래 날짜:").grid(row=3, column=0, sticky="w", pady=5)
        current_date_var = tk.StringVar(value=apt_data.get('max_price_date', ''))
        current_date_entry = ttk.Entry(frame, textvariable=current_date_var, width=20)
        current_date_entry.grid(row=3, column=1, sticky="w", pady=5)
        ttk.Label(frame, text="(형식: YYYY-MM-DD)", font=('Malgun Gothic', 8)).grid(
            row=3, column=2, sticky="w", pady=5)
        
        ttk.Label(frame, text="층:").grid(row=4, column=0, sticky="w", pady=5)
        current_floor_var = tk.StringVar(value=str(apt_data.get('max_price_floor', '')))
        current_floor_entry = ttk.Entry(frame, textvariable=current_floor_var, width=20)
        current_floor_entry.grid(row=4, column=1, sticky="w", pady=5)
        
        ttk.Label(frame, text="동:").grid(row=5, column=0, sticky="w", pady=5)
        current_dong_var = tk.StringVar(value=apt_data.get('max_price_dong', '-'))
        current_dong_entry = ttk.Entry(frame, textvariable=current_dong_var, width=20)
        current_dong_entry.grid(row=5, column=1, sticky="w", pady=5)
        
        # 이전 최고가 정보
        ttk.Label(frame, text="이전 최고가 정보", font=self.font_large).grid(
            row=6, column=0, columnspan=2, pady=(20, 5), sticky="w")
        
        ttk.Label(frame, text="이전 최고가 (만원):").grid(row=7, column=0, sticky="w", pady=5)
        prev_price_var = tk.StringVar(value=str(apt_data.get('prev_max_price', 0)))
        prev_price_entry = ttk.Entry(frame, textvariable=prev_price_var, width=20)
        prev_price_entry.grid(row=7, column=1, sticky="w", pady=5)
        
        ttk.Label(frame, text="이전 거래 날짜:").grid(row=8, column=0, sticky="w", pady=5)
        prev_date_var = tk.StringVar(value=apt_data.get('prev_max_date', ''))
        prev_date_entry = ttk.Entry(frame, textvariable=prev_date_var, width=20)
        prev_date_entry.grid(row=8, column=1, sticky="w", pady=5)
        
        ttk.Label(frame, text="이전 층:").grid(row=9, column=0, sticky="w", pady=5)
        prev_floor_var = tk.StringVar(value=str(apt_data.get('prev_max_floor', '')))
        prev_floor_entry = ttk.Entry(frame, textvariable=prev_floor_var, width=20)
        prev_floor_entry.grid(row=9, column=1, sticky="w", pady=5)
        
        ttk.Label(frame, text="이전 동:").grid(row=10, column=0, sticky="w", pady=5)
        prev_dong_var = tk.StringVar(value=apt_data.get('prev_max_dong', ''))  # 여기 수정됨
        prev_dong_entry = ttk.Entry(frame, textvariable=prev_dong_var, width=20)
        prev_dong_entry.grid(row=10, column=1, sticky="w", pady=5)
        
        # 버튼 프레임
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=11, column=0, columnspan=3, pady=(20, 0))
        
        def save_changes():
            """변경사항 저장"""
            try:
                # 현재 최고가 정보 업데이트
                new_price = int(current_price_var.get())
                apt_data['last_max_price'] = new_price
                apt_data['max_price_date'] = current_date_var.get()
                apt_data['max_price_floor'] = current_floor_var.get()
                apt_data['max_price_dong'] = current_dong_var.get()
                
                # 이전 최고가 정보 업데이트
                prev_price = int(prev_price_var.get()) if prev_price_var.get() else 0
                apt_data['prev_max_price'] = prev_price
                apt_data['prev_max_date'] = prev_date_var.get()
                apt_data['prev_max_floor'] = prev_floor_var.get()
                apt_data['prev_max_dong'] = prev_dong_var.get()
                
                # 업데이트 시간 갱신
                apt_data['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                
                # 리스트 업데이트
                self.monitored_apts[apt_index] = apt_data
                
                # 저장 및 화면 갱신
                self.save_monitored_apts()
                self.update_apt_tree()
                
                messagebox.showinfo("완료", "데이터가 수정되었습니다.")
                edit_window.destroy()
                
            except ValueError as e:
                messagebox.showerror("오류", "가격은 숫자로 입력해주세요.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 오류: {str(e)}")
        
        def swap_prices():
            """현재/이전 최고가 서로 바꾸기"""
            # 값 임시 저장
            temp_price = current_price_var.get()
            temp_date = current_date_var.get()
            temp_floor = current_floor_var.get()
            temp_dong = current_dong_var.get()
            
            # 현재 <- 이전
            current_price_var.set(prev_price_var.get())
            current_date_var.set(prev_date_var.get())
            current_floor_var.set(prev_floor_var.get())
            current_dong_var.set(prev_dong_var.get())
            
            # 이전 <- 임시
            prev_price_var.set(temp_price)
            prev_date_var.set(temp_date)
            prev_floor_var.set(temp_floor)
            prev_dong_var.set(temp_dong)
        
        ttk.Button(button_frame, text="저장", command=save_changes, 
                  style='Accent.TButton').pack(side="left", padx=5)
        ttk.Button(button_frame, text="↔️ 교환", command=swap_prices,
                  style='Accent.TButton').pack(side="left", padx=5)
        ttk.Button(button_frame, text="취소", command=edit_window.destroy).pack(side="left", padx=5)
        
        # 첫 번째 입력 필드에 포커스
        current_price_entry.focus_set()
        current_price_entry.select_range(0, 'end')

    
    def on_active_list_changed(self, event=None):
        # 리스트 변경 시 트리 갱신 + 저장
        self.save_monitored_apts()
        self.save_last_active_list()  # 마지막 활성 리스트 저장
        self.update_apt_tree()
        self.status_var.set(f"리스트 전환: {self.active_list.get()}")
    
    def create_new_list(self):
        name = simpledialog.askstring("새 리스트", "리스트 이름을 입력하세요:", parent=self.root)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.monitored_lists["lists"]:
            messagebox.showinfo("알림", "이미 존재하는 리스트 이름입니다.")
            return
        self.monitored_lists["lists"][name] = []
        self.active_list.set(name)
        self.refresh_list_combobox_values()
        self.save_monitored_apts()
        self.save_last_active_list()  # 마지막 활성 리스트 저장
        self.update_apt_tree()
        self.status_var.set(f"리스트 생성: {name}")
    
    def rename_current_list(self):
        cur = self.active_list.get()
        new_name = simpledialog.askstring("리스트 이름변경", f"'{cur}'의 새 이름:", parent=self.root, initialvalue=cur)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name or new_name == cur:
            return
        if new_name in self.monitored_lists["lists"]:
            messagebox.showinfo("알림", "이미 존재하는 리스트 이름입니다.")
            return
        # 키 변경
        self.monitored_lists["lists"][new_name] = self.monitored_lists["lists"].pop(cur)
        self.active_list.set(new_name)
        self.refresh_list_combobox_values()
        self.save_monitored_apts()
        self.update_apt_tree()
        self.status_var.set(f"리스트 이름 변경: {cur} → {new_name}")
    
    def delete_current_list(self):
        cur = self.active_list.get()
        if len(self.monitored_lists["lists"]) <= 1:
            messagebox.showinfo("알림", "마지막 리스트는 삭제할 수 없습니다.")
            return
        if not messagebox.askyesno("확인", f"'{cur}' 리스트를 삭제하시겠습니까?\n(해당 리스트 내 아파트도 함께 삭제됩니다)"):
            return
        # 다른 리스트 하나로 전환
        names = sorted(self.monitored_lists["lists"].keys())
        fallback = next((n for n in names if n != cur), None)
        self.monitored_lists["lists"].pop(cur, None)
        self.active_list.set(fallback or "기본")
        self.refresh_list_combobox_values()
        self.save_monitored_apts()
        self.save_last_active_list()  # 마지막 활성 리스트 저장
        self.update_apt_tree()
        self.status_var.set(f"리스트 삭제: {cur}")
    
    def move_selected_to_list(self):
        selection = self.apt_tree.selection()
        if not selection:
            messagebox.showinfo("알림", "이동할 아파트를 선택해주세요.")
            return
        cur = self.active_list.get()
        names = [n for n in sorted(self.monitored_lists["lists"].keys()) if n != cur]
        if not names:
            messagebox.showinfo("알림", "이동할 대상 리스트가 없습니다. 먼저 리스트를 추가해 주세요.")
            return
        # 대상 리스트 선택
        target = simpledialog.askstring("선택 이동", f"이동할 리스트 이름을 입력하세요.\n가능: {', '.join(names)}", parent=self.root)
        if not target or target not in self.monitored_lists["lists"] or target == cur:
            return
    
        # 선택 항목 -> dict로 찾아 옮기기
        moved = 0
        cur_list = self.monitored_lists["lists"][cur]
        tgt_list = self.monitored_lists["lists"][target]
    
        # 현재 트리에서 선택된 값으로 매칭
        for item_id in selection:
            vals = self.apt_tree.item(item_id, "values")
            apt_name = vals[0]
            area_str = vals[2]
            import re
            m = re.search(r'(\d+(?:\.\d+)?)', area_str)
            area_val = m.group(1) if m else area_str.replace('㎡','').strip()
            # 현재 리스트에서 동일 항목 찾기
            idx_to_move = None
            for idx, apt in enumerate(cur_list):
                if apt.get('apt_name') == apt_name:
                    try:
                        if float(str(apt.get('area','')).replace('㎡','').strip()) == float(area_val):
                            idx_to_move = idx
                            break
                    except:
                        if str(apt.get('area','')).replace('㎡','').strip() == str(area_val):
                            idx_to_move = idx
                            break
            if idx_to_move is not None:
                tgt_list.append(cur_list.pop(idx_to_move))
                moved += 1
    
        if moved > 0:
            self.save_monitored_apts()
            self.update_apt_tree()
            self.status_var.set(f"{moved}개 항목을 '{cur}' → '{target}'로 이동했습니다.")
        else:
            self.status_var.set("이동할 항목을 찾지 못했습니다.")




    
    def show_cache_statistics(self):
        """캐시 통계 표시"""
        stats_window = tk.Toplevel(self.root)
        stats_window.title("API 캐시 통계")
        stats_window.geometry("400x300")
        stats_window.transient(self.root)
        
        frame = ttk.Frame(stats_window, padding="10")
        frame.pack(fill="both", expand=True)
        
        ttk.Label(frame, text="API 캐시 통계", font=self.font_title).pack(pady=(0, 10))
        
        hit_rate = 0
        if (self.api_call_count + self.cache_hit_count) > 0:
            hit_rate = (self.cache_hit_count / (self.api_call_count + self.cache_hit_count)) * 100
        
        stats_text = f"""
        캐시 크기: {len(self.api_cache)}개
        총 API 호출: {self.api_call_count}회
        캐시 히트: {self.cache_hit_count}회
        캐시 히트율: {hit_rate:.1f}%
        캐시 TTL: {self.cache_ttl.total_seconds() / 3600:.1f}시간
        """
        
        ttk.Label(frame, text=stats_text, font=self.font_normal).pack(pady=10)
        
        def clear_cache():
            if messagebox.askyesno("확인", "캐시를 모두 삭제하시겠습니까?"):
                self.api_cache.clear()
                self.cache_hit_count = 0
                self.api_call_count = 0
                messagebox.showinfo("알림", "캐시가 초기화되었습니다.")
                stats_window.destroy()
        
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(button_frame, text="캐시 초기화", command=clear_cache).pack(side="left")
        ttk.Button(button_frame, text="닫기", command=stats_window.destroy).pack(side="right")

    def save_previous_high_to_db(self, apt_list):
        """이전 신고가 다운로드 시 메모리에 추가 (r7 방식)"""
        if not apt_list:
            return

        try:
            current_time = datetime.now()

            # 메모리 히스토리에 추가
            notification_data = {'timestamp': current_time, 'apt_list': apt_list.copy()}
            self.notifications_history.append(notification_data)

            # 최근 50개 그룹만 유지
            if len(self.notifications_history) > 50:
                self.notifications_history = self.notifications_history[-50:]

            logging.info(f"[이전 신고가 저장] {len(apt_list)}개 단지를 메모리에 추가 완료")

            # JSON 파일로 즉시 저장
            self.save_notifications_history()

        except Exception as e:
            logging.error(f"이전 신고가 저장 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())

    def download_previous_high_prices(self):
        """핑크색으로 표시된 최근 신고가 단지 목록 팝업 및 HTML 다운로드"""
        # 현재 활성 리스트의 핑크색 단지만 조회
        # 핑크색 조건: prev_max_price > 0 AND last_max_price > prev_max_price
        apts_with_prev_high = []

        logging.info("[이전 신고가 다운로드] 현재 리스트의 핑크색 단지 조회 시작")

        try:
            cursor = self.db_conn.cursor()

            # 1. 현재 활성 리스트의 list_id 조회
            active_list_name = self.active_list.get()
            cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (active_list_name,))
            result = cursor.fetchone()

            if not result:
                logging.warning(f"[이전 신고가 다운로드] 활성 리스트 '{active_list_name}'를 찾을 수 없습니다")
                messagebox.showinfo("알림", f"활성 리스트 '{active_list_name}'를 찾을 수 없습니다.")
                return

            active_list_id = result[0]
            logging.info(f"[활성 리스트] '{active_list_name}' (ID: {active_list_id})")

            # 2. 현재 리스트의 핑크색 단지만 조회
            cursor.execute("""
                SELECT
                    id, apt_name, area,
                    prev_max_price, prev_max_date, prev_max_floor, prev_max_dong,
                    last_max_price, max_price_date, max_price_floor, max_price_dong,
                    sido, sigungu, dong, build_year
                FROM apartments
                WHERE list_id = ?
                  AND prev_max_price > 0
                  AND last_max_price > prev_max_price
                ORDER BY last_max_price DESC
            """, (active_list_id,))
            results = cursor.fetchall()
            logging.info(f"[DB 조회] '{active_list_name}' 리스트에서 {len(results)}개 핑크색 단지 발견")

            for row in results:
                (apt_id, apt_name, area,
                 prev_price, prev_date, prev_floor, prev_dong,
                 last_price, max_date, max_floor, max_dong,
                 sido, sigungu, location_dong, build_year) = row

                logging.info(f"✅ 핑크색 단지: {apt_name} ({prev_price:,}만원 → {last_price:,}만원)")

                apts_with_prev_high.append({
                    'apt_name': apt_name,
                    'area': area or '',
                    'old_price': prev_price,
                    'old_date': prev_date or '',
                    'old_floor': prev_floor or '',
                    'old_dong': prev_dong or '',
                    'new_price': last_price,
                    'date': max_date or '',
                    'floor': max_floor or '',
                    'dong': max_dong or '',
                    'sido': sido or '',
                    'sigungu': sigungu or '',
                    'location_dong': location_dong or '',
                    'build_year': build_year or '',
                    'is_young': False  # 기본값
                })

        except Exception as e:
            logging.error(f"핑크색 단지 조회 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())

        logging.info(f"[이전 신고가 다운로드] 총 {len(apts_with_prev_high)}개 핑크색 단지 발견")

        if not apts_with_prev_high:
            messagebox.showinfo("알림", "이전 신고가 기록이 있는 단지가 없습니다.")
            return

        # 신고가 높은 순으로 정렬
        apts_with_prev_high.sort(key=lambda x: x.get('new_price', 0), reverse=True)

        # DB에 저장 (이전 신고가 기록으로 유지)
        self.save_previous_high_to_db(apts_with_prev_high)

        # 팝업창 표시
        self.show_new_max_notification(apts_with_prev_high)

    def show_previous_notifications(self):
        """이전 신고가 히스토리 표시"""
        # DB에서 최신 데이터 다시 로드 (프로그램 재시작 후에도 유지)
        self.notifications_history = self.load_notifications_history()

        # 디버깅: 로드된 데이터 로깅
        total_apts = sum(len(h.get('apt_list', [])) for h in self.notifications_history)
        logging.info(f"[이전 신고가 표시] {len(self.notifications_history)}개 그룹, 총 {total_apts}개 단지")
        print(f"[DEBUG] 이전 신고가: {len(self.notifications_history)}개 그룹, 총 {total_apts}개 단지")

        if not self.notifications_history:
            messagebox.showinfo("알림", "이전 신고가 기록이 없습니다.")
            return
        history_dialog = tk.Toplevel(self.root)
        history_dialog.title("이전 신고가 기록")
        history_dialog.geometry("600x1000")
        history_dialog.transient(self.root)
        history_dialog.grab_set()
        screen_width = history_dialog.winfo_screenwidth()
        screen_height = history_dialog.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 400) // 2
        history_dialog.geometry(f"600x600+{x}+{y}")
        top_frame = ttk.Frame(history_dialog, padding="10")
        top_frame.pack(fill="x")
        ttk.Label(top_frame, text="이전 신고가 기록", font=self.font_title).pack(side="left")
        list_frame = ttk.Frame(history_dialog, padding="10")
        list_frame.pack(fill="both", expand=True)
        columns = ("timestamp", "count", "max_increase", "preview")
        history_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=12)
        history_tree.heading("timestamp", text="발생 시간")
        history_tree.heading("count", text="아파트 수")
        history_tree.heading("max_increase", text="최대 상승률")
        history_tree.heading("preview", text="미리보기")
        history_tree.column("timestamp", width=150)
        history_tree.column("count", width=80)
        history_tree.column("max_increase", width=100)
        history_tree.column("preview", width=250)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=history_tree.yview)
        history_tree.configure(yscrollcommand=scrollbar.set)
        history_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        sorted_history = sorted(self.notifications_history, key=lambda x: x.get('timestamp', datetime.now()), reverse=True)
        for i, notification_item in enumerate(sorted_history):
            timestamp = notification_item.get('timestamp', datetime.now())
            if isinstance(timestamp, str):
                try:
                    timestamp = datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S')
                except:
                    timestamp = datetime.now()
            apt_list = notification_item.get('apt_list', [])
            count = len(apt_list)
            max_increase = 0
            preview_text = ""
            if apt_list:
                for apt in apt_list:
                    if apt.get('old_price', 0) > 0:
                        increase_percent = ((apt.get('new_price', 0) - apt.get('old_price', 0)) / apt.get('old_price', 0)) * 100
                        max_increase = max(max_increase, increase_percent)
                first_apt = apt_list[0]
                preview_text = f"{first_apt.get('apt_name', '')} {first_apt.get('area', '')}㎡"
                if count > 1:
                    preview_text += f" 외 {count-1}개"
            history_tree.insert("", "end", values=(
                timestamp.strftime('%Y-%m-%d %H:%M'),
                f"{count}개",
                f"+{max_increase:.1f}%" if max_increase > 0 else "-",
                preview_text
            ), tags=(str(i),))
        def on_double_click(event):
            selection = history_tree.selection()
            if selection:
                item = history_tree.item(selection[0])
                tag = item['tags'][0] if item['tags'] else '0'
                idx = int(tag)
                if 0 <= idx < len(sorted_history):
                    selected_notification = sorted_history[idx]
                    apt_list = selected_notification.get('apt_list', [])
                    if apt_list:
                        history_dialog.destroy()
                        original_history_count = len(self.notifications_history)
                        self.show_new_max_notification(apt_list)
                        if len(self.notifications_history) > original_history_count:
                            self.notifications_history = self.notifications_history[:-1]
                            self.save_notifications_history()
        history_tree.bind('<Double-1>', on_double_click)
        button_frame = ttk.Frame(history_dialog, padding="10")
        button_frame.pack(fill="x")
        def show_selected():
            selection = history_tree.selection()
            if not selection:
                messagebox.showinfo("알림", "보려는 기록을 선택해주세요.")
                return
            on_double_click(None)
        def delete_selected():
            selection = history_tree.selection()
            if not selection:
                messagebox.showinfo("알림", "삭제할 기록을 선택해주세요.")
                return
            if not messagebox.askyesno("확인", "선택한 신고가 기록을 삭제하시겠습니까?"):
                return
            try:
                indices_to_delete = []
                for item in selection:
                    tag = history_tree.item(item)['tags'][0] if history_tree.item(item)['tags'] else '0'
                    idx = int(tag)
                    if 0 <= idx < len(sorted_history):
                        target_notification = sorted_history[idx]
                        for orig_idx, orig_notification in enumerate(self.notifications_history):
                            if (orig_notification.get('timestamp') == target_notification.get('timestamp') and
                                len(orig_notification.get('apt_list', [])) == len(target_notification.get('apt_list', []))):
                                indices_to_delete.append(orig_idx)
                                break
                for idx in sorted(set(indices_to_delete), reverse=True):
                    if 0 <= idx < len(self.notifications_history):
                        del self.notifications_history[idx]
                self.save_notifications_history()
                for item in selection:
                    history_tree.delete(item)
                messagebox.showinfo("알림", f"{len(indices_to_delete)}개의 기록이 삭제되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"기록 삭제 중 오류가 발생했습니다: {str(e)}")
        def delete_all():
            if not self.notifications_history:
                messagebox.showinfo("알림", "삭제할 기록이 없습니다.")
                return
            if not messagebox.askyesno("확인", f"모든 신고가 기록({len(self.notifications_history)}개)을 삭제하시겠습니까?\n\n이 작업은 되돌릴 수 없습니다."):
                return
            try:
                self.notifications_history = []
                self.save_notifications_history()
                for item in history_tree.get_children():
                    history_tree.delete(item)
                messagebox.showinfo("알림", "모든 신고가 기록이 삭제되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"기록 삭제 중 오류: {str(e)}")
        ttk.Button(button_frame, text="선택 삭제", style='Accent.TButton', command=delete_selected).pack(side="left", padx=5)
        ttk.Button(button_frame, text="전체 삭제", style='Accent.TButton', command=delete_all).pack(side="left", padx=5)
        ttk.Button(button_frame, text="보기",      style='Accent.TButton', command=show_selected).pack(side="right", padx=5)
        ttk.Button(button_frame, text="닫기",      style='Accent.TButton', command=history_dialog.destroy).pack(side="right", padx=5)

    def treeview_sort_column(self, col, reverse):
        self.sort_column = col
        self.sort_reverse = reverse
        self.update_apt_tree()
        # 다음 클릭 때 방향 토글되도록 현재 열에 재바인딩
        self.apt_tree.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))

    def show_settings_dialog(self):
        """설정 대화상자"""
        settings = tk.Toplevel(self.root)
        settings.title("설정")
        settings.geometry("800x400")
        settings.resizable(False, False)
        settings.transient(self.root)
        settings.grab_set()
        settings.configure(bg=self.palette['bg'])
        style = ttk.Style(settings)
        style.theme_use('clam')

        # 다운로드 경로
        ttk.Label(settings, text="다운로드 경로:", background=self.palette['surface'], foreground=self.palette['text_primary']).grid(row=0, column=0, sticky="w", padx=10, pady=10)
        download_path_var = tk.StringVar(value=self.download_path)
        ttk.Entry(settings, textvariable=download_path_var, width=40).grid(row=0, column=1, padx=5, pady=10)
        def select_download_path():
            path = filedialog.askdirectory(initialdir=self.download_path)
            if path:
                download_path_var.set(path)
        ttk.Button(settings, text="찾아보기", style='Accent.TButton', command=select_download_path).grid(row=0, column=2, padx=5, pady=10)

        # 법정동 파일 경로
        ttk.Label(settings, text="법정동 파일 경로:", background=self.palette['surface'], foreground=self.palette['text_primary']).grid(row=1, column=0, sticky="w", padx=10, pady=10)
        lawdong_path_var = tk.StringVar(value=self.lawdong_path)
        ttk.Entry(settings, textvariable=lawdong_path_var, width=40).grid(row=1, column=1, padx=5, pady=10)
        def select_lawdong_path():
            path = filedialog.askopenfilename(initialdir=os.path.dirname(self.lawdong_path), title="법정동 코드 파일 선택", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
            if path:
                lawdong_path_var.set(path)
        ttk.Button(settings, text="찾아보기", style='Accent.TButton', command=select_lawdong_path).grid(row=1, column=2, padx=5, pady=10)

        # DB 파일 경로 (신규 추가)
        ttk.Label(settings, text="DB 파일 경로:", background=self.palette['surface'], foreground=self.palette['text_primary']).grid(row=2, column=0, sticky="w", padx=10, pady=10)
        db_path_var = tk.StringVar(value=self.db_path)
        ttk.Entry(settings, textvariable=db_path_var, width=40).grid(row=2, column=1, padx=5, pady=10)
        def select_db_path():
            path = filedialog.asksaveasfilename(
                initialdir=os.path.dirname(self.db_path),
                initialfile=os.path.basename(self.db_path),
                title="DB 파일 선택",
                defaultextension=".db",
                filetypes=[("SQLite Database", "*.db"), ("All files", "*.*")]
            )
            if path:
                db_path_var.set(path)
        ttk.Button(settings, text="찾아보기", style='Accent.TButton', command=select_db_path).grid(row=2, column=2, padx=5, pady=10)

        # 백업/복원 섹션 (JSON)
        ttk.Label(settings, text="모니터링 리스트 백업:", background=self.palette['surface'], foreground=self.palette['text_primary']).grid(row=3, column=0, sticky="w", padx=10, pady=10)

        backup_restore_frame = ttk.Frame(settings, style='TFrame')
        backup_restore_frame.grid(row=3, column=1, columnspan=2, sticky="w", padx=5, pady=10)

        ttk.Button(backup_restore_frame, text="백업 저장", style='Accent.TButton',
                  command=lambda: self.backup_monitored_lists(parent_window=settings)).pack(side='left', padx=5)
        ttk.Button(backup_restore_frame, text="백업 불러오기", style='Accent.TButton',
                  command=lambda: self.restore_monitored_lists(parent_window=settings)).pack(side='left', padx=5)
        ttk.Button(backup_restore_frame, text="자동 백업 폴더 열기", style='Accent.TButton',
                  command=self.open_backup_folder).pack(side='left', padx=5)

        # DB 백업/복원 섹션 (신규 추가)
        ttk.Label(settings, text="DB 파일 관리:", background=self.palette['surface'], foreground=self.palette['text_primary']).grid(row=4, column=0, sticky="w", padx=10, pady=10)

        db_backup_frame = ttk.Frame(settings, style='TFrame')
        db_backup_frame.grid(row=4, column=1, columnspan=2, sticky="w", padx=5, pady=10)

        ttk.Button(db_backup_frame, text="DB 수동 저장", style='Accent.TButton',
                  command=lambda: self.manual_save_db(parent_window=settings)).pack(side='left', padx=5)
        ttk.Button(db_backup_frame, text="DB 백업 저장", style='Accent.TButton',
                  command=lambda: self.backup_db_file(parent_window=settings)).pack(side='left', padx=5)
        ttk.Button(db_backup_frame, text="DB 불러오기", style='Accent.TButton',
                  command=lambda: self.restore_db_file(parent_window=settings)).pack(side='left', padx=5)


        button_frame = ttk.Frame(settings, style='TFrame')
        button_frame.grid(row=5, column=0, columnspan=3, sticky="e", padx=10, pady=20)
        def save_settings():
            # 다운로드 경로
            new_dp = download_path_var.get()
            if new_dp:
                os.makedirs(new_dp, exist_ok=True)
                self.download_path = new_dp
                self.monitored_apts_file = os.path.join(self.download_path, "monitored_apts.json")

            # 법정동 파일 경로
            new_lp = lawdong_path_var.get()
            if new_lp and os.path.exists(new_lp):
                self.lawdong_path = new_lp

            # DB 파일 경로 (신규 추가)
            new_db_path = db_path_var.get()
            if new_db_path and new_db_path != self.db_path:
                # DB 파일이 실제로 변경된 경우에만 처리

                # 기존 DB 파일이 있으면 백업
                if os.path.exists(self.db_path):
                    import shutil
                    backup_path = self.db_path.replace('.db', '_before_change.db')
                    try:
                        shutil.copy2(self.db_path, backup_path)
                        logging.info(f"기존 DB 백업: {backup_path}")
                    except Exception as e:
                        logging.error(f"DB 백업 실패: {str(e)}")

                # 기존 DB 연결 종료
                if hasattr(self, 'db_conn') and self.db_conn:
                    try:
                        self.db_conn.close()
                        logging.info("기존 DB 연결 종료")
                    except:
                        pass

                # 새 DB 경로 설정
                self.db_path = new_db_path

                # 새 DB 연결 및 초기화
                try:
                    self.db_conn = init_database(self.db_path)
                    logging.info(f"새 DB 연결 완료: {self.db_path}")

                    # 데이터 다시 로드
                    self.monitored_lists = self.load_monitored_apts()
                    self.notifications_history = self.load_notifications_history()
                    self.update_apt_tree()

                    messagebox.showinfo("알림", f"DB 경로가 변경되었습니다.\n새 경로: {self.db_path}")
                except Exception as e:
                    messagebox.showerror("오류", f"DB 연결 실패: {str(e)}")
                    logging.error(f"DB 연결 실패: {str(e)}")
                    return

            # 설정 저장
            settings_data = {
                'download_path': self.download_path,
                'lawdong_path': self.lawdong_path,
                'db_path': self.db_path,  # DB 경로 추가
                'auto_update': self.auto_update_enabled.get(),
                'update_time': self.update_time.get(),
                'active_list': self.active_list.get(),  # 활성 리스트 저장
                'last_search_region': self.last_search_region  # 마지막 검색 지역 저장
            }
            with open(os.path.join(os.getcwd(), 'monitor_settings.json'), 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)

            self.load_lawdong_file()
            messagebox.showinfo("알림", "설정이 저장되었습니다.")
            settings.destroy()
        ttk.Button(button_frame, text="취소", style='Accent.TButton', command=settings.destroy).pack(side='right', padx=5)
        ttk.Button(button_frame, text="저장", style='Accent.TButton', command=save_settings).pack(side='right')

    def setup_scheduler(self):
        """스케줄러 설정"""
        self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduler_thread.start()
    
    def run_scheduler(self):
        """스케줄러 실행"""
        while True:
            schedule.run_pending()
            time.sleep(1)
    
    def toggle_auto_update(self):
        """자동 업데이트 토글 (태그 관리)"""
        if self.auto_update_enabled.get():
            self.apply_update_time()
        else:
            schedule.clear('auto-update')
            self.status_var.set("자동 업데이트가 비활성화되었습니다.")
    
    def apply_update_time(self):
        """업데이트 시간 적용 (태그 관리)"""
        try:
            schedule.clear('auto-update')
            if self.auto_update_enabled.get():
                update_time = self.update_time.get()
                hours, minutes = map(int, update_time.split(':'))
                if not (0 <= hours < 24 and 0 <= minutes < 60):
                    raise ValueError("올바른 시간 형식이 아닙니다.")
                schedule.every().day.at(update_time).do(self.update_all_data).tag('auto-update')
                self.status_var.set(f"자동 업데이트가 {update_time}에 실행되도록 설정되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"업데이트 시간 설정 오류: {str(e)}")
            self.status_var.set("자동 업데이트 설정 중 오류가 발생했습니다.")
        
    def update_apt_tree(self):
        """모니터링 아파트 목록 트리뷰 업데이트 (검색 필터 적용)"""
        # 모두 지우기
        for item in self.apt_tree.get_children():
            self.apt_tree.delete(item)
    
        # 1) 정렬
        sorted_apts = self.get_sorted_apts()
    
        # 2) 검색 필터
        query = getattr(self, 'search_var', tk.StringVar(value="")).get() if hasattr(self, 'search_var') else ""
        filtered_apts = self._filter_apts_by_query(sorted_apts, query)
    
        # 3) 표시

        for apt in filtered_apts:   # 또는 sorted_apts 사용 중이면 거기 맞춰서
            prev_max_price = apt.get('prev_max_price', 0)
            prev_max_price_str = f"{prev_max_price:,}만원" if prev_max_price > 0 else "-"
            prev_date = apt.get('prev_max_date', '')
    
            last_max_price = apt.get('last_max_price', 0)
            last_max_price_str = f"{last_max_price:,}만원" if last_max_price > 0 else "-"
            last_date = apt.get('max_price_date', '-')
            dong_info = apt.get('max_price_dong', '-')
            floor_info = apt.get('max_price_floor', '-')
    
            build_year = apt.get('build_year', '')
            build_year_str = f"{build_year}년" if (build_year and build_year != '분양') else ("분양" if build_year == '분양' else "-")
            location = f"{apt.get('sigungu', '')} {apt.get('dong', '')}"
    
            # ★ 평균평단가 계산: 최근신고가 / 전용면적 * 3.3
            try:
                area_val = float(str(apt.get('area','')).replace('㎡','').strip() or 0)
            except:
                area_val = 0.0
            if last_max_price and area_val > 0:
                avg_py = last_max_price / area_val * 3.3
                # 표시는 반올림 정수(만원/평). 원하면 소수 1자리로 바꿔도 됨: f"{avg_py:,.1f}만원/평"
                avg_py_str = f"{round(avg_py):,}만원/평"
            else:
                avg_py_str = "-"
    
            item_id = self.apt_tree.insert("", "end", values=(
                apt["apt_name"],
                build_year_str,
                f"{apt['area']}㎡",
                location,
                dong_info,
                floor_info,
                prev_max_price_str,
                prev_date,
                last_max_price_str,
                avg_py_str,                 # ★ 추가: 최근 신고가 바로 다음
                last_date,
                apt.get("last_update", "업데이트 필요")
            ))
    
            # ★★★ 핑크색 표시: 최근 7일 이내에 신고가가 갱신된 단지만 ★★★
            if prev_max_price > 0 and last_max_price > prev_max_price:
                # last_update 시간 확인 (최근 7일 이내만 핑크색)
                last_update_str = apt.get('last_update', '')
                is_recent = False
                if last_update_str:
                    try:
                        last_update_time = datetime.strptime(last_update_str, '%Y-%m-%d %H:%M')
                        days_diff = (datetime.now() - last_update_time).days
                        if days_diff <= 7:  # 7일 이내
                            is_recent = True
                    except:
                        pass

                if is_recent:
                    self.apt_tree.item(item_id, tags=('price_up',))
            elif prev_max_price > 0 and last_max_price < prev_max_price:
                self.apt_tree.item(item_id, tags=('price_down',))
    
        self.apt_tree.tag_configure('price_up', background='#FFDDDD', foreground='#3AA6FF')
        self.apt_tree.tag_configure('price_down', background='#DDDDFF')
    
        # 결과 개수 표시
        try:
            total = len(sorted_apts)
            shown = len(filtered_apts)
            self.search_count_var.set(f"{shown}/{total}개 표시")
        except Exception:
            pass
    
        self.adjust_column_widths()


    def get_sorted_apts(self):
        """정렬된 아파트 목록 반환"""
        sorted_apts = self.monitored_apts.copy()
        def get_sort_key(apt):
            if self.sort_column == "apt_name":
                return apt.get("apt_name", "")
            elif self.sort_column == "build_year":
                try:
                    return int(apt.get("build_year", 0))
                except:
                    return 0
            elif self.sort_column == "area":
                try:
                    return float(apt.get("area", 0))
                except:
                    return 0
            elif self.sort_column == "location":
                return f"{apt.get('sigungu', '')} {apt.get('dong', '')}"
            elif self.sort_column == "dong":
                return apt.get("max_price_dong", "")
            elif self.sort_column == "prev_max_price":
                return apt.get("prev_max_price", 0)
            elif self.sort_column == "prev_date":
                return apt.get("prev_max_date", "")
            elif self.sort_column == "last_max_price":
                return apt.get("last_max_price", 0)
            elif self.sort_column == "last_date":
                return apt.get("max_price_date", "")
            elif self.sort_column == "floor":
                try:
                    return int(apt.get("max_price_floor", 0))
                except:
                    return 0
            elif self.sort_column == "last_update":
                return apt.get("last_update", "")
            elif self.sort_column == "avg_py_price":     # ★ 추가
                try:
                    area_val = float(str(apt.get("area","")).replace("㎡","").strip() or 0)
                except:
                    area_val = 0.0
                last_max = apt.get("last_max_price", 0) or 0
                return (last_max / area_val * 3.3) if (last_max and area_val > 0) else -1
            else:
                return 0
    
        sorted_apts.sort(key=get_sort_key, reverse=self.sort_reverse)
        return sorted_apts

    def on_sido_selected(self, event):
        """시/도 선택"""
        sido = self.sido_combobox.get()
        if sido in self.sigungu_dict:
            self.sigungu_combobox['values'] = sorted(self.sigungu_dict[sido])
            self.sigungu_combobox.set("시/군/구 선택")
            self.dong_combobox.set("읍/면/동 선택")
        
    def on_sigungu_selected(self, event):
        """시/군/구 선택"""
        sigungu = self.sigungu_combobox.get()
        logging.info(f"[시군구 선택] {sigungu}")
        print(f"[시군구 선택] {sigungu}")
        if sigungu in self.dong_dict:
            dong_list = self.dong_dict[sigungu]
            gu_list = [dong for dong in dong_list if dong.endswith('구')]
            if gu_list:
                all_items = []
                for gu in sorted(gu_list):
                    all_items.append(gu)
                    gu_key = f"{sigungu}_{gu}"
                    if gu_key in self.dong_dict:
                        gu_dong_list = self.dong_dict[gu_key]
                        for dong in sorted(gu_dong_list):
                            all_items.append(f"  └ {dong}")
                normal_dong_list = [dong for dong in dong_list if not dong.endswith('구')]
                if normal_dong_list:
                    all_items.extend(sorted(normal_dong_list))
                self.dong_combobox['values'] = all_items
            else:
                self.dong_combobox['values'] = sorted(dong_list)
            self.dong_combobox.set("읍/면/동 선택")
        else:
            self.dong_combobox['values'] = []
            self.dong_combobox.set("읍/면/동 선택")
    
    def on_dong_selected(self, event):
        """읍/면/동 선택"""
        dong = self.dong_combobox.get()
        sigungu = self.sigungu_combobox.get()
        if dong.endswith('구'):
            gu_key = f"{sigungu}_{dong}"
            if gu_key in self.dong_dict:
                dong_list = sorted(self.dong_dict[gu_key])
                self.dong_combobox['values'] = dong_list
                self.dong_combobox.set("읍/면/동 선택")
                self.status_var.set(f"{dong}의 하위 동을 선택해주세요.")
            else:
                self.status_var.set("아파트 목록 조회 버튼을 눌러 진행하세요.")
        else:
            self.status_var.set("아파트 목록 조회 버튼을 눌러 진행하세요.")
    
    def show_apt_list(self):
        """아파트 목록 조회 및 표시"""
        sido = self.sido_combobox.get()
        sigungu = self.sigungu_combobox.get()
        dong = self.dong_combobox.get()
        original_dong = dong
        
        # ⭐ 디버그 로그 추가
        print(f"\n{'='*80}")
        print(f"=== show_apt_list 호출 ===")
        print(f"sido: '{sido}'")
        print(f"sigungu: '{sigungu}'")
        print(f"dong: '{dong}'")
        print(f"체크박스: {self.bulk_monitor_enabled.get()}")
        print(f"'선택' not in [sido, sigungu]: {'선택' not in [sido, sigungu]}")
        print(f"'선택' in dong: {'선택' in dong}")
        print(f"{'='*80}\n")  # ← 올바른 따옴표
        
        # ⭐⭐⭐ 이 부분이 핵심! 순서 변경 ⭐⭐⭐
        # 1️⃣ 먼저 체크박스가 켜져있고 시/군/구까지만 선택된 경우 처리
        if self.bulk_monitor_enabled.get() and "선택" not in [sido, sigungu] and "선택" in dong:
            print(">>> ✅ 시/군/구 전체 등록 분기 진입!")
            # 해당 시/군/구의 모든 동 처리
            if messagebox.askyesno("확인",
                f"{sido} {sigungu}의 모든 동을 일괄 등록하시겠습니까?\n"
                "이 작업은 오래 걸릴 수 있습니다."):
                # 마지막 검색 지역 업데이트
                from datetime import datetime
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
                self.last_search_region = f"{sido} {sigungu} 전체"
                self.last_search_label.config(text=f"마지막 검색: {self.last_search_region} ({current_time})")
                self.save_last_search_region()
                self.bulk_add_all_dongs_in_sigungu(sido, sigungu)
            return  # ⭐ 여기서 종료!
        
        # 2️⃣ 체크박스가 켜져있고 시/도만 선택된 경우
        if self.bulk_monitor_enabled.get() and "선택" not in [sido] and "선택" in [sigungu, dong]:
            print(">>> ✅ 시/도 전체 등록 분기 진입!")
            # 해당 시/도의 모든 시군구와 동 처리
            if messagebox.askyesno("확인",
                f"{sido}의 모든 지역을 일괄 등록하시겠습니까?\n"
                "이 작업은 매우 오래 걸릴 수 있습니다. (수 시간 소요 예상)"):
                # 마지막 검색 지역 업데이트
                from datetime import datetime
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
                self.last_search_region = f"{sido} 전체"
                self.last_search_label.config(text=f"마지막 검색: {self.last_search_region} ({current_time})")
                self.save_last_search_region()
                self.bulk_add_all_areas_in_sido(sido)
            return
        
        # 3️⃣ 일반 동 선택 처리
        if dong.startswith("  └ "):
            dong = dong.replace("  └ ", "").strip()
        
        if "선택" in [sido, sigungu, dong]:
            print(">>> ❌ 지역을 모두 선택해주세요")
            messagebox.showerror("오류", "지역을 모두 선택해주세요.")
            return
        
        if dong.startswith("  └ "):
            dong = dong.replace("  └ ", "").strip()
        if "선택" in [sido, sigungu, dong]:
            messagebox.showerror("오류", "지역을 모두 선택해주세요.")
            return
        parent_gu = None
        if original_dong.startswith("  └ "):
            dong_list = self.dong_combobox['values']
            for i, item in enumerate(dong_list):
                if item == original_dong:
                    for j in range(i-1, -1, -1):
                        if not dong_list[j].startswith("  └ ") and dong_list[j].endswith('구'):
                            parent_gu = dong_list[j]
                            break
                    break
        sigungu_code_to_use = None
        if parent_gu:
            if hasattr(self, 'gu_info'):
                if sigungu in self.sigungu_to_full_info:
                    _, original_si, _ = self.sigungu_to_full_info[sigungu]
                else:
                    original_si = sigungu.replace('(경)', '').replace('(충)', '').replace('(전)', '').strip()
                gu_key = f"{sido}_{original_si}_{parent_gu}"
                if gu_key in self.gu_info:
                    sigungu_code_to_use = self.gu_info[gu_key]
            if not sigungu_code_to_use:
                region_key = f"{sigungu}_{parent_gu}_{dong}"
                if region_key in self.region_codes:
                    _, sigungu_code_to_use = self.region_codes[region_key]
        elif dong.endswith('구'):
            if hasattr(self, 'gu_info'):
                if sigungu in self.sigungu_to_full_info:
                    _, original_si, _ = self.sigungu_to_full_info[sigungu]
                else:
                    original_si = sigungu.replace('(경)', '').replace('(충)', '').replace('(전)', '').strip()
                gu_key = f"{sido}_{original_si}_{dong}"
                if gu_key in self.gu_info:
                    sigungu_code_to_use = self.gu_info[gu_key]
        else:
            region_code = self.region_codes.get((sido, sigungu, dong))
            if region_code:
                sigungu_code_to_use = region_code[1]
        if not sigungu_code_to_use:
            if sigungu in self.sigungu_to_full_info:
                _, _, sigungu_code_to_use = self.sigungu_to_full_info[sigungu]
            else:
                messagebox.showerror("오류", "해당 지역의 코드를 찾을 수 없습니다.")
                return
        # 마지막 검색 지역 업데이트
        from datetime import datetime
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
        self.last_search_region = f"{sido} {sigungu} {dong}"
        self.last_search_label.config(text=f"마지막 검색: {self.last_search_region} ({current_time})")
        self.save_last_search_region()

        try:
            self.status_var.set("아파트 목록 검색 중...")
            self.root.update_idletasks()
            apt_list = self.get_apt_list_from_api(sigungu_code_to_use, dong)
            if apt_list:
                if self.bulk_monitor_enabled.get():
                    # ▼▼ 추가: 체크박스 켜짐 → 동 내 조회된 모든 단지를 일괄 추가
                    self.bulk_add_all_complexes_in_dong(
                        sigungu_code=sigungu_code_to_use,
                        dong=dong,
                        sido=sido,
                        sigungu=sigungu,
                        apt_list=apt_list  # get_apt_list_from_api() 반환 그대로 투입
                    )
                else:
                    # 기존 동작: 단일 단지 선택 다이얼로그
                    dialog = AptSelectDialog(
                        self.root, 
                        apt_list,
                        self.service_key,
                        sigungu_code_to_use,
                        dong,
                        sido,
                        sigungu,
                        title=f"{dong} 아파트 목록"
                    )
                    self.root.wait_window(dialog.top)
                    if dialog.result:
                        apt_info = dialog.result
                        apt_info['sigungu_code'] = sigungu_code_to_use
                        self.add_apt_to_monitored(apt_info)
            else:
                messagebox.showinfo("알림", f"{dong}에 거래 내역이 있는 아파트가 없습니다.")

        except Exception as e:
            messagebox.showerror("오류", f"아파트 목록 검색 중 오류: {str(e)}")
            import traceback
            traceback.print_exc()

    def bulk_add_all_areas_in_sido(self, sido):
        """시/도 내 모든 시군구의 모든 동 단지를 일괄 등록"""
        # 진행창
        win = tk.Toplevel(self.root)
        win.title(f"{sido} 전체 단지 등록 중...")
        win.geometry("700x250")
        win.transient(self.root)
        win.grab_set()
        
        ttk.Label(win, text=f"{sido}의 모든 시/군/구와 동을 순회하며 단지를 등록합니다.", 
                 font=self.font_large).pack(pady=10)
        
        progress_label = ttk.Label(win, text="준비 중...")
        progress_label.pack(pady=5)
        
        bar = ttk.Progressbar(win, orient="horizontal", length=660, mode="determinate")
        bar.pack(pady=10, padx=20)
        
        info_label = ttk.Label(win, text="")
        info_label.pack(pady=5)
        
        stats_label = ttk.Label(win, text="")
        stats_label.pack(pady=5)
        
        cancel_flag = [False]
        
        def _cancel():
            cancel_flag[0] = True
            try: win.destroy()
            except: pass
        
        ttk.Button(win, text="중단", command=_cancel).pack(pady=10)
        
        def process():
            try:
                # 해당 시/도의 시군구 목록 가져오기
                if sido not in self.sigungu_dict:
                    messagebox.showerror("오류", f"{sido}의 시군구 목록을 찾을 수 없습니다.")
                    win.destroy()
                    return
                
                sigungu_list = self.sigungu_dict[sido]
                total_sigungus = len(sigungu_list)
                total_dongs_processed = 0
                total_apts_added = 0
                
                for sigungu_idx, sigungu in enumerate(sigungu_list):
                    if cancel_flag[0]:
                        break
                    
                    # 시군구 진행률
                    sigungu_progress = (sigungu_idx / total_sigungus) * 100
                    progress_label.config(text=f"시/군/구: {sigungu_idx + 1}/{total_sigungus} - {sigungu}")
                    
                    # 시군구 코드 얻기
                    sigungu_code = None
                    if sigungu in self.sigungu_to_full_info:
                        _, _, sigungu_code = self.sigungu_to_full_info[sigungu]
                    
                    if not sigungu_code:
                        continue
                    
                    # 해당 시군구의 동 목록 가져오기
                    dong_list = []
                    if sigungu in self.dong_dict:
                        all_dongs = self.dong_dict[sigungu]
                        # 구가 아닌 일반 동만 필터링
                        dong_list = [d for d in all_dongs if not d.endswith('구')]
                    
                    for dong_idx, dong_name in enumerate(dong_list):
                        if cancel_flag[0]:
                            break
                        
                        total_dongs_processed += 1
                        
                        # 전체 진행률
                        dong_progress = ((dong_idx + 1) / len(dong_list)) * (100 / total_sigungus)
                        total_progress = sigungu_progress + dong_progress
                        bar['value'] = total_progress
                        
                        info_label.config(text=f"현재 처리: {sigungu} {dong_name}")
                        stats_label.config(text=f"처리한 동: {total_dongs_processed}개, 추가된 단지: {total_apts_added}개 (추정)")
                        win.update_idletasks()
                        
                        try:
                            # 해당 동의 아파트 목록 조회
                            apt_list = self.get_apt_list_from_api(sigungu_code, dong_name)
                            
                            if apt_list:
                                # 각 동의 단지 추가 (기존 함수 활용)
                                # 추가된 개수를 반환하도록 수정 필요
                                before_count = len(self.monitored_apts)
                                
                                self.bulk_add_all_complexes_in_dong(
                                    sigungu_code=sigungu_code,
                                    dong=dong_name,
                                    sido=sido,
                                    sigungu=sigungu,
                                    apt_list=apt_list
                                )
                                
                                after_count = len(self.monitored_apts)
                                total_apts_added += (after_count - before_count)
                            
                            # API 호출 간격 조절
                            time.sleep(0.3)
                            
                        except Exception as e:
                            logging.error(f"{dong_name} 처리 중 오류: {str(e)}")
                            continue
                    
                    # 시군구 단위로 저장
                    self.save_monitored_apts()
                
                # 완료
                self.update_apt_tree()
                
                messagebox.showinfo("완료", 
                    f"{sido} 전체 단지 등록 완료\n"
                    f"처리한 시/군/구: {total_sigungus}개\n"
                    f"처리한 동: {total_dongs_processed}개\n"
                    f"추가된 단지: {total_apts_added}개")
                
            except Exception as e:
                messagebox.showerror("오류", f"처리 중 오류: {str(e)}")
            finally:
                try: win.destroy()
                except: pass
        
        # 별도 스레드에서 실행
        thread = threading.Thread(target=process, daemon=True)
        thread.start()

            
    def show_bulk_add_result(self, title, total_dongs, added_dongs, skipped_dongs, added_complexes, total_added):
        """일괄 추가 결과를 상세하게 보여주는 팝업"""
        result_win = tk.Toplevel(self.root)
        result_win.title(title)
        result_win.geometry("800x600")
        result_win.transient(self.root)

        # 메인 프레임
        main_frame = ttk.Frame(result_win, padding=10)
        main_frame.pack(fill="both", expand=True)

        # 요약 정보
        summary_frame = ttk.LabelFrame(main_frame, text="요약", padding=10)
        summary_frame.pack(fill="x", pady=(0, 10))

        summary_text = f"""총 처리한 동: {total_dongs}개
추가한 동: {len(added_dongs)}개
건너뛴 동: {len(skipped_dongs)}개
총 추가된 단지: {total_added}개"""

        ttk.Label(summary_frame, text=summary_text, font=self.font_normal).pack(anchor="w")

        # 탭 생성
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True)

        # 탭 1: 추가한 동 목록
        added_dongs_frame = ttk.Frame(notebook)
        notebook.add(added_dongs_frame, text=f"추가한 동 ({len(added_dongs)})")

        added_dongs_text = tk.Text(added_dongs_frame, wrap="word", height=20, font=self.font_normal)
        added_dongs_scroll = ttk.Scrollbar(added_dongs_frame, command=added_dongs_text.yview)
        added_dongs_text.configure(yscrollcommand=added_dongs_scroll.set)
        added_dongs_text.pack(side="left", fill="both", expand=True)
        added_dongs_scroll.pack(side="right", fill="y")

        if added_dongs:
            added_dongs_text.insert("1.0", "\n".join(added_dongs))
        else:
            added_dongs_text.insert("1.0", "추가한 동이 없습니다.")
        added_dongs_text.config(state="disabled")

        # 탭 2: 건너뛴 동 목록
        skipped_dongs_frame = ttk.Frame(notebook)
        notebook.add(skipped_dongs_frame, text=f"건너뛴 동 ({len(skipped_dongs)})")

        skipped_dongs_text = tk.Text(skipped_dongs_frame, wrap="word", height=20, font=self.font_normal)
        skipped_dongs_scroll = ttk.Scrollbar(skipped_dongs_frame, command=skipped_dongs_text.yview)
        skipped_dongs_text.configure(yscrollcommand=skipped_dongs_scroll.set)
        skipped_dongs_text.pack(side="left", fill="both", expand=True)
        skipped_dongs_scroll.pack(side="right", fill="y")

        if skipped_dongs:
            skipped_dongs_text.insert("1.0", "\n".join(skipped_dongs))
        else:
            skipped_dongs_text.insert("1.0", "건너뛴 동이 없습니다.")
        skipped_dongs_text.config(state="disabled")

        # 탭 3: 추가된 단지 목록
        complexes_frame = ttk.Frame(notebook)
        notebook.add(complexes_frame, text=f"추가된 단지 ({len(added_complexes)})")

        complexes_text = tk.Text(complexes_frame, wrap="word", height=20, font=self.font_normal)
        complexes_scroll = ttk.Scrollbar(complexes_frame, command=complexes_text.yview)
        complexes_text.configure(yscrollcommand=complexes_scroll.set)
        complexes_text.pack(side="left", fill="both", expand=True)
        complexes_scroll.pack(side="right", fill="y")

        if added_complexes:
            # 동별로 그룹화해서 표시
            from collections import defaultdict
            complexes_by_dong = defaultdict(list)
            for complex_info in added_complexes:
                complexes_by_dong[complex_info['dong']].append(
                    f"  - {complex_info['apt_name']} ({complex_info['area']}㎡)"
                )

            result_lines = []
            for dong in sorted(complexes_by_dong.keys()):
                result_lines.append(f"[{dong}]")
                result_lines.extend(complexes_by_dong[dong])
                result_lines.append("")  # 빈 줄

            complexes_text.insert("1.0", "\n".join(result_lines))
        else:
            complexes_text.insert("1.0", "추가된 단지가 없습니다.")
        complexes_text.config(state="disabled")

        # 닫기 버튼
        ttk.Button(main_frame, text="닫기", command=result_win.destroy, style='Accent.TButton').pack(pady=(10, 0))

    def bulk_add_all_dongs_in_sigungu(self, sido, sigungu):
        """시/군/구 내 모든 동의 단지를 일괄 등록"""
        # 진행창
        win = tk.Toplevel(self.root)
        win.title(f"{sigungu} 전체 단지 등록 중...")
        win.geometry("600x250")
        win.transient(self.root)
        win.grab_set()
        
        ttk.Label(win, text=f"{sido} {sigungu}의 모든 동을 순회하며 단지를 등록합니다.", 
                 font=self.font_large).pack(pady=10)
        
        progress_label = ttk.Label(win, text="준비 중...")
        progress_label.pack(pady=5)
        
        bar = ttk.Progressbar(win, orient="horizontal", length=560, mode="determinate")
        bar.pack(pady=10, padx=20)
        
        info_label = ttk.Label(win, text="")
        info_label.pack(pady=5)
        
        stats_label = ttk.Label(win, text="")
        stats_label.pack(pady=5)
        
        cancel_flag = [False]
        
        def _cancel():
            cancel_flag[0] = True
            try: win.destroy()
            except: pass
        
        ttk.Button(win, text="중단", command=_cancel).pack(pady=10)
        
        def process():
            try:
                # 시군구 코드 얻기
                sigungu_code = None
                if sigungu in self.sigungu_to_full_info:
                    _, _, sigungu_code = self.sigungu_to_full_info[sigungu]
                
                if not sigungu_code:
                    messagebox.showerror("오류", "시군구 코드를 찾을 수 없습니다.")
                    win.destroy()
                    return
                
                # ===== 핵심 수정 부분: 법정동 데이터에서 동 목록 가져오기 =====
                dong_list = []

                # self.dong_dict에서 해당 시군구의 동 목록 가져오기
                if sigungu in self.dong_dict:
                    dong_list.extend(self.dong_dict[sigungu])

                # 구(區) 하위의 동들도 포함 (예: 수원시(경) -> 영통구 -> 영통동, 매탄동 등)
                for key in self.dong_dict.keys():
                    if isinstance(key, str) and key.startswith(f"{sigungu}_"):
                        # "수원시(경)_영통구" 형태의 키
                        sub_dongs = self.dong_dict[key]
                        dong_list.extend(sub_dongs)

                # 중복 제거 및 정렬
                dong_list = sorted(list(set(dong_list)))

                # 구(區)는 제외 (실제 읍면동만)
                dong_list = [d for d in dong_list if not d.endswith('구')]

                logging.info(f"{sigungu}의 읍면동 목록 ({len(dong_list)}개): {dong_list[:10]}...")
                
                if not dong_list:
                    messagebox.showinfo("알림", f"{sigungu}에 등록 가능한 동이 없습니다.")
                    win.destroy()
                    return
                
                total_dongs = len(dong_list)
                total_added = 0
                total_skipped = 0

                # 상세 정보 저장용
                added_dongs = []  # 추가한 동 목록
                skipped_dongs = []  # 건너뛴 동 목록
                added_complexes = []  # 추가된 단지 목록 (동, 단지명, 면적)

                progress_label.config(text=f"총 {total_dongs}개 동 발견 - 병렬 처리 시작")
                win.update_idletasks()

                # ===== 🚀 병렬 처리로 속도 개선 =====
                from concurrent.futures import ThreadPoolExecutor, as_completed
                import queue

                results_queue = queue.Queue()
                max_workers = 5  # 동시에 5개 동 처리

                def process_dong(dong_name, idx):
                    """단일 동 처리 함수 (병렬 실행용)"""
                    try:
                        if cancel_flag[0]:
                            return None

                        # API 호출
                        apt_list = self.get_apt_list_from_api(sigungu_code, dong_name)

                        if apt_list:
                            result = self.bulk_add_all_complexes_in_dong(
                                sigungu_code=sigungu_code,
                                dong=dong_name,
                                sido=sido,
                                sigungu=sigungu,
                                apt_list=apt_list,
                                silent=True
                            )
                            return {'idx': idx, 'dong': dong_name, 'result': result, 'error': None}
                        else:
                            return {'idx': idx, 'dong': dong_name, 'result': None, 'error': 'no_data'}

                    except Exception as e:
                        logging.error(f"{dong_name} 처리 중 오류: {str(e)}")
                        return {'idx': idx, 'dong': dong_name, 'result': None, 'error': str(e)}

                # 병렬 실행
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # 모든 동에 대해 작업 제출
                    future_to_dong = {
                        executor.submit(process_dong, dong_name, idx): dong_name
                        for idx, dong_name in enumerate(dong_list)
                    }

                    completed = 0

                    # 완료된 작업 순서대로 처리
                    for future in as_completed(future_to_dong):
                        if cancel_flag[0]:
                            executor.shutdown(wait=False, cancel_futures=True)
                            break

                        result_data = future.result()

                        if result_data:
                            completed += 1

                            # 결과 누적
                            if result_data['result']:
                                dong_added = result_data['result'].get('added', 0)
                                dong_skipped = result_data['result'].get('skipped', 0)
                                total_added += dong_added
                                total_skipped += dong_skipped

                                # 동 단위 결과 저장
                                if dong_added > 0:
                                    added_dongs.append(result_data['dong'])
                                    # 추가된 단지 목록 저장
                                    for complex_info in result_data['result'].get('added_list', []):
                                        added_complexes.append({
                                            'dong': result_data['dong'],
                                            'apt_name': complex_info['apt_name'],
                                            'area': complex_info['area']
                                        })
                                else:
                                    skipped_dongs.append(result_data['dong'])
                            elif result_data['error']:
                                total_skipped += 1
                                skipped_dongs.append(result_data['dong'])

                            # UI 업데이트 (10개마다만)
                            if completed % 10 == 0 or completed == total_dongs:
                                progress = (completed / total_dongs) * 100
                                bar['value'] = progress
                                progress_label.config(text=f"동 진행: {completed}/{total_dongs}")
                                info_label.config(text=f"최근 처리: {result_data['dong']}")
                                stats_label.config(text=f"추가: {total_added}개 | 건너뜀: {total_skipped}개")
                                win.update_idletasks()

                            # 20개 동마다 저장
                            if completed % 20 == 0:
                                self.save_monitored_apts()
                
                # 완료
                self.save_monitored_apts()
                self.update_apt_tree()

                # 상세 결과 팝업 표시
                self.show_bulk_add_result(
                    title=f"{sigungu} 전체 단지 등록 완료",
                    total_dongs=total_dongs,
                    added_dongs=added_dongs,
                    skipped_dongs=skipped_dongs,
                    added_complexes=added_complexes,
                    total_added=total_added
                )
                
            except Exception as e:
                messagebox.showerror("오류", f"처리 중 오류: {str(e)}")
                logging.error(f"bulk_add_all_dongs_in_sigungu 오류: {str(e)}")
            finally:
                try: win.destroy()
                except: pass
        
        # 별도 스레드에서 실행
        thread = threading.Thread(target=process, daemon=True)
        thread.start()
    

        
    def get_recent_areas_for_apt(self, sigungu_code, dong, apt_name, months=4):
        """최근 months 개월 캐시에서 해당 단지의 전용면적 목록 추출 (빠르고 가벼움)"""
        from collections import Counter
        areas = []
        now = datetime.now()
        for m in range(months):
            deal_ymd = (now - timedelta(days=30*m)).strftime("%Y%m")
            # 기축
            data_ex = self.get_cached_api_data(sigungu_code, deal_ymd, 'existing')
            for it in data_ex:
                if it.get('apt_name','').strip() == apt_name and it.get('dong','').strip() == dong:
                    try:
                        areas.append(str(int(float(it.get('area',0)))))
                    except:
                        pass
            # 신축/분양권
            data_new = self.get_cached_api_data(sigungu_code, deal_ymd, 'new')
            for it in data_new:
                if it.get('apt_name','').strip() == apt_name and it.get('dong','').strip() == dong:
                    try:
                        areas.append(str(int(float(it.get('area',0)))))
                    except:
                        pass
        if not areas:
            return []
        # 가장 많이 등장한 면적(대표면적) 우선 정렬
        c = Counter(areas)
        return [k for k,_ in c.most_common()]


    
    def bulk_add_all_complexes_in_dong(self, *, sigungu_code, dong, sido, sigungu, apt_list, silent=False):
        """해당 동에서 조회된 모든 '단지'를 대표면적 1개씩 골라 모니터링 목록에 추가
        
        Args:
            silent: True면 진행창과 완료 메시지를 표시하지 않음 (일괄 처리용)
        """
        import re
        
        # ⭐ silent 모드 처리
        if not silent:
            # 진행창 표시
            win = tk.Toplevel(self.root)
            win.title(f"{dong} 전체 단지 추가 중…")
            win.geometry("520x170")
            win.transient(self.root)
            win.grab_set()
            ttk.Label(win, text=f"'{sido} {sigungu} {dong}'의 조회된 모든 단지를 모니터링에 추가합니다.").pack(pady=(10,6))
            bar = ttk.Progressbar(win, orient="horizontal", length=480, mode="determinate")
            bar.pack(pady=6)
            info = ttk.Label(win, text="준비 중…")
            info.pack()
            cancel = [False]
            def _cancel():
                cancel[0] = True
                try: win.destroy()
                except: pass
            ttk.Button(win, text="중단", command=_cancel).pack(pady=(8,6))
        else:
            # silent 모드: 더미 객체 사용
            win = None
            bar = type('obj', (object,), {'__setitem__': lambda *args: None})()
            info = type('obj', (object,), {'config': lambda *args: None, 'pack': lambda *args: None})()
            cancel = [False]
        
        # apt_list 항목에서 단지명 파싱
        def parse_apt_name(s):
            # 형태 예: "[신축] 단지명 [도로/지번] (준공: 2019년)" 또는 "단지명 [도로/지번] (준공: ...)"
            txt = s.strip()
            if txt.startswith("[신축]"):
                txt = txt[5:].strip()
            if '[' in txt:
                return txt.split('[')[0].strip()
            # fallback
            return txt.split('(준공')[0].strip()
        
        # 중복 방지: (apt_name, area_int) 기준
        def _areas_equal(a, b, tol=0.15):
            try:
                fa = float(str(a).replace('㎡','').strip())
                fb = float(str(b).replace('㎡','').strip())
                return abs(fa - fb) <= tol
            except:
                return str(a).replace('㎡','').strip() == str(b).replace('㎡','').strip()
        
        def exists_in_monitor(apt_name, area):
            for apt in self.monitored_apts:
                if apt.get('apt_name') == apt_name and _areas_equal(apt.get('area',''), area):
                    return True
            return False
        
        total = len(apt_list)
        added = 0
        skipped = 0
        added_list = []  # 추가된 단지 목록 (단지명, 면적)

        for i, row in enumerate(apt_list, start=1):
            if cancel[0]:
                break
            
            # 진행률 업데이트
            if not silent:
                bar['value'] = (i/total)*100
                info.config(text=f"[{i}/{total}] {parse_apt_name(row)} 처리 중…")
                win.update_idletasks()
            
            apt_name = parse_apt_name(row)
            
            # ▼▼ 추가: 연식(준공연도/분양) 우선 확보 (캐시 → 목록 문자열 파싱)
            build_year = self.resolve_build_year(
                sigungu_code=sigungu_code,
                dong=dong,
                apt_name=apt_name,
                months=12,
                fallback_text=row  # apt_list의 한 줄 원문
            )
            
            # 면적 후보(최근 4개월) 모두 사용
            areas = self.get_recent_areas_for_apt(sigungu_code, dong, apt_name, months=4)
            if not areas:
                skipped += 1
                continue
            
            # 중복 제거 및 숫자 오름차순 정렬
            def _area_to_float(a):
                try:
                    return float(str(a).replace('㎡','').strip())
                except:
                    return None
            
            uniq_areas = []
            seen = set()
            for a in areas:
                key = str(a).replace('㎡','').strip()
                if key not in seen:
                    seen.add(key)
                    uniq_areas.append(key)
            uniq_areas.sort(key=lambda x: (_area_to_float(x) is None, _area_to_float(x) or 0.0))
            
            # 모든 면적에 대해 추가 시도
            for area in uniq_areas:
                if exists_in_monitor(apt_name, area):
                    skipped += 1
                    continue
                
                apt_info = {
                    'apt_name': apt_name,
                    'jibun_addr': dong,
                    'area': area,              # 문자열 "84" 형태
                    'sido': sido,
                    'sigungu': sigungu,
                    'dong': dong,
                    'sigungu_code': sigungu_code,
                    'build_year': build_year   # ← 캐시/목록에서 확보한 연식 반영 ('', '분양', '2012' 등)
                }
                
                # 빠른 초기 수집: 최근 4개월만(조용히)
                trades = self.collect_apt_data_silent_with_cache(apt_info)
                if not trades:
                    skipped += 1
                    continue
                
                max_trade = max(trades, key=lambda x: x.get('price',0))
                apt_info.update({
                    'prev_max_price': 0,
                    'prev_max_date': '',
                    'prev_max_floor': '',
                    'prev_max_dong': '',
                    'last_max_price': max_trade.get('price',0),
                    'max_price_date': (max_trade.get('date') or datetime.now()).strftime('%Y-%m-%d') if isinstance(max_trade.get('date'), datetime) else str(max_trade.get('date')),
                    'max_price_floor': max_trade.get('floor',''),
                    'max_price_dong':  max_trade.get('dong','-'),
                    'last_update': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'trade_data': trades
                })
                
                self.monitored_apts.append(apt_info)
                added += 1
                added_list.append({'apt_name': apt_name, 'area': area})

        self.save_monitored_apts()
        self.update_apt_tree()

        # ⭐ silent 모드가 아닐 때만 창 닫기 & 메시지 표시
        if not silent:
            try:
                win.destroy()
            except:
                pass
            messagebox.showinfo("완료", f"'{dong}' 단지 일괄 추가 완료\n추가: {added}개, 건너뜀: {skipped}개")

        # ⭐ silent 모드에서는 결과를 반환
        return {'added': added, 'skipped': skipped, 'added_list': added_list}



    
    def test_new_apt_api(self, sigungu_code, deal_ymd):
        """신축 API 응답 구조 테스트"""
        url = (f"https://apis.data.go.kr/1613000/RTMSDataSvcSilvTrade/getRTMSDataSvcSilvTrade"
               f"?serviceKey={self.service_key}"
               f"&LAWD_CD={sigungu_code}"
               f"&DEAL_YMD={deal_ymd}"
               f"&numOfRows=10")
        try:
            response = self.http.get(url, timeout=API_TIMEOUT)
            jitter_sleep()
            if response.status_code == 200:
                print(f"\n=== 신축 API 테스트 결과 ===")
                print(f"상태 코드: {response.status_code}")
                print(f"응답 샘플:\n{response.text[:1000]}")
                root = ET.fromstring(response.text)
                items = root.findall('.//item')
                if items:
                    print(f"\n총 {len(items)}개 항목 발견")
                    print("\n첫 번째 항목 필드:")
                    for child in items[0]:
                        print(f"  {child.tag}: {child.text}")
                else:
                    print("\n항목을 찾을 수 없습니다.")
            else:
                print(f"API 오류: {response.status_code}")
        except Exception as e:
            print(f"테스트 중 오류: {str(e)}")    


    def get_apt_list_from_api(self, sigungu_code, dong):
        """국토부 API에서 아파트 목록 가져오기 (기축 + 신축 통합)"""
        print(f"\n=== get_apt_list_from_api 호출 ===")
        print(f"시군구 코드: {sigungu_code}")
        print(f"동: {dong}")
        
        apt_info = {}
        current_date = datetime.now()
        
        # 최근 3개월만 검색
        for i in range(3):
            search_date = current_date - timedelta(days=30*i)
            deal_ymd = search_date.strftime("%Y%m")
            
            # 기축 아파트 API
            url_existing = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                           f"?serviceKey={self.service_key}"
                           f"&LAWD_CD={sigungu_code}"
                           f"&DEAL_YMD={deal_ymd}"
                           f"&numOfRows=1000")
            
            # 신축 아파트(분양권) API - http로 변경
            url_new = (f"http://apis.data.go.kr/1613000/RTMSDataSvcSilvTrade/getRTMSDataSvcSilvTrade"
                       f"?serviceKey={self.service_key}"
                       f"&LAWD_CD={sigungu_code}"
                       f"&DEAL_YMD={deal_ymd}"
                       f"&numOfRows=1000")
            print(f"\n월 {deal_ymd} API 호출:")
            print(f"URL: {url_existing}")
            for url, apt_type in [(url_existing, "기축"), (url_new, "신축")]:
                try:
                    response = self.http.get(url, timeout=API_TIMEOUT)
                    jitter_sleep(200)
                    print(f"{apt_type} API 응답 상태: {response.status_code}")
                    if response.status_code == 200:
                        root = ET.fromstring(response.text)
                        items = root.findall('.//item')
                        print(f"{apt_type} 조회된 전체 항목 수: {len(items)}")
                        print(f"\n{apt_type} API 응답 샘플 (처음 5개):")
                        for idx, item in enumerate(items[:5]):
                            item_dong = item.findtext('umdNm', '').strip()
                            apt_name = item.findtext('aptNm', '').strip()
                            print(f"  [{idx+1}] 동: {item_dong}, 아파트: {apt_name}")
                        print()
                        dong_count = 0
                        for item in items:
                            item_dong = item.findtext('umdNm', '').strip()
                            if item_dong == dong:
                                dong_count += 1
                                apt_name = item.findtext('aptNm', '').strip()
                                if dong_count <= 3:
                                    print(f"  - {apt_name} ({item_dong})")
                                if apt_name and apt_name not in apt_info:
                                    jibun = item.findtext('jibun', '').strip()
                                    jibun_addr = f"{dong} {jibun}"
                                    road = item.findtext('roadName', '').strip()
                                    road_main = item.findtext('roadNameBonbun', '').strip()
                                    road_sub = item.findtext('roadNameBubun', '').strip()
                                    if road:
                                        road_addr = f"{road} {road_main}"
                                        if road_sub:
                                            road_addr += f"-{road_sub}"
                                    else:
                                        road_addr = jibun_addr
                                    build_year = item.findtext('buildYear', '').strip()
                                    if apt_type == "신축" and not build_year:
                                        build_year = "분양"
                                    apt_info[apt_name] = {
                                        'jibun_addr': jibun_addr,
                                        'road_addr': road_addr,
                                        'build_year': build_year,
                                        'type': apt_type
                                    }
                        print(f"'{dong}'의 {apt_type} 거래 수: {dong_count}")
                except Exception as e:
                    logging.error(f"{apt_type} API 호출 중 오류: {str(e)}")
                    print(f"{apt_type} API 호출 오류: {str(e)}")
                    continue
        print(f"\n수집된 아파트 총 {len(apt_info)}개")
        apt_list = []
        new_apts = []
        for apt_name, info in sorted(apt_info.items()):
            if info['type'] == '신축':
                if info['build_year'] and info['build_year'] != "분양":
                    apt_str = f"[신축] {apt_name} [{info['road_addr']} / {info['jibun_addr']}] (준공: {info['build_year']}년)"
                else:
                    apt_str = f"[신축] {apt_name} [{info['road_addr']} / {info['jibun_addr']}] (분양중)"
                new_apts.append(apt_str)
        existing_apts = [f"{apt_name} [{info['road_addr']} / {info['jibun_addr']}] (준공: {info['build_year']}년)" 
                        for apt_name, info in sorted(apt_info.items()) 
                        if info['type'] == '기축' and info['build_year']]
        no_year_apts = [f"{apt_name} [{info['road_addr']} / {info['jibun_addr']}]" 
                       for apt_name, info in sorted(apt_info.items()) 
                       if info['type'] == '기축' and not info['build_year']]
        logging.info(f"검색 결과 - 신축: {len(new_apts)}개, 기축: {len(existing_apts)}개")
        print(f"최종 결과 - 신축: {len(new_apts)}개, 기축: {len(existing_apts)}개, 연도없음: {len(no_year_apts)}개")
        return new_apts + existing_apts + no_year_apts
    
    def add_apt_to_monitored(self, apt_info):
        """모니터링 아파트 목록에 아파트 추가"""
        for apt in self.monitored_apts:
            if apt['apt_name'] == apt_info['apt_name'] and apt['area'] == apt_info['area']:
                messagebox.showinfo("알림", "이미 모니터링 중인 아파트입니다.")
                return
        self.status_var.set(f"{apt_info['apt_name']} 데이터 수집 중...")
        self.root.update_idletasks()
        try:
            trade_data = self.collect_apt_data_with_cache(apt_info)  # ← 변경!
            if trade_data:
                max_price_trade = max(trade_data, key=lambda x: x['price'])
                max_date_str = max_price_trade['date'].strftime('%Y-%m-%d')
                apt_info['prev_max_price'] = 0
                apt_info['prev_max_date'] = ''
                apt_info['prev_max_floor'] = ''
                apt_info['prev_max_dong'] = ''
                apt_info['last_max_price'] = max_price_trade['price']
                apt_info['max_price_date'] = max_date_str
                apt_info['max_price_floor'] = max_price_trade.get('floor', '')
                apt_info['max_price_dong'] = max_price_trade.get('dong', '-')
                apt_info['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                apt_info['trade_data'] = trade_data
                self.monitored_apts.append(apt_info)
                self.save_monitored_apts()
                self.update_apt_tree()
                self.status_var.set(f"{apt_info['apt_name']} 모니터링 목록에 추가되었습니다.")
            else:
                messagebox.showinfo("알림", f"{apt_info['apt_name']}의 거래 내역이 없습니다.")
                self.status_var.set("거래 내역이 없습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 수집 중 오류: {str(e)}")
            self.status_var.set("오류 발생")
    
    def collect_apt_data(self, apt_info):
        """모든 월을 꼼꼼하게 수집 (기축 + 신축), 병렬"""
        apt_name = apt_info['apt_name']
        target_area = float(apt_info['area'])
        sigungu_code = apt_info['sigungu_code']
        dong = apt_info['dong']
        all_trades = []
        current_date = datetime.now()
        max_months = 240  # 최대 20년
        
        progress_window = tk.Toplevel(self.root)
        progress_window.title("데이터 수집 중...")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        
        ttk.Label(progress_window, text=f"{apt_name} ({target_area}㎡) 실거래 데이터를 수집 중입니다...", wraplength=350).pack(pady=10)
        progress_label = ttk.Label(progress_window, text="0% 완료")
        progress_label.pack(pady=5)
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=350, mode="determinate")
        progress_bar.pack(fill="x", padx=20, pady=10)
        
        cancel_flag = [False]
        def _cancel_collect():
            cancel_flag[0] = True
            try:
                progress_window.destroy()
            except:
                pass
        cancel_button = ttk.Button(progress_window, text="중단", command=_cancel_collect)
        cancel_button.pack(pady=5)
        progress_window.protocol("WM_DELETE_WINDOW", _cancel_collect)
        
        max_workers = min(10, (os.cpu_count() or 4) + 4)
        
        def collect_data():
            try:
                session = build_session()  # 재시도/백오프 적용 세션
                from threading import Lock
                results_lock = Lock()
                all_results = []
            
                def fetch_month_data(month_idx):
                    if cancel_flag[0]:
                        return None
                    search_date = current_date - timedelta(days=30 * month_idx)
                    deal_ymd = search_date.strftime("%Y%m")
                    url_existing = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                                   f"?serviceKey={self.service_key}"
                                   f"&LAWD_CD={sigungu_code}"
                                   f"&DEAL_YMD={deal_ymd}"
                                   f"&numOfRows=1000")
                    url_new = (f"https://apis.data.go.kr/1613000/RTMSDataSvcSilvTrade/getRTMSDataSvcSilvTrade"
                               f"?serviceKey={self.service_key}"
                               f"&LAWD_CD={sigungu_code}"
                               f"&DEAL_YMD={deal_ymd}"
                               f"&numOfRows=1000")
                    month_trades = []
                    for url, api_type in [(url_existing, "existing"), (url_new, "new")]:
                        try:
                            response = session.get(url, timeout=API_TIMEOUT)
                            jitter_sleep(120)
                            if response.status_code != 200:
                                logging.error(f"API 응답 오류: {response.status_code} (월: {deal_ymd})")
                                continue
                            root = ET.fromstring(response.text)
                            items = root.findall('.//item')
                            for item in items:
                                item_apt = item.findtext('aptNm', '').strip()
                                item_dong = item.findtext('umdNm', '').strip()
                                if item_apt == apt_name and item_dong == dong:
                                    area = float(item.findtext('excluUseAr', '0'))
                                    if abs(area - target_area) <= 1:
                                        try:
                                            building_dong = item.findtext('kaptdong', '')
                                            if not building_dong:
                                                jibun = item.findtext('jibun', '').strip()
                                                if jibun and '동' in jibun:
                                                    dong_parts = jibun.split('동')
                                                    if len(dong_parts) > 0 and dong_parts[0].isdigit():
                                                        building_dong = dong_parts[0] + '동'
                                            building_dong = building_dong or '-'
                                            trade = {
                                                'date': datetime(
                                                    int(item.findtext('dealYear')),
                                                    int(item.findtext('dealMonth')),
                                                    int(item.findtext('dealDay', '1'))
                                                ),
                                                'price': int(item.findtext('dealAmount').replace(',', '')),
                                                'floor': int(item.findtext('floor', '0')),
                                                'area': area,
                                                'dong': building_dong,
                                                'type': api_type
                                            }
                                            month_trades.append(trade)
                                        except (ValueError, TypeError) as e:
                                            logging.error(f"데이터 처리 오류: {str(e)}")
                                            continue
                        except Exception as e:
                            logging.error(f"월 데이터 조회 중 오류 (월: {deal_ymd}): {str(e)}")
                    with results_lock:
                        all_results.append({'month_idx': month_idx, 'trades': month_trades})
                    return month_trades
                
                def update_progress():
                    completed = len(all_results)
                    progress = min(100, (completed / max_months) * 100)
                    total_trades = sum(len(result['trades']) for result in all_results)
                    progress_bar['value'] = progress
                    progress_label.config(text=f"{progress:.1f}% 완료 - {total_trades}건 수집됨 ({completed}/{max_months}개월)")
                    progress_window.update_idletasks()
                
                batch_size = 12
                for batch_start in range(0, max_months, batch_size):
                    if cancel_flag[0]:
                        break
                    batch_end = min(batch_start + batch_size, max_months)
                    current_batch = list(range(batch_start, batch_end))
                    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                        futures = {executor.submit(fetch_month_data, month_idx): month_idx for month_idx in current_batch}
                        for future in concurrent.futures.as_completed(futures):
                            month_idx = futures[future]
                            try:
                                _ = future.result()
                                update_progress()
                            except Exception as e:
                                logging.error(f"월 {month_idx} 처리 중 오류: {str(e)}")
                    update_progress()
                    if cancel_flag[0]:
                        break
                    jitter_sleep(300)
                
                all_results.sort(key=lambda x: x['month_idx'])
                for result in all_results:
                    all_trades.extend(result['trades'])
                all_trades.sort(key=lambda x: x['date'])
                progress_bar['value'] = 100
                progress_label.config(text=f"100% 완료 - 총 {len(all_trades)}건 수집됨")
                progress_window.update_idletasks()
                progress_window.after(500, progress_window.destroy)
                    
            except Exception as e:
                messagebox.showerror("오류", f"데이터 수집 중 오류 발생: {str(e)}")
                try:
                    progress_window.destroy()
                except:
                    pass
        
        thread = threading.Thread(target=collect_data, daemon=True)
        thread.start()
        self.root.wait_window(progress_window)
        if cancel_flag[0]:
            return []
        if not all_trades:
            return []
        return sorted(all_trades, key=lambda x: x['date'])


    def collect_apt_data_with_cache(self, apt_info):
        """캐싱을 활용한 아파트 데이터 수집"""
        apt_name = apt_info['apt_name']
        target_area = float(apt_info['area'])
        sigungu_code = apt_info['sigungu_code']
        dong = apt_info['dong']
        all_trades = []
        current_date = datetime.now()
        max_months = 240
        
        # Progress 창
        progress_window = tk.Toplevel(self.root)
        progress_window.title("데이터 수집 중...")
        progress_window.geometry("500x200")
        progress_window.transient(self.root)
        
        ttk.Label(progress_window, text=f"{apt_name} ({target_area}㎡) 실거래 데이터를 수집 중입니다...", 
                 wraplength=450).pack(pady=10)
        progress_label = ttk.Label(progress_window, text="0% 완료")
        progress_label.pack(pady=5)
        
        # ★★★ 캐시 통계 라벨 추가 ★★★
        stats_label = ttk.Label(progress_window, text="API 호출: 0 / 캐시 히트: 0")
        stats_label.pack(pady=5)
        
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", 
                                      length=450, mode="determinate")
        progress_bar.pack(fill="x", padx=20, pady=10)
        
        cancel_flag = [False]
        def _cancel_collect():
            cancel_flag[0] = True
            try:
                progress_window.destroy()
            except:
                pass
        
        cancel_button = ttk.Button(progress_window, text="중단", command=_cancel_collect)
        cancel_button.pack(pady=5)
        progress_window.protocol("WM_DELETE_WINDOW", _cancel_collect)
        
        # 통계 저장
        initial_api_calls = self.api_call_count
        initial_cache_hits = self.cache_hit_count
        
        def collect_data():
            try:
                for month_idx in range(max_months):
                    if cancel_flag[0]:
                        break
                    
                    search_date = current_date - timedelta(days=30 * month_idx)
                    deal_ymd = search_date.strftime("%Y%m")
                    
                    # ★★★ 캐시 활용 ★★★
                    existing_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'existing')
                    filtered_existing = self.filter_apt_data(existing_data, apt_name, dong, target_area)
                    all_trades.extend(filtered_existing)
                    
                    new_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'new')
                    filtered_new = self.filter_apt_data(new_data, apt_name, dong, target_area)
                    all_trades.extend(filtered_new)
                    
                    # 진행률 업데이트
                    progress = min(100, ((month_idx + 1) / max_months) * 100)
                    progress_bar['value'] = progress
                    progress_label.config(text=f"{progress:.1f}% 완료 - {len(all_trades)}건 수집됨")
                    
                    # ★★★ 통계 업데이트 ★★★
                    api_calls = self.api_call_count - initial_api_calls
                    cache_hits = self.cache_hit_count - initial_cache_hits
                    stats_label.config(text=f"API 호출: {api_calls} / 캐시 히트: {cache_hits}")
                    progress_window.update_idletasks()
                    
                    # 최적화: 6개월간 거래 없으면 종료
                    if len(all_trades) > 0 and month_idx > 6:
                        recent_trades = [t for t in all_trades 
                                       if t['date'] > current_date - timedelta(days=180)]
                        if len(recent_trades) == 0:
                            break
                
                all_trades.sort(key=lambda x: x['date'])
                progress_bar['value'] = 100
                progress_label.config(text=f"100% 완료 - 총 {len(all_trades)}건 수집됨")
                
                # 최종 통계
                final_api_calls = self.api_call_count - initial_api_calls
                final_cache_hits = self.cache_hit_count - initial_cache_hits
                stats_label.config(text=f"완료! API 호출: {final_api_calls} / 캐시 히트: {final_cache_hits}")
                
                progress_window.update_idletasks()
                progress_window.after(500, progress_window.destroy)
                    
            except Exception as e:
                messagebox.showerror("오류", f"데이터 수집 중 오류 발생: {str(e)}")
                try:
                    progress_window.destroy()
                except:
                    pass
        
        thread = threading.Thread(target=collect_data, daemon=True)
        thread.start()
        self.root.wait_window(progress_window)
        
        if cancel_flag[0]:
            return []
        return all_trades
    
    def query_data(self, df, query_string):
        """조건에 맞는 데이터 추출"""
        try:
            filtered_df = df.query(query_string)
            return filtered_df
        except Exception as e:
            logging.error(f"데이터 쿼리 중 오류 발생: {str(e)}")
            return df

    # (주의) history 관련 함수들은 미사용/의존성 문제 가능. 필요한 경우 openpyxl 임포트/구현 보완.
    def load_history(self):
        history_list = []
        print(f"\n=== 히스토리 로드 시작 ===")
        print(f"히스토리 경로: {getattr(self, 'history_path', '(미설정)')}")
        if hasattr(self, 'history_path') and os.path.exists(self.history_path):
            try:
                all_files = os.listdir(self.history_path)
                print(f"히스토리 폴더 내 총 {len(all_files)}개 파일 발견")
                for file in all_files:
                    file_path = os.path.join(self.history_path, file)
                    if not os.path.exists(file_path):
                        print(f"파일 없음, 건너뜀: {file_path}")
                        continue
                    try:
                        if file.startswith('history_compare_'):
                            # 비교 파일 메타만 수집 (openpyxl 의존 제거)
                            history_list.append({
                                'file_path': file_path,
                                'apt_name': f"[비교] {file}",
                                'area': "비교분석",
                                'search_date': os.path.getmtime(file_path),
                                'max_trade': "비교분석",
                                'type': 'compare'
                            })
                        elif file.startswith('history_'):
                            # 단일 분석 파일 (메타만)
                            history_list.append({
                                'file_path': file_path,
                                'apt_name': file,
                                'area': "",
                                'search_date': os.path.getmtime(file_path),
                                'max_trade': "정보없음",
                                'type': 'single'
                            })
                    except Exception as e:
                        print(f"파일 처리 중 오류 ({file}): {str(e)}")
                        continue
                print(f"히스토리 로드 완료: {len(history_list)}개 항목")
            except Exception as e:
                print(f"히스토리 로드 중 오류: {str(e)}")
        else:
            print(f"히스토리 폴더가 없습니다: {getattr(self, 'history_path', '(미설정)')}")
        return sorted(history_list, key=lambda x: x['search_date'], reverse=True)

    def update_history_display(self):
        if not hasattr(self, 'history_tree'):
            return
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        print(f"히스토리 목록 업데이트: {len(getattr(self, 'history_list', []))}개 항목")
        sorted_history = sorted(getattr(self, 'history_list', []), key=lambda x: x['search_date'], reverse=True)
        for item in sorted_history:
            search_date = datetime.fromtimestamp(item['search_date'])
            self.history_tree.insert("", "end", values=(
                search_date.strftime("%Y-%m-%d %H:%M"),
                item['apt_name'],
                item['area'],
                item['max_trade']
            ))
        if not getattr(self, 'history_list', []):
            print("히스토리 목록이 비어있습니다.")
        else:
            print(f("히스토리 표시 완료: {len(self.history_list)}개 항목"))
        self.root.update_idletasks()

    def delete_selected_history(self):
        if not hasattr(self, 'history_tree') or not hasattr(self, 'history_list'):
            messagebox.showinfo("알림", "히스토리 UI가 활성화되어 있지 않습니다.")
            return
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showinfo("알림", "삭제할 항목을 선택해주세요.")
            return
        if not messagebox.askyesno("확인", "선택한 히스토리를 삭제하시겠습니까?\n(엑셀 파일과 그래프 이미지가 모두 삭제됩니다)"):
            return
        try:
            deleted_indices = []
            for item in selection:
                idx = self.history_tree.index(item)
                deleted_indices.append(idx)
                if idx < len(self.history_list):
                    history_item = self.history_list[idx]
                    file_path = history_item['file_path']
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            print(f"파일 삭제 실패: {file_path} - {str(e)}")
            deleted_indices.sort(reverse=True)
            for idx in deleted_indices:
                if idx < len(self.history_list):
                    del self.history_list[idx]
            for item in selection:
                self.history_tree.delete(item)
            self.update_history_display()
            messagebox.showinfo("알림", "선택한 히스토리가 삭제되었습니다.")
        except Exception as e:
            print(f"삭제 중 오류: {str(e)}")
            messagebox.showerror("오류", f"히스토리 삭제 중 오류가 발생했습니다: {str(e)}")
    
    def delete_selected_apt(self):
        """선택된 모니터링 아파트 삭제"""
        selection = self.apt_tree.selection()
        if not selection:
            messagebox.showinfo("알림", "삭제할 아파트를 선택해주세요.")
            return
        if messagebox.askyesno("확인", "선택한 아파트를 모니터링 목록에서 삭제하시겠습니까?"):
            deleted_count = 0
            for item_id in selection:
                item_values = self.apt_tree.item(item_id, "values")
                apt_name = item_values[0]
                area_str = item_values[2]
                import re
                area_value = re.search(r'(\d+(?:\.\d+)?)', area_str)
                if area_value:
                    area = area_value.group(1)
                else:
                    area = area_str.replace('㎡', '').strip()
                to_delete = []
                for idx, apt in enumerate(self.monitored_apts):
                    if apt.get('apt_name', '') == apt_name:
                        apt_area = str(apt.get('area', ''))
                        apt_area_clean = apt_area.replace('㎡', '').strip()
                        area_clean = area.replace('㎡', '').strip()
                        try:
                            if float(apt_area_clean) == float(area_clean):
                                to_delete.append(idx)
                        except ValueError:
                            if apt_area_clean == area_clean:
                                to_delete.append(idx)
                for idx in sorted(to_delete, reverse=True):
                    if 0 <= idx < len(self.monitored_apts):
                        del self.monitored_apts[idx]
                        deleted_count += 1
            if deleted_count > 0:
                self.save_monitored_apts()
                self.update_apt_tree()
                self.status_var.set(f"{deleted_count}개 아파트가 모니터링 목록에서 삭제되었습니다.")
            else:
                self.status_var.set("삭제할 항목을 찾을 수 없습니다.")
    
    def clear_all_apts(self):
        """모든 모니터링 아파트 삭제"""
        if not self.monitored_apts:
            messagebox.showinfo("알림", "모니터링 목록이 비어있습니다.")
            return
        if messagebox.askyesno("확인", "모니터링 목록을 모두 삭제하시겠습니까?"):
            self.monitored_apts = []
            self.save_monitored_apts()
            self.update_apt_tree()
            self.status_var.set("모니터링 목록이 모두 삭제되었습니다.")
    
    def recheck_max_prices(self):
        """모니터링 단지의 신고가 재검증 + 새로운 단지 자동 추가 + 연식 보정"""
        if not self.monitored_apts:
            # 모니터링 중인 단지가 없으면 전체 시군구 스캔
            self.scan_and_add_new_complexes()
            return
        
        # 이미 검증된 단지 추적용
        if not hasattr(self, 'rechecked_apts'):
            self.rechecked_apts = set()
        
        # 진행 창
        progress_window = tk.Toplevel(self.root)
        progress_window.title("신고가 재검증 및 신규 단지 탐색 중...")
        progress_window.geometry("600x300")
        progress_window.transient(self.root)
        
        ttk.Label(progress_window, text="신고가 재검증, 신규 단지 탐색, 연식 정보 보정 중...", 
                 wraplength=550).pack(pady=10)
        
        progress_label = ttk.Label(progress_window, text="준비 중...")
        progress_label.pack(pady=5)
        
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", 
                                      length=550, mode="determinate")
        progress_bar.pack(fill="x", padx=20, pady=10)
        
        result_label = ttk.Label(progress_window, text="")
        result_label.pack(pady=5)
        
        debug_text = tk.Text(progress_window, height=8, width=70)
        debug_text.pack(pady=5, padx=10)
        
        cancel_flag = [False]
        def _cancel():
            cancel_flag[0] = True
            try:
                progress_window.destroy()
            except:
                pass
        
        cancel_button = ttk.Button(progress_window, text="중단", command=_cancel)
        cancel_button.pack(pady=5)
        progress_window.protocol("WM_DELETE_WINDOW", _cancel)
        
        def process_recheck():
            try:
                # ========== 1단계: 기존 단지의 연식 보정 ==========
                debug_text.insert('end', "=== 1단계: 기존 단지 연식 보정 ===\n")
                debug_text.see('end')
                progress_window.update_idletasks()
                
                build_year_updated = 0
                for idx, apt in enumerate(self.monitored_apts):
                    if cancel_flag[0]:
                        break
                    
                    # 진행률 업데이트 (1단계는 전체의 30%)
                    progress = (idx / len(self.monitored_apts)) * 30
                    progress_bar['value'] = progress
                    progress_label.config(text=f"연식 보정 중: {idx + 1}/{len(self.monitored_apts)}")
                    progress_window.update_idletasks()
                    
                    # 연식이 비어있거나 유효하지 않은 경우
                    current_build_year = str(apt.get('build_year', '')).strip()
                    if not current_build_year or current_build_year == '':
                        debug_text.insert('end', f"  {apt.get('apt_name', '')} - 연식 정보 없음, 보정 시도...\n")
                        debug_text.see('end')
                        
                        # 연식 정보 보정 시도
                        resolved_year = self.resolve_build_year(
                            sigungu_code=apt.get('sigungu_code', ''),
                            dong=apt.get('dong', ''),
                            apt_name=apt.get('apt_name', ''),
                            months=24,  # 최근 24개월 데이터 검색
                            fallback_text=None
                        )
                        
                        if resolved_year:
                            apt['build_year'] = resolved_year
                            build_year_updated += 1
                            debug_text.insert('end', f"    ✓ 연식 보정 완료: {resolved_year}\n")
                            debug_text.see('end')
                        else:
                            debug_text.insert('end', f"    ✗ 연식 정보를 찾을 수 없음\n")
                            debug_text.see('end')
                
                debug_text.insert('end', f"\n연식 보정 완료: {build_year_updated}개 단지\n\n")
                debug_text.see('end')
                
                # ========== 2단계: 기존 단지들의 시군구 코드 수집 ==========
                debug_text.insert('end', "=== 2단계: 시군구 코드 수집 ===\n")
                sigungu_codes = set()
                for apt in self.monitored_apts:
                    if 'sigungu_code' in apt:
                        sigungu_codes.add(apt['sigungu_code'])
                
                debug_text.insert('end', f"모니터링 중인 시군구 코드: {sigungu_codes}\n\n")
                debug_text.see('end')
                progress_window.update_idletasks()
                
                # ========== 3단계: 각 시군구의 최근 거래 데이터에서 새로운 단지 탐색 ==========
                debug_text.insert('end', "=== 3단계: 신규 단지 탐색 및 신고가 검증 ===\n")
                new_complexes = []
                total_checked = 0
                corrected_count = 0
                skipped_count = 0  # ★ 스킵된 단지 카운트

                current_date = datetime.now()

                start_year = 2019
                months_to_check = (current_date.year - start_year) * 12 + current_date.month

                # ★ 기존 단지 정보를 딕셔너리로 미리 구성 (검색 성능 향상)
                existing_apts_dict = {}
                for apt in self.monitored_apts:
                    apt_key = f"{apt.get('apt_name')}_{apt.get('dong')}_{str(apt.get('area', '')).replace('㎡', '').strip()}"
                    existing_apts_dict[apt_key] = apt

                for sigungu_idx, sigungu_code in enumerate(sigungu_codes):
                    if cancel_flag[0]:
                        break

                    debug_text.insert('end', f"\n시군구 {sigungu_code} 스캔 중...\n")
                    debug_text.see('end')

                    # 해당 시군구의 모든 거래 데이터 수집
                    all_apt_info = {}

                    for month_idx in range(months_to_check):
                        if cancel_flag[0]:
                            break

                        search_date = current_date - timedelta(days=30 * month_idx)
                        deal_ymd = search_date.strftime("%Y%m")

                        # ★ GUI 업데이트 빈도 감소: 10개월마다 한 번만 업데이트
                        if month_idx % 10 == 0:
                            progress = 30 + ((sigungu_idx * months_to_check + month_idx) / (len(sigungu_codes) * months_to_check)) * 60
                            progress_bar['value'] = min(progress, 90)
                            progress_label.config(text=f"데이터 수집 중: {sigungu_idx+1}/{len(sigungu_codes)} 시군구, {month_idx}/{months_to_check} 개월")
                            progress_window.update_idletasks()

                        # 캐시 활용하여 데이터 가져오기
                        existing_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'existing')
                        new_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'new')
                        
                        # 각 데이터에서 단지 정보 추출
                        for item in existing_data + new_data:
                            apt_name = item.get('apt_name', '').strip()
                            dong = item.get('dong', '').strip()
                            area = item.get('area', 0)
                            
                            if not apt_name or not dong:
                                continue
                            
                            # 고유 키 생성
                            apt_key = f"{apt_name}_{dong}_{int(area)}"
                            
                            if apt_key not in all_apt_info:
                                all_apt_info[apt_key] = {
                                    'apt_name': apt_name,
                                    'dong': dong,
                                    'area': str(int(area)),
                                    'sigungu_code': sigungu_code,
                                    'trades': [],
                                    'build_year': item.get('build_year', '')  # ★ API에서 연식 정보도 수집
                                }
                            
                            # 거래 정보 추가
                            try:
                                trade = {
                                    'date': datetime(item['year'], item['month'], item.get('day', 1)),
                                    'price': item['price'],
                                    'floor': item.get('floor', 0),
                                    'building_dong': item.get('kaptdong', '-')
                                }
                                all_apt_info[apt_key]['trades'].append(trade)
                            except:
                                continue
                    
                    debug_text.insert('end', f"  발견된 단지: {len(all_apt_info)}개\n")

                    # 3단계: 기존 모니터링 목록과 비교
                    for apt_key, apt_data in all_apt_info.items():
                        if cancel_flag[0]:
                            break

                        apt_name = apt_data['apt_name']
                        dong = apt_data['dong']
                        area = apt_data['area']

                        # ★ 딕셔너리로 빠르게 기존 단지 확인
                        lookup_key = f"{apt_name}_{dong}_{area}"
                        existing_apt = existing_apts_dict.get(lookup_key)
                        exists = existing_apt is not None

                        # ★ 이미 검증된 단지는 스킵
                        if exists and lookup_key in self.rechecked_apts:
                            skipped_count += 1
                            continue

                        if not exists and apt_data['trades']:
                            # 새로운 단지 발견
                            max_trade = max(apt_data['trades'], key=lambda x: x['price'])
                            
                            # sido, sigungu 정보 찾기
                            sido = ''
                            sigungu = ''
                            for apt in self.monitored_apts:
                                if apt.get('sigungu_code') == sigungu_code:
                                    sido = apt.get('sido', '')
                                    sigungu = apt.get('sigungu', '')
                                    break
                            
                            # ★ API에서 받은 연식 정보 활용 (추가 조회 최소화)
                            build_year = apt_data.get('build_year', '').strip()
                            
                            # ★★★ 신규 단지: prev와 last를 모두 찾은 최고가로 설정 ★★★
                            new_apt = {
                                'apt_name': apt_name,
                                'area': area,
                                'sido': sido,
                                'sigungu': sigungu,
                                'dong': dong,
                                'sigungu_code': sigungu_code,
                                'prev_max_price': max_trade['price'],
                                'prev_max_date': max_trade['date'].strftime('%Y-%m-%d'),
                                'prev_max_floor': max_trade.get('floor', ''),
                                'prev_max_dong': max_trade.get('building_dong', '-'),
                                'last_max_price': max_trade['price'],
                                'max_price_date': max_trade['date'].strftime('%Y-%m-%d'),
                                'max_price_floor': max_trade.get('floor', ''),
                                'max_price_dong': max_trade.get('building_dong', '-'),
                                'last_update': datetime.now().strftime('%Y-%m-%d %H:%M'),
                                'trade_data': apt_data['trades'],
                                'build_year': build_year  # ★ 연식 정보 포함
                            }
                            
                            new_complexes.append(new_apt)
                            year_info = f" (연식: {build_year})" if build_year else " (연식 정보 없음)"
                            debug_text.insert('end', f"  [신규] {apt_name} {area}㎡{year_info} - 최고가: {max_trade['price']:,}만원\n")
                            
                        elif exists and existing_apt and apt_data['trades']:
                            # 기존 단지의 신고가 체크
                            max_trade = max(apt_data['trades'], key=lambda x: x['price'])
                            if max_trade['price'] > existing_apt.get('last_max_price', 0):
                                # ★★★ 신고가 다시 찾기: 역대 최고가를 찾았으므로 prev와 last 모두 같은 값으로 설정 ★★★
                                # 이렇게 해야 '데이터 갱신' 시 정확한 신고가 비교 가능
                                existing_apt['prev_max_price'] = max_trade['price']
                                existing_apt['prev_max_date'] = max_trade['date'].strftime('%Y-%m-%d')
                                existing_apt['prev_max_floor'] = max_trade.get('floor', '')
                                existing_apt['prev_max_dong'] = max_trade.get('building_dong', '-')

                                existing_apt['last_max_price'] = max_trade['price']
                                existing_apt['max_price_date'] = max_trade['date'].strftime('%Y-%m-%d')
                                existing_apt['max_price_floor'] = max_trade.get('floor', '')
                                existing_apt['max_price_dong'] = max_trade.get('building_dong', '-')
                                existing_apt['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M')

                                corrected_count += 1
                                debug_text.insert('end', f"  [신고가] {apt_name} {area}㎡ - {max_trade['price']:,}만원\n")

                            # ★ 검증 완료된 단지는 rechecked_apts에 추가 (다음 실행 시 스킵)
                            self.rechecked_apts.add(lookup_key)

                        total_checked += 1
                        # ★ GUI 업데이트 빈도 감소: 배치 처리
                        if total_checked % 20 == 0:
                            progress = 30 + ((sigungu_idx + (total_checked / len(all_apt_info))) / len(sigungu_codes)) * 60
                            progress_bar['value'] = min(progress, 90)
                            result_label.config(text=f"연식 보정: {build_year_updated}개 / 신규: {len(new_complexes)}개 / 신고가: {corrected_count}개 / 스킵: {skipped_count}개")
                            progress_window.update_idletasks()
                    
                    debug_text.see('end')
                
                # 4단계: 새로운 단지들을 모니터링 목록에 추가
                if new_complexes:
                    self.monitored_apts.extend(new_complexes)
                    debug_text.insert('end', f"\n총 {len(new_complexes)}개 신규 단지 추가됨\n")

                # ★ 중복 아파트 병합 (분양권+준공후 거래 통합)
                before_count = len(self.monitored_apts)
                self.monitored_apts = self.merge_duplicate_apts(self.monitored_apts)
                # monitored_lists에도 반영
                self.monitored_lists["lists"][self.active_list.get()] = self.monitored_apts
                after_count = len(self.monitored_apts)
                if before_count > after_count:
                    debug_text.insert('end', f"중복 병합: {before_count}개 -> {after_count}개 ({before_count - after_count}개 통합)\n")

                # ★ 5단계: 연식 없는 단지들만 배치로 연식 조회
                debug_text.insert('end', "\n=== 4단계: 연식 정보 보정 ===\n")
                debug_text.see('end')
                progress_window.update_idletasks()

                apts_without_year = []
                for apt in self.monitored_apts:
                    build_year = str(apt.get('build_year', '')).strip()
                    if not build_year or build_year == '':
                        apts_without_year.append(apt)

                year_resolved_count = 0
                if apts_without_year:
                    debug_text.insert('end', f"연식 정보가 없는 단지: {len(apts_without_year)}개\n")
                    debug_text.see('end')

                    for idx, apt in enumerate(apts_without_year):
                        if cancel_flag[0]:
                            break

                        if idx % 5 == 0:  # 5개마다 GUI 업데이트
                            progress = 90 + (idx / len(apts_without_year)) * 10
                            progress_bar['value'] = min(progress, 100)
                            progress_label.config(text=f"연식 조회 중: {idx}/{len(apts_without_year)}")
                            progress_window.update_idletasks()

                        resolved_year = self.resolve_build_year(
                            sigungu_code=apt.get('sigungu_code', ''),
                            dong=apt.get('dong', ''),
                            apt_name=apt.get('apt_name', ''),
                            months=24,
                            fallback_text=None
                        )

                        if resolved_year:
                            apt['build_year'] = resolved_year
                            year_resolved_count += 1
                            debug_text.insert('end', f"  ✓ {apt.get('apt_name', '')} - {resolved_year}\n")
                            debug_text.see('end')

                    debug_text.insert('end', f"\n연식 보정 완료: {year_resolved_count}/{len(apts_without_year)}개\n")
                    debug_text.see('end')

                # 저장 및 화면 갱신
                self.save_monitored_apts()
                self.update_apt_tree()

                progress_bar['value'] = 100
                progress_label.config(text="완료!")

                messagebox.showinfo("완료",
                    f"신고가 재검증이 완료되었습니다.\n"
                    f"기존 단지 검증: {total_checked - skipped_count}개\n"
                    f"검증 스킵: {skipped_count}개\n"
                    f"신규 단지: {len(new_complexes)}개 추가\n"
                    f"신고가 갱신: {corrected_count}개\n"
                    f"연식 보정: {build_year_updated + year_resolved_count}개")

                progress_window.destroy()
                
            except Exception as e:
                messagebox.showerror("오류", f"재검증 중 오류 발생: {str(e)}")
                debug_text.insert('end', f"\n오류: {str(e)}\n")
                try:
                    progress_window.destroy()
                except:
                    pass
        
        thread = threading.Thread(target=process_recheck, daemon=True)
        thread.start()
    
    def update_all_data(self):
        """모든 모니터링 아파트 데이터 갱신"""
        if not self.monitored_apts:
            messagebox.showinfo("알림", "모니터링 목록이 비어있습니다.")
            return

        is_auto_update = threading.current_thread() != threading.main_thread()

        # ★★★ 진행 바 팝업 생성 (수동 갱신인 경우만) ★★★
        progress_window = None
        progress_bar = None
        progress_label = None

        if not is_auto_update:
            progress_window = tk.Toplevel(self.root)
            progress_window.title("데이터 갱신 중")
            progress_window.geometry("500x150")
            progress_window.transient(self.root)

            # 화면 중앙에 배치
            progress_window.update_idletasks()
            screen_width = progress_window.winfo_screenwidth()
            screen_height = progress_window.winfo_screenheight()
            x = (screen_width - 500) // 2
            y = (screen_height - 150) // 2
            progress_window.geometry(f"500x150+{x}+{y}")

            ttk.Label(progress_window, text="전체 매물 데이터 갱신 중...",
                     font=("", 12, "bold")).pack(pady=(20, 10))

            progress_label = ttk.Label(progress_window, text="준비 중...")
            progress_label.pack(pady=5)

            progress_bar = ttk.Progressbar(progress_window, orient="horizontal",
                                          length=450, mode="determinate", maximum=100)
            progress_bar.pack(fill="x", padx=25, pady=10)

            self.status_var.set("데이터 갱신 중...")
            self.root.update_idletasks()

        new_max_found = False
        apt_with_new_max = []

        total_apts = len(self.monitored_apts)

        for i, apt_info in enumerate(self.monitored_apts):
            try:
                if not is_auto_update and progress_window:
                    # 진행률 계산
                    progress = (i / total_apts) * 100
                    progress_bar['value'] = progress
                    progress_label.config(text=f"갱신 중: {i+1}/{total_apts} - {apt_info['apt_name']}")
                    progress_window.update_idletasks()

                    self.status_var.set(f"({i+1}/{total_apts}) {apt_info['apt_name']} 데이터 갱신 중...")
                    self.root.update_idletasks()
                
                old_max_price = apt_info.get('last_max_price', 0)
                old_max_date = apt_info.get('max_price_date', '')
                old_max_dong = apt_info.get('max_price_dong', '')
                old_max_floor = apt_info.get('max_price_floor', '')
                
                trade_data = self.collect_apt_data_silent_with_cache(apt_info)
                
                if trade_data:
                    new_max_price_trade = max(trade_data, key=lambda x: x['price'])
                    
                    # date_str을 먼저 정의
                    date_str = ''
                    if isinstance(new_max_price_trade.get('date'), datetime):
                        date_str = new_max_price_trade['date'].strftime('%Y-%m-%d')
                    else:
                        date_str = str(new_max_price_trade.get('date', ''))
                    
                    if new_max_price_trade['price'] > old_max_price:
                        new_max_found = True

                        # 신고가 정보 딕셔너리 생성 - build_year 확실히 포함
                        # 신고가 정보 딕셔너리 생성 - build_year 확실히 포함
                        build_year = str(apt_info.get('build_year', ''))
                        current_year = datetime.now().year
                        is_young = False

                        # 10년 이하 판별
                        if build_year == '분양':
                            is_young = True
                        elif build_year and build_year.isdigit():
                            if current_year - int(build_year) <= 10:
                                is_young = True


                        new_max_info = {
                            'apt_name': apt_info['apt_name'],
                            'area': apt_info['area'],
                            'old_price': old_max_price,
                            'old_date': old_max_date,              # ✨ 추가
                            'old_floor': old_max_floor,            # ✨ 추가
                            'old_dong': old_max_dong,              # ✨ 추가
                            'new_price': new_max_price_trade['price'],
                            'date': date_str,
                            'floor': new_max_price_trade.get('floor', ''),
                            'dong': new_max_price_trade.get('dong', '-'),
                            'sido': apt_info.get('sido', ''),
                            'sigungu': apt_info.get('sigungu', ''),
                            'location_dong': str(apt_info.get('dong', '')) if not isinstance(apt_info.get('dong', ''), dict) else '',
                            'build_year': build_year,
                            'is_young': is_young
                        }

                        # 디버그: location_dong 값 확인
                        if '성남' in apt_info.get('sigungu', ''):
                            logging.info(f"[디버그 location_dong] {apt_info['apt_name']}: sido={apt_info.get('sido')}, sigungu={apt_info.get('sigungu')}, dong={apt_info.get('dong')}")


                        # ✨ 디버그 출력
                        print(f"[디버그] {apt_info['apt_name']}")
                        print(f"  이전 최고가: {old_max_price:,}만원")
                        print(f"  이전 날짜: {old_max_date}")
                        print(f"  이전 층: {old_max_floor}")
                        print(f"  이전 동: {old_max_dong}")


                        apt_with_new_max.append(new_max_info)

                        # 디버깅을 위한 출력
                        logging.info(f"신고가 발견: {apt_info['apt_name']}, 연식: {apt_info.get('build_year', '없음')}")
                        print(f"신고가 리스트 추가: {new_max_info['apt_name']}, 연식: {new_max_info['build_year']}")

                        # ★★★ 신고가 갱신: prev 필드에 기존 max 정보 저장, last 필드에 새로운 max 정보 저장 ★★★
                        apt_info['prev_max_price'] = old_max_price
                        apt_info['prev_max_date'] = old_max_date
                        apt_info['prev_max_dong'] = old_max_dong
                        apt_info['prev_max_floor'] = old_max_floor

                        apt_info['last_max_price'] = new_max_price_trade['price']
                        apt_info['max_price_date'] = date_str
                        apt_info['max_price_floor'] = new_max_price_trade.get('floor', '')
                        apt_info['max_price_dong'] = new_max_price_trade.get('dong', '-')
                    else:
                        # ★★★ 신고가가 아닌 경우: last_max_price가 여전히 최고가이므로 유지 ★★★
                        # 단, 기존 last_max_price보다 높은 거래가 있다면 그것으로 갱신
                        if new_max_price_trade['price'] > apt_info.get('last_max_price', 0):
                            # 이전보다는 낮지만 현재 저장된 max보다는 높은 경우 (데이터 보정)
                            apt_info['last_max_price'] = new_max_price_trade['price']
                            apt_info['max_price_date'] = date_str
                            apt_info['max_price_floor'] = new_max_price_trade.get('floor', '')
                            apt_info['max_price_dong'] = new_max_price_trade.get('dong', '-')

                    # 모든 경우에 공통으로 업데이트
                    apt_info['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                    apt_info['trade_data'] = trade_data
                    
            except Exception as e:
                logging.error(f"{apt_info['apt_name']} 데이터 갱신 중 오류: {str(e)}")
                continue
        
        self.save_monitored_apts()
        self.update_apt_tree()
        
        if new_max_found:
            # 전체 모니터링 단지 기준 지역구별 순위 계산
            regional_rankings = self.calculate_regional_rankings()
            
            # 신고가 단지에 순위 정보 추가
            for apt in apt_with_new_max:
                sido = apt.get('sido', '')
                sigungu = apt.get('sigungu', '').split('(')[0] if apt.get('sigungu') else ''
                region_key = f"{sido} {sigungu}"
                apt_name = apt.get('apt_name', '')
                
                if region_key in regional_rankings:
                    # 84타입 순위
                    if apt_name in regional_rankings[region_key]['84']:
                        apt['region_84_rank'] = regional_rankings[region_key]['84'][apt_name]
                        apt['region_name'] = sigungu
                    
                    # 59타입 순위
                    if apt_name in regional_rankings[region_key]['59']:
                        apt['region_59_rank'] = regional_rankings[region_key]['59'][apt_name]
                        apt['region_name'] = sigungu
            
            # HTML 생성 전 연식 정보 최종 확인
            print("\n=== 신고가 발견 목록 최종 확인 ===")
            for apt in apt_with_new_max:
                rank_info = []
                if apt.get('region_84_rank'):
                    rank_info.append(f"84타입 {apt.get('region_84_rank')}위")
                if apt.get('region_59_rank'):
                    rank_info.append(f"59타입 {apt.get('region_59_rank')}위")
                rank_str = f", 지역순위={', '.join(rank_info)}" if rank_info else ""
                print(f"- {apt['apt_name']}: 연식={apt.get('build_year', '없음')}, 가격={apt['new_price']:,}만원{rank_str}")
            print("================================\n")
            
            self.show_new_max_notification(apt_with_new_max)
        
        # 평형별 순위 이력 저장
        self.save_area_ranking_history()

        # ★★★ 진행 바 팝업 닫기 ★★★
        if not is_auto_update and progress_window:
            try:
                progress_bar['value'] = 100
                progress_label.config(text=f"완료: {total_apts}/{total_apts} 매물 갱신 완료")
                progress_window.update_idletasks()
                self.root.after(500, progress_window.destroy)  # 0.5초 후 닫기
            except:
                pass

        update_time = datetime.now().strftime('%Y-%m-%d %H:%M')
        if is_auto_update:
            logging.info(f"자동 업데이트 완료: {update_time}")
        else:
            self.status_var.set(f"데이터 갱신 완료: {update_time}")
        
    def collect_apt_data_silent_with_cache(self, apt_info):
        """캐싱을 활용한 조용한 데이터 수집"""
        apt_name = apt_info['apt_name']
        target_area = float(apt_info['area'])
        sigungu_code = apt_info['sigungu_code']
        dong = apt_info['dong']
        trades = []
        current_date = datetime.now()
        max_months = 4
        
        try:
            for month in range(max_months):
                search_date = current_date - timedelta(days=30 * month)
                deal_ymd = search_date.strftime("%Y%m")
                
                # ★★★ 캐시 활용 ★★★
                existing_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'existing')
                new_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'new')
                
                trades.extend(self.filter_apt_data(existing_data, apt_name, dong, target_area))
                trades.extend(self.filter_apt_data(new_data, apt_name, dong, target_area))
            
            # 기존 데이터와 병합
            if 'trade_data' in apt_info and apt_info['trade_data']:
                existing_trades = apt_info['trade_data']
                existing_keys = set()

                for trade in existing_trades:
                    if 'date' in trade and isinstance(trade['date'], datetime):
                        # floor와 price가 dict가 아닌지 확인하고 안전하게 변환
                        floor_val = trade.get('floor', 0)
                        price_val = trade.get('price', 0)

                        # dict인 경우 0으로 처리
                        if isinstance(floor_val, dict):
                            logging.warning(f"floor is dict in existing trade: {floor_val}")
                            floor_val = 0
                        if isinstance(price_val, dict):
                            logging.warning(f"price is dict in existing trade: {price_val}")
                            price_val = 0

                        key = (trade['date'].year, trade['date'].month, trade['date'].day,
                              floor_val, price_val)
                        existing_keys.add(key)

                for trade in trades:
                    # floor와 price가 dict가 아닌지 확인하고 안전하게 변환
                    floor_val = trade.get('floor', 0)
                    price_val = trade.get('price', 0)

                    if isinstance(floor_val, dict):
                        logging.warning(f"floor is dict in new trade: {floor_val}")
                        floor_val = 0
                    if isinstance(price_val, dict):
                        logging.warning(f"price is dict in new trade: {price_val}")
                        price_val = 0

                    trade_key = (trade['date'].year, trade['date'].month, trade['date'].day,
                                floor_val, price_val)
                    if trade_key not in existing_keys:
                        existing_trades.append(trade)

                trades = existing_trades
            
            return sorted(trades, key=lambda x: x.get('date', datetime.min))
        except Exception as e:
            logging.error(f"데이터 수집 중 오류: {str(e)}")
            return apt_info.get('trade_data', [])

    def _safe_value(self, value, default=''):
        """dict인 경우 기본값 반환, 아니면 문자열로 변환"""
        if isinstance(value, dict):
            return default
        return str(value) if value is not None else default

    def _safe_apt_data(self, apt):
        """apt 딕셔너리에서 안전하게 값을 추출하여 dict 반환"""
        return {
            'name': self._safe_value(apt.get('apt_name', '')),
            'price': apt.get('new_price', 0) if not isinstance(apt.get('new_price', 0), dict) else 0,
            'date': self._safe_value(apt.get('date', '')),
            'old_price': apt.get('old_price', 0) if not isinstance(apt.get('old_price', 0), dict) else 0,
            'floor': self._safe_value(apt.get('floor', '')),
            'area': self._safe_value(apt.get('area', '')),
            'sido': self._safe_value(apt.get('sido', '')),
            'sigungu': self._safe_value(apt.get('sigungu', '')),
            'location_dong': self._safe_value(apt.get('location_dong', '')),
            'build_year': self._safe_value(apt.get('build_year', ''))
        }

    def build_notification_html(self, apt_list):
        """신고가 알림용 HTML (날짜 필터 + 상승률 필터 포함)"""
        from html import escape
        from collections import Counter
        import re
        import traceback

        # 디버깅: apt_list의 각 항목에서 dict 값 찾기
        for idx, apt in enumerate(apt_list):
            for key, value in apt.items():
                if isinstance(value, dict):
                    logging.error(f"[디버깅] apt_list[{idx}]['{key}']가 dict입니다: {value}")
                    logging.error(f"[디버깅] 해당 아파트: {apt.get('apt_name', 'unknown')}")

        now = datetime.now().strftime('%Y-%m-%d %H:%M')
        current_year = datetime.now().year

        # 신고가 높은 순으로 정렬
        apt_list = sorted(apt_list, key=lambda x: x.get('new_price', 0), reverse=True)
        total = len(apt_list)
        
        # 지역별 건수 집계 및 연식별 카운트
        region_counter = Counter()
        young_count = 0  # 10년 이하
        very_young_count = 0  # 5년 이하
        old_count = 0  # 2000년대 미만 (1999년 이하)
        
        for apt in apt_list:
            sido = apt.get('sido', '')
            sigungu = apt.get('sigungu', '')
            location_dong_raw = apt.get('location_dong', '')
            # location_dong이 dict인 경우 빈 문자열로 처리
            if isinstance(location_dong_raw, dict):
                location_dong = ''
            else:
                location_dong = str(location_dong_raw) if location_dong_raw else ''
            if sido and sigungu:
                sigungu_clean = sigungu.split('(')[0] if '(' in sigungu else sigungu
                # 성남시 특정 동 매핑
                if sido == "경기도" and "성남시" in sigungu:
                    # 분당구 동
                    if location_dong in ['구미동', '금곡동', '대장동', '백현동', '분당동', '서현동', '수내동', '야탑동', '운중동', '정자동', '판교동', '삼평동', '동막동', '궁내동', '율동', '매송동']:
                        sigungu_clean = "성남시 분당구"
                    # 수정구 동
                    elif location_dong in ['고등동', '금토동', '단대동', '복정동', '신흥동', '양지동', '오야동', '태평동', '신촌동', '수진동', '창곡동', '시흥동', '둔전동']:
                        sigungu_clean = "성남시 수정구"
                    # 중원구 동
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
                    
                    # 2000년 미만 체크
                    if year < 2000:
                        old_count += 1
                        apt['is_old'] = True
                    else:
                        apt['is_old'] = False
                    
                    # 기존 연식 체크
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

            # ★ 직전 최고가 날짜와 층 정보 추가
            old_date = escape(str(apt.get('old_date', '')))
            old_floor = escape(str(apt.get('old_floor', '')))
            
            # 지역 정보
            sido = escape(str(apt.get('sido', '')))
            sigungu = escape(str(apt.get('sigungu', '')))
            location_dong_raw = apt.get('location_dong', '')
            # location_dong이 dict인 경우 빈 문자열로 처리
            if isinstance(location_dong_raw, dict):
                location_dong = ''
            else:
                location_dong = escape(str(location_dong_raw)) if location_dong_raw else ''
            location = f"{sido} {sigungu} {location_dong}"

            # 용인시, 고양시, 수원시 특정 동 매핑
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

            # 가격 정보
            inc = new_price - old_price if old_price else 0
            pct = f"{(inc/old_price*100):.1f}%" if old_price else "-"

            # ★ 직전 최고가에 날짜와 층 정보 포함
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
            
            # 연식 표시 설정
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

            # region_name과 rank 값이 dict가 아닌 경우만 사용
            region_name_val = apt.get("region_name", "")
            region_84_rank_val = apt.get("region_84_rank", "")
            region_59_rank_val = apt.get("region_59_rank", "")

            # dict인 경우 빈 문자열로 처리
            if isinstance(region_name_val, dict):
                region_name_val = ""
            if isinstance(region_84_rank_val, dict):
                region_84_rank_val = ""
            if isinstance(region_59_rank_val, dict):
                region_59_rank_val = ""

            # 카카오맵 검색용 쿼리 (시군구 + 동 + 아파트명 + 아파트)
            # 괄호 처리 규칙:
            # 1. "숫자+동" 패턴 (예: 12동, 13동)이 있으면 괄호 전체 무시
            # 2. 그 외 한글이 있으면 한글만 추출하여 검색어에 포함
            # 3. 숫자/기호만 있으면 괄호 무시
            import re
            if '(' in name:
                match = re.search(r'\(([^)]+)\)', name)
                if match:
                    paren_content = match.group(1)
                    # "숫자+동" 패턴 체크 (예: 12동, 13동, 1동 등)
                    if re.search(r'\d+동', paren_content):
                        # 숫자+동 패턴이 있으면 괄호 전체 무시
                        apt_name_for_search = name.split('(')[0].strip()
                    elif re.search(r'[가-힣]', paren_content):
                        # 한글이 있으면 괄호 내용 중 한글만 추출하여 붙임
                        korean_only = re.sub(r'[^가-힣\s]', '', paren_content).strip()
                        apt_name_for_search = name.split('(')[0].strip() + ' ' + korean_only
                    else:
                        # 한글 없으면 괄호 앞부분만
                        apt_name_for_search = name.split('(')[0].strip()
                else:
                    apt_name_for_search = name.split('(')[0].strip()
            else:
                apt_name_for_search = name
            # 시군구 + 동 + 아파트명으로 검색 (동일 단지명 구분을 위해)
            kakao_query = f"{sigungu} {location_dong} {apt_name_for_search} 아파트".strip()

            card = f"""
            <section class="card" data-region="{sido} {region_for_filter}"
                     data-build-year="{data_year}" data-young="{is_young}"
                     data-very-young="{is_very_young}" data-old="{is_old}" data-area="{area}"
                     data-kakao-query="{escape(kakao_query)}" onclick="openKakaoMap(this)" style="cursor:pointer;">
              <h3>{name} {year_badge} <span class="tag">전용면적 {area}㎡</span></h3>
              <div class="card-sub">📍 {location} <span style="color:#007AFF; font-size:0.9em;">🗺️ 지도보기</span></div>
              <div class="grid">
                <div class="k">이번 실거래</div>
                <div><span class="highlight">{new_str}</span> ({date}{f' | {floor}층' if floor else ''})</div>
                <div class="k">직전 최고가</div>
                <div>{old_str}</div>
                <div class="k">변화</div>
                <div class="rise">{inc_str} 상승 ({pct})</div>
                <div class="k">특이사항</div>
                <div>
                    {f'<span class="badge-rank rank-84">{region_name_val} 84타입 NO.{region_84_rank_val} 단지</span>' if region_84_rank_val else ''}
                    {f'<span class="badge-rank rank-59">{region_name_val} 59타입 NO.{region_59_rank_val} 단지</span>' if region_59_rank_val else ''}
                </div>
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
        <title>부태리 신고가 알림 - {escape(now)}</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js"></script>
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
          /* 워터마크 */
          .watermark {{
            position:fixed; top:0; left:0; width:100%; height:100%;
            pointer-events:none; z-index:9999; overflow:hidden;
          }}
          .watermark-text {{
            position:absolute; width:300%; height:300%;
            top:-100%; left:-100%;
            display:flex; flex-wrap:wrap; justify-content:center; align-items:center;
            transform:rotate(-30deg);
          }}
          .watermark-text span {{
            font-size:16px; color:rgba(100,100,100,0.08);
            padding:30px 50px; white-space:nowrap;
            user-select:none; -webkit-user-select:none;
            font-weight:600;
          }}
          @media print {{ .watermark {{ display:block !important; }} }}
          .page{{ max-width:960px; width:92vw; margin:24px auto 80px }}
          .header{{
            background:linear-gradient(135deg,var(--primary),#63A4FF);
            color:#fff; border-radius:16px; padding:20px 20px 16px;
            box-shadow:0 6px 18px rgba(0,0,0,.08);
          }}
          .header h1{{ margin:0 0 8px; font-size:22px; font-weight:800; letter-spacing:-.2px }}
          .header .sub{{ margin:0; opacity:.95; font-size:14px }}

          .chips{{ display:flex; flex-wrap:wrap; gap:8px; margin:12px 0 0 }}
          .chip{{
            display:inline-flex; align-items:center; gap:6px; padding:8px 12px;
            border:1px solid var(--bd); border-radius:999px; background:#fff; color:#111; font-size:13px;
            box-shadow:0 1px 1.5px rgba(0,0,0,.04); cursor:pointer;
            transition: all 0.2s ease;
          }}
          .chip:hover{{ background:#f3f4f6 }}
          .chip.active{{ background:#1976d2; color:#fff; border-color:#1976d2 }}

          .region-chips{{ display:flex; flex-wrap:wrap; gap:8px; margin:10px 0 0 }}
          .region-chip{{
            display:inline-flex; align-items:center; gap:6px; padding:8px 12px;
            border:1px solid var(--bd); border-radius:999px; background:#fff; color:#111; font-size:13px;
            box-shadow:0 1px 1.5px rgba(0,0,0,.04); cursor:pointer;
            transition: all 0.2s ease;
          }}
          .region-chip:hover{{ background:#f3f4f6 }}
          .region-chip.active{{ background:#1976d2; color:#fff; border-color:#1976d2 }}
          .region-chip.zero{{ opacity:.45; }}

          .cards{{ display:grid; grid-template-columns:repeat(12,1fr); gap:12px; margin-top:16px }}
          .card{{
            grid-column:span 12; background:var(--card); border:1px solid var(--bd);
            border-radius:14px; padding:14px; box-shadow:0 2px 12px rgba(0,0,0,.04);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
          }}
          .card:hover{{ transform: translateY(-2px); box-shadow:0 4px 20px rgba(0,0,0,.08) }}
          @media(min-width:720px){{ .card{{ grid-column:span 6 }} }}
          .card.hidden{{ display:none }}
          .card h3{{ margin:0 0 4px; font-size:16px; display:flex; align-items:center; gap:6px }}

          .year-badge {{
            display:inline-flex; align-items:center; padding:2px 8px; border-radius:6px;
            background:#e5e7eb; color:#4b5563; font-size:12px; font-weight:normal;
          }}
          .year-badge.old {{ background:#f3e8ff; color:#6b21a8; font-weight:600; }}         /* 2000년 미만 */
          .year-badge.very-young {{ background:#fce7f3; color:#be185d; font-weight:600; }}  /* 5년 이하 */
          .year-badge.young {{ background:#dbeafe; color:#1e40af; font-weight:600; }}       /* 10년 이하 */
          .year-badge.new {{ background:#fef3c7; color:#92400e; font-weight:600; }}         /* 분양 */

          .tag{{ font-size:12px; padding:2px 8px; border-radius:999px; border:1px solid var(--bd); background:#fff; color:#333 }}
          .card .card-badges{{ display:flex; flex-wrap:wrap; gap:6px; margin:8px 0 6px; }}
          .card .card-sub{{ color:var(--sub); font-size:13px; margin-bottom:6px }}
          .grid{{ display:grid; grid-template-columns:110px 1fr; gap:10px; color:#111; font-size:13px }}
          .grid .k{{ color:var(--sub) }}
          .highlight{{ color:var(--primary-dark); font-weight:700 }}
          .rise{{ color:#e11d48; font-weight:700 }}
          .badge-rank{{
            display:inline-flex; align-items:center; gap:4px; padding:4px 10px; margin-left:6px;
            border-radius:999px; color:#fff; font-size:12px; font-weight:600;
            box-shadow:0 2px 4px rgba(0,0,0,0.2); animation:pulse 2s infinite;
          }}
          .badge-rank.rank-84{{ background:linear-gradient(135deg,#fbbf24,#f59e0b); border:1px solid #f59e0b; }}
          .badge-rank.rank-59{{ background:linear-gradient(135deg,#8b5cf6,#7c3aed); border:1px solid #7c3aed; }}
          @keyframes pulse{{ 0%,100%{{opacity:1;transform:scale(1)}} 50%{{opacity:.9;transform:scale(1.02)}} }}

          /* 투기과열지구/토지거래허가구역 배지 */
          .badge-speculation{{
            display:inline-flex; align-items:center; gap:4px; padding:6px 12px;
            border-radius:8px; font-size:12px; font-weight:700; letter-spacing:-0.3px;
            box-shadow:0 3px 8px rgba(0,0,0,0.25); animation:glow 2s ease-in-out infinite;
          }}
          .badge-speculation.hot{{
            background:linear-gradient(135deg,#dc2626,#b91c1c);
            color:#fff; border:2px solid #991b1b;
          }}
          .badge-speculation.permit{{
            background:linear-gradient(135deg,#ea580c,#c2410c);
            color:#fff; border:2px solid #9a3412;
          }}
          @keyframes glow{{
            0%,100%{{box-shadow:0 3px 8px rgba(220,38,38,0.4),0 0 15px rgba(220,38,38,0.2)}}
            50%{{box-shadow:0 3px 12px rgba(220,38,38,0.6),0 0 25px rgba(220,38,38,0.4)}}
          }}

          .toolbar{{ display:flex; gap:8px; align-items:center; margin:8px 0 4px; flex-wrap:wrap }}
          .toolbar input[type="date"], .toolbar input[type="number"]{{
            padding:10px 12px; border:1px solid var(--bd); border-radius:10px; background:#fff; font-size:13px; width:120px;
          }}
          .toolbar .btn{{
            display:inline-flex; align-items:center; padding:10px 12px; border:1px solid var(--bd);
            border-radius:999px; background:#fff; font-size:13px; cursor:pointer; transition: all 0.2s ease;
          }}
          .toolbar .btn:hover{{ background:#f3f4f6 }}

          footer{{ margin:24px 0 0; color:#6b7280; font-size:12px; text-align:center }}

          /* 신고가 분위 그래프 모달 */
          .quintile-modal {{
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            overflow: auto;
            background-color: rgba(0, 0, 0, 0.5);
          }}
          .quintile-modal-content {{
            background-color: #fff;
            margin: 5% auto;
            border-radius: 20px;
            width: 90%;
            max-width: 900px;
            max-height: 80vh;
            overflow: hidden;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            display: flex;
            flex-direction: column;
          }}
          .quintile-modal-header {{
            padding: 20px 30px;
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: #fff;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }}
          .quintile-modal-header h2 {{
            margin: 0;
            font-size: 1.5em;
          }}
          .quintile-modal-close {{
            font-size: 2em;
            font-weight: bold;
            cursor: pointer;
            color: #fff;
            line-height: 1;
            transition: transform 0.2s;
          }}
          .quintile-modal-close:hover {{
            transform: scale(1.2);
          }}
          .quintile-modal-body {{
            padding: 30px;
            overflow-y: auto;
            flex: 1;
          }}
          .quintile-description {{
            color: #6b7280;
            font-size: 14px;
            margin-bottom: 20px;
            text-align: center;
          }}
          #quintileChart {{
            max-height: 400px;
          }}
          .quintile-stats {{
            margin-top: 20px;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
          }}
          .quintile-stat-card {{
            padding: 15px;
            border-radius: 12px;
            background: #f8f9fa;
            text-align: center;
          }}
          .quintile-stat-card h4 {{
            margin: 0 0 8px;
            font-size: 13px;
            color: #6b7280;
          }}
          .quintile-stat-card .value {{
            font-size: 24px;
            font-weight: bold;
            color: #111;
          }}
          .quintile-stat-card .range {{
            font-size: 11px;
            color: #9ca3af;
            margin-top: 4px;
          }}
        </style>
        </head>
        <body>
          <!-- 워터마크 -->
          <div class="watermark"><div class="watermark-text">{"".join(['<span>부태리 ⓒ 2025</span>' for _ in range(100)])}</div></div>
          <div class="page">
            <header class="header">
              <h1>부태리 신고가 알림</h1>
              <p class="sub">{escape(now)} 기준 · 총 {total}개 아파트</p>

              <!-- 상단 칩: 연식별 필터 및 그래프 버튼 -->
              <div class="chips">
                <span class="chip" id="chip-old">2000년대 미만 {old_count}건</span>
                <span class="chip" id="chip-very-young">입주 5년 이하 {very_young_count}건</span>
                <span class="chip" id="chip-young">입주 10년 이하 {young_count}건</span>
                <span class="chip chip-graph" id="chip-quintile-graph" style="background: linear-gradient(135deg, #667eea, #764ba2); color: #fff; border: none; font-weight: 600;">
                  📊 신고가 분위그래프
                </span>
              </div>

              <!-- 지역 칩 (동적) -->
              <div class="region-chips" id="regionChips"></div>
              <!-- 하위 지역 칩 (동적) -->
              <div class="region-chips" id="subRegionChips" style="margin-top:8px;"></div>
            </header>

            <!-- 상승률 필터 툴바 -->
            <div class="toolbar">
              <label style="font-size:13px;color:var(--sub)">상승률(%)</label>
              <input id="rateMin" type="number" step="0.1" min="0" placeholder="최소"/>
              <span style="color:var(--sub)">~</span>
              <input id="rateMax" type="number" step="0.1" min="0" placeholder="최대"/>
              <button class="btn" id="btnRate5">≥ 5%</button>
              <button class="btn" id="btnRate10">≥ 10%</button>
              <button class="btn" id="btnRateReset">상승률 초기화</button>
            </div>

            <!-- 면적 필터 툴바 -->
            <div class="toolbar">
              <label style="font-size:13px;color:var(--sub)">면적(㎡)</label>
              <input id="areaMin" type="number" step="1" min="0" placeholder="최소"/>
              <span style="color:var(--sub)">~</span>
              <input id="areaMax" type="number" step="1" min="0" placeholder="최대"/>
              <button class="btn" id="btnArea60">≤60㎡</button>
              <button class="btn" id="btnArea60_85">60~85㎡</button>
              <button class="btn" id="btnArea85_135">85~135㎡</button>
              <button class="btn" id="btnArea135">≥135㎡</button>
              <button class="btn" id="btnAreaReset">면적 초기화</button>
            </div>

            <!-- 날짜 필터 툴바 -->
            <div class="toolbar">
              <label style="font-size:13px;color:var(--sub)">기간</label>
              <input id="dateFrom" type="date" lang="en-CA" placeholder="YYYY-MM-DD" inputmode="numeric" pattern="\d{{4}}-\d{{2}}-\d{{2}}"/>
              <span style="color:var(--sub)">~</span>
              <input id="dateTo" type="date" lang="en-CA" placeholder="YYYY-MM-DD" inputmode="numeric" pattern="\d{{4}}-\d{{2}}-\d{{2}}"/>
              <button class="btn" id="btnToday">오늘</button>
              <button class="btn" id="btn7">최근 7일</button>
              <button class="btn" id="btn30">최근 30일</button>
              <button class="btn" id="btnDateReset">날짜 초기화</button>

              <span style="width:8px"></span>
              <button class="btn" id="resetBtn">전체 초기화</button>
            </div>

            <section class="cards" id="cards">
              {cards}
            </section>

            <footer>© 부태리의 실거래가 모니터링 시스템</footer>
          </div>

          <!-- 신고가 분위 그래프 모달 -->
          <div id="quintileModal" class="quintile-modal">
            <div class="quintile-modal-content">
              <div class="quintile-modal-header">
                <h2>📊 신고가 분위 분석</h2>
                <span class="quintile-modal-close" onclick="closeQuintileModal()">&times;</span>
              </div>
              <div class="quintile-modal-body">
                <p class="quintile-description">이번 갱신에서 발생한 신고가를 5억 단위로 구간을 나누어 가격대별 분포를 보여줍니다. 막대를 클릭하면 해당 가격대의 아파트 목록을 볼 수 있습니다.</p>
                <canvas id="quintileChart"></canvas>
                <div id="quintileStats" class="quintile-stats"></div>
              </div>
            </div>
          </div>

          <!-- 가격대별 아파트 목록 모달 -->
          <div id="aptListModal" class="quintile-modal">
            <div class="quintile-modal-content">
              <div class="quintile-modal-header">
                <h2 id="aptListModalTitle">아파트 목록</h2>
                <span class="quintile-modal-close" onclick="closeAptListModal()">&times;</span>
              </div>
              <div class="quintile-modal-body">
                <div id="aptListContent"></div>
              </div>
            </div>
          </div>

        <script>
        // 카카오맵 열기 함수
        function openKakaoMap(element) {{
          const query = element.getAttribute('data-kakao-query');
          if (query) {{
            const url = 'https://map.kakao.com/?q=' + encodeURIComponent(query);
            window.open(url, '_blank', 'width=1200,height=800');
          }}
        }}

        (function(){{
          const $  = (s, r=document)=>r.querySelector(s);
          const $$ = (s, r=document)=>Array.from(r.querySelectorAll(s));

          const cards = $$('.card');
          const headerSub = $('.header .sub');
          const chipsWrap = $('#regionChips');
          const subRegionChipsWrap = $('#subRegionChips');

          /* === 날짜/상승률/면적 필터 요소 === */
          const dateFrom = $('#dateFrom');
          const dateTo   = $('#dateTo');
          const btnToday = $('#btnToday');
          const btn7     = $('#btn7');
          const btn30    = $('#btn30');
          const btnDateReset = $('#btnDateReset');

          const rateMin = $('#rateMin');
          const rateMax = $('#rateMax');
          const btnRate5 = $('#btnRate5');
          const btnRate10 = $('#btnRate10');
          const btnRateReset = $('#btnRateReset');

          const areaMin = $('#areaMin');
          const areaMax = $('#areaMax');
          const btnArea60 = $('#btnArea60');
          const btnArea60_85 = $('#btnArea60_85');
          const btnArea85_135 = $('#btnArea85_135');
          const btnArea135 = $('#btnArea135');
          const btnAreaReset = $('#btnAreaReset');

          /* === 전체 초기화 버튼 === */
          const resetBtn = $('#resetBtn');

          /* === 연식 칩 상태 === */
          let selectedSido = '__ALL__';  // 선택된 시도 ('__ALL__' 또는 '__SIDO__경기도' 등)
          let selectedRegions = new Set();  // 다중 선택된 상세 지역들
          let filterYoung = false;
          let filterVeryYoung = false;
          let filterOld = false;

          /* ===== 유틸 ===== */
          function toInputValue(d){{
            const y = d.getFullYear();
            const m = String(d.getMonth()+1).padStart(2,'0');
            const day = String(d.getDate()).padStart(2,'0');
            return `${{y}}-${{m}}-${{day}}`;
          }}
          function todayStr(){{ return toInputValue(new Date()); }}
          function daysAgo(n){{
            const d = new Date();
            d.setDate(d.getDate() - n + 1); // 오늘 포함 n일
            return d;
          }}

          // '이번 실거래 (YYYY-MM-DD | ...)' 날짜 파싱
          function parseTradeDate(card){{
            const priceSpan = card.querySelector('.grid .highlight');
            if(!priceSpan) return null;
            const cell = priceSpan.parentElement;
            const text = (cell.textContent || '').trim();
            const m = text.match(/\((\d{{4}}-\d{{2}}-\d{{2}})[^\)]*\)/);
            return m ? new Date(m[1]) : null;
          }}

          // 변화 영역에서 상승률 % 파싱 → 숫자
          function parseRisePct(card){{
            const changeEl = card.querySelector('.grid .rise');
            if(!changeEl) return null;
            const text = changeEl.textContent || '';
            const m = text.match(/(\d+(?:\.\d+)?)%/);
            return m ? parseFloat(m[1]) : null;
          }}

          // 동 이름으로 구 매핑 (성남시, 고양시, 용인시, 수원시, 안양시)
          const dongToGuMap = {{
            // 성남시
            '삼평동': '분당구', '정자동': '분당구', '서현동': '분당구', '판교동': '분당구',
            '이매동': '분당구', '구미동': '분당구', '야탑동': '분당구', '금곡동': '분당구',
            '수내동': '분당구', '분당동': '분당구', '운중동': '분당구', '백현동': '분당구',
            '신흥동': '수정구', '태평동': '수정구', '수진동': '수정구', '단대동': '수정구',
            '산성동': '수정구', '양지동': '수정구', '복정동': '수정구', '신촌동': '수정구',
            '은행동': '중원구', '상대원동': '중원구', '금광동': '중원구', '성남동': '중원구',
            '중앙동': '중원구', '하대원동': '중원구', '도촌동': '중원구',

            // 고양시
            '대화동': '일산서구', '주엽동': '일산서구', '탄현동': '일산서구', '일산동': '일산동구',
            '마두동': '일산동구', '백석동': '일산동구', '장항동': '일산서구', '식사동': '일산동구',
            '풍동': '일산서구', '화정동': '덕양구', '행신동': '덕양구', '원당동': '덕양구',

            // 용인시
            '기흥동': '기흥구', '보라동': '기흥구', '구갈동': '기흥구', '신갈동': '기흥구',
            '상현동': '수지구', '동천동': '수지구', '죽전동': '수지구', '풍덕천동': '수지구',
            '삼가동': '처인구', '역북동': '처인구', '남사면': '처인구',

            // 수원시
            '영통동': '영통구', '매탄동': '영통구', '원천동': '영통구', '광교동': '영통구',
            '이의동': '영통구', '망포동': '영통구',
            '인계동': '팔달구', '매교동': '팔달구', '화서동': '팔달구',
            '천천동': '장안구', '조원동': '장안구', '연무동': '장안구',
            '권선동': '권선구', '서둔동': '권선구', '구운동': '권선구',

            // 안양시
            '만안동': '만안구', '석수동': '만안구', '박달동': '만안구',
            '비산동': '동안구', '평촌동': '동안구', '관양동': '동안구', '호계동': '동안구'
          }};

          // 상세 지역 추출 함수
          function extractDetailedRegion(card){{
            const cardSub = card.querySelector('.card-sub');
            if(!cardSub) return card.dataset.region || '';

            const baseRegion = card.dataset.region || '';
            const text = (cardSub.textContent || '').replace('📍', '').trim();

            // 서울/부산/대구/인천/광주/대전/울산 등 광역시는 그대로 반환 (이미 구까지 표현됨)
            if(baseRegion.includes('특별시') || baseRegion.includes('광역시')){{
              return baseRegion;
            }}

            // 도 지역에서 시 단위인 경우만 처리 (성남시, 고양시, 용인시, 수원시 등)
            if(baseRegion.includes('도') && baseRegion.includes('시')){{
              // 동 추출: "경기도 성남시(경) 삼평동" -> "삼평동"
              const dongMatch = text.match(/([가-힣]+동|[가-힣]+면)/);
              if(dongMatch){{
                const dong = dongMatch[1];

                // 시 정보를 먼저 확인하여 정확한 매핑
                let gu = null;
                if(baseRegion.includes('성남시')){{
                  // 성남시
                  const seongnamDongs = {{
                    '분당동':'분당구','수내동':'분당구','수내1동':'분당구','수내2동':'분당구','정자동':'분당구','정자1동':'분당구','정자2동':'분당구','정자3동':'분당구',
                    '서현동':'분당구','서현1동':'분당구','서현2동':'분당구','이매동':'분당구','이매1동':'분당구','이매2동':'분당구',
                    '야탑동':'분당구','야탑1동':'분당구','야탑2동':'분당구','야탑3동':'분당구','판교동':'분당구','삼평동':'분당구',
                    '백현동':'분당구','금곡동':'분당구','구미동':'분당구','구미1동':'분당구','운중동':'분당구',
                    '신흥동':'수정구','태평동':'수정구','태평1동':'수정구','태평2동':'수정구','태평3동':'수정구','태평4동':'수정구',
                    '수진동':'수정구','수진1동':'수정구','수진2동':'수정구','단대동':'수정구','산성동':'수정구',
                    '양지동':'수정구','복정동':'수정구','신촌동':'수정구','시흥동':'수정구','고등동':'수정구',
                    '성남동':'중원구','중앙동':'중원구','금광동':'중원구','금광1동':'중원구','금광2동':'중원구',
                    '은행동':'중원구','은행1동':'중원구','은행2동':'중원구','상대원동':'중원구','상대원1동':'중원구','상대원2동':'중원구','상대원3동':'중원구',
                    '하대원동':'중원구','도촌동':'중원구'
                  }};
                  gu = seongnamDongs[dong];
                }} else if(baseRegion.includes('수원시')){{
                  // 수원시
                  const suwonDongs = {{
                    '장안구':'장안구','파장동':'장안구','율천동':'장안구','정자동':'장안구','정자1동':'장안구','정자2동':'장안구','정자3동':'장안구',
                    '영화동':'장안구','송죽동':'장안구','조원동':'장안구','조원1동':'장안구','조원2동':'장안구','연무동':'장안구',
                    '세류동':'권선구','세류1동':'권선구','세류2동':'권선구','세류3동':'권선구','평동':'권선구','서둔동':'권선구','구운동':'권선구',
                    '금곡동':'권선구','오매동':'권선구','권선동':'권선구','권선1동':'권선구','권선2동':'권선구','곡선동':'권선구','입북동':'권선구',
                    '매교동':'팔달구','매산동':'팔달구','고등동':'팔달구','화서동':'팔달구','화서1동':'팔달구','화서2동':'팔달구',
                    '지동':'팔달구','우만동':'팔달구','우만1동':'팔달구','우만2동':'팔달구','행궁동':'팔달구',
                    '매탄동':'영통구','매탄1동':'영통구','매탄2동':'영통구','매탄3동':'영통구','매탄4동':'영통구',
                    '원천동':'영통구','영통동':'영통구','영통1동':'영통구','영통2동':'영통구','영통3동':'영통구',
                    '태장동':'영통구','망포동':'영통구','망포1동':'영통구','망포2동':'영통구','광교동':'영통구','광교1동':'영통구','광교2동':'영통구'
                  }};
                  gu = suwonDongs[dong];
                }} else if(baseRegion.includes('안양시')){{
                  // 안양시
                  const anyangDongs = {{
                    '안양동':'만안구','안양1동':'만안구','안양2동':'만안구','안양3동':'만안구','안양4동':'만안구','안양5동':'만안구',
                    '안양6동':'만안구','안양7동':'만안구','안양8동':'만안구','안양9동':'만안구','석수동':'만안구','석수1동':'만안구',
                    '석수2동':'만안구','박달동':'만안구','박달1동':'만안구','박달2동':'만안구',
                    '비산동':'동안구','비산1동':'동안구','비산2동':'동안구','비산3동':'동안구','달안동':'동안구','관양동':'동안구',
                    '인덕원동':'동안구','부림동':'동안구','평안동':'동안구','평촌동':'동안구','귀인동':'동안구',
                    '호계동':'동안구','호계1동':'동안구','호계2동':'동안구','호계3동':'동안구',
                    '범계동':'동안구','신촌동':'동안구','부산동':'동안구'
                  }};
                  gu = anyangDongs[dong];
                }} else if(baseRegion.includes('부천시')){{
                  // 부천시
                  const bucheonDongs = {{
                    '심곡동':'원미구','심곡1동':'원미구','심곡2동':'원미구','심곡3동':'원미구','원미동':'원미구','원미1동':'원미구',
                    '원미2동':'원미구','소사동':'원미구','역곡동':'원미구','역곡1동':'원미구','역곡2동':'원미구','춘의동':'원미구',
                    '도당동':'원미구','중동':'원미구','중1동':'원미구','중2동':'원미구','중3동':'원미구','중4동':'원미구',
                    '상동':'원미구','상1동':'원미구','상2동':'원미구','상3동':'원미구'
                  }};
                  gu = bucheonDongs[dong];
                }} else if(baseRegion.includes('용인시')){{
                  // 용인시
                  const yonginDongs = {{
                    '포곡읍':'처인구','모현읍':'처인구','이동읍':'처인구','원삼면':'처인구','백암면':'처인구','양지면':'처인구',
                    '중앙동':'처인구','삼가동':'처인구','역북동':'처인구','남사면':'처인구',
                    '신갈동':'기흥구','영덕동':'기흥구','영덕1동':'기흥구','영덕2동':'기흥구','상갈동':'기흥구','보라동':'기흥구',
                    '기흥동':'기흥구','서농동':'기흥구','마북동':'기흥구','동백동':'기흥구','동백1동':'기흥구','동백2동':'기흥구',
                    '동백3동':'기흥구','구갈동':'기흥구',
                    '죽전동':'수지구','죽전1동':'수지구','죽전2동':'수지구','신봉동':'수지구','상현동':'수지구','상현1동':'수지구',
                    '상현2동':'수지구','죽전3동':'수지구','성복동':'수지구','풍덕천동':'수지구','동천동':'수지구'
                  }};
                  gu = yonginDongs[dong];
                }} else if(baseRegion.includes('고양시')){{
                  // 고양시
                  const goyangDongs = {{
                    '주교동':'덕양구','원신동':'덕양구','흥도동':'덕양구','성사동':'덕양구','성사1동':'덕양구','성사2동':'덕양구',
                    '효자동':'덕양구','삼송동':'덕양구','삼송1동':'덕양구','삼송2동':'덕양구','고양동':'덕양구','관산동':'덕양구',
                    '능곡동':'덕양구','향동동':'덕양구','행주동':'덕양구','행주1동':'덕양구','행주2동':'덕양구','행신동':'덕양구',
                    '행신1동':'덕양구','행신2동':'덕양구','행신3동':'덕양구','행신4동':'덕양구','화전동':'덕양구','대덕동':'덕양구',
                    '백석동':'일산동구','중산동':'일산동구','중산1동':'일산동구','중산2동':'일산동구','장항동':'일산동구','장항1동':'일산동구',
                    '장항2동':'일산동구','식사동':'일산동구','마두동':'일산동구','마두1동':'일산동구','마두2동':'일산동구','풍산동':'일산동구',
                    '산황동':'일산서구','일산동':'일산서구','일산1동':'일산서구','일산2동':'일산서구','일산3동':'일산서구','탄현동':'일산서구',
                    '탄현1동':'일산서구','탄현2동':'일산서구','주엽동':'일산서구','주엽1동':'일산서구','주엽2동':'일산서구',
                    '대화동':'일산서구','송포동':'일산서구','송산동':'일산서구'
                  }};
                  gu = goyangDongs[dong];
                }} else if(baseRegion.includes('오산시')){{
                  // 오산시
                  const osanDongs = {{
                    '성곡동':'오산시','원동':'오산시','원종동':'오산시','신흥동':'오산시'
                  }};
                  gu = osanDongs[dong];
                }}

                if(gu){{
                  // "경기도 성남시" -> "경기도 성남시 분당구"
                  return `${{baseRegion}} ${{gu}}`;
                }}
              }}
            }}

            return baseRegion;
          }}

          // 투기과열지구/토지거래허가구역 판정
          // 서울 전체, 경기 과천·광명·성남(분당·수정·중원구)·수원(영통·장안·팔달구)·안양(동안구)·용인(수지구)·의왕시·하남시
          function isSpeculationArea(card){{
            // data-region-detail 속성 사용 (구 정보 포함)
            const detailedRegion = card.dataset.regionDetail || '';

            let isHot = false;
            let isPermit = false;

            // 서울 전체
            if(detailedRegion.includes('서울특별시')){{
              isHot = true;
              isPermit = true;
            }}
            // 경기 과천
            else if(detailedRegion.includes('과천시')){{
              isHot = true;
              isPermit = true;
            }}
            // 경기 광명
            else if(detailedRegion.includes('광명시')){{
              isHot = true;
              isPermit = true;
            }}
            // 경기 성남 (분당·수정·중원구)
            else if(detailedRegion.includes('성남시')){{
              if(detailedRegion.includes('분당구') || detailedRegion.includes('수정구') || detailedRegion.includes('중원구')){{
                isHot = true;
                isPermit = true;
              }}
              // 동 이름으로 구 판별
              else {{
                const seongnamBundang = ['구미동', '금곡동', '대장동', '백현동', '분당동', '서현동', '수내동', '야탑동', '운중동', '정자동', '판교동', '삼평동', '동막동', '궁내동', '율동', '매송동'];
                const seongnamSujeong = ['고등동', '금토동', '단대동', '복정동', '신흥동', '양지동', '오야동', '태평동', '신촌동', '수진동', '창곡동', '시흥동', '둔전동'];
                const seongnamJungwon = ['갈현동', '도촌동', '상대원동', '성남동', '은행동', '중앙동', '하대원동', '금광동', '여수동'];

                for(let dong of seongnamBundang) {{
                  if(detailedRegion.includes(dong)) {{
                    isHot = true;
                    isPermit = true;
                    break;
                  }}
                }}
                if(!isHot) {{
                  for(let dong of seongnamSujeong) {{
                    if(detailedRegion.includes(dong)) {{
                      isHot = true;
                      isPermit = true;
                      break;
                    }}
                  }}
                }}
                if(!isHot) {{
                  for(let dong of seongnamJungwon) {{
                    if(detailedRegion.includes(dong)) {{
                      isHot = true;
                      isPermit = true;
                      break;
                    }}
                  }}
                }}
              }}
            }}
            // 경기 수원 (영통·장안·팔달구)
            else if(detailedRegion.includes('수원시')){{
              if(detailedRegion.includes('영통구') || detailedRegion.includes('장안구') || detailedRegion.includes('팔달구')){{
                isHot = true;
                isPermit = true;
              }}
            }}
            // 경기 안양 (동안구)
            else if(detailedRegion.includes('안양시')){{
              if(detailedRegion.includes('동안구')){{
                isHot = true;
                isPermit = true;
              }}
            }}
            // 경기 용인 (수지구)
            else if(detailedRegion.includes('용인시')){{
              if(detailedRegion.includes('수지구')){{
                isHot = true;
                isPermit = true;
              }}
            }}
            // 경기 의왕시
            else if(detailedRegion.includes('의왕시')){{
              isHot = true;
              isPermit = true;
            }}
            // 경기 하남시
            else if(detailedRegion.includes('하남시')){{
              isHot = true;
              isPermit = true;
            }}

            return {{ isHot, isPermit }};
          }}

          // 모든 카드에 상세 지역 정보 설정 및 배지 추가
          cards.forEach(card => {{
            const detailedRegion = extractDetailedRegion(card);
            card.dataset.regionDetail = detailedRegion;

            // 시도 정보 추출 (서울특별시, 경기도, 인천광역시 등)
            const sidoMatch = detailedRegion.match(/^(.+?특별시|.+?광역시|.+?특별자치시|.+?도|.+?특별자치도)/);
            if(sidoMatch){{
              card.dataset.sido = sidoMatch[1];
            }}

            // 투기과열지구/토지거래허가구역 배지 추가
            const {{ isHot, isPermit }} = isSpeculationArea(card);
            if(isHot || isPermit){{
              const h3 = card.querySelector('h3');
              const cardSub = card.querySelector('.card-sub');

              if(h3 && cardSub){{
                // h3 다음, card-sub 이전에 배지 라인 삽입
                const badgesDiv = document.createElement('div');
                badgesDiv.className = 'card-badges';

                // 투기과열지구 배지
                if(isHot){{
                  const hotBadge = document.createElement('span');
                  hotBadge.className = 'badge-speculation hot';
                  hotBadge.textContent = '🔥 투기과열지구';
                  badgesDiv.appendChild(hotBadge);
                }}
                // 토지거래허가구역 배지
                if(isPermit){{
                  const permitBadge = document.createElement('span');
                  permitBadge.className = 'badge-speculation permit';
                  permitBadge.textContent = '⚠️ 토지거래허가';
                  badgesDiv.appendChild(permitBadge);
                }}

                // h3 다음에 삽입
                h3.parentNode.insertBefore(badgesDiv, cardSub);
              }}
            }}
          }});

          // 초기 지역 목록 (구까지 포함)
          const allRegions = Array.from(new Set(cards.map(c => c.dataset.regionDetail || '').filter(Boolean))).sort();

          // 시도 목록 추출
          const allSidos = Array.from(new Set(cards.map(c => c.dataset.sido || '').filter(Boolean))).sort();

          // 지역칩 렌더 (처음 1회)
          function renderRegionChips(){{
            chipsWrap.innerHTML = '';

            // '전체 보기' 칩
            const allChip = document.createElement('span');
            allChip.className = 'region-chip active';
            allChip.dataset.region = '__ALL__';
            allChip.dataset.type = 'all';
            allChip.textContent = '전체 보기';
            chipsWrap.appendChild(allChip);

            // 시도별 전체 칩 (서울 전체, 경기 전체 등)
            allSidos.forEach(sido=>{{
              const sidoChip = document.createElement('span');
              sidoChip.className = 'region-chip';
              sidoChip.dataset.region = `__SIDO__${{sido}}`;
              sidoChip.dataset.type = 'sido';
              sidoChip.dataset.sido = sido;

              // "서울특별시" -> "서울 전체", "경기도" -> "경기 전체"
              const sidoName = sido.replace('특별시', '').replace('광역시', '').replace('특별자치시', '').replace('도', '').replace('특별자치도', '');
              sidoChip.textContent = `${{sidoName}} 전체 0건`;
              chipsWrap.appendChild(sidoChip);
            }});

            // 상세 지역 칩
            allRegions.forEach(r=>{{
              const chip = document.createElement('span');
              chip.className = 'region-chip';
              chip.dataset.region = r;
              chip.dataset.type = 'detail';
              chip.textContent = `${{r}} 0건`;
              chipsWrap.appendChild(chip);
            }});
          }}

          // 지역칩 숫자 갱신(지역 제외 모든 필터 반영) + ★ 거래수 많은 순으로 재정렬
          function refreshRegionChips(){{
            // 상세 지역 카운팅
            const counts = new Map(allRegions.map(r=>[r,0]));

            // 시도별 카운팅
            const sidoCounts = new Map(allSidos.map(s=>[s,0]));

            cards.forEach(card=>{{
              if(nonRegionMatch(card)){{ // 지역 조건 제외
                const r = card.dataset.regionDetail || '';
                const s = card.dataset.sido || '';

                if(counts.has(r)) counts.set(r, counts.get(r)+1);
                if(sidoCounts.has(s)) sidoCounts.set(s, sidoCounts.get(s)+1);
              }}
            }});

            // ★ 거래수가 많은 순으로 지역 정렬
            const sortedRegions = Array.from(counts.entries())
              .sort((a, b) => b[1] - a[1])  // 거래수 내림차순
              .map(entry => entry[0]);

            // ★ 1줄: 메인 메뉴 칩 (전체 보기, 서울 전체, 경기 전체, 인천 전체)
            chipsWrap.innerHTML = '';

            // '전체 보기' 칩
            const allChip = document.createElement('span');
            allChip.className = 'region-chip' + (selectedSido === '__ALL__' ? ' active' : '');
            allChip.dataset.region = '__ALL__';
            allChip.dataset.type = 'all';
            allChip.textContent = '전체 보기';
            chipsWrap.appendChild(allChip);

            // 시도별 전체 칩 추가
            allSidos.forEach(sido=>{{
              const n = sidoCounts.get(sido) || 0;
              const sidoChip = document.createElement('span');
              const sidoKey = `__SIDO__${{sido}}`;
              sidoChip.className = 'region-chip' + (selectedSido === sidoKey ? ' active' : '') + (n === 0 ? ' zero' : '');
              sidoChip.dataset.region = sidoKey;
              sidoChip.dataset.type = 'sido';
              sidoChip.dataset.sido = sido;

              // "서울특별시" -> "서울 전체", "경기도" -> "경기 전체"
              const sidoName = sido.replace('특별시', '').replace('광역시', '').replace('특별자치시', '').replace('도', '').replace('특별자치도', '');
              sidoChip.textContent = `${{sidoName}} 전체 ${{n}}건`;
              chipsWrap.appendChild(sidoChip);
            }});

            // ★ 2줄: 서브 지역 칩 (시도가 선택되었을 때만)
            subRegionChipsWrap.innerHTML = '';

            if(selectedSido.startsWith('__SIDO__')){{
              const currentSido = selectedSido.replace('__SIDO__', '');

              // 선택된 시도의 지역만 필터링
              const regionsInSido = sortedRegions.filter(r => r.startsWith(currentSido));

              regionsInSido.forEach(r=>{{
                const n = counts.get(r) || 0;
                const chip = document.createElement('span');
                const isSelected = selectedRegions.has(r);
                chip.className = 'region-chip' + (isSelected ? ' active' : '') + (n === 0 ? ' zero' : '');
                chip.dataset.region = r;
                chip.dataset.type = 'detail';
                chip.textContent = `${{r}} ${{n}}건`;
                subRegionChipsWrap.appendChild(chip);
              }});
            }}
          }}

          // 헤더 카운트 갱신
          function setHeaderCount(n){{
            if(!headerSub) return;
            // 원문에서 날짜·총 N개 패턴 앞부분만 추출
            const base = headerSub.textContent.replace(/· 총\s+\d+\s*개\s*아파트(?:\s*\(필터 적용\))?/,'').trim();
            const filterOn = (
              dateFrom.value || dateTo.value || filterYoung || filterVeryYoung || filterOld ||
              selectedSido !== '__ALL__' || selectedRegions.size > 0 || rateMin.value || rateMax.value ||
              areaMin.value || areaMax.value
            );
            const filterBadge = filterOn ? ' (필터 적용)' : '';
            headerSub.textContent = `${{base}} · 총 ${{n}}개 아파트${{filterBadge}}`;
          }}

          /* ===== 필터 판정 로직 ===== */
          function nonRegionMatch(card){{
            // 날짜
            const fromVal = dateFrom.value ? new Date(dateFrom.value) : null;
            const toVal   = dateTo.value ? new Date(dateTo.value) : null;
            const d = parseTradeDate(card);
            const passDate =
              (!fromVal || (d && d >= fromVal)) &&
              (!toVal   || (d && d <= toVal));

            // 연식
            const yOldOn   = filterOld;
            const yVYOn    = filterVeryYoung;
            const yYoungOn = filterYoung;
            const yearFilterOn = yOldOn || yVYOn || yYoungOn;
            const isOld   = card.dataset.old === '1';
            const isVY    = card.dataset.veryYoung === '1';
            const isYoung = card.dataset.young === '1';
            const passYear = !yearFilterOn || (
              (yOldOn   && isOld) ||
              (yVYOn    && isVY)  ||
              (yYoungOn && isYoung)
            );

            // 상승률
            const minPct = rateMin.value !== '' ? parseFloat(rateMin.value) : null;
            const maxPct = rateMax.value !== '' ? parseFloat(rateMax.value) : null;
            const rateFilterOn = (minPct !== null) || (maxPct !== null);
            const pct = parseRisePct(card);
            const passRate = !rateFilterOn || (
              (minPct === null || (pct !== null && pct >= minPct)) &&
              (maxPct === null || (pct !== null && pct <= maxPct))
            );

            // 면적
            const minArea = areaMin.value !== '' ? parseFloat(areaMin.value) : null;
            const maxArea = areaMax.value !== '' ? parseFloat(areaMax.value) : null;
            const areaFilterOn = (minArea !== null) || (maxArea !== null);
            const area = card.dataset.area ? parseFloat(card.dataset.area) : null;
            const passArea = !areaFilterOn || (
              (minArea === null || (area !== null && area >= minArea)) &&
              (maxArea === null || (area !== null && area <= maxArea))
            );

            return passDate && passYear && passRate && passArea;
          }}

          function applyFilter(){{
            const shown = cards.reduce((acc, card)=>{{
              let regionOk = false;
              if(selectedSido === '__ALL__'){{
                regionOk = true;
              }} else if(selectedSido.startsWith('__SIDO__')){{
                // 시도 선택 시
                const sido = selectedSido.replace('__SIDO__', '');
                const cardSido = card.dataset.sido || '';
                if(selectedRegions.size === 0){{
                  // 상세 지역 선택 없으면 해당 시도 전체
                  regionOk = (cardSido === sido);
                }} else {{
                  regionOk = selectedRegions.has(card.dataset.regionDetail);
                }}
              }}

              const match = nonRegionMatch(card) && regionOk;
              card.classList.toggle('hidden', !match);
              return acc + (match ? 1 : 0);
            }}, 0);

            setHeaderCount(shown);
            refreshRegionChips(); // 카드 표시 후 지역칩 숫자도 즉시 갱신
          }}

          /* ===== 지역칩: 이벤트 위임 ===== */
          renderRegionChips();
          chipsWrap.addEventListener('click', (e)=>{{
            const chip = e.target.closest('.region-chip');
            if(!chip) return;
            const region = chip.dataset.region;
            if(!region) return;

            // 전체 보기 또는 시도 선택 시
            if(chip.dataset.type === 'all' || chip.dataset.type === 'sido'){{
              selectedSido = region;
              selectedRegions.clear(); // 시도가 변경되면 상세 지역 선택 초기화

              // 칩 활성화 상태 업데이트
              $$('.region-chip', chipsWrap).forEach(c=>{{
                c.classList.toggle('active', c.dataset.region === selectedSido);
              }});
            }}

            applyFilter();
          }});

          // 서브 지역 칩 클릭 이벤트
          subRegionChipsWrap.addEventListener('click', (e)=>{{
            const chip = e.target.closest('.region-chip');
            if(!chip) return;
            const region = chip.dataset.region;
            if(!region) return;

            // 상세 지역 다중 선택 토글
            if(selectedRegions.has(region)){{
              selectedRegions.delete(region);
            }} else {{
              selectedRegions.add(region);
            }}

            // 칩 활성화 상태 업데이트
            $$('.region-chip', subRegionChipsWrap).forEach(c=>{{
              c.classList.toggle('active', selectedRegions.has(c.dataset.region));
            }});

            applyFilter();
          }});

          /* ===== 연식 칩 ===== */
          const oldChip = $('#chip-old');
          const veryYoungChip = $('#chip-very-young');
          const youngChip = $('#chip-young');

          if (oldChip){{
            oldChip.addEventListener('click', ()=>{{
              filterOld = !filterOld; filterVeryYoung = false; filterYoung = false;
              oldChip.classList.toggle('active', filterOld);
              veryYoungChip.classList.remove('active');
              youngChip.classList.remove('active');
              applyFilter();
            }});
          }}
          if (veryYoungChip){{
            veryYoungChip.addEventListener('click', ()=>{{
              filterVeryYoung = !filterVeryYoung; filterYoung = false; filterOld = false;
              veryYoungChip.classList.toggle('active', filterVeryYoung);
              youngChip.classList.remove('active');
              oldChip.classList.remove('active');
              applyFilter();
            }});
          }}
          if (youngChip){{
            youngChip.addEventListener('click', ()=>{{
              filterYoung = !filterYoung; filterVeryYoung = false; filterOld = false;
              youngChip.classList.toggle('active', filterYoung);
              veryYoungChip.classList.remove('active');
              oldChip.classList.remove('active');
              applyFilter();
            }});
          }}

          /* ===== 날짜 입력/프리셋 ===== */
          dateFrom.addEventListener('change', ()=>{{
            if(dateTo.value && dateFrom.value && dateFrom.value > dateTo.value){{
              dateTo.value = dateFrom.value;
            }}
            applyFilter();
          }});
          dateTo.addEventListener('change', ()=>{{
            if(dateFrom.value && dateTo.value && dateTo.value < dateFrom.value){{
              dateFrom.value = dateTo.value;
            }}
            applyFilter();
          }});

          $('#btnToday').addEventListener('click', ()=>{{
            const t = todayStr();
            dateFrom.value = t; dateTo.value = t;
            applyFilter();
          }});
          $('#btn7').addEventListener('click', ()=>{{
            dateFrom.value = toInputValue(daysAgo(7));
            dateTo.value   = todayStr();
            applyFilter();
          }});
          $('#btn30').addEventListener('click', ()=>{{
            dateFrom.value = toInputValue(daysAgo(30));
            dateTo.value   = todayStr();
            applyFilter();
          }});
          btnDateReset.addEventListener('click', ()=>{{
            dateFrom.value = '';
            dateTo.value = '';
            applyFilter();
          }});

          /* ===== 상승률 입력/프리셋 ===== */
          function onRateChange(){{ applyFilter(); }}
          rateMin.addEventListener('input', onRateChange);
          rateMax.addEventListener('input', onRateChange);
          btnRate5.addEventListener('click', ()=>{{ rateMin.value = '5';  rateMax.value = ''; applyFilter(); }});
          btnRate10.addEventListener('click',()=>{{ rateMin.value = '10'; rateMax.value = ''; applyFilter(); }});
          btnRateReset.addEventListener('click', ()=>{{ rateMin.value = ''; rateMax.value = ''; applyFilter(); }});

          /* ===== 면적 입력/프리셋 ===== */
          function onAreaChange(){{ applyFilter(); }}
          areaMin.addEventListener('input', onAreaChange);
          areaMax.addEventListener('input', onAreaChange);
          btnArea60.addEventListener('click', ()=>{{ areaMin.value = ''; areaMax.value = '60'; applyFilter(); }});
          btnArea60_85.addEventListener('click', ()=>{{ areaMin.value = '60'; areaMax.value = '85'; applyFilter(); }});
          btnArea85_135.addEventListener('click', ()=>{{ areaMin.value = '85'; areaMax.value = '135'; applyFilter(); }});
          btnArea135.addEventListener('click', ()=>{{ areaMin.value = '135'; areaMax.value = ''; applyFilter(); }});
          btnAreaReset.addEventListener('click', ()=>{{ areaMin.value = ''; areaMax.value = ''; applyFilter(); }});

          /* ===== 전체 초기화 ===== */
          resetBtn.addEventListener('click', ()=>{{
            dateFrom.value = '';
            dateTo.value   = '';
            rateMin.value  = '';
            rateMax.value  = '';
            areaMin.value  = '';
            areaMax.value  = '';
            filterOld = false;
            filterVeryYoung = false;
            filterYoung = false;

            oldChip.classList.remove('active');
            veryYoungChip.classList.remove('active');
            youngChip.classList.remove('active');

            selectedSido = '__ALL__';
            selectedRegions.clear();
            $$('.region-chip', chipsWrap).forEach(c=> c.classList.toggle('active', c.dataset.region === '__ALL__'));
            applyFilter();
          }});

          // 최초 렌더 시 필터 적용
          applyFilter();
        }})();

        // 신고가 분위 그래프 관련 함수
        const aptData = {json.dumps([self._safe_apt_data(apt) for apt in apt_list], ensure_ascii=False)};
        let quintileChartInstance = null;

        function openQuintileModal() {{
          const modal = document.getElementById('quintileModal');
          modal.style.display = 'block';

          // 그래프 생성
          if (!quintileChartInstance) {{
            createQuintileChart();
          }}
        }}

        function closeQuintileModal() {{
          const modal = document.getElementById('quintileModal');
          modal.style.display = 'none';
        }}

        function createQuintileChart() {{
          // 이 HTML에 표시된 모든 신고가 데이터를 사용
          if (aptData.length === 0) {{
            document.getElementById('quintileStats').innerHTML = '<p style="text-align:center;color:#999;">데이터가 없습니다.</p>';
            return;
          }}

          // 최대값 찾기
          const allPrices = aptData.map(apt => apt.price);
          const maxPrice = Math.max(...allPrices);

          // 5억 단위로 구간 생성 (0억~5억, 5억~10억, 10억~15억, ...)
          const interval = 50000; // 5억 = 50,000만원
          const maxBracket = Math.ceil(maxPrice / interval) * interval;

          // 구간 경계값 배열 생성
          const brackets = [];
          for (let i = 0; i <= maxBracket; i += interval) {{
            brackets.push(i);
          }}

          // 구간별 카운트
          const counts = new Array(brackets.length - 1).fill(0);
          aptData.forEach(apt => {{
            const price = apt.price;
            for (let i = 0; i < brackets.length - 1; i++) {{
              if (price >= brackets[i] && price < brackets[i + 1]) {{
                counts[i]++;
                break;
              }}
            }}
            // 정확히 최대값인 경우 마지막 구간에 포함
            if (price === maxPrice && price === brackets[brackets.length - 1]) {{
              counts[counts.length - 1]++;
            }}
          }});

          // 라벨 생성 (예: "0~5억", "5~10억", ...)
          const labels = [];
          for (let i = 0; i < brackets.length - 1; i++) {{
            const start = brackets[i] / 10000; // 억 단위로 변환
            const end = brackets[i + 1] / 10000;
            labels.push(`${{start}}~${{end}}억`);
          }}

          // 색상 배열 생성 (구간 수만큼)
          const colors = [];
          const borderColors = [];
          const baseColors = [
            'rgba(59, 130, 246, 0.8)',
            'rgba(99, 102, 241, 0.8)',
            'rgba(139, 92, 246, 0.8)',
            'rgba(168, 85, 247, 0.8)',
            'rgba(217, 70, 239, 0.8)',
            'rgba(236, 72, 153, 0.8)',
            'rgba(239, 68, 68, 0.8)',
            'rgba(249, 115, 22, 0.8)',
            'rgba(234, 179, 8, 0.8)',
            'rgba(132, 204, 22, 0.8)',
            'rgba(34, 197, 94, 0.8)'
          ];
          const baseBorderColors = [
            'rgba(59, 130, 246, 1)',
            'rgba(99, 102, 241, 1)',
            'rgba(139, 92, 246, 1)',
            'rgba(168, 85, 247, 1)',
            'rgba(217, 70, 239, 1)',
            'rgba(236, 72, 153, 1)',
            'rgba(239, 68, 68, 1)',
            'rgba(249, 115, 22, 1)',
            'rgba(234, 179, 8, 1)',
            'rgba(132, 204, 22, 1)',
            'rgba(34, 197, 94, 1)'
          ];
          for (let i = 0; i < counts.length; i++) {{
            colors.push(baseColors[i % baseColors.length]);
            borderColors.push(baseBorderColors[i % baseBorderColors.length]);
          }}

          // 차트 생성
          const ctx = document.getElementById('quintileChart').getContext('2d');
          quintileChartInstance = new Chart(ctx, {{
            type: 'bar',
            data: {{
              labels: labels,
              datasets: [{{
                label: '신고가 건수',
                data: counts,
                backgroundColor: colors,
                borderColor: borderColors,
                borderWidth: 2,
                borderRadius: 8
              }}]
            }},
            options: {{
              responsive: true,
              maintainAspectRatio: true,
              onClick: function(event, activeElements) {{
                if (activeElements.length > 0) {{
                  const index = activeElements[0].index;
                  const bracketStart = brackets[index];
                  const bracketEnd = brackets[index + 1];
                  const label = labels[index];
                  showAptListForBracket(bracketStart, bracketEnd, label);
                }}
              }},
              plugins: {{
                legend: {{
                  display: false
                }},
                title: {{
                  display: true,
                  text: '신고가 단지의 가격대별 분포 (5억 단위)',
                  font: {{
                    size: 16,
                    weight: 'bold'
                  }},
                  padding: {{
                    top: 10,
                    bottom: 20
                  }}
                }},
                tooltip: {{
                  backgroundColor: 'rgba(0, 0, 0, 0.8)',
                  padding: 12,
                  callbacks: {{
                    label: function(context) {{
                      return context.parsed.y + '건 (클릭하여 목록 보기)';
                    }}
                  }}
                }}
              }},
              scales: {{
                y: {{
                  beginAtZero: true,
                  ticks: {{
                    stepSize: 1,
                    callback: function(value) {{
                      return value + '건';
                    }}
                  }}
                }}
              }}
            }}
          }});

          // 통계 정보 표시
          let statsHtml = '';
          for (let i = 0; i < counts.length; i++) {{
            const start = (brackets[i] / 10000).toFixed(1);
            const end = (brackets[i + 1] / 10000).toFixed(1);
            statsHtml += `
              <div class="quintile-stat-card">
                <h4>${{labels[i]}}</h4>
                <div class="value">${{counts[i]}}건</div>
                <div class="range">${{brackets[i].toLocaleString()}} ~ ${{brackets[i + 1].toLocaleString()}}만원</div>
              </div>
            `;
          }}
          document.getElementById('quintileStats').innerHTML = statsHtml;
        }}

        // 특정 가격대의 아파트 목록 표시
        function showAptListForBracket(bracketStart, bracketEnd, label) {{
          // 해당 가격대의 아파트 필터링
          const filteredApts = aptData.filter(apt => {{
            return apt.price >= bracketStart && apt.price < bracketEnd;
          }});

          // 가격 높은 순으로 정렬
          filteredApts.sort((a, b) => b.price - a.price);

          // 모달 제목 설정
          document.getElementById('aptListModalTitle').textContent = `${{label}} 신고가 아파트 (${{filteredApts.length}}건)`;

          // 아파트 목록 HTML 생성
          let html = '<div style="padding: 10px;">';

          if (filteredApts.length === 0) {{
            html += '<p style="text-align:center; color:#999; padding:40px;">해당 가격대의 아파트가 없습니다.</p>';
          }} else {{
            html += '<table style="width:100%; border-collapse:collapse; font-size:14px;">';
            html += '<thead><tr style="background:#f3f4f6; border-bottom:2px solid #ddd;">';
            html += '<th style="padding:12px; text-align:left;">아파트명</th>';
            html += '<th style="padding:12px; text-align:left;">위치</th>';
            html += '<th style="padding:12px; text-align:center;">면적</th>';
            html += '<th style="padding:12px; text-align:center;">연식</th>';
            html += '<th style="padding:12px; text-align:right;">신고가</th>';
            html += '<th style="padding:12px; text-align:right;">이전가</th>';
            html += '<th style="padding:12px; text-align:center;">날짜</th>';
            html += '</tr></thead><tbody>';

            filteredApts.forEach((apt, idx) => {{
              const increase = apt.price - apt.old_price;
              const increaseRate = apt.old_price > 0 ? ((increase / apt.old_price) * 100).toFixed(1) : 0;
              const bgColor = idx % 2 === 0 ? '#fff' : '#f9fafb';

              html += `<tr style="border-bottom:1px solid #eee; background:${{bgColor}};">`;
              html += `<td style="padding:10px; font-weight:600;">${{apt.name}}</td>`;
              html += `<td style="padding:10px;">${{apt.sido}} ${{apt.sigungu}}<br><small style="color:#666;">${{apt.location_dong}}</small></td>`;
              html += `<td style="padding:10px; text-align:center;">${{apt.area}}㎡</td>`;
              html += `<td style="padding:10px; text-align:center;">${{apt.build_year}}</td>`;
              html += `<td style="padding:10px; text-align:right; font-weight:700; color:#dc2626;">${{apt.price.toLocaleString()}}만원</td>`;
              html += `<td style="padding:10px; text-align:right;">${{apt.old_price.toLocaleString()}}만원<br><small style="color:#059669;">+${{increase.toLocaleString()}}만원 (${{increaseRate}}%)</small></td>`;
              html += `<td style="padding:10px; text-align:center; font-size:12px;">${{apt.date}}<br>${{apt.floor}}</td>`;
              html += '</tr>';
            }});

            html += '</tbody></table>';
          }}

          html += '</div>';
          document.getElementById('aptListContent').innerHTML = html;

          // 모달 표시
          document.getElementById('aptListModal').style.display = 'block';
        }}

        // 아파트 목록 모달 닫기
        function closeAptListModal() {{
          document.getElementById('aptListModal').style.display = 'none';
        }}

        // 그래프 버튼 클릭 이벤트
        document.getElementById('chip-quintile-graph').addEventListener('click', openQuintileModal);

        // 모달 배경 클릭 시 닫기
        window.onclick = function(event) {{
          const quintileModal = document.getElementById('quintileModal');
          const aptListModal = document.getElementById('aptListModal');
          if (event.target === quintileModal) {{
            closeQuintileModal();
          }}
          if (event.target === aptListModal) {{
            closeAptListModal();
          }}
        }};

        </script>
        </body>
        </html>"""


        return html_content

    def save_area_ranking_history(self):
        """84㎡, 59㎡ 평형의 현재 순위를 DB에 저장"""
        try:
            cursor = self.db_conn.cursor()
            active_list_name = self.active_list.get()
            cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (active_list_name,))
            result = cursor.fetchone()

            if not result:
                return

            active_list_id = result[0]

            # 84㎡와 59㎡ 각각 처리
            for area_type in ['84', '59']:
                target_area = 84 if area_type == '84' else 59
                tolerance = 0.5

                # 현재 평형의 아파트 조회
                cursor.execute("""
                    SELECT apt_name, area, last_max_price, sido, sigungu, dong
                    FROM apartments
                    WHERE list_id = ?
                    ORDER BY last_max_price DESC
                """, (active_list_id,))

                all_apts = cursor.fetchall()

                # 평형 필터링 및 순위 매기기
                rank = 1
                for row in all_apts:
                    apt_name, area, price, sido, sigungu, dong = row

                    try:
                        area_num = float(str(area).replace('㎡', '').strip()) if area else 0
                    except:
                        continue

                    if abs(area_num - target_area) <= tolerance:
                        # 순위 이력 저장
                        cursor.execute("""
                            INSERT INTO area_ranking_history
                            (list_id, area_type, apt_name, area, ranking, price, sido, sigungu, dong, timestamp)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                        """, (active_list_id, area_type, apt_name, area, rank, price or 0, sido, sigungu, dong))

                        rank += 1

            self.db_conn.commit()
            logging.info(f"평형별 순위 이력 저장 완료")

        except Exception as e:
            logging.error(f"순위 이력 저장 중 오류: {str(e)}")

    def build_area_ranking_html(self, area_type):
        """모니터링 중인 모든 아파트를 대상으로 평형별 가격 순위 HTML 생성
        area_type: '84' 또는 '59'
        """
        import re
        from html import escape

        # DB에서 모니터링 중인 모든 아파트 조회
        try:
            cursor = self.db_conn.cursor()
            active_list_name = self.active_list.get()
            cursor.execute("SELECT id FROM monitoring_lists WHERE name = ?", (active_list_name,))
            result = cursor.fetchone()

            if not result:
                return None

            active_list_id = result[0]

            cursor.execute("""
                SELECT apt_name, area, last_max_price, max_price_date, max_price_floor, max_price_dong,
                       sido, sigungu, dong, build_year, sigungu_code
                FROM apartments
                WHERE list_id = ?
                ORDER BY last_max_price DESC
            """, (active_list_id,))

            all_apts = cursor.fetchall()
        except Exception as e:
            logging.error(f"평형별 순위 조회 중 오류: {str(e)}")
            return None

        # 평형 필터링 (±0.5㎡)
        target_area = 84 if area_type == '84' else 59
        tolerance = 0.5

        filtered_apts = []
        for row in all_apts:
            apt_name, area, price, date, floor, dong, sido, sigungu, location_dong, build_year, sigungu_code = row

            # 면적 파싱
            try:
                area_num = float(str(area).replace('㎡', '').strip()) if area else 0
            except:
                continue

            if abs(area_num - target_area) <= tolerance:
                filtered_apts.append({
                    'apt_name': apt_name,
                    'area': f"{int(area_num)}㎡" if area_num else '',
                    'new_price': price or 0,
                    'date': date or '',
                    'floor': floor or '',
                    'dong': dong or '',
                    'sido': sido or '',
                    'sigungu': sigungu or '',
                    'location_dong': location_dong or '',
                    'build_year': build_year or '',
                    'sigungu_code': sigungu_code or ''
                })

        # 가격 순으로 정렬 (높은 가격부터)
        filtered_apts.sort(key=lambda x: x['new_price'], reverse=True)

        # ========== 중복 단지 제거 (띄어쓰기 무시) ==========
        # 띄어쓰기 제거한 이름을 키로 사용하여 중복 체크
        unique_apts = {}
        for apt in filtered_apts:
            # 띄어쓰기, 대소문자 제거한 정규화된 이름
            normalized_name = apt['apt_name'].replace(' ', '').replace('-', '').lower()

            # 같은 정규화 이름이 있으면 가격이 더 높은 것만 유지
            if normalized_name in unique_apts:
                if apt['new_price'] > unique_apts[normalized_name]['new_price']:
                    logging.info(f"[중복 제거] '{unique_apts[normalized_name]['apt_name']}' ({unique_apts[normalized_name]['new_price']:,}만원) → '{apt['apt_name']}' ({apt['new_price']:,}만원) 유지")
                    unique_apts[normalized_name] = apt
                else:
                    logging.info(f"[중복 제거] '{apt['apt_name']}' ({apt['new_price']:,}만원) 제거 (기존: '{unique_apts[normalized_name]['apt_name']}' {unique_apts[normalized_name]['new_price']:,}만원)")
            else:
                unique_apts[normalized_name] = apt

        # 중복 제거된 리스트로 교체 (가격 순으로 다시 정렬)
        filtered_apts = sorted(unique_apts.values(), key=lambda x: x['new_price'], reverse=True)
        # ==========================================================

        # 이전 순위 조회 (가장 최근 2번째 데이터와 비교)
        try:
            cursor.execute("""
                SELECT apt_name, ranking, price
                FROM area_ranking_history
                WHERE list_id = ? AND area_type = ?
                AND timestamp < (
                    SELECT MAX(timestamp) FROM area_ranking_history
                    WHERE list_id = ? AND area_type = ?
                )
                AND timestamp = (
                    SELECT MAX(timestamp) FROM area_ranking_history
                    WHERE list_id = ? AND area_type = ?
                    AND timestamp < (
                        SELECT MAX(timestamp) FROM area_ranking_history
                        WHERE list_id = ? AND area_type = ?
                    )
                )
            """, (active_list_id, area_type, active_list_id, area_type,
                  active_list_id, area_type, active_list_id, area_type))

            prev_rankings = {}
            for row in cursor.fetchall():
                prev_apt_name, prev_rank, prev_price = row
                prev_rankings[prev_apt_name] = {'rank': prev_rank, 'price': prev_price}

        except Exception as e:
            logging.error(f"이전 순위 조회 중 오류: {str(e)}")
            prev_rankings = {}

        # 각 아파트에 이전 순위 정보 추가
        for idx, apt in enumerate(filtered_apts, 1):
            apt['current_rank'] = idx
            prev_info = prev_rankings.get(apt['apt_name'])
            if prev_info:
                apt['prev_rank'] = prev_info['rank']
                apt['prev_price'] = prev_info['price']
                apt['rank_change'] = prev_info['rank'] - idx  # 양수면 상승, 음수면 하락
                apt['price_change'] = apt['new_price'] - prev_info['price']
            else:
                apt['prev_rank'] = None
                apt['prev_price'] = None
                apt['rank_change'] = None
                apt['price_change'] = None

        now = datetime.now().strftime('%Y-%m-%d %H:%M')
        total = len(filtered_apts)
        area_name = "전용면적 84㎡ (국평 25평형)" if area_type == '84' else "전용면적 59㎡ (국평 18평형)"

        # ========== 인덱스 기반 데이터 압축 (용량 최적화) ==========
        MAX_RANK_DISPLAY = 3000
        display_apts = filtered_apts[:MAX_RANK_DISPLAY]

        # 인덱스 테이블
        apt_name_set = {}
        sigungu_set = {}
        dong_set = {}
        location_dong_set = {}

        apt_names_list = []
        sigungus_list = []
        dongs_list = []
        location_dongs_list = []

        def get_index(value, value_set, value_list):
            if not value:
                return -1
            val_str = str(value) if not isinstance(value, dict) else ''
            if not val_str:
                return -1
            if val_str not in value_set:
                value_set[val_str] = len(value_list)
                value_list.append(val_str)
            return value_set[val_str]

        # 압축 데이터 생성
        # [apt_name_idx, sigungu_idx, dong_idx, location_dong_idx, new_price, date, floor,
        #  build_year, current_rank, rank_change, prev_rank, price_change, prev_price, sido]
        compressed_apts = []
        for apt in display_apts:
            apt_idx = get_index(apt['apt_name'], apt_name_set, apt_names_list)
            sigungu_idx = get_index(apt['sigungu'], sigungu_set, sigungus_list)
            dong_idx = get_index(apt['dong'], dong_set, dongs_list)
            loc_dong_idx = get_index(apt['location_dong'], location_dong_set, location_dongs_list)

            compressed_apts.append([
                apt_idx,                          # 0: apt_name index
                sigungu_idx,                      # 1: sigungu index
                dong_idx,                         # 2: dong index
                loc_dong_idx,                     # 3: location_dong index
                apt['new_price'] or 0,            # 4: price
                apt['date'] or '',                # 5: date
                apt['floor'] or '',               # 6: floor
                apt['build_year'] or '',          # 7: build_year
                apt['current_rank'],              # 8: current_rank
                apt.get('rank_change'),           # 9: rank_change (None 가능)
                apt.get('prev_rank'),             # 10: prev_rank (None 가능)
                apt.get('price_change'),          # 11: price_change (None 가능)
                apt.get('prev_price'),            # 12: prev_price (None 가능)
                apt['sido'] or ''                 # 13: sido
            ])

        # JSON 변환
        apt_names_json = json.dumps(apt_names_list, ensure_ascii=False)
        sigungus_json = json.dumps(sigungus_list, ensure_ascii=False)
        dongs_json = json.dumps(dongs_list, ensure_ascii=False)
        location_dongs_json = json.dumps(location_dongs_list, ensure_ascii=False)
        apts_json = json.dumps(compressed_apts, ensure_ascii=False)

        # 압축률 로깅
        original_data = json.dumps(display_apts, ensure_ascii=False, default=str)
        original_size = len(original_data)
        compressed_size = len(apt_names_json) + len(sigungus_json) + len(dongs_json) + len(location_dongs_json) + len(apts_json)
        compression_ratio = (1 - compressed_size / original_size) * 100 if original_size > 0 else 0
        logging.info(f"[평형별 순위] 데이터 압축: {original_size:,} -> {compressed_size:,} bytes ({compression_ratio:.1f}% 감소)")
        # ========== 압축 완료 ==========

        # 더 이상 cards_html 생성하지 않음 (JavaScript에서 동적 생성)
        area_display = f"{target_area}㎡"

        html_content = f"""<!DOCTYPE html>
        <html lang="ko">
        <head>
        <meta charset="utf-8"/>
        <meta content="width=device-width, initial-scale=1" name="viewport"/>
        <title>{area_name} 가격 순위 - {escape(now)}</title>
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
          /* 워터마크 */
          .watermark {{
            position:fixed; top:0; left:0; width:100%; height:100%;
            pointer-events:none; z-index:9999; overflow:hidden;
          }}
          .watermark-text {{
            position:absolute; width:300%; height:300%;
            top:-100%; left:-100%;
            display:flex; flex-wrap:wrap; justify-content:center; align-items:center;
            transform:rotate(-30deg);
          }}
          .watermark-text span {{
            font-size:16px; color:rgba(100,100,100,0.08);
            padding:30px 50px; white-space:nowrap;
            user-select:none; -webkit-user-select:none;
            font-weight:600;
          }}
          @media print {{ .watermark {{ display:block !important; }} }}
          .page{{ max-width:960px; width:92vw; margin:24px auto 80px }}
          .header{{
            background:linear-gradient(135deg,var(--primary),#63A4FF);
            color:#fff; border-radius:16px; padding:20px 20px 16px;
            box-shadow:0 6px 18px rgba(0,0,0,.08);
          }}
          .header h1{{ margin:0 0 8px; font-size:22px; font-weight:800; letter-spacing:-.2px }}
          .header .sub{{ margin:0; opacity:.95; font-size:14px }}

          .cards{{ display:flex; flex-direction:column; gap:16px; margin-top:16px }}
          .card{{
            background:var(--card); border:1px solid var(--bd);
            border-radius:14px; box-shadow:0 2px 12px rgba(0,0,0,.04);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            overflow:hidden;
          }}
          .card:hover{{ transform: translateY(-2px); box-shadow:0 4px 20px rgba(0,0,0,.08) }}

          .card-header{{
            background:linear-gradient(135deg,#f8fafc,#e2e8f0);
            padding:16px 20px;
            border-bottom:2px solid #cbd5e1;
          }}
          .card-header h3{{
            margin:0; font-size:18px; font-weight:700;
            display:flex; align-items:center; gap:8px; flex-wrap:wrap;
          }}

          .rank-badge{{
            display:inline-flex; align-items:center; padding:4px 10px;
            border-radius:999px; font-size:12px; font-weight:600;
            box-shadow:0 2px 4px rgba(0,0,0,0.15);
          }}
          .rank-badge.up{{ background:linear-gradient(135deg,#3b82f6,#2563eb); color:#fff; }}
          .rank-badge.down{{ background:linear-gradient(135deg,#ef4444,#dc2626); color:#fff; }}
          .rank-badge.same{{ background:linear-gradient(135deg,#94a3b8,#64748b); color:#fff; }}
          .rank-badge.new{{ background:linear-gradient(135deg,#fbbf24,#f59e0b); color:#fff; }}

          .divider{{
            height:2px;
            background:repeating-linear-gradient(
              90deg,
              #cbd5e1 0px,
              #cbd5e1 10px,
              transparent 10px,
              transparent 20px
            );
          }}

          .card-body{{
            padding:16px 20px;
            display:flex; flex-direction:column; gap:10px;
          }}
          .info-row{{
            font-size:14px; line-height:1.6;
            display:flex; align-items:center; gap:8px;
          }}
          .info-row strong{{ color:var(--primary-dark); font-weight:700; font-size:16px; }}

          .regulation-badges{{
            margin-top:8px;
            display:flex; gap:6px; flex-wrap:wrap;
          }}

          .badge-speculation{{
            display:inline-flex; align-items:center; gap:4px; padding:6px 12px;
            border-radius:8px; font-size:12px; font-weight:700; letter-spacing:-0.3px;
            box-shadow:0 3px 8px rgba(0,0,0,0.25); animation:glow 2s ease-in-out infinite;
          }}
          .badge-speculation.hot{{
            background:linear-gradient(135deg,#dc2626,#b91c1c);
            color:#fff; border:2px solid #991b1b;
          }}
          .badge-speculation.permit{{
            background:linear-gradient(135deg,#ea580c,#c2410c);
            color:#fff; border:2px solid #9a3412;
          }}
          @keyframes glow{{
            0%,100%{{box-shadow:0 3px 8px rgba(220,38,38,0.4),0 0 15px rgba(220,38,38,0.2)}}
            50%{{box-shadow:0 3px 12px rgba(220,38,38,0.6),0 0 25px rgba(220,38,38,0.4)}}
          }}

          .year-badge{{
            display:inline-flex; align-items:center; padding:4px 10px;
            border-radius:6px; font-size:12px; font-weight:600;
            background:linear-gradient(135deg,#94a3b8,#64748b);
            color:#fff; box-shadow:0 2px 4px rgba(0,0,0,0.15);
          }}
          .year-badge.very-young{{
            background:linear-gradient(135deg,#22c55e,#16a34a);
          }}
          .year-badge.young{{
            background:linear-gradient(135deg,#3b82f6,#2563eb);
          }}
          .year-badge.new-building{{
            background:linear-gradient(135deg,#f59e0b,#d97706);
            animation:pulse 2s ease-in-out infinite;
          }}
          @keyframes pulse{{
            0%,100%{{transform:scale(1)}}
            50%{{transform:scale(1.05)}}
          }}

          .price-filter{{
            margin-top:16px;
            display:flex; gap:8px; flex-wrap:wrap;
          }}
          .price-btn{{
            display:inline-flex; align-items:center; padding:8px 16px;
            border:1px solid rgba(255,255,255,0.3); border-radius:999px;
            background:rgba(255,255,255,0.2); color:#fff; font-size:13px; font-weight:600;
            cursor:pointer; transition:all 0.2s ease;
          }}
          .price-btn:hover{{
            background:rgba(255,255,255,0.3); transform:translateY(-1px);
          }}
          .price-btn.active{{
            background:#fff; color:var(--primary); border-color:#fff;
          }}

          .search-bar{{
            max-width:960px; width:92vw; margin:16px auto;
            position:relative; display:flex; align-items:center;
          }}
          .search-bar input{{
            width:100%; padding:12px 40px 12px 16px;
            border:1px solid var(--bd); border-radius:12px;
            background:var(--card); color:var(--fg);
            font-size:14px; font-family:inherit;
            box-shadow:0 2px 8px rgba(0,0,0,.05);
            transition:all 0.2s ease;
          }}
          .search-bar input:focus{{
            outline:none; border-color:var(--primary);
            box-shadow:0 0 0 3px rgba(0,122,255,0.1);
          }}
          .search-bar input::placeholder{{
            color:var(--sub);
          }}
          .clear-btn{{
            position:absolute; right:12px;
            width:24px; height:24px;
            border:none; border-radius:50%;
            background:var(--sub); color:#fff;
            font-size:14px; cursor:pointer;
            display:none; align-items:center; justify-content:center;
            transition:all 0.2s ease;
          }}
          .clear-btn:hover{{
            background:var(--fg); transform:scale(1.1);
          }}
          .clear-btn.show{{
            display:flex;
          }}

          footer{{ margin:24px 0 0; color:#6b7280; font-size:12px; text-align:center }}
        </style>
        </head>
        <body>
          <!-- 워터마크 -->
          <div class="watermark"><div class="watermark-text">{"".join(['<span>부태리 ⓒ 2025</span>' for _ in range(100)])}</div></div>
          <div class="page">
            <header class="header">
              <h1>🏆 {area_name} 가격 순위</h1>
              <p class="sub">{escape(now)} 기준 · 상위 {min(total, MAX_RANK_DISPLAY)}위 (전체 {total}개 단지)</p>

              <div class="price-filter">
                <button class="price-btn active" data-range="all">전체</button>
                <button class="price-btn" data-range="under10">10억 이하</button>
                <button class="price-btn" data-range="under15">15억 이하</button>
                <button class="price-btn" data-range="15to25">15억 초과~25억 이하</button>
                <button class="price-btn" data-range="30s">30억대</button>
              </div>
            </header>

            <!-- 검색 바 -->
            <div class="search-bar">
              <input type="text" id="searchInput" placeholder="🔍 단지명 검색..." />
              <button id="searchClear" class="clear-btn">✕</button>
            </div>

            <section class="cards" id="cards">
              <!-- 카드가 JavaScript에서 동적으로 생성됩니다 -->
              <div id="loadingIndicator" style="text-align:center; padding:40px; color:#666;">
                <div style="font-size:24px; margin-bottom:10px;">⏳</div>
                <div>데이터 로딩 중...</div>
              </div>
            </section>

            <footer>© 부태리의 실거래가 모니터링 시스템</footer>
          </div>

        <script>
        // ========== 압축된 데이터 (인덱스 기반) ==========
        const aptNames = {apt_names_json};
        const sigungus = {sigungus_json};
        const dongs = {dongs_json};
        const locationDongs = {location_dongs_json};
        const compressedApts = {apts_json};
        const areaDisplay = "{area_display}";
        const currentYear = new Date().getFullYear();

        console.log('[평형별 순위] 압축 데이터 로드:', compressedApts.length, '개');

        // 투기과열지구 판정 함수
        function checkRegulation(sido, sigungu, locationDong) {{
          let isHot = false, isPermit = false;

          if (sido.includes('서울')) {{ isHot = true; isPermit = true; }}
          else if (sigungu.includes('과천') || sigungu.includes('광명') || sigungu.includes('의왕') || sigungu.includes('하남')) {{
            isHot = true; isPermit = true;
          }}
          else if (sigungu.includes('성남')) {{
            if (sigungu.includes('분당구') || sigungu.includes('수정구') || sigungu.includes('중원구')) {{
              isHot = true; isPermit = true;
            }} else {{
              const bundangDongs = ['구미동','금곡동','대장동','백현동','분당동','서현동','수내동','야탑동','운중동','정자동','판교동','삼평동','동막동','궁내동','율동','매송동'];
              const sujeongDongs = ['고등동','금토동','단대동','복정동','신흥동','양지동','오야동','태평동','신촌동','수진동','창곡동','시흥동','둔전동'];
              const jungwonDongs = ['갈현동','도촌동','상대원동','성남동','은행동','중앙동','하대원동','금광동','여수동'];
              if (bundangDongs.includes(locationDong) || sujeongDongs.includes(locationDong) || jungwonDongs.includes(locationDong)) {{
                isHot = true; isPermit = true;
              }}
            }}
          }}
          else if (sigungu.includes('수원') && (sigungu.includes('영통구') || sigungu.includes('장안구') || sigungu.includes('팔달구'))) {{
            isHot = true; isPermit = true;
          }}
          else if (sigungu.includes('안양') && sigungu.includes('동안구')) {{ isHot = true; isPermit = true; }}
          else if (sigungu.includes('용인') && sigungu.includes('수지구')) {{ isHot = true; isPermit = true; }}

          return {{ isHot, isPermit }};
        }}

        // 가격대 분류 함수
        function getPriceRange(price) {{
          if (price <= 100000) return 'under10';
          if (price <= 150000) return 'under15';
          if (price <= 250000) return '15to25';
          if (price <= 400000) return '30s';
          return 'over30';
        }}

        // 카드 HTML 생성 함수
        function createCardHTML(apt) {{
          const aptName = apt[0] >= 0 ? aptNames[apt[0]] : '';
          const sigungu = apt[1] >= 0 ? sigungus[apt[1]] : '';
          const dong = apt[2] >= 0 ? dongs[apt[2]] : '';
          const locationDong = apt[3] >= 0 ? locationDongs[apt[3]] : '';
          const price = apt[4];
          const date = apt[5] || '-';
          const floor = apt[6] || '';
          const buildYear = apt[7] || '';
          const rank = apt[8];
          const rankChange = apt[9];
          const prevRank = apt[10];
          const priceChange = apt[11];
          const prevPrice = apt[12];
          const sido = apt[13] || '';

          // 순위 배지
          let rankBadge = '';
          if (rankChange !== null) {{
            if (rankChange > 0) rankBadge = `<span class="rank-badge up">[▲${{rankChange}}]</span>`;
            else if (rankChange < 0) rankBadge = `<span class="rank-badge down">[▼${{Math.abs(rankChange)}}]</span>`;
            else rankBadge = '<span class="rank-badge same">[-]</span>';
          }} else {{
            rankBadge = '<span class="rank-badge new">[★NEW]</span>';
          }}

          // 연식 배지
          let yearBadge = '';
          if (buildYear) {{
            if (buildYear === '분양') {{
              yearBadge = '<span class="year-badge new-building">분양</span>';
            }} else {{
              const age = currentYear - parseInt(buildYear);
              if (age <= 5) yearBadge = `<span class="year-badge very-young">${{buildYear}}년 (${{age}}년차)</span>`;
              else if (age <= 10) yearBadge = `<span class="year-badge young">${{buildYear}}년 (${{age}}년차)</span>`;
              else yearBadge = `<span class="year-badge">${{buildYear}}년 (${{age}}년차)</span>`;
            }}
          }}

          // 규제 배지
          const {{ isHot, isPermit }} = checkRegulation(sido, sigungu, locationDong);
          let regulationBadges = '';
          if (isHot) regulationBadges += '<span class="badge-speculation hot">🔥 투기과열지구</span> ';
          if (isPermit) regulationBadges += '<span class="badge-speculation permit">🚧 토지거래허가구역</span>';

          // 순위 정보
          let rankInfo = '';
          if (rankChange !== null) {{
            if (rankChange > 0) rankInfo = `📊 순위: ${{prevRank}}위 → ${{rank}}위 (${{rankChange}}단계 상승)`;
            else if (rankChange < 0) rankInfo = `📊 순위: ${{prevRank}}위 → ${{rank}}위 (${{Math.abs(rankChange)}}단계 하락)`;
            else rankInfo = '📊 순위: 변동 없음';
          }} else {{
            rankInfo = '🆕 신규 진입';
          }}

          // 가격 정보
          let priceInfo = '';
          if (priceChange !== null && prevPrice) {{
            const pct = prevPrice > 0 ? (Math.abs(priceChange) / prevPrice * 100).toFixed(1) : 0;
            if (priceChange > 0) priceInfo = `📈 가격: ${{prevPrice.toLocaleString()}}만원 → ${{price.toLocaleString()}}만원 (+${{priceChange.toLocaleString()}}만원, +${{pct}}%)`;
            else if (priceChange < 0) priceInfo = `📉 가격: ${{prevPrice.toLocaleString()}}만원 → ${{price.toLocaleString()}}만원 (${{priceChange.toLocaleString()}}만원, -${{pct}}%)`;
            else priceInfo = '📈 가격: 변동 없음';
          }} else {{
            priceInfo = `📈 가격: ${{price.toLocaleString()}}만원 (신규)`;
          }}

          // 카카오맵 쿼리 생성
          // 괄호 처리: "숫자+동" 패턴이면 괄호 무시, 그 외 한글은 추출
          let aptNameForSearch = aptName;
          if (aptName.includes('(')) {{
            const match = aptName.match(/\(([^)]+)\)/);
            if (match) {{
              const parenContent = match[1];
              // "숫자+동" 패턴 체크 (예: 12동, 13동)
              if (/\d+동/.test(parenContent)) {{
                // 숫자+동 패턴이면 괄호 전체 무시
                aptNameForSearch = aptName.split('(')[0].trim();
              }} else if (/[가-힣]/.test(parenContent)) {{
                const koreanOnly = parenContent.replace(/[^가-힣\s]/g, '').trim();
                aptNameForSearch = aptName.split('(')[0].trim() + ' ' + koreanOnly;
              }} else {{
                aptNameForSearch = aptName.split('(')[0].trim();
              }}
            }}
          }}
          const kakaoQuery = `${{sigungu}} ${{locationDong}} ${{aptNameForSearch}} 아파트`.trim();

          const priceRange = getPriceRange(price);
          const location = `${{sido}} ${{sigungu}} ${{locationDong}}`;
          const floorStr = floor ? `${{floor}}층` : '-';
          const dongStr = dong && dong !== '-' ? dong : '';

          return `
            <div class="card" data-price-range="${{priceRange}}" data-kakao-query="${{kakaoQuery}}" onclick="openKakaoMap(this)" style="cursor:pointer;">
              <div class="card-header">
                <h3>🏆 ${{rank}}위 ${{rankBadge}} ${{aptName}} ${{yearBadge}} 전용면적 ${{areaDisplay}}</h3>
                ${{regulationBadges ? `<div class="regulation-badges">${{regulationBadges}}</div>` : ''}}
              </div>
              <div class="divider"></div>
              <div class="card-body">
                <div class="info-row">💰 신고가: <strong>${{price.toLocaleString()}}만원</strong></div>
                <div class="info-row">${{rankInfo}}</div>
                <div class="info-row">${{priceInfo}}</div>
                <div class="info-row">📍 ${{location}} <span style="color:#007AFF; font-size:0.9em;">🗺️ 지도보기</span></div>
                <div class="info-row">📅 ${{date}} | ${{floorStr}}${{dongStr ? ' | ' + dongStr : ''}}</div>
              </div>
            </div>
          `;
        }}

        // 카카오맵 열기 함수
        function openKakaoMap(element) {{
          const query = element.getAttribute('data-kakao-query');
          if (query) {{
            const url = 'https://map.kakao.com/?q=' + encodeURIComponent(query);
            window.open(url, '_blank', 'width=1200,height=800');
          }}
        }}

        // 카드 렌더링 (Lazy Loading - 청크 단위)
        let renderedCount = 0;
        const CHUNK_SIZE = 100;
        let allCards = [];
        let currentRange = 'all';
        let searchQuery = '';

        function renderChunk() {{
          const cardsContainer = document.getElementById('cards');
          const end = Math.min(renderedCount + CHUNK_SIZE, compressedApts.length);

          for (let i = renderedCount; i < end; i++) {{
            const cardHTML = createCardHTML(compressedApts[i]);
            allCards.push(cardHTML);
          }}

          if (renderedCount === 0) {{
            // 첫 번째 청크: 로딩 표시 제거하고 카드 표시
            cardsContainer.innerHTML = allCards.join('');
          }} else {{
            // 추가 청크: 기존에 추가
            cardsContainer.innerHTML = allCards.join('');
          }}

          renderedCount = end;

          if (renderedCount < compressedApts.length) {{
            // 다음 청크 예약
            requestAnimationFrame(renderChunk);
          }} else {{
            // 렌더링 완료 - 필터 초기화
            applyFilters();
            console.log('[평형별 순위] 렌더링 완료:', renderedCount, '개');
          }}
        }}

        function applyFilters() {{
          const cards = document.querySelectorAll('.card');
          let visibleCount = 0;

          cards.forEach(card => {{
            const priceMatch = currentRange === 'all' || card.dataset.priceRange === currentRange;
            let searchMatch = true;
            if (searchQuery) {{
              const cardHeader = card.querySelector('.card-header h3');
              const cardText = cardHeader ? cardHeader.textContent.toLowerCase() : '';
              searchMatch = cardText.includes(searchQuery.toLowerCase());
            }}

            if (priceMatch && searchMatch) {{
              card.style.display = '';
              visibleCount++;
            }} else {{
              card.style.display = 'none';
            }}
          }});

          const subText = document.querySelector('.header .sub');
          if (subText) {{
            const now = subText.textContent.split(' 기준')[0];
            subText.textContent = `${{now}} 기준 · 총 ${{visibleCount}}개 아파트`;
          }}
        }}

        function applyPriceFilter(range) {{
          currentRange = range;
          document.querySelectorAll('.price-btn').forEach(btn => {{
            btn.classList.toggle('active', btn.dataset.range === range);
          }});
          applyFilters();
        }}

        // 이벤트 리스너 설정
        document.querySelectorAll('.price-btn').forEach(btn => {{
          btn.addEventListener('click', () => applyPriceFilter(btn.dataset.range));
        }});

        const searchInput = document.getElementById('searchInput');
        const searchClear = document.getElementById('searchClear');

        searchInput.addEventListener('input', (e) => {{
          searchQuery = e.target.value.trim();
          searchClear.classList.toggle('show', !!searchQuery);
          applyFilters();
        }});

        searchClear.addEventListener('click', () => {{
          searchInput.value = '';
          searchQuery = '';
          searchClear.classList.remove('show');
          applyFilters();
        }});

        // 페이지 로드 시 렌더링 시작
        window.addEventListener('load', () => {{
          renderChunk();
        }});
        </script>
        </body>
        </html>"""

        return html_content

    def export_notification_html(self, apt_list, *, ask_path=False, silent=False):
        """신고가 알림 HTML 파일 저장 및 열기
        - ask_path=True 이면 저장 위치를 사용자에게 물음
        - silent=True 이면 메시지박스 없이 조용히 저장
        """
        import traceback
        try:
            html_str = self.build_notification_html(apt_list)
            # 저장 경로
            if ask_path:
                initdir = os.path.join(self.download_path, "reports")
                os.makedirs(initdir, exist_ok=True)
                from tkinter import filedialog
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                default = f"부태리신고가_{ts}.html"
                filepath = filedialog.asksaveasfilename(
                    initialdir=initdir, initialfile=default,
                    title="신고가 HTML 저장",
                    defaultextension=".html",
                    filetypes=[("HTML files","*.html"), ("All files","*.*")]
                )
                if not filepath:
                    return None
            else:
                reports_dir = os.path.join(self.download_path, "reports")
                os.makedirs(reports_dir, exist_ok=True)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                filepath = os.path.join(reports_dir, f"부태리신고가_{ts}.html")
            # 쓰기
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(html_str)
            # 열기
            try:
                if os.name == "nt":
                    os.startfile(filepath)  # Windows
                else:
                    webbrowser.open("file://" + filepath)
            except Exception:
                pass
            if not silent:
                messagebox.showinfo("저장 완료", f"HTML 저장: {filepath}")
            logging.info(f"[HTML 저장] {filepath}")
            return filepath
        except Exception as e:
            logging.error(f"HTML 저장 중 오류: {str(e)}")
            logging.error(f"상세 오류: {traceback.format_exc()}")
            if not silent:
                messagebox.showerror("오류", f"HTML 저장 실패: {str(e)}")
            return None

    def export_area_ranking_html(self, area_type):
        """평형별 가격 순위 HTML 다운로드
        area_type: '84' 또는 '59'
        """
        try:
            html_str = self.build_area_ranking_html(area_type)
            if not html_str:
                messagebox.showwarning("알림", "순위 데이터를 생성할 수 없습니다.")
                return None

            # 저장 경로
            reports_dir = os.path.join(self.download_path, "reports")
            os.makedirs(reports_dir, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            area_name = "84평형" if area_type == '84' else "59평형"
            filepath = os.path.join(reports_dir, f"부태리_{area_name}순위_{ts}.html")

            # 쓰기
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(html_str)

            # 열기
            try:
                if os.name == "nt":
                    os.startfile(filepath)  # Windows
                else:
                    webbrowser.open("file://" + filepath)
            except Exception:
                pass

            messagebox.showinfo("저장 완료", f"HTML 저장: {filepath}")
            logging.info(f"[평형별 순위 HTML 저장] {filepath}")
            return filepath
        except Exception as e:
            logging.error(f"평형별 순위 HTML 저장 중 오류: {str(e)}")
            messagebox.showerror("오류", f"HTML 저장 실패: {str(e)}")
            return None

    def export_trade_volume_ranking_html(self):
        """연도별 거래량 순위 HTML 다운로드"""
        # 연도 입력 다이얼로그
        year_dialog = tk.Toplevel(self.root)
        year_dialog.title("거래량 순위 - 연도 선택")
        year_dialog.geometry("300x150")
        year_dialog.transient(self.root)
        year_dialog.grab_set()

        # 화면 중앙에 배치
        year_dialog.update_idletasks()
        screen_width = year_dialog.winfo_screenwidth()
        screen_height = year_dialog.winfo_screenheight()
        x = (screen_width - 300) // 2
        y = (screen_height - 150) // 2
        year_dialog.geometry(f"300x150+{x}+{y}")

        ttk.Label(year_dialog, text="조회할 연도를 입력하세요:", font=("", 10)).pack(pady=(20, 10))

        year_var = tk.StringVar(value=str(datetime.now().year))
        year_entry = ttk.Entry(year_dialog, textvariable=year_var, font=("", 12), width=10, justify='center')
        year_entry.pack(pady=5)
        year_entry.focus()

        def on_ok():
            try:
                year = int(year_var.get())
                current_year = datetime.now().year
                if year < 2000 or year > current_year:
                    messagebox.showerror("오류", f"유효한 연도를 입력하세요 (2000-{current_year})")
                    return

                year_dialog.destroy()

                # 저장 경로
                reports_dir = os.path.join(self.download_path, "reports")
                os.makedirs(reports_dir, exist_ok=True)

                # 모든 연도를 매번 새로 생성
                html_str = self.build_trade_volume_ranking_html(year)
                if not html_str:
                    messagebox.showwarning("알림", "거래량 데이터를 생성할 수 없습니다.")
                    return

                # 타임스탬프 포함하여 매번 새로 저장
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                filepath = os.path.join(reports_dir, f"부태리_거래량순위_{year}년_{ts}.html")
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(html_str)

                # 열기
                try:
                    if os.name == "nt":
                        os.startfile(filepath)
                    else:
                        webbrowser.open("file://" + filepath)
                except Exception:
                    pass

                messagebox.showinfo("저장 완료", f"HTML 저장: {filepath}")
                logging.info(f"[거래량 순위 HTML 저장] {filepath}")

            except ValueError:
                messagebox.showerror("오류", "숫자를 입력하세요")
            except Exception as e:
                logging.error(f"거래량 순위 HTML 저장 중 오류: {str(e)}")
                messagebox.showerror("오류", f"HTML 저장 실패: {str(e)}")

        ttk.Button(year_dialog, text="확인", command=on_ok).pack(pady=10)
        year_entry.bind('<Return>', lambda e: on_ok())

    def build_trade_volume_ranking_html(self, year):
        """연도별 거래량 순위 HTML 생성 (해당 연도 전체 매매 거래 집계)"""
        from html import escape
        from collections import defaultdict

        # 단지별 거래량 집계 (평형 무관, 해당 연도 전체)
        trade_volumes = defaultdict(lambda: {'count': 0, 'sido': '', 'sigungu': '', 'dong': '', 'build_year': ''})

        # 진행 바 팝업
        progress_window = tk.Toplevel(self.root)
        progress_window.title("거래량 집계 중")
        progress_window.geometry("500x150")
        progress_window.transient(self.root)

        # 화면 중앙에 배치
        progress_window.update_idletasks()
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
        x = (screen_width - 500) // 2
        y = (screen_height - 150) // 2
        progress_window.geometry(f"500x150+{x}+{y}")

        ttk.Label(progress_window, text=f"{year}년 거래량 집계 중...",
                 font=("", 12, "bold")).pack(pady=(20, 10))

        progress_label = ttk.Label(progress_window, text="준비 중...")
        progress_label.pack(pady=5)

        progress_bar = ttk.Progressbar(progress_window, orient="horizontal",
                                      length=450, mode="determinate", maximum=100)
        progress_bar.pack(fill="x", padx=25, pady=10)

        # 단지별로 그룹화 (평형 무관하게 합산하기 위해)
        complex_dict = {}
        for apt in self.monitored_apts:
            apt_name = apt.get('apt_name', '')
            if apt_name not in complex_dict:
                complex_dict[apt_name] = {
                    'sigungu_code': apt.get('sigungu_code'),
                    'sido': apt.get('sido', ''),
                    'sigungu': apt.get('sigungu', ''),
                    'dong': apt.get('dong', ''),
                    'build_year': apt.get('build_year', '')
                }

        total = len(complex_dict)
        processed = 0

        # 단지별로 해당 연도 전체 거래 조회 (평형 무관)
        for apt_name, info in complex_dict.items():
            processed += 1
            if processed % 5 == 0:
                progress = (processed / total) * 100
                progress_bar['value'] = progress
                progress_label.config(text=f"집계 중: {processed}/{total} - {apt_name}")
                progress_window.update_idletasks()

            sigungu_code = info['sigungu_code']
            dong = info['dong']

            # 해당 연도의 12개월 데이터 조회 (캐시 활용)
            for month in range(1, 13):
                deal_ymd = f"{year}{month:02d}"

                existing_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'existing')
                new_data = self.get_cached_api_data(sigungu_code, deal_ymd, 'new')

                # ★★★ 평형 무관하게 해당 단지의 모든 거래 카운트 ★★★
                for item in existing_data + new_data:
                    if item.get('apt_name', '').strip() == apt_name and item.get('dong', '').strip() == dong:
                        trade_volumes[apt_name]['count'] += 1
                        trade_volumes[apt_name]['sido'] = info['sido']
                        trade_volumes[apt_name]['sigungu'] = info['sigungu']
                        trade_volumes[apt_name]['dong'] = info['dong']
                        trade_volumes[apt_name]['build_year'] = info['build_year']

        progress_bar['value'] = 100
        progress_label.config(text=f"완료: {total}/{total}")
        progress_window.update_idletasks()
        self.root.after(300, progress_window.destroy)

        # 거래량 순으로 정렬
        sorted_apts = sorted(trade_volumes.items(), key=lambda x: x[1]['count'], reverse=True)

        if not sorted_apts:
            return None

        # 지역별로 분류하여 상위 20개씩만 선택
        seoul_apts = []
        gyeonggi_apts = []
        incheon_apts = []

        for apt_name, info in sorted_apts:
            sido = info['sido']
            if "서울" in sido and len(seoul_apts) < 20:
                seoul_apts.append((apt_name, info))
            elif "경기" in sido and len(gyeonggi_apts) < 20:
                gyeonggi_apts.append((apt_name, info))
            elif "인천" in sido and len(incheon_apts) < 20:
                incheon_apts.append((apt_name, info))

            # 세 지역 모두 20개씩 채워지면 조기 종료
            if len(seoul_apts) >= 20 and len(gyeonggi_apts) >= 20 and len(incheon_apts) >= 20:
                break

        # 선택된 아파트들을 하나로 합치고 거래량 순으로 재정렬
        selected_apts = seoul_apts + gyeonggi_apts + incheon_apts
        selected_apts.sort(key=lambda x: x[1]['count'], reverse=True)

        if not selected_apts:
            return None

        now = datetime.now().strftime('%Y-%m-%d %H:%M')

        # HTML 생성 - data 속성에 지역 정보 추가
        cards_html = ""
        for rank, (apt_name, info) in enumerate(selected_apts, 1):
            count = info['count']
            sido = info['sido']
            sigungu = info['sigungu']
            dong_info = info['dong']
            build_year = info['build_year']

            # 연식 처리
            build_year_str = ""
            if build_year:
                build_year_clean = str(build_year).strip()
                if build_year_clean == '분양':
                    build_year_str = f'<span class="year-badge new-building">분양</span>'
                else:
                    try:
                        year_num = int(build_year_clean)
                        age = datetime.now().year - year_num
                        if age <= 5:
                            build_year_str = f'<span class="year-badge very-young">{build_year_clean}년 ({age}년차)</span>'
                        elif age <= 10:
                            build_year_str = f'<span class="year-badge young">{build_year_clean}년 ({age}년차)</span>'
                        else:
                            build_year_str = f'<span class="year-badge">{build_year_clean}년 ({age}년차)</span>'
                    except:
                        build_year_str = f'<span class="year-badge">{build_year_clean}년</span>'

            location = f"{sido} {sigungu} {dong_info}"

            # 순위 배지
            rank_badge_class = ""
            rank_emoji = "🏆"
            if rank == 1:
                rank_badge_class = "rank-gold"
                rank_emoji = "🥇"
            elif rank == 2:
                rank_badge_class = "rank-silver"
                rank_emoji = "🥈"
            elif rank == 3:
                rank_badge_class = "rank-bronze"
                rank_emoji = "🥉"

            # 지역 구분 (서울/경기/인천)
            region_class = ""
            if "서울" in sido:
                region_class = "seoul"
            elif "경기" in sido:
                region_class = "gyeonggi"
            elif "인천" in sido:
                region_class = "incheon"

            cards_html += f"""
            <div class="card" data-region="{region_class}" data-count="{count}" data-rank="{rank}">
              <div class="card-header {rank_badge_class}">
                <h3><span class="rank-number">{rank_emoji} {rank}위</span> {escape(apt_name)} {build_year_str}</h3>
              </div>
              <div class="divider"></div>
              <div class="card-body">
                <div class="info-row">📊 거래량: <strong>{count}건</strong></div>
                <div class="info-row">📍 {escape(location)}</div>
              </div>
            </div>
            """

        html_content = f"""<!DOCTYPE html>
        <html lang="ko">
        <head>
        <meta charset="utf-8"/>
        <meta content="width=device-width, initial-scale=1" name="viewport"/>
        <title>{year}년 거래량 순위 - {escape(now)}</title>
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

          .region-filter{{
            margin-top:16px;
            display:flex; gap:8px; flex-wrap:wrap;
          }}
          .region-btn{{
            display:inline-flex; align-items:center; padding:8px 16px;
            border:1px solid rgba(255,255,255,0.3); border-radius:999px;
            background:rgba(255,255,255,0.2); color:#fff; font-size:13px; font-weight:600;
            cursor:pointer; transition:all 0.2s ease;
          }}
          .region-btn:hover{{
            background:rgba(255,255,255,0.3); transform:translateY(-1px);
          }}
          .region-btn.active{{
            background:#fff; color:var(--primary); border-color:#fff;
          }}

          .cards{{ display:flex; flex-direction:column; gap:16px; margin-top:16px }}
          .card{{
            background:var(--card); border:1px solid var(--bd);
            border-radius:14px; box-shadow:0 2px 12px rgba(0,0,0,.04);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            overflow:hidden;
          }}
          .card:hover{{ transform: translateY(-2px); box-shadow:0 4px 20px rgba(0,0,0,.08) }}

          .card-header{{
            background:linear-gradient(135deg,#f8fafc,#e2e8f0);
            padding:16px 20px;
            border-bottom:2px solid #cbd5e1;
          }}
          .card-header.rank-gold{{
            background:linear-gradient(135deg,#fef3c7,#fbbf24);
            border-bottom:2px solid #f59e0b;
          }}
          .card-header.rank-silver{{
            background:linear-gradient(135deg,#f1f5f9,#cbd5e1);
            border-bottom:2px solid #94a3b8;
          }}
          .card-header.rank-bronze{{
            background:linear-gradient(135deg,#fed7aa,#fb923c);
            border-bottom:2px solid #ea580c;
          }}
          .card-header h3{{
            margin:0; font-size:18px; font-weight:700;
            display:flex; align-items:center; gap:8px; flex-wrap:wrap;
          }}

          .divider{{
            height:2px;
            background:repeating-linear-gradient(
              90deg,
              #cbd5e1 0px,
              #cbd5e1 10px,
              transparent 10px,
              transparent 20px
            );
          }}

          .card-body{{
            padding:16px 20px;
            display:flex; flex-direction:column; gap:10px;
          }}
          .info-row{{
            font-size:14px; line-height:1.6;
            display:flex; align-items:center; gap:8px;
          }}
          .info-row strong{{ color:var(--primary-dark); font-weight:700; font-size:16px; }}

          .year-badge{{
            display:inline-flex; align-items:center; padding:4px 10px;
            border-radius:6px; font-size:12px; font-weight:600;
            background:linear-gradient(135deg,#94a3b8,#64748b);
            color:#fff; box-shadow:0 2px 4px rgba(0,0,0,0.15);
          }}
          .year-badge.very-young{{
            background:linear-gradient(135deg,#22c55e,#16a34a);
          }}
          .year-badge.young{{
            background:linear-gradient(135deg,#3b82f6,#2563eb);
          }}
          .year-badge.new-building{{
            background:linear-gradient(135deg,#f59e0b,#d97706);
            animation:pulse 2s ease-in-out infinite;
          }}
          @keyframes pulse{{
            0%,100%{{transform:scale(1)}}
            50%{{transform:scale(1.05)}}
          }}
        </style>
        </head>
        <body>
        <div class="page">
          <div class="header">
            <h1>📊 {year}년 거래량 순위</h1>
            <p class="sub">{escape(now)} 기준 · 평형 무관 · <span id="countDisplay">전체</span></p>
            <div class="region-filter">
              <button class="region-btn active" data-region="all">전체</button>
              <button class="region-btn" data-region="seoul">서울</button>
              <button class="region-btn" data-region="gyeonggi">경기</button>
              <button class="region-btn" data-region="incheon">인천</button>
            </div>
          </div>
          <div class="cards">
            {cards_html}
          </div>
        </div>

        <script>
        const regionBtns = document.querySelectorAll('.region-btn');
        const cards = document.querySelectorAll('.card');
        const countDisplay = document.getElementById('countDisplay');
        let currentRegion = 'all';

        function applyFilter() {{
          let visibleCards = [];

          cards.forEach(card => {{
            const cardRegion = card.dataset.region;
            if(currentRegion === 'all' || cardRegion === currentRegion) {{
              card.style.display = '';
              visibleCards.push(card);
            }} else {{
              card.style.display = 'none';
            }}
          }});

          // 순위 재계산 및 표시 (상위 20개만)
          visibleCards.forEach((card, index) => {{
            const rankNumber = card.querySelector('.rank-number');
            const originalRank = parseInt(card.dataset.rank);
            const newRank = index + 1;

            // 20위 이후는 숨김
            if(newRank > 20) {{
              card.style.display = 'none';
              return;
            }}

            // 순위 이모지 결정
            let rankEmoji = '🏆';
            if(newRank === 1) rankEmoji = '🥇';
            else if(newRank === 2) rankEmoji = '🥈';
            else if(newRank === 3) rankEmoji = '🥉';

            rankNumber.textContent = `${{rankEmoji}} ${{newRank}}위`;

            // 헤더 클래스 변경
            const header = card.querySelector('.card-header');
            header.classList.remove('rank-gold', 'rank-silver', 'rank-bronze');
            if(newRank === 1) header.classList.add('rank-gold');
            else if(newRank === 2) header.classList.add('rank-silver');
            else if(newRank === 3) header.classList.add('rank-bronze');
          }});

          // 카운트 표시 업데이트
          const regionNames = {{
            'all': '전체',
            'seoul': '서울',
            'gyeonggi': '경기',
            'incheon': '인천'
          }};
          const displayCount = Math.min(visibleCards.length, 20);
          countDisplay.textContent = `${{regionNames[currentRegion]}} TOP ${{displayCount}}`;
        }}

        regionBtns.forEach(btn => {{
          btn.addEventListener('click', () => {{
            regionBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentRegion = btn.dataset.region;
            applyFilter();
          }});
        }});

        // 초기 카운트 표시
        applyFilter();
        </script>
        </body>
        </html>
        """

        return html_content

    def export_price_distribution_html(self):
        """모니터링 중인 단지들의 저장된 거래 데이터를 기반으로 가격대별 분위 HTML 생성"""
        try:
            # 현재 활성 리스트 이름 확인 - self.active_list (tk.StringVar) 사용
            list_name = self.active_list.get() if hasattr(self, 'active_list') else ""
            logging.info(f"[가격대별 분위] active_list에서 가져온 리스트: '{list_name}'")

            # 서울 수도권인 경우에만 지역 선택 다이얼로그 표시
            region_filter = None  # None이면 전체
            logging.info(f"[가격대별 분위] 현재 리스트: '{list_name}'")

            # 서울과 수도권이 모두 포함된 경우에만 True
            is_seoul_sudogwon = ("서울" in list_name) and ("수도권" in list_name)
            logging.info(f"[가격대별 분위] 서울수도권 여부: {is_seoul_sudogwon}")

            if is_seoul_sudogwon:
                region_filter = self._show_region_selection_dialog()
                if region_filter == "cancel":
                    return  # 사용자가 취소함

            # 저장 경로
            reports_dir = os.path.join(self.download_path, "reports")
            os.makedirs(reports_dir, exist_ok=True)

            # HTML 생성 (DB에서 거래 데이터 조회)
            html_str = self.build_price_distribution_html(region_filter=region_filter)

            if not html_str:
                messagebox.showwarning("알림", "가격대별 거래 데이터를 생성할 수 없습니다.")
                return

            # 타임스탬프 포함하여 저장
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            # 지역명 결정
            if region_filter:
                region_name = region_filter
            elif "서울" in list_name and "수도권" in list_name:
                region_name = "서울수도권"
            else:
                region_name = list_name.replace(" ", "_")
            filepath = os.path.join(reports_dir, f"{region_name}_가격대별_거래분위_{ts}.html")
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(html_str)

            # 열기
            try:
                if os.name == "nt":
                    os.startfile(filepath)
                else:
                    webbrowser.open("file://" + filepath)
            except Exception:
                pass

            messagebox.showinfo("저장 완료", f"HTML 저장: {filepath}\n\n모니터링 중인 단지들의 거래 데이터를 기반으로 생성되었습니다.")
            logging.info(f"[가격대별 분위 HTML 저장] {filepath}")

        except Exception as e:
            logging.error(f"가격대별 분위 HTML 저장 중 오류: {str(e)}")
            messagebox.showerror("오류", f"HTML 저장 실패: {str(e)}")

    def _show_region_selection_dialog(self):
        """서울 수도권 리스트일 때 지역 선택 다이얼로그 표시"""
        dialog = tk.Toplevel(self.root)
        dialog.title("지역 선택")
        dialog.geometry("300x250")
        dialog.transient(self.root)
        dialog.grab_set()

        # 화면 중앙에 위치
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 300) // 2
        y = (dialog.winfo_screenheight() - 250) // 2
        dialog.geometry(f"+{x}+{y}")

        result = {"value": "cancel"}

        ttk.Label(dialog, text="가격대별 분위 분석 지역을 선택하세요",
                  font=("맑은 고딕", 11, "bold")).pack(pady=15)

        def select_region(region):
            result["value"] = region
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10, fill="x", padx=20)

        ttk.Button(btn_frame, text="🗺️ 전체 (서울+경기+인천)",
                   command=lambda: select_region(None), width=30).pack(pady=5)
        ttk.Button(btn_frame, text="🏙️ 서울만",
                   command=lambda: select_region("서울"), width=30).pack(pady=5)
        ttk.Button(btn_frame, text="🏘️ 경기도만",
                   command=lambda: select_region("경기"), width=30).pack(pady=5)
        ttk.Button(btn_frame, text="🌊 인천만",
                   command=lambda: select_region("인천"), width=30).pack(pady=5)

        ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=20, pady=10)

        ttk.Button(dialog, text="취소",
                   command=lambda: select_region("cancel")).pack(pady=5)

        dialog.wait_window()
        return result["value"]

    def build_price_distribution_html(self, region_filter=None):
        """현재 모니터링 중인 아파트 리스트의 거래 데이터를 사용하여 HTML 생성

        Args:
            region_filter: 지역 필터 (None=전체, "서울", "경기", "인천")
        """
        from html import escape

        # 모니터링 중인 아파트가 없으면 종료
        if not self.monitored_apts:
            messagebox.showwarning("알림", "모니터링 중인 아파트가 없습니다.")
            return None

        # 전체 거래 데이터 리스트
        all_trades = []
        today = datetime.now()

        # 최근 6개월 기준 날짜 계산
        from dateutil.relativedelta import relativedelta
        six_months_ago = today - relativedelta(months=6)
        logging.info(f"[가격대별 분위] 거래 데이터 기간: {six_months_ago.strftime('%Y-%m-%d')} ~ {today.strftime('%Y-%m-%d')} (최근 6개월)")
        if region_filter:
            logging.info(f"[가격대별 분위] 지역 필터: {region_filter}")

        # 현재 활성 리스트 이름 가져오기 - self.active_list (tk.StringVar) 사용
        list_name = self.active_list.get() if hasattr(self, 'active_list') else "모니터링 리스트"

        # 지역 필터가 있으면 리스트 이름에 추가
        if region_filter:
            list_name = f"{list_name} ({region_filter})"

        logging.info(f"[가격대별 분위] 활성 리스트: {list_name}")
        logging.info(f"[가격대별 분위] 모니터링 중인 아파트 수: {len(self.monitored_apts)}개")

        try:
            # 중복 제거를 위한 Set
            seen_trades = set()

            # 지역 필터 디버깅: 첫 번째 아파트 정보 확인
            if region_filter and self.monitored_apts:
                sample_apt = self.monitored_apts[0]
                logging.info(f"[가격대별 분위] 샘플 아파트 키: {list(sample_apt.keys())}")
                logging.info(f"[가격대별 분위] 샘플 sigungu: '{sample_apt.get('sigungu', '')}'")
                logging.info(f"[가격대별 분위] 샘플 address: '{sample_apt.get('address', '')}'")
                logging.info(f"[가격대별 분위] 샘플 sido: '{sample_apt.get('sido', '')}'")

            # 각 모니터링 아파트의 거래 데이터 수집
            for apt_info in self.monitored_apts:
                apt_name = apt_info.get('apt_name', '')
                sigungu = apt_info.get('sigungu', '')
                dong = apt_info.get('dong', '')
                address = apt_info.get('address', '')  # 주소 필드도 확인
                sido = apt_info.get('sido', '')  # 시도 필드도 확인
                trade_data = apt_info.get('trade_data', [])

                if not trade_data:
                    continue

                # 지역 필터 적용 - 주소(address), 시도(sido), 시군구(sigungu) 순서로 체크
                if region_filter:
                    # 지역 정보 확인을 위한 문자열 결합
                    location_str = f"{sido} {sigungu} {address}".lower()

                    if region_filter == "서울":
                        if "서울" not in location_str:
                            continue
                    elif region_filter == "경기":
                        if "경기" not in location_str:
                            continue
                    elif region_filter == "인천":
                        if "인천" not in location_str:
                            continue

                # 해당 아파트의 최고가 찾기 (신고가 판별용)
                max_price_for_apt = 0
                for t in trade_data:
                    try:
                        p = t.get('price', 0)
                        if isinstance(p, (int, float)) and p > max_price_for_apt:
                            max_price_for_apt = int(p)
                    except:
                        pass

                # 거래 데이터 파싱
                for trade in trade_data:
                    try:
                        # 거래 날짜
                        trade_date = trade.get('date')
                        if not trade_date:
                            continue

                        # 날짜 형식 통일
                        if isinstance(trade_date, str):
                            trade_date_str = trade_date.replace('.', '-').replace('/', '-')
                            if ' ' in trade_date_str:
                                trade_date_str = trade_date_str.split()[0]
                        elif isinstance(trade_date, datetime):
                            trade_date_str = trade_date.strftime('%Y-%m-%d')
                        else:
                            trade_date_str = str(trade_date)

                        # 최근 6개월 이내 거래만 포함
                        try:
                            trade_dt = datetime.strptime(trade_date_str, '%Y-%m-%d')
                            if trade_dt < six_months_ago:
                                continue  # 6개월 이전 거래는 스킵
                        except:
                            pass  # 날짜 파싱 실패시 일단 포함

                        # 가격
                        price = trade.get('price', 0)
                        if not price:
                            continue

                        deal_amount = int(price)

                        # 면적, 층, 동
                        area = trade.get('area', '')
                        floor = trade.get('floor', '')
                        trade_dong = trade.get('dong', '')

                        # dict가 아닌지 확인하고 안전하게 변환
                        if isinstance(area, dict):
                            logging.warning(f"area is dict in trade: {area} for {apt_name}")
                            area = ''
                        if isinstance(floor, dict):
                            logging.warning(f"floor is dict in trade: {floor} for {apt_name}")
                            floor = ''
                        if isinstance(trade_dong, dict):
                            logging.warning(f"trade_dong is dict in trade: {trade_dong} for {apt_name}")
                            trade_dong = ''

                        # 거래 고유 식별자 생성
                        trade_key = (sigungu, apt_name, dong, trade_date_str, deal_amount,
                                    str(area), str(floor), str(trade_dong))

                        # 중복 체크
                        if trade_key in seen_trades:
                            continue

                        seen_trades.add(trade_key)

                        # 신고가 여부 판별
                        is_highest = (deal_amount == max_price_for_apt)

                        # 거래 데이터 추가
                        all_trades.append({
                            'date': trade_date_str,
                            'price': deal_amount,
                            'apt_name': apt_name,
                            'dong': dong,
                            'sigungu': sigungu,
                            'area': str(area) if area else '',
                            'floor': str(floor) if floor else '',
                            'trade_dong': str(trade_dong) if trade_dong else '',
                            'is_highest': is_highest
                        })

                    except Exception as e:
                        logging.error(f"거래 데이터 파싱 중 오류: {str(e)} - 아파트: {apt_name}")
                        continue

            logging.info(f"[가격대별 분위] 중복 제거 후 데이터: {len(all_trades)}건")

            # 최종 데이터 통계
            if all_trades:
                final_prices = [t['price'] for t in all_trades]
                final_min = min(final_prices)
                final_max = max(final_prices)
                logging.info(f"[가격대별 분위] 최종 가격 범위: {final_min:,}만원 ~ {final_max:,}만원 ({final_min/10000:.1f}억 ~ {final_max/10000:.1f}억)")

                # 지역별 통계
                regions = {}
                for t in all_trades:
                    region = t['sigungu']
                    regions[region] = regions.get(region, 0) + 1
                logging.info(f"[가격대별 분위] 지역별 분포: {regions}")

                # 샘플 데이터 로깅 (처음 5건)
                logging.info(f"[가격대별 분위] 샘플 데이터:")
                for i, trade in enumerate(sorted(all_trades, key=lambda x: x['price'], reverse=True)[:5]):
                    logging.info(f"  {i+1}. {trade['sigungu']} {trade['apt_name']} - {trade['price']:,}만원 - {trade['date']}")

            if len(all_trades) == 0:
                messagebox.showwarning("알림", "유효한 거래 데이터가 없습니다.\n먼저 '데이터 갱신'을 실행해주세요.")
                return None

        except Exception as e:
            logging.error(f"가격대별 분위 데이터 수집 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            return None

        # HTML 생성
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # ========== 인덱스 기반 데이터 압축 (용량 최적화) ==========
        # 중복되는 문자열을 인덱스 테이블로 분리하여 용량 대폭 감소
        apt_name_set = {}  # apt_name -> index
        sigungu_set = {}   # sigungu -> index
        dong_set = {}      # dong -> index

        apt_names_list = []
        sigungus_list = []
        dongs_list = []

        def get_index(value, value_set, value_list):
            """값의 인덱스를 반환. 없으면 새로 추가"""
            if not value:
                return -1
            if value not in value_set:
                value_set[value] = len(value_list)
                value_list.append(value)
            return value_set[value]

        # 압축된 거래 데이터 생성
        # 형식: [date, price, apt_idx, dong_idx, sigungu_idx, area, floor, trade_dong, is_highest]
        compressed_trades = []
        for t in all_trades:
            apt_idx = get_index(t['apt_name'], apt_name_set, apt_names_list)
            sigungu_idx = get_index(t['sigungu'], sigungu_set, sigungus_list)
            dong_idx = get_index(t['dong'], dong_set, dongs_list)

            # 면적과 층은 숫자만 추출
            area_val = t.get('area', '')
            floor_val = t.get('floor', '')
            trade_dong_val = t.get('trade_dong', '')

            compressed_trades.append([
                t['date'],           # 0: 날짜
                t['price'],          # 1: 가격
                apt_idx,             # 2: 아파트명 인덱스
                dong_idx,            # 3: 동 인덱스
                sigungu_idx,         # 4: 시군구 인덱스
                area_val,            # 5: 면적
                floor_val,           # 6: 층
                trade_dong_val,      # 7: 거래동
                1 if t.get('is_highest') else 0  # 8: 신고가 여부 (1/0)
            ])

        # JSON 변환
        apt_names_json = json.dumps(apt_names_list, ensure_ascii=False)
        sigungus_json = json.dumps(sigungus_list, ensure_ascii=False)
        dongs_json = json.dumps(dongs_list, ensure_ascii=False)
        trades_json = json.dumps(compressed_trades, ensure_ascii=False)

        # 압축률 로깅
        original_size = len(json.dumps(all_trades, ensure_ascii=False))
        compressed_size = len(apt_names_json) + len(sigungus_json) + len(dongs_json) + len(trades_json)
        compression_ratio = (1 - compressed_size / original_size) * 100
        logging.info(f"[가격대별 분위] 데이터 압축: {original_size:,} -> {compressed_size:,} bytes ({compression_ratio:.1f}% 감소)")
        # ========== 압축 완료 ==========

        # 데이터 범위 계산 (최소/최대 날짜)
        if all_trades:
            min_date_str = min(t['date'] for t in all_trades)
            max_date_str = max(t['date'] for t in all_trades)
        else:
            min_date_str = ""
            max_date_str = ""

        html_content = f"""
        <!DOCTYPE html>
        <html lang="ko">
        <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{escape(list_name)} 가격대별 거래 분위 분석</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js"></script>
        <style>
          * {{ margin:0; padding:0; box-sizing:border-box; }}
          body {{
            font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;
            background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
            padding:20px;
            min-height:100vh;
          }}
          .container {{
            max-width:1400px;
            margin:0 auto;
            background:#fff;
            border-radius:20px;
            padding:40px;
            box-shadow:0 20px 60px rgba(0,0,0,0.3);
          }}
          .created-date {{
            position:fixed;
            top:10px;
            right:20px;
            background:rgba(255,255,255,0.95);
            padding:8px 16px;
            border-radius:8px;
            font-size:12px;
            color:#666;
            box-shadow:0 2px 10px rgba(0,0,0,0.1);
            z-index:1000;
          }}
          /* 워터마크 */
          .watermark {{
            position:fixed; top:0; left:0; width:100%; height:100%;
            pointer-events:none; z-index:9999; overflow:hidden;
          }}
          .watermark-text {{
            position:absolute; width:300%; height:300%;
            top:-100%; left:-100%;
            display:flex; flex-wrap:wrap; justify-content:center; align-items:center;
            transform:rotate(-30deg);
          }}
          .watermark-text span {{
            font-size:16px; color:rgba(100,100,100,0.08);
            padding:30px 50px; white-space:nowrap;
            user-select:none; -webkit-user-select:none;
            font-weight:600;
          }}
          @media print {{ .watermark {{ display:block !important; }} }}
          .header {{
            text-align:center;
            margin-bottom:40px;
            padding-bottom:30px;
            border-bottom:3px solid #667eea;
          }}
          h1 {{
            font-size:2.5em;
            color:#1a1a2e;
            margin-bottom:15px;
            text-shadow:2px 2px 4px rgba(0,0,0,0.1);
          }}
          .subtitle {{
            font-size:1.1em;
            color:#666;
            margin-bottom:10px;
          }}
          .date-info {{
            font-size:0.95em;
            color:#999;
            font-style:italic;
          }}
          .chart-container {{
            position:relative;
            height:600px;
            margin:30px 0;
          }}
          .stats-grid {{
            display:grid;
            grid-template-columns:repeat(auto-fit, minmax(250px, 1fr));
            gap:20px;
            margin-top:40px;
          }}
          .stat-card {{
            background:linear-gradient(135deg,#667eea,#764ba2);
            color:#fff;
            padding:25px;
            border-radius:15px;
            text-align:center;
            box-shadow:0 5px 15px rgba(0,0,0,0.2);
          }}
          .stat-card h3 {{
            font-size:1.2em;
            margin-bottom:15px;
            opacity:0.9;
          }}
          .stat-card .value {{
            font-size:2.5em;
            font-weight:bold;
            margin-bottom:10px;
          }}
          .stat-card .label {{
            font-size:0.9em;
            opacity:0.8;
          }}
          .period-selector {{
            background:#f8f9fa;
            padding:25px;
            border-radius:15px;
            margin-bottom:30px;
          }}
          .period-selector h3 {{
            margin-bottom:20px;
            color:#333;
            font-size:1.2em;
          }}
          .period-row {{
            display:flex;
            align-items:center;
            gap:15px;
            margin-bottom:15px;
            flex-wrap:wrap;
          }}
          .period-row label {{
            font-weight:600;
            min-width:80px;
            color:#555;
          }}
          .period-row input[type="date"] {{
            padding:8px 12px;
            border:2px solid #ddd;
            border-radius:8px;
            font-size:0.95em;
          }}
          .period-row button {{
            padding:10px 25px;
            background:linear-gradient(135deg,#667eea,#764ba2);
            color:#fff;
            border:none;
            border-radius:8px;
            cursor:pointer;
            font-weight:600;
            transition:transform 0.2s;
          }}
          .period-row button:hover {{
            transform:scale(1.05);
          }}
          .legend-custom {{
            display:flex;
            justify-content:center;
            gap:40px;
            margin-bottom:20px;
            flex-wrap:wrap;
          }}
          .legend-item {{
            display:flex;
            align-items:center;
            gap:10px;
            font-size:1.1em;
          }}
          .legend-color {{
            width:30px;
            height:20px;
            border-radius:5px;
          }}
          .recent {{ background:rgba(102, 126, 234, 0.8); }}
          .previous {{ background:rgba(255, 159, 64, 0.8); }}

          /* Modal styles */
          .modal {{
            display:none;
            position:fixed;
            z-index:1000;
            left:0;
            top:0;
            width:100%;
            height:100%;
            overflow:auto;
            background-color:rgba(0,0,0,0.5);
          }}
          .modal-content {{
            background-color:#fff;
            margin:5% auto;
            border-radius:15px;
            width:90%;
            max-width:900px;
            max-height:80vh;
            overflow:hidden;
            box-shadow:0 20px 60px rgba(0,0,0,0.3);
            display:flex;
            flex-direction:column;
          }}
          .modal-header {{
            padding:20px 30px;
            background:linear-gradient(135deg,#667eea,#764ba2);
            color:#fff;
            display:flex;
            justify-content:space-between;
            align-items:center;
          }}
          .modal-header h2 {{
            margin:0;
            font-size:1.5em;
          }}
          .modal-close {{
            font-size:2em;
            font-weight:bold;
            cursor:pointer;
            color:#fff;
            line-height:1;
            transition:transform 0.2s;
          }}
          .modal-close:hover {{
            transform:scale(1.2);
          }}
          .modal-body {{
            padding:30px;
            overflow-y:auto;
            flex:1;
          }}
          .trade-table {{
            width:100%;
            border-collapse:collapse;
            margin-top:10px;
          }}
          .trade-table th {{
            background:#f8f9fa;
            padding:12px;
            text-align:left;
            font-weight:600;
            color:#333;
            border-bottom:2px solid #ddd;
            position:sticky;
            top:0;
          }}
          .trade-table td {{
            padding:10px 12px;
            border-bottom:1px solid #eee;
          }}
          .trade-table tr:hover {{
            background:#f8f9fa;
          }}
          .trade-table tr.highest-price {{
            background:linear-gradient(90deg, #fff3cd 0%, #ffeeba 100%);
            border-left:4px solid #ffc107;
          }}
          .trade-table tr.highest-price:hover {{
            background:linear-gradient(90deg, #ffecb5 0%, #ffe8a1 100%);
          }}
          .trade-table tr.highest-price td {{
            font-weight:600;
          }}
          .highest-badge {{
            display:inline-block;
            background:linear-gradient(135deg, #f59e0b, #d97706);
            color:#fff;
            padding:2px 8px;
            border-radius:12px;
            font-size:0.75em;
            font-weight:bold;
            margin-left:8px;
            vertical-align:middle;
          }}
          .trade-count {{
            color:#667eea;
            font-weight:bold;
            font-size:1.1em;
            margin-bottom:10px;
          }}
          .highest-count {{
            color:#f59e0b;
            font-weight:bold;
            margin-left:15px;
          }}
          .sort-buttons {{
            display:flex;
            gap:8px;
            margin:10px 0 15px 0;
          }}
          .sort-btn {{
            padding:6px 14px;
            border:2px solid #ddd;
            border-radius:20px;
            background:#fff;
            cursor:pointer;
            font-size:0.85em;
            transition:all 0.2s;
          }}
          .sort-btn:hover {{
            border-color:#667eea;
            color:#667eea;
          }}
          .sort-btn.active {{
            background:linear-gradient(135deg,#667eea,#764ba2);
            color:#fff;
            border-color:transparent;
          }}
        </style>
        </head>
        <body>
        <!-- 워터마크 -->
        <div class="watermark"><div class="watermark-text">{"".join(['<span>부태리 ⓒ 2025</span>' for _ in range(100)])}</div></div>
        <div class="created-date">📅 제작일: {escape(now)}</div>
        <div class="container">
          <div class="header">
            <h1>📊 {escape(list_name)} 가격대별 거래 분위 분석</h1>
            <p class="subtitle">{escape(list_name)} - 최근 3개월 거래 데이터</p>
            <p class="date-info">생성 시간: {escape(now)} | 데이터: {min_date_str} ~ {max_date_str} ({len(all_trades):,}건)</p>
          </div>

          <div class="period-selector">
            <h3>🗓️ 기간 선택</h3>
            <div class="period-row">
              <label>기간 1:</label>
              <input type="date" id="period1Start" value="{(today - timedelta(days=30)).strftime('%Y-%m-%d')}" min="{min_date_str}" max="{max_date_str}">
              <span>~</span>
              <input type="date" id="period1End" value="{today.strftime('%Y-%m-%d')}" min="{min_date_str}" max="{max_date_str}">
            </div>
            <div class="period-row">
              <label>기간 2:</label>
              <input type="date" id="period2Start" value="{(today - timedelta(days=60)).strftime('%Y-%m-%d')}" min="{min_date_str}" max="{max_date_str}">
              <span>~</span>
              <input type="date" id="period2End" value="{(today - timedelta(days=31)).strftime('%Y-%m-%d')}" min="{min_date_str}" max="{max_date_str}">
            </div>
            <div class="period-row">
              <button onclick="updateChart()">📈 그래프 업데이트</button>
              <button onclick="setRecentPeriods()">⏱️ 최근 30일 vs 31-60일</button>
              <button onclick="setMonthlyPeriods()">📅 이번달 vs 지난달</button>
            </div>
          </div>

          <div class="legend-custom">
            <div class="legend-item">
              <div class="legend-color recent"></div>
              <span id="period1Label">기간 1</span>
            </div>
            <div class="legend-item">
              <div class="legend-color previous"></div>
              <span id="period2Label">기간 2</span>
            </div>
          </div>

          <div class="chart-container">
            <canvas id="priceChart"></canvas>
          </div>

          <!-- 거래 내역 팝업 모달 -->
          <div id="tradeModal" class="modal">
            <div class="modal-content">
              <div class="modal-header">
                <h2 id="modalTitle">거래 내역</h2>
                <span class="modal-close" onclick="closeTradeModal()">&times;</span>
              </div>
              <div class="modal-body">
                <div id="tradeList"></div>
              </div>
            </div>
          </div>

          <div class="stats-grid">
            <div class="stat-card">
              <h3 id="stat1Title">기간 1</h3>
              <div class="value" id="period1Total">0</div>
              <div class="label">총 거래 건수</div>
            </div>
            <div class="stat-card">
              <h3 id="stat2Title">기간 2</h3>
              <div class="value" id="period2Total">0</div>
              <div class="label">총 거래 건수</div>
            </div>
            <div class="stat-card">
              <h3>변화율</h3>
              <div class="value" id="changeRate">0%</div>
              <div class="label">기간 대비</div>
            </div>
          </div>
        </div>

        <script>
          // ========== 압축된 거래 데이터 (용량 최적화) ==========
          // 인덱스 테이블
          const aptNames = {apt_names_json};
          const sigungus = {sigungus_json};
          const dongs = {dongs_json};

          // 압축된 거래 데이터: [date, price, apt_idx, dong_idx, sigungu_idx, area, floor, trade_dong, is_highest]
          const compressedTrades = {trades_json};

          let priceChart = null;
          let maxPrice = 200000;
          let maxPriceBracket = 200000;

          console.log('[가격대별 분위] 압축 데이터 로드 완료');
          console.log('[가격대별 분위] 전체 데이터 건수:', compressedTrades.length);
          console.log('[가격대별 분위] 인덱스 테이블: 아파트', aptNames.length, '개, 시군구', sigungus.length, '개, 동', dongs.length, '개');

          // 데이터 체크
          if (!compressedTrades || compressedTrades.length === 0) {{
            console.error('[가격대별 분위] 오류: 거래 데이터가 없습니다!');
            document.querySelector('.container').innerHTML = '<div style="text-align:center; padding:100px;"><h2 style="color:#dc2626;">⚠️ 거래 데이터가 없습니다</h2><p style="margin-top:20px; color:#666;">먼저 "데이터 갱신"을 실행해주세요.</p></div>';
          }} else {{
            // 데이터에서 최대 가격 찾기 (압축 데이터에서 직접 - index 1이 price)
            maxPrice = Math.max(...compressedTrades.map(t => t[1]));
            maxPriceBracket = Math.ceil(maxPrice / 10000) * 10000;
            console.log('[가격대별 분위] 최대 가격:', maxPrice, '만원 (', (maxPrice/10000), '억)');
            console.log('[가격대별 분위] 최대 구간:', maxPriceBracket, '만원 (', (maxPriceBracket/10000), '억)');
          }}

          // 압축 데이터를 원본 형식으로 복원 (Lazy Loading - 모달 표시 시에만 사용)
          function decompressTrade(t) {{
            return {{
              date: t[0],
              price: t[1],
              apt_name: t[2] >= 0 ? aptNames[t[2]] : '',
              dong: t[3] >= 0 ? dongs[t[3]] : '',
              sigungu: t[4] >= 0 ? sigungus[t[4]] : '',
              area: t[5] || '',
              floor: t[6] || '',
              trade_dong: t[7] || '',
              is_highest: t[8] === 1
            }};
          }}

          // 가격대별로 데이터 집계하는 함수 (압축 데이터 직접 사용 - 복원 불필요)
          function aggregateByPriceBracket(trades, startDate, endDate) {{
            const brackets = {{}};
            const start = new Date(startDate);
            const end = new Date(endDate);

            trades.forEach(t => {{
              // t[0] = date, t[1] = price
              const tradeDate = new Date(t[0]);
              if (tradeDate >= start && tradeDate <= end) {{
                // 1억 단위로 가격대 결정
                const bracket = Math.floor(t[1] / 10000) * 10000;
                brackets[bracket] = (brackets[bracket] || 0) + 1;
              }}
            }});

            return brackets;
          }}

          // 가격대 라벨 생성 (동적으로 최대값까지)
          function generatePriceLabels() {{
            const labels = [];
            for (let i = 10000; i <= maxPriceBracket; i += 10000) {{
              labels.push((i / 10000) + '억');
            }}
            console.log('[가격대별 분위] 생성된 라벨 수:', labels.length, '개 (1억~' + (maxPriceBracket/10000) + '억)');
            return labels;
          }}

          // 차트 클릭 이벤트 핸들러
          function handleChartClick(event, activeElements) {{
            if (activeElements.length === 0) return;

            const element = activeElements[0];
            const datasetIndex = element.datasetIndex;
            const index = element.index;

            // 클릭한 라벨에서 가격대 추출 (예: "5억" -> 50000)
            const priceLabel = priceChart.data.labels[index];
            const priceBracket = parseInt(priceLabel.replace('억', '')) * 10000;

            console.log('[클릭] 라벨:', priceLabel, '가격대:', priceBracket, '만원');

            // 해당 기간 가져오기
            let startDate, endDate, periodLabel;
            if (datasetIndex === 0) {{
              startDate = document.getElementById('period1Start').value;
              endDate = document.getElementById('period1End').value;
              periodLabel = document.getElementById('period1Label').textContent;
            }} else {{
              startDate = document.getElementById('period2Start').value;
              endDate = document.getElementById('period2End').value;
              periodLabel = document.getElementById('period2Label').textContent;
            }}

            // 해당 가격대의 거래 필터링 (압축 데이터에서 필터링 후 복원 - Lazy Loading)
            const start = new Date(startDate);
            const end = new Date(endDate);

            // 압축 데이터에서 필터링: t[0]=date, t[1]=price
            const filteredCompressed = compressedTrades.filter(t => {{
              const tradeDate = new Date(t[0]);
              const tradeBracket = Math.floor(t[1] / 10000) * 10000;
              return tradeDate >= start && tradeDate <= end && tradeBracket === priceBracket;
            }});

            // 필터링된 데이터만 복원 (Lazy Loading - 필요한 것만 복원)
            const filteredTrades = filteredCompressed.map(decompressTrade);

            console.log('[클릭] 기간:', startDate, '~', endDate, '필터된 거래:', filteredTrades.length, '건');

            // 모달 표시
            showTradeModal(priceLabel, periodLabel, filteredTrades);
          }}

          // 현재 모달에 표시중인 거래 데이터 (정렬용)
          let currentModalTrades = [];
          let currentSortType = 'date';  // 기본 정렬: 날짜순

          // 거래 내역 모달 표시
          function showTradeModal(priceLabel, periodLabel, trades) {{
            const modal = document.getElementById('tradeModal');
            const modalTitle = document.getElementById('modalTitle');
            const tradeList = document.getElementById('tradeList');

            modalTitle.textContent = `${{priceLabel}} 거래 내역 (${{periodLabel}})`;

            // 현재 거래 데이터 저장
            currentModalTrades = [...trades];
            currentSortType = 'date';  // 기본값 날짜순

            // 테이블 렌더링
            renderTradeTable('date');

            modal.style.display = 'block';
          }}

          // 정렬 방식에 따라 테이블 렌더링
          function renderTradeTable(sortType) {{
            const tradeList = document.getElementById('tradeList');
            const trades = [...currentModalTrades];
            currentSortType = sortType;

            // 신고가 거래 수 계산
            const highestCount = trades.filter(t => t.is_highest).length;

            // 거래 내역 테이블 생성
            let html = `<div class="trade-count">총 ${{trades.length}}건의 거래`;
            if (highestCount > 0) {{
              html += `<span class="highest-count">🏆 신고가 ${{highestCount}}건</span>`;
            }}
            html += `</div>`;

            // 정렬 버튼 추가
            html += `
              <div class="sort-buttons">
                <button class="sort-btn ${{sortType === 'date' ? 'active' : ''}}" onclick="renderTradeTable('date')">📅 날짜순</button>
                <button class="sort-btn ${{sortType === 'price' ? 'active' : ''}}" onclick="renderTradeTable('price')">💰 가격순</button>
                <button class="sort-btn ${{sortType === 'highest' ? 'active' : ''}}" onclick="renderTradeTable('highest')">🏆 신고가 우선</button>
              </div>
            `;

            if (trades.length === 0) {{
              html += '<p style="text-align:center; color:#999; padding:20px;">해당 기간에 거래 내역이 없습니다.</p>';
            }} else {{
              // 정렬 적용
              if (sortType === 'date') {{
                // 날짜순 (최신순)
                trades.sort((a, b) => new Date(b.date) - new Date(a.date));
              }} else if (sortType === 'price') {{
                // 가격순 (높은순)
                trades.sort((a, b) => b.price - a.price);
              }} else if (sortType === 'highest') {{
                // 신고가 우선, 그 다음 날짜순
                trades.sort((a, b) => {{
                  if (a.is_highest && !b.is_highest) return -1;
                  if (!a.is_highest && b.is_highest) return 1;
                  return new Date(b.date) - new Date(a.date);
                }});
              }}

              html += `
                <table class="trade-table">
                  <thead>
                    <tr>
                      <th>거래일</th>
                      <th>시군구</th>
                      <th>동</th>
                      <th>아파트명</th>
                      <th>면적(㎡)</th>
                      <th>층</th>
                      <th>동</th>
                      <th>거래가격</th>
                    </tr>
                  </thead>
                  <tbody>
              `;

              trades.forEach(trade => {{
                const rowClass = trade.is_highest ? 'highest-price' : '';
                const badge = trade.is_highest ? '<span class="highest-badge">신고가</span>' : '';
                const priceColor = trade.is_highest ? '#d97706' : '#667eea';

                html += `
                  <tr class="${{rowClass}}">
                    <td>${{trade.date}}</td>
                    <td>${{trade.sigungu || '-'}}</td>
                    <td>${{trade.dong || '-'}}</td>
                    <td>${{trade.apt_name}}${{badge}}</td>
                    <td>${{trade.area || '-'}}</td>
                    <td>${{trade.floor || '-'}}</td>
                    <td>${{trade.trade_dong || '-'}}</td>
                    <td style="font-weight:bold; color:${{priceColor}};">${{trade.price.toLocaleString()}}만원</td>
                  </tr>
                `;
              }});

              html += `
                  </tbody>
                </table>
              `;
            }}

            tradeList.innerHTML = html;
          }}

          // 모달 닫기
          function closeTradeModal() {{
            document.getElementById('tradeModal').style.display = 'none';
          }}

          // 모달 배경 클릭 시 닫기
          window.onclick = function(event) {{
            const modal = document.getElementById('tradeModal');
            if (event.target === modal) {{
              closeTradeModal();
            }}
          }}

          // 그래프 업데이트 함수
          function updateChart() {{
            const period1Start = document.getElementById('period1Start').value;
            const period1End = document.getElementById('period1End').value;
            const period2Start = document.getElementById('period2Start').value;
            const period2End = document.getElementById('period2End').value;

            // 날짜 유효성 검사
            if (!period1Start || !period1End || !period2Start || !period2End) {{
              alert('모든 날짜를 입력해주세요.');
              return;
            }}

            // 기간 라벨 업데이트
            const period1Label = period1Start + ' ~ ' + period1End;
            const period2Label = period2Start + ' ~ ' + period2End;
            document.getElementById('period1Label').textContent = period1Label;
            document.getElementById('period2Label').textContent = period2Label;
            document.getElementById('stat1Title').textContent = period1Label;
            document.getElementById('stat2Title').textContent = period2Label;

            // 데이터 집계 (선택한 기간 내 거래만 집계됨)
            const period1Data = aggregateByPriceBracket(compressedTrades, period1Start, period1End);
            const period2Data = aggregateByPriceBracket(compressedTrades, period2Start, period2End);

            console.log('[가격대별 분위] period1Data 키:', Object.keys(period1Data));
            console.log('[가격대별 분위] period2Data 키:', Object.keys(period2Data));

            // 선택한 두 기간 내 실제 거래가 있는 가격대만 추출 (값이 0보다 큰 것만)
            const period1Brackets = Object.entries(period1Data).filter(([k, v]) => v > 0).map(([k, v]) => parseInt(k));
            const period2Brackets = Object.entries(period2Data).filter(([k, v]) => v > 0).map(([k, v]) => parseInt(k));

            // 두 기간의 가격대 합집합 (중복 제거)
            const allBracketsSet = new Set([...period1Brackets, ...period2Brackets]);
            const sortedBrackets = Array.from(allBracketsSet).sort((a, b) => a - b);

            console.log('[가격대별 분위] 선택 기간 내 실제 거래 가격대:', sortedBrackets.map(b => (b/10000) + '억'));

            if (sortedBrackets.length === 0) {{
              console.warn('[가격대별 분위] 선택한 기간에 거래 데이터가 없습니다.');
              document.querySelector('.chart-container').innerHTML = '<div style="text-align:center; padding:100px; color:#999;">선택한 기간에 거래 데이터가 없습니다.</div>';
              return;
            }}

            // 라벨 및 데이터 배열 생성
            const labels = [];
            const period1Counts = [];
            const period2Counts = [];

            sortedBrackets.forEach(bracket => {{
              labels.push((bracket / 10000) + '억');
              period1Counts.push(period1Data[bracket] || 0);
              period2Counts.push(period2Data[bracket] || 0);
            }});

            console.log('[가격대별 분위] X축 라벨:', labels);
            console.log('[가격대별 분위] 최소~최대:', labels[0], '~', labels[labels.length - 1]);

            console.log('[가격대별 분위] 표시 구간 수:', labels.length, '개');
            console.log('[가격대별 분위] 기간1 데이터:', period1Counts);
            console.log('[가격대별 분위] 기간2 데이터:', period2Counts);

            // 통계 계산 (기간1이 최근, 기간2가 과거)
            const period1Total = period1Counts.reduce((a, b) => a + b, 0);
            const period2Total = period2Counts.reduce((a, b) => a + b, 0);

            // 변화율: (최근 - 과거) / 과거 * 100
            const changeRate = period2Total > 0
              ? ((period1Total - period2Total) / period2Total * 100).toFixed(1)
              : (period1Total > 0 ? '+100.0' : '0.0');

            document.getElementById('period1Total').textContent = period1Total.toLocaleString();
            document.getElementById('period2Total').textContent = period2Total.toLocaleString();
            document.getElementById('changeRate').textContent = (changeRate > 0 ? '+' : '') + changeRate + '%';

            console.log('[가격대별 분위] 변화율: 기간1(' + period1Total + ') vs 기간2(' + period2Total + ') = ' + changeRate + '%');

            // 차트 업데이트
            if (priceChart) {{
              priceChart.data.labels = labels;
              priceChart.data.datasets[0].label = period1Label;
              priceChart.data.datasets[0].data = period1Counts;
              priceChart.data.datasets[1].label = period2Label;
              priceChart.data.datasets[1].data = period2Counts;
              priceChart.update();
            }} else {{
              const ctx = document.getElementById('priceChart').getContext('2d');
              priceChart = new Chart(ctx, {{
                type: 'bar',
                data: {{
                  labels: labels,
                  datasets: [
                    {{
                      label: period1Label,
                      data: period1Counts,
                      backgroundColor: 'rgba(102, 126, 234, 0.8)',
                      borderColor: 'rgba(102, 126, 234, 1)',
                      borderWidth: 2,
                      borderRadius: 8,
                    }},
                    {{
                      label: period2Label,
                      data: period2Counts,
                      backgroundColor: 'rgba(255, 159, 64, 0.8)',
                      borderColor: 'rgba(255, 159, 64, 1)',
                      borderWidth: 2,
                      borderRadius: 8,
                    }}
                  ]
                }},
                options: {{
                  responsive: true,
                  maintainAspectRatio: false,
                  onClick: handleChartClick,
                  interaction: {{
                    mode: 'nearest',
                    intersect: true,
                  }},
                  plugins: {{
                    legend: {{
                      display: false
                    }},
                    title: {{
                      display: true,
                      text: '가격대별 거래 건수 비교',
                      font: {{
                        size: 18,
                        weight: 'bold'
                      }},
                      padding: {{
                        top: 10,
                        bottom: 30
                      }}
                    }},
                    tooltip: {{
                      backgroundColor: 'rgba(0, 0, 0, 0.8)',
                      padding: 12,
                      titleFont: {{
                        size: 14
                      }},
                      bodyFont: {{
                        size: 13
                      }},
                      callbacks: {{
                        label: function(context) {{
                          let label = context.dataset.label || '';
                          if (label) {{
                            label += ': ';
                          }}
                          label += context.parsed.y + '건';
                          return label;
                        }}
                      }}
                    }}
                  }},
                  scales: {{
                    x: {{
                      grid: {{
                        display: false
                      }},
                      ticks: {{
                        font: {{
                          size: 12
                        }}
                      }}
                    }},
                    y: {{
                      type: 'logarithmic',
                      min: 0.5,  // 1건짜리 바도 보이게 (min:1이면 1건은 높이가 0)
                      grid: {{
                        color: 'rgba(0, 0, 0, 0.05)'
                      }},
                      ticks: {{
                        font: {{
                          size: 12
                        }},
                        callback: function(value) {{
                          // 로그 스케일에서 깔끔한 숫자만 표시
                          if (value === 1 || value === 2 || value === 5 ||
                              value === 10 || value === 20 || value === 50 ||
                              value === 100 || value === 200 || value === 500 ||
                              value === 1000 || value === 2000 || value === 5000) {{
                            return value + '건';
                          }}
                          return '';
                        }}
                      }}
                    }}
                  }}
                }}
              }});
            }}
          }}

          // 빠른 설정 함수들
          function setRecentPeriods() {{
            const today = new Date();
            const date30 = new Date(today.getTime() - 30*24*60*60*1000);
            const date60 = new Date(today.getTime() - 60*24*60*60*1000);
            const date31 = new Date(today.getTime() - 31*24*60*60*1000);

            document.getElementById('period1Start').value = date30.toISOString().split('T')[0];
            document.getElementById('period1End').value = today.toISOString().split('T')[0];
            document.getElementById('period2Start').value = date60.toISOString().split('T')[0];
            document.getElementById('period2End').value = date31.toISOString().split('T')[0];

            updateChart();
          }}

          function setMonthlyPeriods() {{
            const today = new Date();
            console.log('=== setMonthlyPeriods 호출 ===');
            console.log('오늘 날짜:', today);

            // 이번달: 1일부터 오늘까지
            const thisMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
            const thisMonthEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate());

            console.log('이번달 시작:', thisMonthStart);
            console.log('이번달 끝 (오늘):', thisMonthEnd);

            // 지난달: 1일부터 마지막 날까지
            const lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
            const lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0); // 이번달 0일 = 지난달 마지막 날

            console.log('지난달 시작:', lastMonthStart);
            console.log('지난달 끝:', lastMonthEnd);

            // 날짜를 YYYY-MM-DD 형식으로 변환 (로컬 시간 기준)
            const formatDate = (date) => {{
              const year = date.getFullYear();
              const month = String(date.getMonth() + 1).padStart(2, '0');
              const day = String(date.getDate()).padStart(2, '0');
              return `${{year}}-${{month}}-${{day}}`;
            }};

            const period1Start = formatDate(thisMonthStart);
            const period1End = formatDate(thisMonthEnd);
            const period2Start = formatDate(lastMonthStart);
            const period2End = formatDate(lastMonthEnd);

            console.log('Period 1 (이번달):', period1Start, '~', period1End);
            console.log('Period 2 (지난달):', period2Start, '~', period2End);

            document.getElementById('period1Start').value = period1Start;
            document.getElementById('period1End').value = period1End;
            document.getElementById('period2Start').value = period2Start;
            document.getElementById('period2End').value = period2End;

            updateChart();
          }}

          // 페이지 로드 시 초기 차트 생성
          window.addEventListener('load', function() {{
            updateChart();
          }});
        </script>
        </body>
        </html>
        """

        return html_content


    def show_new_max_notification(self, apt_list):
        """신고가 발견 시 9:16 비율 알림 창 (페이지네이션 + 캡처 기능)"""
        if not apt_list:
            return

        # 신고가 높은 순으로 정렬
        apt_list = sorted(apt_list, key=lambda x: x.get('new_price', 0), reverse=True)
        current_time = datetime.now()

        # 로깅 추가
        logging.info(f"[신고가 알림] {len(apt_list)}개 단지 신고가 발생, 시간: {current_time}")
        for apt in apt_list:
            logging.info(f"  - {apt.get('apt_name')}: {apt.get('new_price', 0):,}만원")

        # 메모리에 추가 및 JSON 저장 (r7 방식)
        try:
            # 중복 체크 (최근 1분 이내 동일한 타임스탬프가 있는지)
            is_duplicate = False
            cutoff_time = current_time - timedelta(minutes=1)
            for existing in self.notifications_history:
                existing_ts = existing.get('timestamp', datetime.now())
                if existing_ts > cutoff_time and len(existing.get('apt_list', [])) == len(apt_list):
                    is_duplicate = True
                    break

            if not is_duplicate:
                # 메모리에 추가
                notification_data = {'timestamp': current_time, 'apt_list': apt_list.copy()}
                self.notifications_history.append(notification_data)

                # 최근 50개 그룹만 유지
                if len(self.notifications_history) > 50:
                    self.notifications_history = self.notifications_history[-50:]

                logging.info(f"[메모리 추가] {len(apt_list)}개 단지를 notifications_history에 추가 (총 {len(self.notifications_history)}개 그룹)")

                # 즉시 JSON 파일로 저장
                self.save_notifications_history()
                logging.info(f"[JSON 저장] notifications_history.json 파일 저장 완료")

                # apartments 테이블의 prev_max 필드 업데이트 (핑크색 표시용)
                try:
                    cursor = self.db_conn.cursor()
                    for apt in apt_list:
                        apt_name = apt.get('apt_name', '')
                        cursor.execute("SELECT id FROM apartments WHERE apt_name = ? LIMIT 1", (apt_name,))
                        result = cursor.fetchone()
                        apt_id = result[0] if result else None

                        if apt_id:
                            prev_price = apt.get('old_price', 0) or apt.get('prev_max_price', 0)
                            cursor.execute("""
                                UPDATE apartments
                                SET prev_max_price = ?,
                                    prev_max_date = ?,
                                    prev_max_floor = ?,
                                    prev_max_dong = ?
                                WHERE id = ?
                            """, (
                                prev_price,
                                apt.get('old_date', ''),
                                apt.get('old_floor', ''),
                                apt.get('old_dong', ''),
                                apt_id
                            ))
                    self.db_conn.commit()
                except Exception as db_error:
                    logging.warning(f"apartments 테이블 업데이트 중 오류 (무시됨): {str(db_error)}")
            else:
                logging.info(f"[중복 건너뜀] 최근 1분 이내 동일한 신고가 기록 존재")

        except Exception as e:
            logging.error(f"신고가 히스토리 저장 중 오류: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
        apts_per_page = 2
        total_pages = (len(apt_list) + apts_per_page - 1) // apts_per_page
        current_page = 0
        notification_window = tk.Toplevel(self.root)
        notification_window.title("📱 부태리의 신고가")
        base_width = 400
        base_height = int(base_width * 16 / 9)
        screen_width = notification_window.winfo_screenwidth()
        screen_height = notification_window.winfo_screenheight()
        max_width = min(500, int(screen_width * 0.4))
        max_height = min(900, int(screen_height * 0.85))
        if base_height > max_height:
            window_height = max_height
            window_width = int(window_height * 9 / 16)
        elif base_width > max_width:
            window_width = max_width
            window_height = int(window_width * 16 / 9)
        else:
            window_width = base_width
            window_height = base_height
        min_width = 300
        min_height = int(min_width * 16 / 9)
        if window_width < min_width:
            window_width = min_width
            window_height = min_height
        elif window_height < min_height:
            window_height = min_height
            window_width = int(window_height * 9 / 16)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        notification_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        notification_window.resizable(False, False)
        notification_window.transient(self.root)
        notification_window.grab_set()
        notification_window.attributes('-topmost', True)
        try:
            dpi = notification_window.winfo_fpixels('1i')
            scale_factor = max(1.0, dpi / 96.0)
        except:
            scale_factor = 1.0
        def scaled_font(family, size, weight='normal'):
            scaled_size = max(8, int(size * scale_factor))
            return (family, scaled_size, weight)
        colors = {
            'primary': '#007AFF',
            'primary_dark': '#0051D2',
            'background': '#F2F2F7',
            'surface': '#FFFFFF',
            'card_bg': '#FFFFFF',
            'error': '#FF3B30',
            'error_light': '#FFEBEE',
            'success': '#34C759',
            'text_primary': '#000000',
            'text_secondary': '#8E8E93',
            'separator': '#E5E5EA',
            'capture': '#FF9500'
        }
        main_frame = tk.Frame(notification_window, bg=colors['background'])
        main_frame.pack(fill='both', expand=True)
        header_height = int(window_height * 0.16)
        header_frame = tk.Frame(main_frame, bg=colors['primary'], height=header_height)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        header_content = tk.Frame(header_frame, bg=colors['primary'])
        header_content.pack(expand=True)
        header_main = tk.Frame(header_content, bg=colors['primary'])
        header_main.pack(expand=True)
        icon_size = max(24, int(28 * scale_factor))
        icon_label = tk.Label(header_main, text="🔔", font=scaled_font('Segoe UI Emoji', icon_size), bg=colors['primary'], fg='white')
        icon_label.pack(pady=(5, 3))
        title_label = tk.Label(header_main, text="부태리의 신고가", font=scaled_font('Malgun Gothic', 18, 'bold'), bg=colors['primary'], fg='white')
        title_label.pack()
        page_info_label = tk.Label(header_main, text="", font=scaled_font('Malgun Gothic', 10), bg=colors['primary'], fg='#E3F2FD')
        page_info_label.pack(pady=(3, 10))
        content_height = int(window_height * 0.68)
        content_container = tk.Frame(main_frame, bg=colors['background'], height=content_height)
        content_container.pack(fill='both', expand=True, padx=15, pady=10)
        content_container.pack_propagate(False)
        content_frame = tk.Frame(content_container, bg=colors['background'])
        content_frame.pack(fill='both', expand=True)
        def update_page():
            for widget in content_frame.winfo_children():
                widget.destroy()
            start_idx = current_page * apts_per_page
            end_idx = min(start_idx + apts_per_page, len(apt_list))
            current_apts = apt_list[start_idx:end_idx]
            if total_pages > 1:
                page_text = f"{current_page + 1}/{total_pages} 페이지 | {len(apt_list)}개 아파트"
            else:
                page_text = f"{len(apt_list)}개 아파트에서 최고가 갱신"
            page_info_label.config(text=page_text)
            for i, apt in enumerate(current_apts):
                card_frame = tk.Frame(content_frame, bg=colors['card_bg'], relief='flat', bd=0, highlightbackground=colors['separator'], highlightthickness=1)
                card_frame.pack(fill='x', pady=(0, 6 if i < len(current_apts) - 1 else 0))
                card_content = tk.Frame(card_frame, bg=colors['card_bg'])
                card_content.pack(fill='both', expand=True, padx=12, pady=10)
                apt_header = tk.Frame(card_content, bg=colors['card_bg'])
                apt_header.pack(fill='x', pady=(0, 6))
                apt_name_size = max(10, int(12 * scale_factor))
                apt_name_label = tk.Label(apt_header, text=f"🏢 {apt['apt_name']}", font=scaled_font('Malgun Gothic', apt_name_size, 'bold'), bg=colors['card_bg'], fg=colors['text_primary'], anchor='w')
                apt_name_label.pack(anchor='w')
                
                # 지역 정보 추가
                location_text = ""
                if 'sido' in apt and 'sigungu' in apt and 'dong' in apt:
                    location_text = f"📍 {apt['sido']} {apt['sigungu']} {apt['dong']}"
                elif 'location' in apt:
                    location_text = f"📍 {apt['location']}"
                    
                if location_text:
                    location_size = max(7, int(9 * scale_factor))
                    location_label = tk.Label(apt_header, text=location_text, font=scaled_font('Malgun Gothic', location_size), bg=colors['card_bg'], fg=colors['text_secondary'], anchor='w')
                    location_label.pack(anchor='w', pady=(1, 0))
                
                area_size = max(8, int(10 * scale_factor))
                area_label = tk.Label(apt_header, text=f"📐 {apt['area']}㎡", font=scaled_font('Malgun Gothic', area_size), bg=colors['card_bg'], fg=colors['text_secondary'], anchor='w')
                area_label.pack(anchor='w', pady=(1, 0))
                price_section = tk.Frame(card_content, bg=colors['card_bg'])
                price_section.pack(fill='x', pady=(0, 6))

                # old_price 또는 prev_max_price 키 지원 (순서 중요!)
                old_price = apt.get('old_price', 0) or apt.get('prev_max_price', 0)
                if old_price > 0:
                    old_price_frame = tk.Frame(price_section, bg=colors['card_bg'])
                    old_price_frame.pack(fill='x', pady=(0, 3))
                    old_label_size = max(8, int(9 * scale_factor))
                    tk.Label(old_price_frame, text="이전 최고가",
                             font=scaled_font('Malgun Gothic', old_label_size),
                             bg=colors['card_bg'], fg=colors['text_secondary']).pack(side='left')

                    # ✨ 날짜/층 정보 추가
                    old_detail = f"{old_price:,}만원"
                    if apt.get('old_date'):
                        old_detail += f" ({apt['old_date']}"
                        if apt.get('old_floor'):
                            old_detail += f" | {apt['old_floor']}층"
                        if apt.get('old_dong') and apt['old_dong'] != '-':
                            old_detail += f" | {apt['old_dong']}"
                        old_detail += ")"
                    
                    tk.Label(old_price_frame, text=old_detail, 
                             font=scaled_font('Malgun Gothic', old_label_size), 
                             bg=colors['card_bg'], fg=colors['text_secondary']).pack(side='right')
                new_price_frame = tk.Frame(price_section, bg=colors['card_bg'])
                new_price_frame.pack(fill='x')
                new_label_size = max(9, int(11 * scale_factor))
                tk.Label(new_price_frame, text="📈 신고가", font=scaled_font('Malgun Gothic', new_label_size, 'bold'), bg=colors['card_bg'], fg=colors['error']).pack(side='left')
                new_price_size = max(11, int(13 * scale_factor))

                # new_price 또는 new_max_price 키 지원 (순서 중요!)
                new_price = apt.get('new_price', 0) or apt.get('new_max_price', 0)
                tk.Label(new_price_frame, text=f"{new_price:,}만원", font=scaled_font('Malgun Gothic', new_price_size, 'bold'), bg=colors['card_bg'], fg=colors['error']).pack(side='right')
                if old_price > 0:
                    increase = new_price - old_price
                    increase_percent = (increase / old_price) * 100
                    highlight_frame = tk.Frame(card_content, bg=colors['error_light'], relief='flat', bd=1)
                    highlight_frame.pack(fill='x', pady=(6, 0))
                    highlight_content = tk.Frame(highlight_frame, bg=colors['error_light'])
                    highlight_content.pack(fill='x', padx=6, pady=4)
                    increase_size = max(8, int(10 * scale_factor))
                    increase_label = tk.Label(highlight_content, text=f"🔥 {increase:,}만원 상승 (+{increase_percent:.1f}%)", font=scaled_font('Malgun Gothic', increase_size, 'bold'), bg=colors['error_light'], fg='#C62828')
                    increase_label.pack(anchor='w')
                    detail_info = f"📅 {apt['date']}"
                    if apt.get('floor'):
                        detail_info += f" | {apt['floor']}층"
                    if apt.get('dong') and apt['dong'] != '-':
                        detail_info += f" | {apt['dong']}"
                    detail_size = max(7, int(8 * scale_factor))
                    detail_label = tk.Label(highlight_content, text=detail_info, font=scaled_font('Malgun Gothic', detail_size), bg=colors['error_light'], fg=colors['text_secondary'])
                    detail_label.pack(anchor='w', pady=(1, 0))
            if total_pages > 1:
                prev_button.config(state='normal' if current_page > 0 else 'disabled')
                next_button.config(state='normal' if current_page < total_pages - 1 else 'disabled')
            if total_pages > 1:
                page_display.config(text=f"{current_page + 1} / {total_pages}")
            else:
                page_display.config(text="")
        
        bottom_height = int(window_height * 0.16)
        bottom_container = tk.Frame(main_frame, bg=colors['background'], height=bottom_height)
        bottom_container.pack(fill='x', padx=15, pady=(0, 15))
        bottom_container.pack_propagate(False)
        if total_pages > 1:
            nav_frame = tk.Frame(bottom_container, bg=colors['background'])
            nav_frame.pack(fill='x', pady=(0, 8))
            def go_prev():
                nonlocal current_page
                if current_page > 0:
                    current_page -= 1
                    update_page()
            prev_button = tk.Button(nav_frame, text="◀", font=scaled_font('Malgun Gothic', 12, 'bold'), bg=colors['primary'], fg='white', activebackground=colors['primary_dark'], relief='flat', bd=0, width=4, height=1, cursor='hand2', command=go_prev)
            prev_button.pack(side='left')
            page_display = tk.Label(nav_frame, text="", font=scaled_font('Malgun Gothic', 11, 'bold'), bg=colors['background'], fg=colors['text_primary'])
            page_display.pack(side='left', expand=True)
            def go_next():
                nonlocal current_page
                if current_page < total_pages - 1:
                    current_page += 1
                    update_page()
            next_button = tk.Button(nav_frame, text="▶", font=scaled_font('Malgun Gothic', 12, 'bold'), bg=colors['primary'], fg='white', activebackground=colors['primary_dark'], relief='flat', bd=0, width=4, height=1, cursor='hand2', command=go_next)
            next_button.pack(side='right')
        else:
            prev_button = tk.Button(bottom_container)
            next_button = tk.Button(bottom_container)
            page_display = tk.Label(bottom_container)
            prev_button.pack_forget()
            next_button.pack_forget()
            page_display.pack_forget()
        def capture_screen():
            try:
                notification_window.update_idletasks()
                time.sleep(0.1)
                x = notification_window.winfo_rootx()
                y = notification_window.winfo_rooty()
                width = notification_window.winfo_width()
                try:
                    from ctypes import windll
                    windll.shcore.SetProcessDpiAwareness(1)
                except:
                    pass
                header_y = header_frame.winfo_rooty()
                header_height = header_frame.winfo_height()
                content_y = content_container.winfo_rooty()
                content_height = content_container.winfo_height()
                capture_y = header_y
                capture_bottom = content_y + content_height
                print(f"창 위치: x={x}, y={y}")
                print(f"창 크기: width={width}, height={notification_window.winfo_height()}")
                print(f"헤더 위치: y={header_y}, height={header_height}")
                print(f"콘텐츠 위치: y={content_y}, height={content_height}")
                print(f"캡처 영역: y={capture_y}, height={capture_bottom - capture_y}")
                try:
                    bbox = (x, capture_y, x + width, capture_bottom)
                    screenshot = ImageGrab.grab(bbox=bbox)
                    if screenshot.size[0] <= 0 or screenshot.size[1] <= 0:
                        raise ValueError("캡처된 이미지 크기가 유효하지 않습니다.")
                    capture_folder = os.path.join(self.download_path, "captures")
                    os.makedirs(capture_folder, exist_ok=True)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    page_info = f"_page{current_page + 1}of{total_pages}" if total_pages > 1 else ""
                    filename = f"부태리신고가_{timestamp}{page_info}.png"
                    filepath = os.path.join(capture_folder, filename)
                    screenshot.save(filepath, 'PNG', quality=95)
                    capture_button.config(text="✓", fg=colors['success'])
                    notification_window.after(3000, lambda: capture_button.config(text="📸", fg='white'))
                    print(f"📸 화면 캡처 성공: {filepath}")
                    print(f"   캡처 크기: {screenshot.size[0]}x{screenshot.size[1]}")
                except ImportError:
                    capture_button.config(text="❌", fg=colors['error'])
                    notification_window.after(3000, lambda: capture_button.config(text="📸", fg='white'))
                    messagebox.showerror("오류", "화면 캡처를 위해 Pillow 라이브러리가 필요합니다.\npip install Pillow 명령으로 설치해주세요.")
            except Exception as e:
                capture_button.config(text="❌", fg=colors['error'])
                notification_window.after(3000, lambda: capture_button.config(text="📸", fg='white'))
                print(f"캡처 오류: {str(e)}")
                import traceback
                traceback.print_exc()
                messagebox.showerror("오류", f"화면 캡처 중 오류가 발생했습니다:\n{str(e)}")
        # --- 버튼 행 컨테이너 ---
        button_frame = tk.Frame(bottom_container, bg=colors['background'])
        button_frame.pack(fill='x')
        button_size = max(12, int(14 * scale_factor))
        
        # 1) 확인 버튼
        confirm_button = tk.Button(
            button_frame,
            text="확인",
            font=scaled_font('Malgun Gothic', button_size, 'bold'),
            bg=colors['primary'], fg='white', activebackground=colors['primary_dark'],
            activeforeground='white', relief='flat', bd=0, pady=12, cursor='hand2',
            command=notification_window.destroy
        )
        confirm_button.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        # 2) (원하는 위치에 따라 pack 순서 조절) 캡처 버튼
        capture_button = tk.Button(
            button_frame,
            text="📸",
            font=scaled_font('Segoe UI Emoji', 14),
            bg=colors['capture'], fg='white', activebackground='#FF7A00',
            relief='flat', bd=0, width=4, pady=12, cursor='hand2',
            command=capture_screen
        )
        capture_button.pack(side='right')
        
        # 3) HTML 저장 버튼 (지금 순서면 📸 왼쪽에 위치)
        html_button = tk.Button(
            button_frame,
            text="📝 HTML 저장",
            font=scaled_font('Malgun Gothic', button_size, 'bold'),
            bg='#10B981', fg='white', activebackground='#0E9F6E', relief='flat', bd=0,
            pady=12, cursor='hand2',
            command=lambda: self.export_notification_html(apt_list, ask_path=False, silent=False)
        )
        html_button.pack(side='right', padx=(6, 0))
        
        # --- 호버 핸들러 (버튼 만든 뒤 정의해도 되고, 앞에 있어도 됩니다) ---
        def on_enter(e):
            if e.widget == confirm_button:
                confirm_button.config(bg=colors['primary_dark'])
            elif e.widget == capture_button:
                capture_button.config(bg='#FF7A00')
            elif e.widget == html_button:
                html_button.config(bg='#0E9F6E')
        
        def on_leave(e):
            if e.widget == confirm_button:
                confirm_button.config(bg=colors['primary'])
            elif e.widget == capture_button:
                capture_button.config(bg=colors['capture'])
            elif e.widget == html_button:
                html_button.config(bg='#10B981')
        
        # --- 바인딩 (버튼 '생성 후'에 있어야 함) ---
        confirm_button.bind("<Enter>", on_enter)
        confirm_button.bind("<Leave>", on_leave)
        capture_button.bind("<Enter>", on_enter)
        capture_button.bind("<Leave>", on_leave)
        html_button.bind("<Enter>", on_enter)
        html_button.bind("<Leave>", on_leave)
        if total_pages > 1:
            prev_button.bind("<Enter>", on_enter)
            prev_button.bind("<Leave>", on_leave)
            next_button.bind("<Enter>", on_enter)
            next_button.bind("<Leave>", on_leave)
        def on_key_press(event):
            if event.keysym == 'Escape' or event.keysym == 'Return':
                notification_window.destroy()
            elif event.keysym == 'Left' and total_pages > 1 and current_page > 0:
                go_prev()
            elif event.keysym == 'Right' and total_pages > 1 and current_page < total_pages - 1:
                go_next()
            elif event.keysym == 'F12' or (event.state & 0x4 and event.keysym == 's'):
                capture_screen()
        notification_window.bind('<Key>', on_key_press)
        notification_window.focus_set()
        update_page()
        notification_window.after(100, lambda: notification_window.focus_force())
        try:
            title = "📱 부태리의 신고가"
            message = f"{len(apt_list)}개 아파트에서 신고가가 발견되었습니다!"
            notification.notify(title=title, message=message, app_name="실거래가 모니터", timeout=10)
        except Exception as e:
            logging.error(f"시스템 알림 표시 중 오류: {str(e)}")
        print(f"📱 부태리 신고가 알림창 크기: {window_width}x{window_height} (9:16 비율)")
        print(f"📄 총 {total_pages}페이지, 페이지당 {apts_per_page}개 아파트")
        print(f"📸 캡처 기능: F12 키 또는 우측 하단 📸 버튼")
    
    def on_closing(self):
        """프로그램 종료 시 처리"""
        try:
            # 먼저 종료 여부를 묻고
            if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
                # 사용자가 확인하면 저장 작업 수행
                try:
                    self.save_monitored_apts()
                    self.save_notifications_history()
                except Exception as e:
                    logging.error(f"데이터 저장 중 오류: {str(e)}")

                # DB 연결 종료
                try:
                    if hasattr(self, 'db_conn') and self.db_conn:
                        self.db_conn.close()
                        logging.info("DB 연결이 종료되었습니다.")
                except Exception as e:
                    logging.error(f"DB 연결 종료 중 오류: {str(e)}")

                # 창 종료
                self.root.destroy()
        except Exception as e:
            logging.error(f"종료 처리 중 오류: {str(e)}")
            try:
                if hasattr(self, 'db_conn') and self.db_conn:
                    self.db_conn.close()
            except:
                pass
            self.root.destroy()


class AptSelectDialog:
    def __init__(self, parent, apt_list, service_key, sigungu_code, dong, sido, sigungu, title="아파트 선택"):
        self.parent = parent
        self.service_key = service_key
        self.sigungu_code = sigungu_code
        self.dong = dong
        self.sido = sido
        self.sigungu = sigungu
        self.apt_list = apt_list
        self.result = None
        self.selected_apt = None
        
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.attributes('-topmost', True)
        
        width = 800
        height = 500
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.top.geometry(f"{width}x{height}+{x}+{y}")
        
        self.font_normal = ('Malgun Gothic', 9)
        self.font_large = ('Malgun Gothic', 11)
        self.font_button = ('Malgun Gothic', 9)
        
        search_frame = ttk.Frame(self.top, padding="5")
        search_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(search_frame, text=f"{self.dong} 아파트 목록", font=self.font_large).pack(side='left')
        ttk.Label(search_frame, text="검색:", font=self.font_normal).pack(side='left', padx=(20, 0))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_apartments)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40, font=self.font_normal)
        search_entry.pack(side='left', fill='x', expand=True, padx=5)
        
        list_frame = ttk.Frame(self.top, padding="5")
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=self.font_normal)
        self.listbox.pack(fill='both', expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        self.update_listbox(apt_list)
        
        button_frame = ttk.Frame(self.top, padding="5")
        button_frame.pack(fill='x', padx=5, pady=5)
        select_button = ttk.Button(button_frame, text="선택", command=self.on_button_select)
        select_button.pack(side='right', padx=5)
        cancel_button = ttk.Button(button_frame, text="취소", command=self.top.destroy)
        cancel_button.pack(side='right', padx=5)
        self.listbox.bind('<Double-Button-1>', self.on_select)
    
    def update_listbox(self, items):
        self.listbox.delete(0, tk.END)
        for item in items:
            self.listbox.insert(tk.END, item)
    
    def filter_apartments(self, *args):
        search_text = self.search_var.get().lower()
        filtered_list = [apt for apt in self.apt_list if search_text in apt.lower()]
        self.update_listbox(filtered_list)
    
    def on_button_select(self):
        if not self.listbox.curselection():
            messagebox.showinfo("알림", "아파트를 선택해주세요.")
            return
        self.on_select(None)
    
    def on_select(self, event):
        if self.listbox.curselection():
            full_text = self.listbox.get(self.listbox.curselection())
            is_new_apt = False
            if full_text.startswith("[신축]"):
                is_new_apt = True
                full_text = full_text[5:].strip()
            if '[' in full_text and ']' in full_text:
                address_info = full_text[full_text.find('[')+1:full_text.find(']')]
                addr_parts = address_info.split(' / ')
                if len(addr_parts) > 1:
                    jibun_addr = addr_parts[1]
                else:
                    jibun_addr = addr_parts[0]
                simple_addr = ' '.join(jibun_addr.split()[-2:])
            else:
                simple_addr = ""
            apt_name = full_text.split('[')[0].strip()
            build_year = ""
            if "(준공:" in full_text:
                build_year_part = full_text.split("(준공:")[1].strip()
                build_year = build_year_part.split("년")[0].strip()
            elif "(분양중)" in full_text:
                build_year = "분양"
            self.selected_apt = apt_name
            self.simple_addr = simple_addr
            self.build_year = build_year
            self.is_new_apt = is_new_apt
            self.show_area_dialog()
    
    def show_area_dialog(self):
        """전용면적 목록 다이얼로그 표시"""
        # 전용면적 목록 가져오기
        area_list = self.get_areas_for_apt(self.selected_apt)
        
        if not area_list:
            messagebox.showinfo("알림", "해당 아파트의 전용면적 정보를 찾을 수 없습니다.")
            return
        
        area_dialog = tk.Toplevel(self.top)
        area_dialog.title(f"{self.selected_apt} - 전용면적 선택")
        area_dialog.attributes('-topmost', True)
        
        width = 300
        height = 200
        x = self.top.winfo_x() + 50
        y = self.top.winfo_y() + 50
        area_dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        list_frame = ttk.Frame(area_dialog, padding="5")
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=self.font_normal)
        listbox.pack(fill='both', expand=True)
        scrollbar.config(command=listbox.yview)
        
        for area in sorted(area_list, key=lambda x: float(x)):
            listbox.insert(tk.END, f"{area}㎡")
        
        # 첫 번째 항목 선택 (사용자 편의)
        if len(area_list) > 0:
            listbox.selection_set(0)
        
        def on_area_select(event=None):
            # 선택한 항목이 없으면 첫 번째 항목 자동 선택
            if not listbox.curselection() and len(area_list) > 0:
                listbox.selection_set(0)
                
            if not listbox.curselection():
                messagebox.showinfo("알림", "전용면적을 선택해주세요.")
                return
                
            selected_area = listbox.get(listbox.curselection())
            area_value = selected_area.replace('㎡', '').strip()
            
            # 아파트 정보 구성
            apt_info = {
                'apt_name': self.selected_apt,
                'jibun_addr': self.simple_addr,
                'area': area_value,
                'sido': self.sido,
                'sigungu': self.sigungu,
                'dong': self.dong,
                'sigungu_code': self.sigungu_code,
                'build_year': self.build_year
            }
            
            # 결과 저장
            self.result = apt_info
            
            # 창 닫기
            area_dialog.destroy()
            self.top.destroy()
        
        # 이벤트 바인딩
        listbox.bind('<Double-1>', on_area_select)
        
        # 선택 버튼 프레임
        button_frame = ttk.Frame(area_dialog, padding="5")
        button_frame.pack(fill='x', padx=5, pady=5)
        
        # 선택 버튼
        select_button = ttk.Button(button_frame, text="선택", command=on_area_select)
        select_button.pack(side='right', padx=5)
        
        cancel_button = ttk.Button(button_frame, text="취소", command=area_dialog.destroy)
        cancel_button.pack(side='right', padx=5)
        
        # Enter 키 이벤트 바인딩 추가
        listbox.bind('<Return>', on_area_select)
        area_dialog.bind('<Return>', on_area_select)
        
        # 다이얼로그가 닫힐 때 처리
        def on_dialog_close():
            area_dialog.destroy()
        
        area_dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        
        # 포커스 설정
        listbox.focus_set()
        
        # 모달 다이얼로그로 처리
        area_dialog.transient(self.top)
        area_dialog.grab_set()
        self.top.wait_window(area_dialog)
        
    def get_areas_for_apt(self, apt_name):
        """해당 아파트의 전용면적 목록 가져오기 (최근 3개월)"""
        areas = set()
        current_date = datetime.now()
        
        # [신축] 태그 제거
        if apt_name.startswith("[신축] "):
            apt_name = apt_name[6:]  # "[신축] " 제거
        
        # 진행 상황 표시 창
        progress_window = tk.Toplevel(self.top)
        progress_window.title("데이터 수집 중...")
        progress_window.geometry("300x100")
        progress_window.transient(self.top)
        progress_window.grab_set()
        
        ttk.Label(progress_window, text=f"{apt_name} 전용면적 정보를 수집 중입니다...").pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        progress_bar.pack(fill="x", padx=20, pady=10)
        
        cancel_flag = [False]
        progress_window.protocol("WM_DELETE_WINDOW", lambda: setattr(cancel_flag, 0, True) or progress_window.destroy())
        
        if not hasattr(self, 'apt_area_cache'):
            self.apt_area_cache = {}
        
        if apt_name in self.apt_area_cache:
            progress_bar['value'] = 100
            progress_window.update_idletasks()
            time.sleep(0.3)
            progress_window.destroy()
            return self.apt_area_cache[apt_name]
        
        concurrent_requests = 12
        
        session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=concurrent_requests,
            pool_maxsize=concurrent_requests * 2,
            max_retries=1
        )
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        
        # 최근 3개월만 조회
        max_months = 4
        
        def collect_areas():
            nonlocal areas
            
            try:
                all_requests = []
                for month in range(max_months):
                    search_date = current_date - timedelta(days=30 * month)
                    deal_ymd = search_date.strftime("%Y%m")
                    
                    # 기축 매매 API
                    trade_url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                               f"?serviceKey={self.service_key}"
                               f"&LAWD_CD={self.sigungu_code}"
                               f"&DEAL_YMD={deal_ymd}"
                               f"&numOfRows=1000")
                    
                    # 신축(분양권) 매매 API
                    new_trade_url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcSilvTrade/getRTMSDataSvcSilvTrade"
                                   f"?serviceKey={self.service_key}"
                                   f"&LAWD_CD={self.sigungu_code}"
                                   f"&DEAL_YMD={deal_ymd}"
                                   f"&numOfRows=1000")
                    
                    # 기축 전세 API
                    rent_url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"
                               f"?serviceKey={self.service_key}"
                               f"&LAWD_CD={self.sigungu_code}"
                               f"&DEAL_YMD={deal_ymd}"
                               f"&numOfRows=1000")
                    
                    all_requests.append((trade_url, 'trade', month))
                    all_requests.append((new_trade_url, 'new_trade', month))
                    all_requests.append((rent_url, 'rent', month))
                
                with concurrent.futures.ThreadPoolExecutor(max_workers=concurrent_requests) as executor:
                    future_to_request = {
                        executor.submit(session.get, url, timeout=2): (req_type, month_idx) 
                        for url, req_type, month_idx in all_requests
                    }
                    
                    total_requests = len(all_requests)
                    processed = 0
                    
                    for future in concurrent.futures.as_completed(future_to_request):
                        processed += 1
                        req_type, month_idx = future_to_request[future]
                        
                        progress = min(100, (processed / total_requests) * 100)
                        progress_bar['value'] = progress
                        progress_window.update_idletasks()
                        
                        if len(areas) >= 5:
                            continue
                            
                        if cancel_flag[0]:
                            continue
                        
                        try:
                            response = future.result()
                            
                            if response.status_code == 200:
                                try:
                                    root = ET.fromstring(response.text)
                                    
                                    for item in root.findall('.//item'):
                                        item_apt = item.findtext('aptNm', '').strip()
                                        
                                        if item_apt == apt_name:
                                            item_area = float(item.findtext('excluUseAr', '0'))
                                            if item_area > 0:
                                                # 소수점 없는 정수로 변환
                                                area_int = str(int(item_area))
                                                areas.add(area_int)
                                except ET.ParseError:
                                    pass
                        
                        except Exception as e:
                            continue
                
                progress_bar['value'] = 100
                progress_window.update_idletasks()
                
                self.apt_area_cache[apt_name] = sorted(list(areas), key=float)
                
            except Exception as e:
                logging.error(f"전용면적 정보 수집 중 오류: {str(e)}")
            
            time.sleep(0.3)
            if not cancel_flag[0]:
                progress_window.destroy()
        

        thread = threading.Thread(target=collect_areas)
        thread.daemon = True
        thread.start()
        
        self.top.wait_window(progress_window)
        
        if cancel_flag[0] or not areas:
            return []
        
        if apt_name not in self.apt_area_cache:
            self.apt_area_cache[apt_name] = sorted(list(areas), key=float)
        
        return sorted(list(areas), key=float)


def main():
    app = RealEstateMonitorApp()
    app.root.mainloop()

if __name__ == "__main__":
    main()



