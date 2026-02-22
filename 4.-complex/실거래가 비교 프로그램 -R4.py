
"""
실거래가 비교 프로그램 -R4.py

수정 내역 (2026-01-07) - R4:
1. 중복 지역명 처리 로직 개선 🔧
   - 시도 약어 생성 방식 개선: 첫 글자 → 전체 시도명 (접미사 제거)
   - "대구 중구" vs "대전 중구" 구분 문제 해결
   - 표시 형식: 중구(대구), 중구(대전) 으로 명확히 구분
   - 광역시/특별시/도 접미사 제거하여 간결하고 명확한 표시

2. load_lawdong_file 함수 간소화
   - 3단계 구조(시-구-동) 처리 로직 제거
   - 2단계 구조(시도-시군구-동)만 지원하여 안정성 향상
   - 코드 가독성 및 유지보수성 개선

수정 내역 (2025-12-17) - R2:
1. get_apt_list_from_api 함수를 신고가 프로그램(real_estate_monitor_db_fixed_v3.py) 방식으로 개선
   - 기축 아파트와 신축 아파트(분양권) API 모두 조회
   - 더 상세한 디버그 로그 출력 추가
   - API 호출 간격 조절 (time.sleep 추가)

2. '시-구-동' 3단계 행정구역 구조 완벽 지원
   - load_lawdong_file 함수 개선: 시 아래 구, 구 아래 동 계층 구조 파싱
   - gu_info 딕셔너리 추가: 구 정보 저장
   - on_sigungu_selected 함수 개선: 구 아래 하위 동을 "  └ " 형식으로 표시

3. show_apt_list 함수 대폭 개선
   - '구' 선택 시 하위 모든 동의 아파트 목록을 한번에 조회
   - 하위 동 (  └ 로 시작) 선택 시 정확한 코드 매칭
   - [동명] 형식으로 동 정보 포함하여 아파트 표시

4. 그래프 비주얼 개선 (seaborn 스타일)
   - seaborn 라이브러리 통합: 더 세련된 색상 팔레트 (husl)
   - 그리드 스타일 개선: 연한 회색, 데이터 뒤로 배치
   - 어노테이션 스타일 개선: 더 두꺼운 테두리, 곡선 화살표
   - adjustText 라이브러리 지원: 설치 시 레이블 자동 조정으로 겹침 방지
   - 그래프 크기 확대: 18x9 (기존 16x8) - 더 넓은 화면

5. GUI 디자인 완전 개선 🎨
   - 현대적인 색상 테마: Flat Design 스타일 (블루/그린 계열)
   - ttk 스타일 커스터마이징: clam 테마 기반 커스텀 스타일
   - 버튼 스타일 3종: 기본/액센트(Accent)/성공(Success)
   - 이모지 추가: 직관적인 UI (📊, 🔍, ✅, 📈, 💰, 🏠 등)
   - 프레임 카드 스타일: 흰색 배경에 깔끔한 테두리
   - 폰트 크기 증가: 10pt → 12pt (가독성 향상)
   - 여백 및 패딩 개선: 더 넓은 간격으로 여유 있는 레이아웃
   - 창 크기 확대: 1200x900 → 1300x950

6. 데이터 수집 진행 상황 표시 개선 ⏳
   - 단지 선택 시 자동 데이터 수집의 진행 상황을 프로그레스바로 표시
   - 단계별 상태 메시지: 준비 → 매매 수집 → 전세 수집 → 완료
   - 수집 건수 실시간 표시 (매매 XX건, 전세 XX건)
   - 완료 후 명확한 안내 메시지: "🎉 데이터 수집 완료! 이제 그래프 생성하기 버튼을 눌러주세요"
   - 데이터 없을 시 경고 메시지 표시
   - 이모지로 직관적인 상태 표시 (🔄 수집중, 📊 매매, 🏠 전세, ✅ 완료)

7. 해결된 문제
   - '시' 다음에 '구'로 나눠지는 지역 (예: 경기도 성남시 분당구)의 단지 목록 조회 문제 완전 해결
   - 구 단위로 선택 시 모든 하위 동의 단지를 한번에 불러올 수 있음
   - 레이블이 겹치는 문제 개선 (adjustText 설치 시 자동 조정)
   - 눈에 잘 들어오는 색상과 배치로 사용성 대폭 향상
   - 데이터 수집 진행 상황을 알 수 없던 문제 해결
"""

import re  # 기존 import 구문들과 함께 추가
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib
matplotlib.use('Agg')  # GUI 없이 이미지만 생성하는 백엔드 설정 (tkinter 충돌 방지)
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from PIL import Image, ImageTk
import pandas as pd
import numpy as np
import os
import json
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import logging
import shutil
from matplotlib import font_manager
import time
import concurrent.futures
import threading
import gc  # 가비지 컬렉션 추가
import seaborn as sns  # seaborn 추가
try:
    from adjustText import adjust_text  # 레이블 자동 조정
    ADJUSTTEXT_AVAILABLE = True
except ImportError:
    ADJUSTTEXT_AVAILABLE = False
    print("adjustText 라이브러리가 설치되지 않았습니다. pip install adjustText 로 설치하세요.")


# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# API 타임아웃 설정 (connect, read)
# connect: 서버 연결까지 대기 시간
# read: 데이터 읽기 대기 시간
API_TIMEOUT = (5, 15)  # (connect timeout, read timeout)

# 항상 최상위에 표시되는 커스텀 messagebox 함수들
def show_topmost_info(title, message, parent=None):
    """항상 최상위에 표시되는 정보 메시지박스"""
    if parent:
        parent.attributes('-topmost', True)
    result = messagebox.showinfo(title, message, parent=parent)
    if parent:
        parent.attributes('-topmost', False)
    return result

def show_topmost_warning(title, message, parent=None):
    """항상 최상위에 표시되는 경고 메시지박스"""
    if parent:
        parent.attributes('-topmost', True)
    result = messagebox.showwarning(title, message, parent=parent)
    if parent:
        parent.attributes('-topmost', False)
    return result

def show_topmost_error(title, message, parent=None):
    """항상 최상위에 표시되는 오류 메시지박스"""
    if parent:
        parent.attributes('-topmost', True)
    result = messagebox.showerror(title, message, parent=parent)
    if parent:
        parent.attributes('-topmost', False)
    return result

def ask_topmost_yesno(title, message, parent=None):
    """항상 최상위에 표시되는 예/아니오 메시지박스"""
    if parent:
        parent.attributes('-topmost', True)
    result = messagebox.askyesno(title, message, parent=parent)
    if parent:
        parent.attributes('-topmost', False)
    return result

class RealEstateAnalyzerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("📊 부태리의 실거래가 차트")  # 이모지 추가
        self.root.geometry("1400x1200")  # 크기 증가 (세로 길이 더 늘림)

        # 현대적인 색상 테마 정의
        self.colors = {
            'primary': '#2C3E50',      # 다크 블루 (제목)
            'secondary': '#3498DB',    # 밝은 블루 (액센트)
            'success': '#27AE60',      # 그린 (성공)
            'warning': '#F39C12',      # 오렌지 (경고)
            'danger': '#E74C3C',       # 레드 (위험)
            'light': '#ECF0F1',        # 밝은 회색 (배경)
            'dark': '#34495E',         # 다크 그레이
            'white': '#FFFFFF',        # 흰색
            'bg_main': '#F8F9FA',      # 메인 배경
            'bg_frame': '#FFFFFF',     # 프레임 배경
            'text_primary': '#2C3E50', # 주요 텍스트
            'text_secondary': '#7F8C8D', # 보조 텍스트
            'border': '#BDC3C7'        # 테두리
        }

        # 배경색 설정
        self.root.configure(bg=self.colors['bg_main'])

        # 종료 플래그 (백그라운드 스레드에서 확인)
        self.is_closing = False

        # 안전한 UI 업데이트 헬퍼 함수
        def safe_after(delay, func):
            """종료 중이 아닐 때만 UI 업데이트 실행"""
            if not self.is_closing:
                try:
                    self.root.after(delay, func)
                except:
                    pass

        self.safe_after = safe_after

        # 기본 설정
        self.download_path = "C:\\Download"
        self.history_path = os.path.join(self.download_path, "history")
        self.lawdong_path = "C:/law-dong/law-dong.txt"
        self.image_path = os.path.join(self.download_path, "graph.jpg")
        self.excel_path = None
        
        # 그래프 옵션 변수 추가 - 여기에 정의
        self.show_monthly_avg = tk.BooleanVar(value=True)
        self.show_monthly_max = tk.BooleanVar(value=True)
        self.show_scatter_plot = tk.BooleanVar(value=True)
        
        # 전세 그래프 옵션 추가
        self.show_jeonse = tk.BooleanVar(value=True)  # 전세 데이터 표시 여부

        # 전세가 옵션 변수들 아래에 추가
        self.show_jeonse_scatter_plot = tk.BooleanVar(value=True)

        # 세부정보 가져오기 옵션 변수 추가 (기본값: False)
        self.show_complex_info = tk.BooleanVar(value=False)


        # 새로운 전세가 옵션 변수 추가
        self.show_jeonse_monthly_avg = tk.BooleanVar(value=True)  # 전세 월평균가격 표시 여부
        self.show_jeonse_monthly_max = tk.BooleanVar(value=True)  # 전세 월최고가 표시 여부
        self.show_jeonse_scatter_plot = tk.BooleanVar(value=True)  # 전세 점도표 표시 여부

        # 전세 데이터 수집 여부 옵션 추가 (기본값: True)
        self.collect_jeonse_data = tk.BooleanVar(value=True)  # 전세 데이터 수집 여부
        
        # 설정 파일 로드
        self.load_settings()
        
        # 폴더 생성
        for path in [self.download_path, self.history_path]:
            if not os.path.exists(path):
                os.makedirs(path)
        
        # 폰트 설정
        self.setup_fonts()

        # ttk 스타일 설정
        self.setup_styles()

        # 법정동 코드 관련 변수 초기화
        self.region_codes = {}
        self.sido_list = []
        self.sigungu_dict = {}
        self.dong_dict = {}
        
        # 법정동 파일 로드
        self.load_lawdong_file()
        
        # API 키 설정
        self.service_key = "Vs5lXsSo6iEI8no3pP%2FT0udWF9s7Cc8oP1SIWnEI5F4h6dKq92fLvnKmxkoWGJxSeW2%2FSOLQECGxOJzWcjJEXQ%3D%3D"
        
        # GUI 설정
        self.setup_gui()
        
        # 히스토리 로드
        self.history_list = self.load_history()
        self.update_history_display()

        # 윈도우 닫기 프로토콜 설정 (프로그램 종료 시 자동 저장)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 선택된 아파트 목록 저장 변수 추가
        self.selected_apts = []

        # 아파트 목록 캐시 (지역별로 저장)
        # 키: (sigungu_code, dong) 튜플
        # 값: 아파트 목록 리스트
        self.apt_list_cache = {}

        # 거래 데이터 파일 캐시 경로 초기화 (설정에서 로드됨)
        self.trade_cache_path = None


    def load_settings(self):
        """설정 파일 로드 (단지정보 경로 포함)"""
        settings_file = os.path.join(os.getcwd(), 'real_estate_analyzer_settings.json')
        
        # 기본 설정값 정의
        default_settings = {
            'download_path': os.path.join(os.path.expanduser('~'), 'Documents', 'RealEstateAnalyzer'),
            'history_path': os.path.join(os.path.expanduser('~'), 'Documents', 'RealEstateAnalyzer', 'history'),
            'lawdong_path': os.path.join(os.getcwd(), 'data', 'law-dong.txt'),
            'complex_info_path': os.path.join(os.getcwd(), 'data', 'complex_info.xlsx'),  # 기본 단지정보 경로 추가
            'trade_cache_path': os.path.join(os.path.expanduser('~'), 'Documents', 'RealEstateAnalyzer', 'trade_cache'),  # 거래 데이터 캐시 경로 추가
            'graph_options': {
                'show_monthly_avg': True,
                'show_monthly_max': True,
                'show_scatter_plot': True,
                'show_jeonse': True,
                'show_jeonse_monthly_avg': True,
                'show_jeonse_monthly_max': True,
                'show_jeonse_scatter_plot': True,
                'collect_jeonse_data': True  # 전세 데이터 수집 여부
            }
        }
        
        # 설정 파일이 있으면 로드, 없으면 기본값 사용
        if os.path.exists(settings_file):
            try:
                with open(settings_file, 'r', encoding='utf-8') as f:
                    settings_data = json.load(f)
                    
                    # 기본 설정에 저장된 설정 병합 (누락된 설정은 기본값 유지)
                    self.download_path = settings_data.get('download_path', default_settings['download_path'])
                    self.history_path = settings_data.get('history_path', default_settings['history_path'])
                    self.lawdong_path = settings_data.get('lawdong_path', default_settings['lawdong_path'])
                    self.complex_info_path = settings_data.get('complex_info_path', default_settings['complex_info_path'])  # 단지정보 경로 로드
                    self.trade_cache_path = settings_data.get('trade_cache_path', default_settings['trade_cache_path'])  # 거래 데이터 캐시 경로 로드
                    
                    # 그래프 옵션 로드
                    # 그래프 옵션 로드 부분 수정
                    if 'graph_options' in settings_data:
                        graph_options = settings_data['graph_options']
                        self.show_monthly_avg.set(graph_options.get('show_monthly_avg', True))
                        self.show_monthly_max.set(graph_options.get('show_monthly_max', True))
                        self.show_scatter_plot.set(graph_options.get('show_scatter_plot', True))
                        self.show_jeonse.set(graph_options.get('show_jeonse', True))
                        self.show_jeonse_monthly_avg.set(graph_options.get('show_jeonse_monthly_avg', True))
                        self.show_jeonse_monthly_max.set(graph_options.get('show_jeonse_monthly_max', True))
                        self.show_jeonse_scatter_plot.set(graph_options.get('show_jeonse_scatter_plot', True))

                        # 세부정보 옵션 추가
                        self.show_complex_info.set(graph_options.get('show_complex_info', False))

                        # 전세 데이터 수집 여부 옵션 추가
                        self.collect_jeonse_data.set(graph_options.get('collect_jeonse_data', True))
            except Exception as e:
                print(f"설정 파일 로드 중 오류: {str(e)}")
                self._apply_default_settings(default_settings)
        else:
            self._apply_default_settings(default_settings)
            
            # 기본 폴더 생성
            for path in [self.download_path, self.history_path, self.trade_cache_path]:
                if not os.path.exists(path):
                    os.makedirs(path)

            # 단지정보 파일 경로 디렉토리 생성
            complex_info_dir = os.path.dirname(self.complex_info_path)
            if not os.path.exists(complex_info_dir):
                os.makedirs(complex_info_dir)

    def load_complex_info(self):
        """단지정보 파일 로드"""
        if not hasattr(self, 'complex_info_path') or not os.path.exists(self.complex_info_path):
            return None
        
        try:
            import pandas as pd
            # 엑셀 파일 로드 - 첫 번째 행을 열 이름으로 사용
            df = pd.read_excel(self.complex_info_path, engine='openpyxl')
            
            # 열 이름을 알파벳으로 설정
            # 엑셀의 열 이름(A, B, C, ...)을 DataFrame의 열 이름으로 설정
            alphabet_columns = []
            for i in range(df.shape[1]):
                if i < 26:
                    col_name = chr(ord('a') + i)
                else:
                    # 26열 이상인 경우 aa, ab, ac... 로 설정
                    first_letter = chr(ord('a') + (i // 26) - 1)
                    second_letter = chr(ord('a') + (i % 26))
                    col_name = first_letter + second_letter
                alphabet_columns.append(col_name)
            
            df.columns = alphabet_columns
            
            return df
        except Exception as e:
            print(f"단지정보 파일 로드 중 오류: {str(e)}")
            return None
    
    def _apply_default_settings(self, default_settings):
        """기본 설정값 적용"""
        self.download_path = default_settings['download_path']
        self.history_path = default_settings['history_path']
        self.lawdong_path = default_settings['lawdong_path']
        self.complex_info_path = default_settings['complex_info_path']  # 단지정보 경로 추가
        self.trade_cache_path = default_settings['trade_cache_path']  # 거래 데이터 캐시 경로 추가
        
        # 그래프 옵션 설정
        graph_options = default_settings['graph_options']
        self.show_monthly_avg.set(graph_options['show_monthly_avg'])
        self.show_monthly_max.set(graph_options['show_monthly_max'])
        self.show_scatter_plot.set(graph_options['show_scatter_plot'])
        self.show_jeonse.set(graph_options['show_jeonse'])
        self.show_jeonse_monthly_avg.set(graph_options['show_jeonse_monthly_avg'])
        self.show_jeonse_monthly_max.set(graph_options['show_jeonse_monthly_max'])
        self.show_jeonse_scatter_plot.set(graph_options['show_jeonse_scatter_plot'])
        self.collect_jeonse_data.set(graph_options.get('collect_jeonse_data', True))


    def setup_fonts(self):
        """폰트 설정"""
        # 기본 폰트 설정 (크기 증가)
        self.font_normal = ('Malgun Gothic', 10)
        self.font_large = ('Malgun Gothic', 12)
        self.font_title = ('Malgun Gothic', 18, 'bold')
        self.font_button = ('Malgun Gothic', 10)

        # matplotlib 폰트 설정
        plt.rcParams['font.family'] = 'Malgun Gothic'
        plt.rcParams['axes.unicode_minus'] = False  # 마이너스 기호 깨짐 방지

    def setup_styles(self):
        """ttk 스타일 커스터마이징"""
        style = ttk.Style()
        style.theme_use('clam')  # 현대적인 테마

        # 프레임 스타일
        style.configure('TFrame', background=self.colors['bg_main'])
        style.configure('Card.TFrame', background=self.colors['bg_frame'],
                       relief='flat', borderwidth=0)

        # 라벨프레임 스타일
        style.configure('TLabelframe', background=self.colors['bg_frame'],
                       foreground=self.colors['primary'], borderwidth=2,
                       relief='solid')
        style.configure('TLabelframe.Label', background=self.colors['bg_frame'],
                       foreground=self.colors['primary'], font=self.font_large)

        # 라벨 스타일
        style.configure('TLabel', background=self.colors['bg_frame'],
                       foreground=self.colors['text_primary'], font=self.font_normal)
        style.configure('Title.TLabel', background=self.colors['bg_main'],
                       foreground=self.colors['primary'], font=self.font_title)
        style.configure('Subtitle.TLabel', background=self.colors['bg_frame'],
                       foreground=self.colors['text_secondary'], font=self.font_normal)

        # 버튼 스타일
        style.configure('TButton', font=self.font_button, borderwidth=0,
                       focuscolor='none', padding=10)
        style.map('TButton',
                 background=[('active', self.colors['secondary']),
                            ('!active', self.colors['primary'])],
                 foreground=[('active', self.colors['white']),
                            ('!active', self.colors['white'])])

        # 액센트 버튼 (주요 동작)
        style.configure('Accent.TButton', font=self.font_button,
                       borderwidth=0, focuscolor='none', padding=10)
        style.map('Accent.TButton',
                 background=[('active', '#2980B9'),
                            ('!active', self.colors['secondary'])],
                 foreground=[('active', self.colors['white']),
                            ('!active', self.colors['white'])])

        # 성공 버튼
        style.configure('Success.TButton', font=self.font_button,
                       borderwidth=0, focuscolor='none', padding=10)
        style.map('Success.TButton',
                 background=[('active', '#229954'),
                            ('!active', self.colors['success'])],
                 foreground=[('active', self.colors['white']),
                            ('!active', self.colors['white'])])

        # 콤보박스 스타일
        style.configure('TCombobox', fieldbackground=self.colors['white'],
                       background=self.colors['white'],
                       foreground=self.colors['text_primary'],
                       arrowcolor=self.colors['secondary'])

        # 체크버튼 스타일
        style.configure('TCheckbutton', background=self.colors['bg_frame'],
                       foreground=self.colors['text_primary'], font=self.font_normal)

        # 프로그레스바 스타일
        style.configure('TProgressbar', thickness=20,
                       troughcolor=self.colors['light'],
                       background=self.colors['secondary'])
        
    def load_lawdong_file(self):
        """법정동 코드 파일 로드 - 동일 이름 시군구 구분 개선 (간결한 표시명)"""
        try:
            if not os.path.exists(self.lawdong_path):
                show_topmost_error("오류", "법정동 코드 파일이 존재하지 않습니다.", parent=self.root)
                return False
                
            for encoding in ['cp949', 'euc-kr', 'utf-8']:
                try:
                    with open(self.lawdong_path, 'r', encoding=encoding) as file:
                        # 법정동 정보를 완전히 저장
                        law_dong_data = []
                        
                        for line in file:
                            parts = line.strip().split('\t')
                            if len(parts) < 2:
                                continue
                            
                            code = parts[0].strip()
                            name = parts[1].strip()
                            
                            # 폐지된 동 필터링
                            if any('폐지' in part for part in parts):
                                continue
                                
                            # 법정동 코드에서 시도, 시군구, 읍면동 코드 추출
                            sido_code = code[:2]  # 앞 2자리는 시도코드
                            sigungu_code = code[2:5]  # 다음 3자리는 시군구코드
                            dong_code = code[5:]  # 나머지 5자리는 읍면동코드
                            
                            # 코드와 이름을 리스트에 저장
                            law_dong_data.append({
                                'code': code,
                                'name': name,
                                'sido_code': sido_code,
                                'sigungu_code': sigungu_code, 
                                'dong_code': dong_code
                            })
                        
                        # 시도 목록 초기화
                        self.sido_list = []
                        self.sigungu_dict = {}
                        self.dong_dict = {}
                        self.region_codes = {}

                        # 시군구 이름 정보 저장을 위한 매핑
                        self.sigungu_to_full_info = {}  # 시군구이름 -> (시도, 시도코드, 시군구코드)
                        self.special_sigungu_names = {}  # 중복 이름 시군구 관리 (강서구, 중구 등)

                        # 시도 정보 추출 (앞 2자리 코드가 같고 나머지가 0인 항목)
                        sido_data = [item for item in law_dong_data if item['code'].endswith('00000000')]
                        for sido in sido_data:
                            sido_name = sido['name']
                            sido_code = sido['sido_code']
                            self.sido_list.append(sido_name)
                            self.sigungu_dict[sido_name] = []

                        # 시군구 정보 추출 및 중복 이름 식별
                        sigungu_data = [item for item in law_dong_data if item['dong_code'] == '00000' and not item['code'].endswith('00000000')]

                        # 모든 시군구 이름과 그 개수 확인 (중복 확인용)
                        sigungu_name_count = {}
                        for item in sigungu_data:
                            names = item['name'].split()
                            if len(names) >= 2:
                                sigungu_name = names[1]  # 두 번째 부분이 시군구명
                                sigungu_name_count[sigungu_name] = sigungu_name_count.get(sigungu_name, 0) + 1

                        # 중복된 이름을 가진 시군구 목록 생성
                        duplicate_sigungu_names = {name for name, count in sigungu_name_count.items() if count > 1}

                        # 각 시도의 시군구 정보 처리
                        for item in sigungu_data:
                            names = item['name'].split()
                            if len(names) >= 2:
                                sido_name = names[0]  # 첫 부분이 시도명
                                sigungu_name = names[1]  # 두 번째 부분이 시군구명

                                if sido_name in self.sido_list:
                                    # 시군구 코드 추출
                                    sigungu_full_code = f"{item['sido_code']}{item['sigungu_code']}"

                                    # 표시 이름 설정 - 기본적으로는 시군구명만 사용
                                    display_name = sigungu_name

                                    # 중복된 이름을 가진 시군구는 특별 처리
                                    if sigungu_name in duplicate_sigungu_names:
                                        # 시도명 약어 생성 (충돌 방지를 위해 첫 2글자 사용)
                                        sido_abbr = sido_name[:2]  # 첫 2글자 사용 (대구광역시 -> 대구, 대전광역시 -> 대전)
                                        if "특별시" in sido_name:
                                            sido_abbr = sido_name.replace("특별시", "")  # 서울특별시 -> 서울
                                        elif "광역시" in sido_name:
                                            sido_abbr = sido_name.replace("광역시", "")  # 대구광역시 -> 대구
                                        elif "특별자치시" in sido_name:
                                            sido_abbr = sido_name.replace("특별자치시", "")  # 세종특별자치시 -> 세종
                                        elif "특별자치도" in sido_name:
                                            sido_abbr = sido_name.replace("특별자치도", "")  # 제주특별자치도 -> 제주
                                        elif sido_name.endswith("도"):
                                            sido_abbr = sido_name.replace("도", "")  # 경기도 -> 경기

                                        # 시군구 표시명에 시도 약어 추가
                                        display_name = f"{sigungu_name}({sido_abbr})"

                                        # 특별 시군구 목록에 추가
                                        self.special_sigungu_names[display_name] = (sido_name, sigungu_name)

                                    # 시군구 정보 저장
                                    self.sigungu_to_full_info[display_name] = (sido_name, sigungu_name, sigungu_full_code)

                                    # 시도별 시군구 목록에 추가
                                    if display_name not in self.sigungu_dict[sido_name]:
                                        self.sigungu_dict[sido_name].append(display_name)
                                        self.dong_dict[display_name] = []

                        # 읍면동 정보 추출
                        for item in law_dong_data:
                            if item['dong_code'] != '00000' and not item['code'].endswith('00000'):
                                names = item['name'].split()
                                if len(names) >= 3:
                                    sido_name = names[0]  # 첫 부분이 시도명
                                    sigungu_name = names[1]  # 두 번째 부분이 시군구명
                                    dong_name = names[2]  # 세 번째 부분이 읍면동명

                                    # 시군구 표시명 찾기
                                    display_name = sigungu_name

                                    # 중복 이름 시군구 처리
                                    if sigungu_name in duplicate_sigungu_names:
                                        # 시도명 약어 생성 (충돌 방지를 위해 첫 2글자 사용)
                                        sido_abbr = sido_name[:2]  # 첫 2글자 사용 (대구광역시 -> 대구, 대전광역시 -> 대전)
                                        if "특별시" in sido_name:
                                            sido_abbr = sido_name.replace("특별시", "")  # 서울특별시 -> 서울
                                        elif "광역시" in sido_name:
                                            sido_abbr = sido_name.replace("광역시", "")  # 대구광역시 -> 대구
                                        elif "특별자치시" in sido_name:
                                            sido_abbr = sido_name.replace("특별자치시", "")  # 세종특별자치시 -> 세종
                                        elif "특별자치도" in sido_name:
                                            sido_abbr = sido_name.replace("특별자치도", "")  # 제주특별자치도 -> 제주
                                        elif sido_name.endswith("도"):
                                            sido_abbr = sido_name.replace("도", "")  # 경기도 -> 경기

                                        display_name = f"{sigungu_name}({sido_abbr})"

                                    # 해당 시군구가 존재하는지 확인
                                    if display_name in self.dong_dict:
                                        # 동이 아직 추가되지 않았다면 추가
                                        if dong_name not in self.dong_dict[display_name]:
                                            self.dong_dict[display_name].append(dong_name)

                                            # 지역 코드 저장
                                            sigungu_code_5digits = f"{item['sido_code']}{item['sigungu_code']}"
                                            self.region_codes[(sido_name, display_name, dong_name)] = (item['code'], sigungu_code_5digits)
                        
                        # 중복 제거 및 정렬
                        self.sido_list = sorted(set(self.sido_list))
                        for sido in self.sido_list:
                            self.sigungu_dict[sido] = sorted(set(self.sigungu_dict[sido]))
                            
                        for sigungu in self.dong_dict:
                            self.dong_dict[sigungu] = sorted(set(self.dong_dict[sigungu]))
                        
                        return True
                        
                except UnicodeDecodeError:
                    continue
            
            show_topmost_error("오류", "법정동 코드 파일을 읽을 수 없습니다. 인코딩을 확인해주세요.", parent=self.root)
            return False
                    
        except Exception as e:
            show_topmost_error("오류", f"법정동 코드 파일 로드 중 오류: {str(e)}", parent=self.root)
            import traceback
            traceback.print_exc()
            return False

    
    # 1. 히스토리 로드 함수 수정 - 그래프 이미지 파일 기준으로 변경
    # 1. 히스토리 로드 함수 수정 - 그래프 이미지 파일 기준으로 변경
    def load_history(self):
        """저장된 히스토리 목록 로드 - 그래프 이미지 파일 기준으로 수정된 버전 (개선)"""
        history_list = []
        if os.path.exists(self.download_path):
            try:
                # 이미지 파일 기준으로 히스토리 생성
                image_extensions = ['.jpg', '.jpeg', '.png']
                image_files = {}
                
                for f in os.listdir(self.download_path):
                    file_lower = f.lower()
                    if any(f.endswith(ext) for ext in image_extensions) and f != "graph.jpg":
                        file_path = os.path.join(self.download_path, f)
                        image_files[f] = os.path.getmtime(file_path)
                
                print(f"[디버그] 발견된 이미지 파일 수: {len(image_files)}")
                
                for image_file in image_files:
                    file_path = os.path.join(self.download_path, image_file)
                    try:
                        # 파일명에서 아파트 정보 추출
                        print(f"[디버그] 처리 중인 파일: {image_file}")
                        
                        if image_file.startswith("multi_apt_comparison_"):
                            # 다중 아파트 비교 그래프인 경우
                            chart_title = "다중 아파트 비교 분석"
                            
                            # 파일명에서 아파트명과 면적 추출
                            parts = image_file[len("multi_apt_comparison_"):].split('_vs_')
                            apt_parts = []
                            for part in parts:
                                if '_' in part and 'm2' in part:
                                    try:
                                        apt_name = '_'.join(part.split('_')[:-1])  # 마지막 부분(면적) 제외
                                        area = part.split('_')[-1].replace('m2', '').split('.')[0]  # 확장자 제거
                                        apt_name = apt_name.replace('_', ' ')  # 언더스코어를 공백으로 치환
                                        apt_parts.append(f"{apt_name} ({area}㎡)")
                                    except:
                                        apt_parts.append(part.split('.')[0])  # 확장자 제거
                            
                            apt_name = " vs ".join(apt_parts)
                            area = apt_name  # 비교아파트 정보를 area 필드에 저장
                            
                        else:
                            # 단일 아파트 그래프인 경우
                            chart_title = "아파트 실거래가 분석"
                            name_part = image_file.split('.')[0]  # 확장자 제거
                            
                            # 이름과 면적 분리
                            if '_' in name_part and 'm2' in name_part:
                                parts = name_part.split('_')
                                if len(parts) >= 2 and 'm2' in parts[-1]:
                                    apt_name = '_'.join(parts[:-1])  # 마지막 부분(면적) 제외
                                    area_part = parts[-1].replace('m2', '')  # m2 제거
                                    apt_name = apt_name.replace('_', ' ')  # 언더스코어를 공백으로 치환
                                    area = f"{apt_name} ({area_part}㎡)"  # 비교아파트 정보 형식으로 저장
                                else:
                                    apt_name = name_part.replace('_', ' ')
                                    area = apt_name
                            else:
                                apt_name = name_part.replace('_', ' ')
                                area = apt_name
                        
                        # 관련 엑셀 파일 찾기
                        excel_filename = None
                        if area and '(' in area and ')' in area:
                            # 면적 정보가 있는 경우
                            area_match = re.search(r'\((\d+)㎡\)', area)
                            if area_match:
                                area_num = area_match.group(1)
                                apt_name_for_excel = apt_name.replace(' ', '_')
                                excel_filename = f"{apt_name_for_excel}_{area_num}m2_매매.xlsx"
                                if not os.path.exists(os.path.join(self.download_path, excel_filename)):
                                    excel_filename = f"{apt_name_for_excel}_{area_num}m2.xlsx"
                        
                        excel_path = None
                        if excel_filename:
                            excel_path = os.path.join(self.download_path, excel_filename)
                            if not os.path.exists(excel_path):
                                excel_path = None
                        
                        history_item = {
                            'image_path': file_path,
                            'excel_path': excel_path,
                            'apt_name': apt_name,
                            'area': area,  # 비교아파트 정보
                            'search_date': image_files[image_file],
                            'chart_title': chart_title,
                            'type': 'multi' if "multi_apt_comparison_" in image_file else 'single'
                        }
                        
                        history_list.append(history_item)
                        print(f"[디버그] 히스토리 항목 추가: {apt_name} - {area}")
                        
                    except Exception as e:
                        print(f"[디버그] 이미지 파일 처리 중 오류 ({image_file}): {str(e)}")
            except Exception as e:
                print(f"[디버그] 히스토리 로드 중 오류: {str(e)}")
                        
        print(f"[디버그] 총 히스토리 항목 수: {len(history_list)}")
        return sorted(history_list, key=lambda x: x['search_date'], reverse=True)
    
    def collect_jeonse_data_btn(self):
        """전세거래가 수집 버튼 처리 함수 - 매매가 수집과 동일한 고속 병렬 처리 방식 적용"""
        if not self.selected_apts:
            show_topmost_error("오류", "선택된 아파트가 없습니다.", parent=self.root)
            return
            
        try:
            # 진행 상태 표시
            self.update_progress(5, "전세 데이터 수집 준비 중...")

            # 그래프 생성 버튼 비활성화 (수집 시작) - 회색으로 표시됨
            self.graph_button.config(state="disabled", text="📊 데이터 수집 중...")
            
            # 전세 데이터 수집이 필요한 아파트만 필터링
            apts_to_collect = []
            for apt_info in self.selected_apts:
                if 'jeonse_data' not in apt_info or not apt_info['jeonse_data']:
                    apts_to_collect.append(apt_info)
                else:
                    # 이미 데이터가 있는 경우
                    apt_name = apt_info['apt_name']
                    area = apt_info['area']
                    self.status_label.config(text=f"{apt_name} ({area}㎡) 전세 데이터가 이미 있습니다.")
            
            if not apts_to_collect:
                show_topmost_info("알림", "모든 선택된 아파트의 전세 데이터가 이미 수집되어 있습니다.", parent=self.root)
                self.update_progress(0, "")
                
                # 데이터가 이미 있으면 그래프 버튼 활성화
                has_data = any(('trades_data' in apt and apt['trades_data']) or ('jeonse_data' in apt and apt['jeonse_data']) 
                             for apt in self.selected_apts)
                if has_data:
                    self.graph_button.config(state="normal")
                return
                
            # 진행 상황 창 생성
            progress_window = tk.Toplevel(self.root)
            progress_window.title("아파트 전세 데이터 수집 중...")
            progress_window.geometry("500x300")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # 전체 진행 상황
            main_label = ttk.Label(progress_window, text="전체 진행 상황:", font=('Malgun Gothic', 11))
            main_label.pack(pady=(10, 5))
            
            main_progress = ttk.Progressbar(progress_window, orient="horizontal", length=450, mode="determinate")
            main_progress.pack(padx=20, pady=5)
            
            main_status = ttk.Label(progress_window, text="0% 완료")
            main_status.pack(pady=5)
            
            # 각 아파트별 진행 상황을 보여줄 프레임
            apt_frame = ttk.Frame(progress_window)
            apt_frame.pack(fill="both", expand=True, padx=20, pady=10)
            
            # 스크롤 가능한 영역 생성
            canvas = tk.Canvas(apt_frame)
            scrollbar = ttk.Scrollbar(apt_frame, orient="vertical", command=canvas.yview)
            
            scrollable_frame = ttk.Frame(canvas)
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # 각 아파트별 진행바 생성
            apt_progress_bars = {}
            apt_labels = {}
            
            for i, apt_info in enumerate(apts_to_collect):
                apt_name = apt_info['apt_name']
                area = apt_info['area']
                
                apt_frame = ttk.Frame(scrollable_frame)
                apt_frame.pack(fill="x", pady=5)
                
                label = ttk.Label(apt_frame, text=f"{apt_name} ({area}㎡):", width=35, anchor="w")
                label.pack(side="left", padx=5)
                
                progress = ttk.Progressbar(apt_frame, orient="horizontal", length=200, mode="determinate")
                progress.pack(side="left", padx=5, fill="x", expand=True)
                
                status = ttk.Label(apt_frame, text="대기 중...", width=20)
                status.pack(side="left", padx=5)
                
                apt_progress_bars[apt_name + str(area)] = progress
                apt_labels[apt_name + str(area)] = status
            
            # 취소 버튼
            cancel_btn = ttk.Button(progress_window, text="취소", command=progress_window.destroy)
            cancel_btn.pack(pady=10)
            
            # 취소 플래그
            cancel_flag = [False]
            progress_window.protocol("WM_DELETE_WINDOW", 
                                  lambda: setattr(cancel_flag, 0, True) or progress_window.destroy())
            
            # 각 아파트별 데이터 수집 함수 정의
            def collect_apt_data(apt_info, lock):
                if cancel_flag[0]:
                    return None
                        
                apt_name = apt_info['apt_name']
                area = apt_info['area']
                key = apt_name + str(area)
                
                # UI 업데이트
                def update_ui(progress_val, message):
                    with lock:
                        if not cancel_flag[0]:
                            apt_progress_bars[key]['value'] = progress_val
                            apt_labels[key].config(text=message)
                            progress_window.update_idletasks()
                
                update_ui(0, "수집 시작...")
                
                # 데이터 수집 - 초고속 최적화
                jeonse_trades = []
                current_date = datetime.now()
                
                # 최적화된 데이터 수집 설정 - 매매가 수집과 동일하게 설정
                max_years = 30  # 최대 30년(최대 기간으로 설정)
                months_per_year = 12
                max_months = max_years * months_per_year
                
                # 초고속 병렬 처리 설정
                batch_size = 120  # 한 번에 10년(120개월)씩 처리
                concurrent_requests = 24  # 동시에 24개 요청 처리
                consecutive_empty_years = 0  # 연속으로 데이터가 없는 년도 수
                
                # API 쿼리 최소화를 위한 설정
                sigungu_code = apt_info['sigungu_code']
                dong = apt_info['dong']
                target_area = float(area)
                
                # 서비스 키 설정
                service_key = "Vs5lXsSo6iEI8no3pP%2FT0udWF9s7Cc8oP1SIWnEI5F4h6dKq92fLvnKmxkoWGJxSeW2%2FSOLQECGxOJzWcjJEXQ%3D%3D"
                
                # 스레딩을 위한 세션 생성 및 최적화
                session = requests.Session()
                adapter = requests.adapters.HTTPAdapter(
                    pool_connections=concurrent_requests,
                    pool_maxsize=concurrent_requests * 2,
                    max_retries=1
                )
                session.mount('http://', adapter)
                session.mount('https://', adapter)
                
                # 최근 1년 데이터 우선 조회 (빠른 피드백을 위해)
                recent_year_trades = []
                futures = []
                
                with concurrent.futures.ThreadPoolExecutor(max_workers=concurrent_requests) as executor:
                    # 최근 1년 요청 준비
                    for i in range(12):
                        search_date = current_date - timedelta(days=30 * i)
                        deal_ymd = search_date.strftime("%Y%m")
                        
                        url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"
                               f"?serviceKey={service_key}"
                               f"&LAWD_CD={sigungu_code}"
                               f"&DEAL_YMD={deal_ymd}"
                               f"&numOfRows=2000")
                        
                        # 비동기 요청 제출 (타임아웃 10초로 증가)
                        futures.append((search_date, executor.submit(session.get, url, timeout=API_TIMEOUT)))
                    
                    # 결과 수집
                    for i, (search_date, future) in enumerate(futures):
                        progress_val = (i / 12) * 10  # 0-10% 진행률 할당
                        year_month = search_date.strftime("%Y-%m")
                        
                        try:
                            response = future.result()
                            monthly_trades = []
                            
                            if response.status_code == 200:
                                try:
                                    root = ET.fromstring(response.text)
                                    items = root.findall('.//item')
                                    
                                    # 데이터 필터링 최적화 - 딕셔너리 캐싱 적용
                                    filtered_items = []
                                    for item in items:
                                        item_apt = item.findtext('aptNm', '').strip()
                                        item_dong = item.findtext('umdNm', '').strip()
                                        
                                        if item_apt == apt_name and item_dong == dong:
                                            filtered_items.append(item)
                                    
                                    # 필터링된 아이템만 상세 처리
                                    for item in filtered_items:
                                        item_area = float(item.findtext('excluUseAr', '0'))
                                        monthly_rent = item.findtext('monthlyRent', '0').replace(',', '')
                                        is_monthly = monthly_rent and int(monthly_rent) > 0
                                        
                                        # 면적 필터링 및 전세만 처리
                                        if abs(item_area - target_area) <= 1 and not is_monthly:
                                            deposit = item.findtext('deposit', '0').replace(',', '')
                                            
                                            try:
                                                trade = {
                                                    'date': datetime(
                                                        int(item.findtext('dealYear', '0')),
                                                        int(item.findtext('dealMonth', '0')),
                                                        int(item.findtext('dealDay', '1') or '1')
                                                    ),
                                                    'price': int(deposit),  # 전세금액
                                                    'floor': int(item.findtext('floor', '0')),
                                                    'area': item_area,
                                                    'rent_type': '전세'
                                                }
                                                monthly_trades.append(trade)
                                            except (ValueError, TypeError) as e:
                                                # 숫자 변환 오류 무시하고 계속
                                                continue
                                except ET.ParseError:
                                    # XML 파싱 오류 무시하고 계속
                                    pass
                                
                                if monthly_trades:
                                    recent_year_trades.extend(monthly_trades)
                                    update_ui(progress_val, f"{year_month}: {len(monthly_trades)}건")
                                else:
                                    update_ui(progress_val, f"{year_month}: 데이터 없음")
                            
                        except Exception as e:
                            update_ui(progress_val, f"{year_month}: 오류")
                
                # 최근 1년 거래 데이터 추가
                jeonse_trades.extend(recent_year_trades)
                update_ui(10, f"최근 1년: {len(recent_year_trades)}건")
                
                # 나머지 29년 데이터 병렬 처리 (10년 블록 단위)
                remaining_years = list(range(1, max_years))
                
                # 10년 단위로 분할하여 처리
                year_blocks = [remaining_years[i:i+10] for i in range(0, len(remaining_years), 10)]
                
                for block_idx, year_block in enumerate(year_blocks):
                    if cancel_flag[0] or consecutive_empty_years >= 2:  # 연속 2년간 거래 없으면 중단
                        break
                        
                    block_trades = []
                    block_progress_start = 10 + (block_idx * (90 / len(year_blocks)))
                    block_progress_end = 10 + ((block_idx + 1) * (90 / len(year_blocks)))
                    
                    update_ui(block_progress_start, f"{year_block[0]}~{year_block[-1]}년 조회 시작")
                    
                    # 각 블록 내 월별 요청 생성
                    months_to_request = []
                    for year in year_block:
                        for month in range(1, 13):
                            # 미래 날짜 건너뛰기
                            if current_date.year - year < 1970 or (current_date.year - year == current_date.year and month > current_date.month):
                                continue
                                
                            try:
                                search_date = current_date.replace(
                                    year=current_date.year - year,
                                    month=month,
                                    day=1
                                )
                                deal_ymd = search_date.strftime("%Y%m")
                                months_to_request.append((search_date, deal_ymd))
                            except ValueError:
                                # 날짜 변환 오류 무시하고 계속
                                continue
                    
                    # 병렬 처리 설정에 따라 더 많은 동시 요청 처리
                    with concurrent.futures.ThreadPoolExecutor(max_workers=concurrent_requests) as executor:
                        # 모든 월 데이터 요청 준비
                        futures = []
                        for search_date, deal_ymd in months_to_request:
                            url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"
                                   f"?serviceKey={service_key}"
                                   f"&LAWD_CD={sigungu_code}"
                                   f"&DEAL_YMD={deal_ymd}"
                                   f"&numOfRows=2000")
                            
                            # 비동기 요청 제출 (작업 큐에 추가)
                            futures.append((search_date, executor.submit(session.get, url, timeout=1)))
                        
                        # 준비 완료 메시지 업데이트
                        update_ui(block_progress_start, f"{year_block[0]}~{year_block[-1]}년: {len(futures)}개월 요청 중")
                        
                        # 비동기 요청 처리를 위한 큐 생성
                        from queue import Queue
                        response_queue = Queue()
                        
                        # 응답 처리 스레드
                        def process_responses():
                            processed = 0
                            total = len(futures)
                            while processed < total and not cancel_flag[0]:
                                try:
                                    # 타임아웃 설정 (큐에서 대기 시간 제한)
                                    idx, response = response_queue.get(timeout=0.5)
                                    processed += 1
                                    
                                    progress_val = block_progress_start + (processed / total) * (block_progress_end - block_progress_start)
                                    
                                    try:
                                        monthly_trades = []
                                        
                                        if response.status_code == 200:
                                            try:
                                                # XML 파싱 최적화
                                                root = ET.fromstring(response.text)
                                                
                                                # 필터링을 먼저 수행하여 처리할 항목 줄이기
                                                items = []
                                                for item in root.findall('.//item'):
                                                    item_apt = item.findtext('aptNm', '').strip()
                                                    item_dong = item.findtext('umdNm', '').strip()
                                                    
                                                    if item_apt == apt_name and item_dong == dong:
                                                        items.append(item)
                                                
                                                # 필터링된 항목만 처리
                                                for item in items:
                                                    item_area = float(item.findtext('excluUseAr', '0'))
                                                    monthly_rent = item.findtext('monthlyRent', '0').replace(',', '')
                                                    is_monthly = monthly_rent and int(monthly_rent) > 0
                                                    
                                                    # 면적 필터링 및 전세만 처리
                                                    if abs(item_area - target_area) <= 1 and not is_monthly:
                                                        deposit = item.findtext('deposit', '0').replace(',', '')
                                                        
                                                        try:
                                                            year = int(item.findtext('dealYear', '0'))
                                                            month = int(item.findtext('dealMonth', '0'))
                                                            day = int(item.findtext('dealDay', '1') or '1')
                                                            price = int(deposit)
                                                            floor = int(item.findtext('floor', '0'))
                                                            
                                                            trade = {
                                                                'date': datetime(year, month, day),
                                                                'price': price,
                                                                'floor': floor,
                                                                'area': item_area,
                                                                'rent_type': '전세'
                                                            }
                                                            monthly_trades.append(trade)
                                                        except (ValueError, TypeError) as e:
                                                            # 숫자 변환 오류 무시하고 계속
                                                            continue
                                            except ET.ParseError:
                                                # XML 파싱 오류 무시하고 계속
                                                pass
                                            
                                            if monthly_trades:
                                                block_trades.extend(monthly_trades)
                                        
                                        update_ui(progress_val, f"처리중: {processed}/{total}")
                                        
                                    except Exception as e:
                                        # 개별 응답 처리 중 오류가 발생해도 전체 프로세스는 계속 진행
                                        continue
                                        
                                    # 큐 작업 완료 신호
                                    response_queue.task_done()
                                    
                                except Exception as e:
                                    # 큐 대기 타임아웃이나 다른 오류는 무시하고 계속 진행
                                    continue
                        
                        # 응답 처리 스레드 시작
                        processor = threading.Thread(target=process_responses)
                        processor.daemon = True
                        processor.start()
                        
                        # 응답이 완료되는 대로 큐에 추가
                        for i, future in enumerate(concurrent.futures.as_completed([f[1] for f in futures])):
                            if cancel_flag[0]:
                                break
                                
                            try:
                                response = future.result()
                                # 결과를 큐에 추가
                                response_queue.put((i, response))
                            except Exception as e:
                                # 요청 실패해도 계속 진행
                                continue
                        
                        # 모든 응답 처리가 완료될 때까지 대기
                        processor.join(timeout=API_TIMEOUT)  # 최대 10초 대기
                    
                    # 해당 블록에 데이터가 있는지 확인
                    if block_trades:
                        jeonse_trades.extend(block_trades)
                        consecutive_empty_years = 0
                        update_ui(block_progress_end, f"{year_block[0]}~{year_block[-1]}년: {len(block_trades)}건")
                    else:
                        consecutive_empty_years += 1
                        update_ui(block_progress_end, f"{year_block[0]}~{year_block[-1]}년: 데이터 없음")
                        
                # 최종 진행 상태 업데이트
                update_ui(100, f"완료 - {len(jeonse_trades)}건")
                
                if jeonse_trades:
                    try:
                        # 날짜 순으로 정렬
                        jeonse_trades = sorted(jeonse_trades, key=lambda x: x['date'])
                        
                        # 데이터프레임 생성 및 엑셀 저장
                        df = pd.DataFrame(jeonse_trades)
                        
                        # 엑셀 저장은 메인 스레드에서 처리 (UI 작업)
                        return {'apt_info': apt_info, 'trades': jeonse_trades, 'df': df}
                    except Exception as e:
                        update_ui(100, f"저장 오류: {str(e)[:20]}")
                        return {'apt_info': apt_info, 'error': str(e)}
                else:
                    update_ui(100, "거래 데이터 없음")
                    return {'apt_info': apt_info, 'trades': []}
            
            # 전체 진행 상황 업데이트 함수
            def update_main_progress(value, message):
                if not cancel_flag[0]:
                    main_progress['value'] = value
                    main_status.config(text=message)
                    progress_window.update_idletasks()
            
            # 병렬 처리 실행 - 아파트 병렬 처리 수준 상향
            def run_parallel_collection():
                results = []
                lock = threading.Lock()
                
                with concurrent.futures.ThreadPoolExecutor(max_workers=min(len(apts_to_collect), 3)) as executor:
                    # 각 아파트별 데이터 수집 함수 제출
                    future_to_apt = {
                        executor.submit(collect_apt_data, apt_info, lock): apt_info
                        for apt_info in apts_to_collect
                    }
                    
                    # 완료된 작업 처리
                    for i, future in enumerate(concurrent.futures.as_completed(future_to_apt)):
                        if cancel_flag[0]:
                            executor.shutdown(wait=False)
                            break
                            
                        apt_info = future_to_apt[future]
                        progress_val = ((i + 1) / len(apts_to_collect)) * 100
                        
                        try:
                            result = future.result()
                            if result:
                                results.append(result)
                                update_main_progress(progress_val, f"{i+1}/{len(apts_to_collect)} 완료")
                        except Exception as e:
                            print(f"데이터 수집 중 오류: {str(e)}")
                            update_main_progress(progress_val, f"오류: {str(e)[:50]}")
                
                # 모든 작업 완료 후 데이터 저장 및 업데이트
                update_main_progress(100, "데이터 저장 중...")
                
                for result in results:
                    if 'error' in result:
                        continue
                        
                    apt_info = result['apt_info']
                    trades = result['trades']
                    df = result.get('df')
                    
                    # 선택된 아파트 정보 업데이트
                    for i, selected_apt in enumerate(self.selected_apts):
                        if selected_apt['apt_name'] == apt_info['apt_name'] and str(selected_apt['area']) == str(apt_info['area']):
                            # 전세 데이터 저장
                            self.selected_apts[i]['jeonse_data'] = trades
                            
                            # 엑셀 파일 저장
                            if df is not None and not df.empty:
                                try:
                                    excel_path = self.save_apt_data_to_excel(df, apt_info, "jeonse")
                                    self.selected_apts[i]['jeonse_excel_path'] = excel_path
                                except Exception as e:
                                    print(f"엑셀 저장 중 오류: {str(e)}")
                            break
                
                # 진행 창 닫기 (약간의 지연 후)
                if not cancel_flag[0]:
                    self.root.after(500, progress_window.destroy)

                    # 히스토리 업데이트 (메인 스레드에서)
                    def update_ui():
                        self.history_list = self.load_history()
                        self.update_history_display()

                        # 상태 업데이트
                        self.update_progress(100, "전세 데이터 수집 완료!")

                        # 그래프 생성 버튼 활성화 (텍스트도 복원)
                        self.graph_button.config(state="normal", text="📊 그래프 생성하기")

                    self.root.after(0, update_ui)

                    # 3초 후 진행바 초기화
                    self.root.after(3000, lambda: self.update_progress(0, ""))
                else:
                    # 취소된 경우에도 그래프 버튼 활성화 (텍스트도 복원)
                    self.root.after(0, lambda: self.graph_button.config(state="normal", text="📊 그래프 생성하기"))

            # 별도 스레드에서 병렬 처리 실행
            threading.Thread(target=run_parallel_collection, daemon=True).start()

        except Exception as e:
            show_topmost_error("오류", f"전세 데이터 수집 준비 중 오류 발생: {str(e)}", parent=self.root)
            self.update_progress(0, "")
            # 오류 발생 시에도 그래프 버튼 활성화 (텍스트도 복원)
            self.graph_button.config(state="normal", text="📊 그래프 생성하기")


    
    def collect_jeonse_data(self, apt_info):
        """선택한 아파트의 전세 데이터 수집 함수 - 인코딩된 서비스 키 사용"""
        apt_name = apt_info['apt_name']
        target_area = apt_info['area']
        sigungu_code = apt_info['sigungu_code']
        dong = apt_info['dong']
        
        # 전월세 API를 위한 인코딩된 서비스 키 직접 지정
        jeonse_service_key = "Vs5lXsSo6iEI8no3pP%2FT0udWF9s7Cc8oP1SIWnEI5F4h6dKq92fLvnKmxkoWGJxSeW2%2FSOLQECGxOJzWcjJEXQ%3D%3D"
        
        jeonse_trades = []
        current_date = datetime.now()
        
        # 데이터 수집 설정
        max_months = 36  # 최대 3년(36개월)
        consecutive_empty_months_limit = 12  # 연속 12개월(1년) 동안 데이터가 없으면 중단
        
        # 진행 상황을 표시할 창 생성
        progress_window = tk.Toplevel(self.root)
        progress_window.title("전세 데이터 수집 중...")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        ttk.Label(progress_window, 
                  text=f"{apt_name} ({target_area}㎡) 전세 데이터를 수집 중입니다...",
                  wraplength=350).pack(pady=10)
        
        progress_label = ttk.Label(progress_window, text="0% 완료")
        progress_label.pack(pady=5)
        
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=350, mode="determinate")
        progress_bar.pack(fill="x", padx=20, pady=10)
        
        cancel_button = ttk.Button(progress_window, text="중단", command=progress_window.destroy)
        cancel_button.pack(pady=5)
        
        # 취소 플래그
        cancel_flag = [False]
        progress_window.protocol("WM_DELETE_WINDOW", lambda: setattr(cancel_flag, 0, True) or progress_window.destroy())
        
        # 디버깅용 로그 파일 설정
        import logging
        log_file = os.path.join(self.download_path, f"jeonse_api_debug_{apt_name}_{target_area}.log")
        logging.basicConfig(filename=log_file, level=logging.INFO, 
                            format='%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        
        # 데이터 수집 함수
        def collect_data():
            nonlocal jeonse_trades
            consecutive_empty_months = 0  # 연속으로 데이터가 없는 월 수
            
            try:
                logging.info("전세 데이터 수집 시작")
                logging.info(f"아파트: {apt_name}, 면적: {target_area}㎡, 지역코드: {sigungu_code}, 동: {dong}")
                
                for month in range(max_months):
                    if cancel_flag[0]:
                        break
                        
                    # 진행 상태 업데이트
                    progress = min(100, (month / max_months) * 100)
                    progress_bar['value'] = progress
                    
                    # 수집한 데이터 개수에 따라 진행 상태 메시지 업데이트
                    if jeonse_trades:
                        progress_label.config(text=f"{progress:.1f}% 완료 - {len(jeonse_trades)}건 수집됨 ({month+1}/{max_months}개월)")
                    else:
                        progress_label.config(text=f"{progress:.1f}% 완료 ({month+1}/{max_months}개월)")
                    
                    progress_window.update_idletasks()
                    
                    # 현재 조회할 월 계산
                    search_date = current_date - timedelta(days=30 * month)
                    deal_ymd = search_date.strftime("%Y%m")
                    
                    # API 호출 - 전월세 API 사용
                    url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"
                           f"?serviceKey={jeonse_service_key}"  # 인코딩된 전월세 API 키 사용
                           f"&LAWD_CD={sigungu_code}"
                           f"&DEAL_YMD={deal_ymd}"
                           f"&numOfRows=1000")
                    
                    try:
                        # 디버깅 로그
                        logging.info(f"API 요청: {url}")
                        
                        response = requests.get(url, timeout=API_TIMEOUT)  # 타임아웃 설정
                        
                        # 디버깅 로그
                        logging.info(f"응답 상태: {response.status_code}")
                        
                        if response.status_code == 200:
                            # 디버깅 로그 (응답 일부만)
                            logging.info(f"응답 내용 일부: {response.text[:500]}")
                            
                            root = ET.fromstring(response.text)
                            
                            # 오류 확인
                            return_code = root.findtext(".//returnCode")
                            return_msg = root.findtext(".//returnMsg")
                            if return_code and return_code != "00":
                                logging.info(f"API 오류: {return_code} - {return_msg}")
                                continue
                            
                            items = root.findall('.//item')
                            logging.info(f"총 항목 수: {len(items)}")
                            
                            # 전세 항목 필터링
                            monthly_trades = []
                            
                            for item in items:
                                item_apt = item.findtext('aptNm', '').strip()
                                item_dong = item.findtext('umdNm', '').strip()
                                rent_gbn = item.findtext('rentGbn', '').strip()
                                
                                # 해당 아파트, 동, 전세 조건 확인
                                if item_apt == apt_name and item_dong == dong and rent_gbn == '전세':
                                    item_area = float(item.findtext('excluUseAr', '0'))
                                    
                                    # 면적 필터링 (±1㎡ 오차 허용)
                                    if abs(item_area - float(target_area)) <= 1:
                                        # 전세금 추출
                                        deposit = item.findtext('deposit', '0').replace(',', '')
                                        logging.info(f"전세 데이터 발견: {item_apt}, {item_area}㎡, {deposit}만원")
                                        
                                        # 거래 정보 구성
                                        try:
                                            trade = {
                                                'date': datetime(
                                                    int(item.findtext('dealYear', '0')),
                                                    int(item.findtext('dealMonth', '0')),
                                                    int(item.findtext('dealDay', '1'))
                                                ),
                                                'price': int(deposit),  # 전세금액
                                                'floor': int(item.findtext('floor', '0')),
                                                'area': item_area,
                                                'rent_type': '전세'
                                            }
                                            monthly_trades.append(trade)
                                        except Exception as e:
                                            logging.error(f"거래 정보 처리 중 오류: {str(e)}")
                            
                            # 이번 달 데이터 개수 확인
                            if monthly_trades:
                                consecutive_empty_months = 0  # 데이터가 있으면 카운터 리셋
                                jeonse_trades.extend(monthly_trades)  # 전체 거래 목록에 추가
                                
                                # 진행 상태 업데이트
                                progress_label.config(text=f"{progress:.1f}% 완료 - {len(jeonse_trades)}건 수집됨 ({month+1}/{max_months}개월)")
                                progress_window.update_idletasks()
                                
                                logging.info(f"{deal_ymd}: {len(monthly_trades)}건 데이터 추가됨")
                            else:
                                consecutive_empty_months += 1  # 데이터가 없으면 카운터 증가
                                logging.info(f"{deal_ymd}: 데이터 없음")
                            
                            # 일정 기간 연속으로 데이터가 없으면 조기 종료
                            if consecutive_empty_months >= consecutive_empty_months_limit:
                                logging.info(f"연속 {consecutive_empty_months}개월 동안 데이터 없음 - 조기 종료")
                                break
                                
                        else:
                            logging.error(f"API 응답 오류: {response.status_code} - {response.text}")
                            
                        # 잠시 대기하여 API 서버 부하 방지
                        time.sleep(0.2)
                        
                    except Exception as e:
                        logging.error(f"API 호출 중 오류: {str(e)}")
                        continue
                
                # 진행 상태 100%로 설정
                progress_bar['value'] = 100
                progress_label.config(text=f"100% 완료 - 총 {len(jeonse_trades)}건 수집됨")
                progress_window.update_idletasks()
                
                # 결과 로그
                logging.info(f"전세 데이터 수집 완료: 총 {len(jeonse_trades)}건")
                
                # 잠시 후 창 닫기
                time.sleep(0.5)
                if not cancel_flag[0]:
                    progress_window.destroy()
                    
            except Exception as e:
                logging.error(f"데이터 수집 중 오류 발생: {str(e)}")
                import traceback
                logging.error(traceback.format_exc())
                show_topmost_error("오류", f"데이터 수집 중 오류 발생: {str(e)}", parent=self.root)
                progress_window.destroy()
        
        # 별도 스레드로 데이터 수집 실행
        thread = threading.Thread(target=collect_data)
        thread.daemon = True
        thread.start()
        
        # 창이 닫힐 때까지 대기
        self.root.wait_window(progress_window)
        
        # 결과 반환
        if cancel_flag[0]:
            return []  # 취소된 경우 빈 목록 반환
        
        # 결과가 없는 경우
        if not jeonse_trades:
            show_topmost_info("알림", f"{apt_name} ({target_area}㎡)의 전세 거래 데이터가 없습니다.", parent=self.root)
            return []
            
        # 결과를 날짜순으로 정렬하여 반환
        return sorted(jeonse_trades, key=lambda x: x['date'])

            
    def collect_selected_apt_data(self, data_type="purchase"):
        """선택된 아파트의 거래 데이터 수집 및 엑셀 저장 (스레딩 오류 수정)"""
        if not self.selected_apts:
            show_topmost_error("오류", "선택된 아파트가 없습니다.", parent=self.root)
            return
            
        try:
            # 데이터 유형에 따라 필드와 메시지 설정
            if data_type == "purchase":
                data_field = 'trades_data'
                api_type = "매매"
                api_endpoint = "http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
            else:  # jeonse
                data_field = 'jeonse_data'
                api_type = "전세"
                api_endpoint = "http://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"
            
            # 진행 상태 표시
            self.update_progress(5, f"{api_type} 데이터 수집 준비 중...")

            # 그래프 버튼 비활성화 (데이터 수집 시작) - 회색으로 표시됨
            self.graph_button.config(state="disabled", text="📊 데이터 수집 중...")
            
            # 데이터 수집이 필요한 아파트만 필터링
            apts_to_collect = []
            for apt_info in self.selected_apts:
                if data_field not in apt_info or not apt_info[data_field]:
                    apts_to_collect.append(apt_info)
                else:
                    # 이미 데이터가 있는 경우
                    apt_name = apt_info['apt_name']
                    area = apt_info['area']
                    self.status_label.config(text=f"{apt_name} ({area}㎡) {api_type} 데이터가 이미 있습니다.")
            
            if not apts_to_collect:
                show_topmost_info("알림", f"모든 선택된 아파트의 {api_type} 데이터가 이미 수집되어 있습니다.", parent=self.root)
                self.update_progress(0, "")
                
                # 데이터가 이미 있으면 그래프 버튼 활성화
                has_data = any(('trades_data' in apt and apt['trades_data']) or ('jeonse_data' in apt and apt['jeonse_data']) 
                              for apt in self.selected_apts)
                if has_data:
                    self.graph_button.config(state="normal")
                return
                
            # 진행 상황 창 생성
            progress_window = tk.Toplevel(self.root)
            progress_window.title(f"아파트 {api_type} 데이터 수집 중...")
            progress_window.geometry("500x300")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # 전체 진행 상황
            main_label = ttk.Label(progress_window, text="전체 진행 상황:", font=('Malgun Gothic', 11))
            main_label.pack(pady=(10, 5))
            
            main_progress = ttk.Progressbar(progress_window, orient="horizontal", length=450, mode="determinate")
            main_progress.pack(padx=20, pady=5)
            
            main_status = ttk.Label(progress_window, text="0% 완료")
            main_status.pack(pady=5)
            
            # 각 아파트별 진행 상황을 보여줄 프레임
            apt_frame = ttk.Frame(progress_window)
            apt_frame.pack(fill="both", expand=True, padx=20, pady=10)
            
            # 스크롤 가능한 영역 생성
            canvas = tk.Canvas(apt_frame)
            scrollbar = ttk.Scrollbar(apt_frame, orient="vertical", command=canvas.yview)
            
            scrollable_frame = ttk.Frame(canvas)
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # 각 아파트별 진행바 생성
            apt_progress_bars = {}
            apt_labels = {}
            
            for i, apt_info in enumerate(apts_to_collect):
                apt_name = apt_info['apt_name']
                area = apt_info['area']
                
                apt_frame = ttk.Frame(scrollable_frame)
                apt_frame.pack(fill="x", pady=5)
                
                label = ttk.Label(apt_frame, text=f"{apt_name} ({area}㎡):", width=35, anchor="w")
                label.pack(side="left", padx=5)
                
                progress = ttk.Progressbar(apt_frame, orient="horizontal", length=200, mode="determinate")
                progress.pack(side="left", padx=5, fill="x", expand=True)
                
                status = ttk.Label(apt_frame, text="대기 중...", width=20)
                status.pack(side="left", padx=5)
                
                apt_progress_bars[apt_name + str(area)] = progress
                apt_labels[apt_name + str(area)] = status
            
            # 취소 버튼
            cancel_btn = ttk.Button(progress_window, text="취소", command=progress_window.destroy)
            cancel_btn.pack(pady=10)
            
            # 취소 플래그
            cancel_flag = [False]
            progress_window.protocol("WM_DELETE_WINDOW", 
                                  lambda: setattr(cancel_flag, 0, True) or progress_window.destroy())
            
            # 각 아파트별 데이터 수집 함수 정의
            def collect_apt_data(apt_info, lock):
                if cancel_flag[0]:
                    return None
                        
                apt_name = apt_info['apt_name']
                area = apt_info['area']
                key = apt_name + str(area)
                
                # UI 업데이트
                def update_ui(progress_val, message):
                    with lock:
                        if not cancel_flag[0]:
                            apt_progress_bars[key]['value'] = progress_val
                            apt_labels[key].config(text=message)
                            progress_window.update_idletasks()
                
                update_ui(0, "수집 시작...")
                
                # 데이터 수집 - 스레딩 오류 수정
                trades = []
                current_date = datetime.now()
                
                # 설정
                max_years = 30  # 최대 30년
                months_per_year = 12
                max_months = max_years * months_per_year
                
                # 병렬 처리 설정
                concurrent_requests = 24  # 동시에 24개 요청 처리
                
                # API 쿼리 최소화를 위한 설정
                sigungu_code = apt_info['sigungu_code']
                dong = apt_info['dong']
                
                # 스레딩을 위한 세션 생성 및 최적화
                session = requests.Session()
                adapter = requests.adapters.HTTPAdapter(
                    pool_connections=concurrent_requests,
                    pool_maxsize=concurrent_requests * 2,
                    max_retries=1
                )
                session.mount('http://', adapter)
                session.mount('https://', adapter)
                
                # 최근 2년(24개월) 데이터 먼저 수집
                recent_trades = []
                
                # 최근 2년 데이터 수집 함수
                def collect_recent_data():
                    monthly_data = []
                    
                    # 스레드 풀 생성
                    with concurrent.futures.ThreadPoolExecutor(max_workers=concurrent_requests) as executor:
                        # 최근 24개월 요청 준비
                        future_results = {}
                        
                        for i in range(24):
                            search_date = current_date - timedelta(days=30 * i)
                            deal_ymd = search_date.strftime("%Y%m")
                            
                            url = (f"{api_endpoint}"
                                   f"?serviceKey={self.service_key}"
                                   f"&LAWD_CD={sigungu_code}"
                                   f"&DEAL_YMD={deal_ymd}"
                                   f"&numOfRows=2000")
                            
                            # 비동기 요청 제출
                            future = executor.submit(session.get, url, timeout=API_TIMEOUT)
                            future_results[future] = (i, search_date)
                        
                        # 결과 수집
                        for i, future in enumerate(concurrent.futures.as_completed(future_results.keys())):
                            idx, search_date = future_results[future]
                            progress_val = (i / 24) * 20
                            year_month = search_date.strftime("%Y-%m")
                            
                            try:
                                response = future.result()
                                
                                if response.status_code == 200:
                                    try:
                                        root = ET.fromstring(response.text)
                                        items = root.findall('.//item')
                                        
                                        monthly_trades = []
                                        
                                        for item in items:
                                            item_apt = item.findtext('aptNm', '').strip()
                                            item_dong = item.findtext('umdNm', '').strip()
                                            
                                            if item_apt == apt_name and item_dong == dong:
                                                item_area = float(item.findtext('excluUseAr', '0'))
                                                
                                                if abs(item_area - float(area)) <= 1:
                                                    if data_type == "purchase":
                                                        # 매매 데이터
                                                        try:
                                                            deal_amount = item.findtext('dealAmount', '0').replace(',', '')
                                                            trade = {
                                                                'date': datetime(
                                                                    int(item.findtext('dealYear', '0')),
                                                                    int(item.findtext('dealMonth', '0')),
                                                                    int(item.findtext('dealDay', '1') or '1')
                                                                ),
                                                                'price': int(deal_amount),
                                                                'floor': int(item.findtext('floor', '0')),
                                                                'area': item_area
                                                            }
                                                            monthly_trades.append(trade)
                                                        except (ValueError, TypeError):
                                                            continue
                                                    else:
                                                        # 전세 데이터
                                                        try:
                                                            monthly_rent = item.findtext('monthlyRent', '0').replace(',', '')
                                                            is_monthly = monthly_rent and int(monthly_rent) > 0
                                                            
                                                            if not is_monthly:
                                                                deposit = item.findtext('deposit', '0').replace(',', '')
                                                                trade = {
                                                                    'date': datetime(
                                                                        int(item.findtext('dealYear', '0')),
                                                                        int(item.findtext('dealMonth', '0')),
                                                                        int(item.findtext('dealDay', '1') or '1')
                                                                    ),
                                                                    'price': int(deposit),
                                                                    'floor': int(item.findtext('floor', '0')),
                                                                    'area': item_area,
                                                                    'rent_type': '전세'
                                                                }
                                                                monthly_trades.append(trade)
                                                        except (ValueError, TypeError):
                                                            continue
                                        
                                        if monthly_trades:
                                            monthly_data.extend(monthly_trades)
                                            update_ui(progress_val, f"{year_month}: {len(monthly_trades)}건")
                                        else:
                                            update_ui(progress_val, f"{year_month}: 데이터 없음")
                                    except ET.ParseError:
                                        update_ui(progress_val, f"{year_month}: XML 오류")
                                else:
                                    update_ui(progress_val, f"{year_month}: 응답오류 {response.status_code}")
                            except Exception as e:
                                update_ui(progress_val, f"{year_month}: 요청오류")
                    
                    return monthly_data
                
                # 최근 데이터 수집
                recent_trades = collect_recent_data()
                trades.extend(recent_trades)
                
                update_ui(20, f"최근 2년: {len(recent_trades)}건")
                
                # 가용성 분석 - 최근 2년 내 데이터 존재 여부
                if not recent_trades:
                    update_ui(95, "최근 2년간 거래 없음")
                    update_ui(100, "데이터 없음")
                    return {'apt_info': apt_info, 'trades': []}
                
                # 나머지 과거 데이터 수집 (5년 단위 블록)
                remaining_years = list(range(2, max_years))
                year_blocks = [remaining_years[i:i+5] for i in range(0, len(remaining_years), 5)]
                
                consecutive_empty_blocks = 0
                
                for block_idx, year_block in enumerate(year_blocks):
                    if cancel_flag[0] or consecutive_empty_blocks >= 1:
                        break
                    
                    block_progress_start = 20 + (block_idx * (80 / len(year_blocks)))
                    block_progress_end = 20 + ((block_idx + 1) * (80 / len(year_blocks)))
                    
                    update_ui(block_progress_start, f"{year_block[0]}~{year_block[-1]}년 조회 중")
                    
                    # 블록 내 월별 데이터 수집 함수
                    def collect_block_data(year_block):
                        block_data = []
                        
                        months_to_request = []
                        for year in year_block:
                            for month in range(1, 13):
                                try:
                                    search_date = current_date.replace(
                                        year=current_date.year - year,
                                        month=month,
                                        day=1
                                    )
                                    deal_ymd = search_date.strftime("%Y%m")
                                    months_to_request.append((search_date, deal_ymd))
                                except ValueError:
                                    continue
                        
                        # 스레드 풀로 블록 데이터 수집
                        with concurrent.futures.ThreadPoolExecutor(max_workers=concurrent_requests) as executor:
                            future_results = {}
                            
                            for search_date, deal_ymd in months_to_request:
                                url = (f"{api_endpoint}"
                                       f"?serviceKey={self.service_key}"
                                       f"&LAWD_CD={sigungu_code}"
                                       f"&DEAL_YMD={deal_ymd}"
                                       f"&numOfRows=2000")
                                
                                future = executor.submit(session.get, url, timeout=API_TIMEOUT)
                                future_results[future] = (search_date, deal_ymd)
                            
                            # 결과 처리
                            processed = 0
                            for future in concurrent.futures.as_completed(future_results.keys()):
                                search_date, deal_ymd = future_results[future]
                                processed += 1
                                year_month = search_date.strftime("%Y-%m")
                                
                                progress_val = block_progress_start + (processed / len(months_to_request)) * (block_progress_end - block_progress_start)
                                
                                try:
                                    response = future.result()
                                    monthly_trades = []
                                    
                                    if response.status_code == 200:
                                        try:
                                            root = ET.fromstring(response.text)
                                            items = root.findall('.//item')
                                            
                                            for item in items:
                                                item_apt = item.findtext('aptNm', '').strip()
                                                item_dong = item.findtext('umdNm', '').strip()
                                                
                                                if item_apt == apt_name and item_dong == dong:
                                                    item_area = float(item.findtext('excluUseAr', '0'))
                                                    
                                                    if abs(item_area - float(area)) <= 1:
                                                        if data_type == "purchase":
                                                            # 매매 데이터
                                                            try:
                                                                deal_amount = item.findtext('dealAmount', '0').replace(',', '')
                                                                trade = {
                                                                    'date': datetime(
                                                                        int(item.findtext('dealYear', '0')),
                                                                        int(item.findtext('dealMonth', '0')),
                                                                        int(item.findtext('dealDay', '1') or '1')
                                                                    ),
                                                                    'price': int(deal_amount),
                                                                    'floor': int(item.findtext('floor', '0')),
                                                                    'area': item_area
                                                                }
                                                                monthly_trades.append(trade)
                                                            except (ValueError, TypeError):
                                                                continue
                                                        else:
                                                            # 전세 데이터
                                                            try:
                                                                monthly_rent = item.findtext('monthlyRent', '0').replace(',', '')
                                                                is_monthly = monthly_rent and int(monthly_rent) > 0
                                                                
                                                                if not is_monthly:
                                                                    deposit = item.findtext('deposit', '0').replace(',', '')
                                                                    trade = {
                                                                        'date': datetime(
                                                                            int(item.findtext('dealYear', '0')),
                                                                            int(item.findtext('dealMonth', '0')),
                                                                            int(item.findtext('dealDay', '1') or '1')
                                                                        ),
                                                                        'price': int(deposit),
                                                                        'floor': int(item.findtext('floor', '0')),
                                                                        'area': item_area,
                                                                        'rent_type': '전세'
                                                                    }
                                                                    monthly_trades.append(trade)
                                                            except (ValueError, TypeError):
                                                                continue
                                            
                                            if monthly_trades:
                                                block_data.extend(monthly_trades)
                                                if processed % 10 == 0 or len(monthly_trades) > 0:
                                                    update_ui(progress_val, f"{year_month}: {len(monthly_trades)}건")
                                        except ET.ParseError:
                                            continue
                                except Exception as e:
                                    continue
                        
                        return block_data
                    
                    # 블록 데이터 수집 실행
                    block_trades = collect_block_data(year_block)
                    
                    if block_trades:
                        trades.extend(block_trades)
                        consecutive_empty_blocks = 0
                        update_ui(block_progress_end, f"{year_block[0]}~{year_block[-1]}년: {len(block_trades)}건")
                    else:
                        consecutive_empty_blocks += 1
                        update_ui(block_progress_end, f"{year_block[0]}~{year_block[-1]}년: 데이터 없음")
                
                # 최종 진행 상태 업데이트
                if trades:
                    update_ui(95, "데이터 정리 중...")
                    trades = sorted(trades, key=lambda x: x['date'])
                    update_ui(100, f"완료 - {len(trades)}건")
                    
                    try:
                        # 데이터프레임 생성
                        df = pd.DataFrame(trades)
                        return {'apt_info': apt_info, 'trades': trades, 'df': df}
                    except Exception as e:
                        update_ui(100, f"저장 오류: {str(e)[:20]}")
                        return {'apt_info': apt_info, 'error': str(e)}
                else:
                    update_ui(100, "거래 데이터 없음")
                    return {'apt_info': apt_info, 'trades': []}
            
            # 전체 진행 상황 업데이트 함수
            def update_main_progress(value, message):
                if not cancel_flag[0]:
                    main_progress['value'] = value
                    main_status.config(text=message)
                    progress_window.update_idletasks()
            
            # 병렬 처리 실행 - 아파트 병렬 처리 수준은 3개로 제한 (안정성 위해)
            def run_parallel_collection():
                results = []
                lock = threading.Lock()
                
                with concurrent.futures.ThreadPoolExecutor(max_workers=min(len(apts_to_collect), 3)) as executor:
                    # 각 아파트별 데이터 수집 함수 제출
                    future_to_apt = {
                        executor.submit(collect_apt_data, apt_info, lock): apt_info
                        for apt_info in apts_to_collect
                    }
                    
                    # 완료된 작업 처리
                    for i, future in enumerate(concurrent.futures.as_completed(future_to_apt)):
                        if cancel_flag[0]:
                            executor.shutdown(wait=False)
                            break
                            
                        apt_info = future_to_apt[future]
                        progress_val = ((i + 1) / len(apts_to_collect)) * 100
                        
                        try:
                            result = future.result()
                            if result:
                                results.append(result)
                                update_main_progress(progress_val, f"{i+1}/{len(apts_to_collect)} 완료")
                        except Exception as e:
                            print(f"데이터 수집 중 오류: {str(e)}")
                            update_main_progress(progress_val, f"오류: {str(e)[:50]}")
                
                # 모든 작업 완료 후 데이터 저장 및 업데이트
                update_main_progress(100, "데이터 저장 중...")
                
                for result in results:
                    if 'error' in result:
                        continue
                        
                    apt_info = result['apt_info']
                    trades = result['trades']
                    df = result.get('df')
                    
                    # 선택된 아파트 정보 업데이트
                    for i, selected_apt in enumerate(self.selected_apts):
                        if selected_apt['apt_name'] == apt_info['apt_name'] and str(selected_apt['area']) == str(apt_info['area']):
                            # 데이터 유형에 따라 저장할 필드 결정
                            self.selected_apts[i][data_field] = trades
                            
                            # 엑셀 파일 저장
                            if df is not None and not df.empty:
                                try:
                                    excel_path = self.save_apt_data_to_excel(df, apt_info, data_type)
                                    if data_type == "purchase":
                                        self.selected_apts[i]['excel_path'] = excel_path
                                    else:
                                        self.selected_apts[i]['jeonse_excel_path'] = excel_path
                                except Exception as e:
                                    print(f"엑셀 저장 중 오류: {str(e)}")
                            break
                
                # 진행 창 닫기 (약간의 지연 후)
                if not cancel_flag[0]:
                    self.root.after(500, progress_window.destroy)

                    # 히스토리 업데이트 (메인 스레드에서)
                    def update_ui():
                        self.history_list = self.load_history()
                        self.update_history_display()

                        # 상태 업데이트
                        self.update_progress(100, f"{api_type} 데이터 수집 완료!")

                        # 그래프 생성 버튼 활성화 (텍스트도 복원)
                        self.graph_button.config(state="normal", text="📊 그래프 생성하기")

                    self.root.after(0, update_ui)

                    # 3초 후 진행바 초기화
                    self.root.after(3000, lambda: self.update_progress(0, ""))
                else:
                    # 취소된 경우에도 그래프 버튼 활성화 (텍스트도 복원)
                    self.root.after(0, lambda: self.graph_button.config(state="normal", text="📊 그래프 생성하기"))

            # 별도 스레드에서 병렬 처리 실행
            threading.Thread(target=run_parallel_collection, daemon=True).start()

        except Exception as e:
            show_topmost_error("오류", f"데이터 수집 준비 중 오류 발생: {str(e)}", parent=self.root)
            self.update_progress(0, "")
            # 오류 발생 시에도 그래프 버튼 활성화 (텍스트도 복원)
            self.graph_button.config(state="normal", text="📊 그래프 생성하기")

    
    # 2. 히스토리 표시 업데이트 함수 수정
    # 히스토리 표시 업데이트 함수 수정
    # 히스토리 표시 업데이트 함수 수정
    def update_history_display(self):
        """히스토리 목록 업데이트 - 더 간소화된 컬럼"""
        # 기존 항목 삭제
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # 최신 순으로 정렬
        sorted_history = sorted(self.history_list, key=lambda x: x['search_date'], reverse=True)
        
        # 새로운 항목 추가
        for item in sorted_history:
            search_date = datetime.fromtimestamp(item['search_date'])
            
            # 비교아파트 정보 - 기존 area 필드를 활용
            compare_apt = item['area']
            if item['type'] == 'multi' and 'vs' in item['apt_name']:
                compare_apt = item['apt_name']  # 다중 비교인 경우 전체 비교 정보 표시
            
            # 더 간소화된 컬럼으로 표시 (차트제목 열 제거)
            self.history_tree.insert("", "end", values=(
                search_date.strftime("%Y-%m-%d %H:%M"),
                compare_apt
            ))
    
    def setup_gui(self):
        """GUI 구성 - 매매/전세 수집 버튼 제거 및 자동 데이터 수집 적용"""
        # 메인 프레임 (좌측: 컨텐츠, 우측: 히스토리)
        self.root.columnconfigure(0, weight=3)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 좌측 메인 컨텐츠 프레임 (카드 스타일)
        main_frame = ttk.Frame(self.root, padding="15", style='Card.TFrame')
        main_frame.grid(row=0, column=0, sticky="nsew", padx=(10,5), pady=10)

        # 우측 히스토리 프레임 (개선된 스타일)
        history_frame = ttk.LabelFrame(self.root, text="📋 검색 히스토리", padding="15")
        history_frame.grid(row=0, column=1, sticky="nsew", padx=(5,10), pady=10)

        # 제목 레이블 (더 크고 눈에 띄게)
        title_label = ttk.Label(main_frame, text="📊 부태리의 실거래가 차트",
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, pady=(0, 15), sticky="w")

        # 설정 버튼 (우측 상단 - 독립 컬럼)
        settings_button = ttk.Button(main_frame, text="⚙ 설정", width=8,
                                     command=self.show_settings_dialog)
        settings_button.grid(row=0, column=1, sticky="e", pady=(0, 15), padx=(0, 0))

        # 안내 레이블 (더 보기 좋게)
        guide_text = "1️⃣ 지역 선택 → 아파트 선택\n2️⃣ 자동으로 매매/전세 데이터 수집\n3️⃣ 그래프 옵션 선택 후 '생성' 버튼 클릭"
        self.guide_label = ttk.Label(main_frame,
                                    text=guide_text,
                                    style='Subtitle.TLabel',
                                    wraplength=400)
        self.guide_label.grid(row=1, column=0, columnspan=2, pady=(0,10))

        # 지역 검색 프레임 (개선된 스타일)
        region_frame = ttk.LabelFrame(main_frame, text="🔍 지역 검색", padding=15)
        region_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(5,10))
        
        # 지역 선택 콤보박스들
        ttk.Label(region_frame, text="시/도:").grid(row=0, column=0, sticky="w", pady=5)
        self.sido_combobox = ttk.Combobox(region_frame, width=20)

        # 주요 도시 우선 정렬
        priority_cities = ['서울특별시', '경기도', '인천광역시', '부산광역시', '대전광역시', '대구광역시', '광주광역시']
        sorted_sido_list = []

        # 우선 순위 도시 먼저 추가
        for city in priority_cities:
            if city in self.sido_list:
                sorted_sido_list.append(city)

        # 나머지 도시 알파벳 순으로 추가
        for city in sorted(self.sido_list):
            if city not in priority_cities:
                sorted_sido_list.append(city)

        self.sido_combobox['values'] = sorted_sido_list
        self.sido_combobox.set("시/도 선택")
        self.sido_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.sido_combobox.bind('<<ComboboxSelected>>', self.on_sido_selected)
        
        ttk.Label(region_frame, text="시/군/구:").grid(row=1, column=0, sticky="w", pady=5)
        self.sigungu_combobox = ttk.Combobox(region_frame, width=20)
        self.sigungu_combobox.set("시/군/구 선택")
        self.sigungu_combobox.grid(row=1, column=1, padx=5, pady=5)
        self.sigungu_combobox.bind('<<ComboboxSelected>>', self.on_sigungu_selected)
        
        ttk.Label(region_frame, text="읍/면/동:").grid(row=2, column=0, sticky="w", pady=5)
        self.dong_combobox = ttk.Combobox(region_frame, width=20)
        self.dong_combobox.set("읍/면/동 선택")
        self.dong_combobox.grid(row=2, column=1, padx=5, pady=5)
        self.dong_combobox.bind('<<ComboboxSelected>>', self.on_dong_selected)

        # 아파트 목록 조회 버튼 (액센트 스타일) - 변수로 저장
        self.apt_list_button = ttk.Button(region_frame, text="🏢 아파트 목록 조회",
                                          command=self.show_apt_list,
                                          style='Accent.TButton')
        self.apt_list_button.grid(row=3, column=0, columnspan=2, pady=15)

        # 선택된 아파트 정보 표시 프레임 (개선된 스타일)
        selected_apt_frame = ttk.LabelFrame(main_frame, text="✅ 선택된 아파트 목록", padding=15)
        selected_apt_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(5,10))
        
        # 선택된 아파트 목록을 표시할 리스트박스 추가
        apt_list_frame = ttk.Frame(selected_apt_frame)
        apt_list_frame.pack(fill="both", expand=True)
        
        self.selected_apt_listbox = tk.Listbox(apt_list_frame, height=4, font=self.font_normal)
        self.selected_apt_listbox.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(apt_list_frame, orient="vertical", command=self.selected_apt_listbox.yview)
        self.selected_apt_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        
        # 데이터 수집 옵션 프레임
        collect_options_frame = ttk.Frame(selected_apt_frame)
        collect_options_frame.pack(fill="x", padx=5, pady=5)

        # 전세 데이터 포함 조회 체크박스 (크고 눈에 띄게)
        jeonse_collect_check = ttk.Checkbutton(
            collect_options_frame,
            text="🏠 전세 데이터 포함 조회",
            variable=self.collect_jeonse_data,
            command=self.toggle_jeonse_options
        )
        jeonse_collect_check.pack(side="left", padx=5)

        # 안내 레이블
        info_label = ttk.Label(
            collect_options_frame,
            text="※ 전세 데이터가 많으면 API 제한이 발생할 수 있습니다",
            foreground="gray",
            font=("Malgun Gothic", 8)
        )
        info_label.pack(side="left", padx=10)

        # 아파트 선택 관리 버튼 프레임 (매매/전세 버튼 제거)
        apt_btn_frame = ttk.Frame(selected_apt_frame)
        apt_btn_frame.pack(fill="x", padx=5, pady=5)

        # 선택 삭제 버튼
        delete_apt_btn = ttk.Button(apt_btn_frame, text="선택 삭제", command=self.delete_selected_apt)
        delete_apt_btn.pack(side="left", padx=5)

        # 모두 삭제 버튼
        clear_apt_btn = ttk.Button(apt_btn_frame, text="모두 삭제", command=self.clear_all_apts)
        clear_apt_btn.pack(side="left", padx=5)

        # 캐시 초기화 버튼
        clear_cache_btn = ttk.Button(apt_btn_frame, text="🗑 캐시 초기화", command=self.clear_cache)
        clear_cache_btn.pack(side="right", padx=5)
        
        # 그래프 옵션 프레임 (개선된 스타일)
        graph_options_frame = ttk.LabelFrame(main_frame, text="📈 그래프 옵션", padding=15)
        graph_options_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(5,10))

        # 매매가 옵션 프레임
        purchase_frame = ttk.LabelFrame(graph_options_frame, text="💰 매매가", padding=10)
        purchase_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        # 매매가 체크박스 생성
        ttk.Checkbutton(purchase_frame, text="📊 월평균", variable=self.show_monthly_avg).pack(side="left", padx=10, pady=5)
        ttk.Checkbutton(purchase_frame, text="📈 월최고", variable=self.show_monthly_max).pack(side="left", padx=10, pady=5)
        ttk.Checkbutton(purchase_frame, text="⚫ 점도표", variable=self.show_scatter_plot).pack(side="left", padx=10, pady=5)

        # 전세가 옵션 프레임
        jeonse_frame = ttk.LabelFrame(graph_options_frame, text="🏠 전세가", padding=10)
        jeonse_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        # 전세가 체크박스 생성
        ttk.Checkbutton(jeonse_frame, text="📊 월평균", variable=self.show_jeonse_monthly_avg).pack(side="left", padx=10, pady=5)
        ttk.Checkbutton(jeonse_frame, text="📈 월최고", variable=self.show_jeonse_monthly_max).pack(side="left", padx=10, pady=5)
        ttk.Checkbutton(jeonse_frame, text="⚫ 점도표", variable=self.show_jeonse_scatter_plot).pack(side="left", padx=10, pady=5)

        # 그래프 생성 버튼 (더 크고 눈에 띄게)
        # 초기 상태: 비활성화 (데이터 수집 완료 후 활성화)
        self.graph_button = ttk.Button(main_frame, text="📊 그래프 생성하기",
                                      command=self.create_graph_only,
                                      style='Success.TButton',
                                      state="disabled")
        self.graph_button.grid(row=5, column=0, columnspan=2, pady=15, ipady=10)
        
        # 프로그레스 바와 상태 레이블
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=6, column=0, columnspan=2, pady=10, padx=5, sticky="ew")
        
        self.status_label = ttk.Label(main_frame, text="", wraplength=500)
        self.status_label.grid(row=7, column=0, columnspan=2, pady=5, sticky="ew")
        
        # 히스토리 영역 구성
        history_btn_frame = ttk.Frame(history_frame)
        history_btn_frame.pack(side="bottom", fill="x", pady=(5, 0))
        
        # 폴더 열기 버튼
        open_folder_btn = ttk.Button(history_btn_frame, text="저장 폴더 열기", command=self.open_download_folder)
        open_folder_btn.pack(side="left", padx=5)
        
        # 삭제 버튼
        delete_btn = ttk.Button(history_btn_frame, text="선택 삭제", command=self.delete_selected_history)
        delete_btn.pack(side="right", padx=5)
        
        delete_all_btn = ttk.Button(history_btn_frame, text="전체 삭제", command=self.delete_all_history)
        delete_all_btn.pack(side="right", padx=5)
        
        # 히스토리 트리뷰
        self.history_tree = ttk.Treeview(history_frame, 
                                      columns=("date", "compare_apt"),
                                      show="headings",
                                      height=24)
        
        # 컬럼 헤더 설정
        self.history_tree.heading("date", text="검색일시")
        self.history_tree.heading("compare_apt", text="비교아파트")
        
        # 컬럼 너비 조정
        self.history_tree.column("date", width=120)
        self.history_tree.column("compare_apt", width=270)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(history_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        
        self.history_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 히스토리 클릭 이벤트 바인딩
        self.history_tree.bind('<Double-1>', self.on_history_select)
        
        # 하단 정보 레이블 - 만든이 정보 추가 및 클릭 이벤트 설정
        self.footer_label = ttk.Label(main_frame, text="만든이 부태리", 
                                  font=self.font_normal, foreground="blue", cursor="hand2")
        self.footer_label.grid(row=8, column=0, columnspan=2, pady=(0, 5), sticky="s")
        
        # 클릭 이벤트: 블로그 URL로 이동
        self.footer_label.bind("<Button-1>", self.open_blog)


        # 그래프 옵션 프레임에서 전세가 옵션 프레임 아래에 추가
        # 세부정보 가져오기 옵션 추가
        detail_frame = ttk.LabelFrame(graph_options_frame, text="부가옵션", padding=5)
        detail_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        
        # 세부정보 가져오기 체크박스 (기본값: False로 설정)
        ttk.Checkbutton(detail_frame, text="단지세부정보 가져오기 (시간이 걸릴 수 있습니다)", 
                       variable=self.show_complex_info).pack(side="left", padx=10, pady=5)

        
    
    def open_blog(self, event):
        """만든이 부태리 블로그 열기"""
        import webbrowser
        blog_url = "https://blog.naver.com/landlover333"
        webbrowser.open_new(blog_url)
    
    
    def create_graph_only(self):
        """거래 데이터를 사용하여 그래프만 생성 (데이터 수집 제외)"""
        if not self.selected_apts:
            show_topmost_error("오류", "선택된 아파트가 없습니다.", parent=self.root)
            return
        
        # 데이터가 수집되었는지 확인 (매매거래 또는 전세거래 데이터 중 하나만 있으면 됨)
        apts_with_data = [apt for apt in self.selected_apts 
                         if ('trades_data' in apt and apt['trades_data']) or 
                            ('jeonse_data' in apt and apt['jeonse_data'])]
        
        if not apts_with_data:
            show_topmost_error("오류", "수집된 거래 데이터가 없습니다. 먼저 '매매거래가 수집' 또는 '전세거래가 수집' 버튼을 클릭하세요.", parent=self.root)
            return
        
        try:
            # 그래프 생성 버튼 비활성화 (처리 시작)
            self.graph_button.config(state="disabled")
            
            self.update_progress(10, "그래프 생성 준비 중...")
            
            # 데이터프레임 및 월별 데이터 준비
            apt_dfs = []        # 매매 데이터프레임
            monthly_dfs = []    # 매매 월별 데이터프레임
            jeonse_dfs = []     # 전세 데이터프레임
            monthly_jeonse_dfs = [] # 전세 월별 데이터프레임
            
            for i, apt_info in enumerate(apts_with_data):
                # 매매 거래 데이터 처리
                if 'trades_data' in apt_info and apt_info['trades_data']:
                    trades = apt_info['trades_data']
                    # 데이터 정렬
                    trades = sorted(trades, key=lambda x: x['date'])
                    
                    # 데이터프레임 생성
                    df = pd.DataFrame(trades)
                    df['apt_name'] = apt_info['apt_name']
                    df['area'] = apt_info['area']
                    df['data_type'] = 'purchase'  # 매매 데이터 타입 표시
                    
                    apt_dfs.append(df)
                    
                    # 월별 평균가격 계산
                    df_monthly = df.groupby(pd.Grouper(key='date', freq='ME')).agg({'price': 'mean'}).reset_index()
                    df_monthly['apt_name'] = apt_info['apt_name']
                    df_monthly['area'] = apt_info['area']
                    df_monthly['data_type'] = 'purchase'  # 매매 데이터 타입 표시
                    monthly_dfs.append(df_monthly)
                
                # 전세 거래 데이터 처리
                if 'jeonse_data' in apt_info and apt_info['jeonse_data']:
                    jeonse_trades = apt_info['jeonse_data']
                    # 데이터 정렬
                    jeonse_trades = sorted(jeonse_trades, key=lambda x: x['date'])
                    
                    # 데이터프레임 생성
                    jeonse_df = pd.DataFrame(jeonse_trades)
                    jeonse_df['apt_name'] = apt_info['apt_name']
                    jeonse_df['area'] = apt_info['area']
                    jeonse_df['data_type'] = 'jeonse'  # 전세 데이터 타입 표시
                    
                    jeonse_dfs.append(jeonse_df)
                    
                    # 월별 평균가격 계산
                    jeonse_monthly = jeonse_df.groupby(pd.Grouper(key='date', freq='ME')).agg({'price': 'mean'}).reset_index()
                    jeonse_monthly['apt_name'] = apt_info['apt_name']
                    jeonse_monthly['area'] = apt_info['area']
                    jeonse_monthly['data_type'] = 'jeonse'  # 전세 데이터 타입 표시
                    monthly_jeonse_dfs.append(jeonse_monthly)
            
            self.update_progress(30, "그래프 생성 중...")
            
            # 그래프 생성
            try:
                self.create_multi_chart(apt_dfs, monthly_dfs, jeonse_dfs, monthly_jeonse_dfs)
                self.update_progress(90, "그래프 생성 완료!")
                
                # 그래프 이미지 표시
                if os.path.exists(self.image_path):
                    os.startfile(self.image_path)
                    
                # 3초 후 메시지 초기화
                self.root.after(3000, lambda: self.update_progress(0, ""))
            except Exception as e:
                show_topmost_error("오류", f"그래프 생성 중 오류 발생: {str(e)}", parent=self.root)
                self.update_progress(0, "")
                import traceback
                traceback.print_exc()
            finally:
                # 그래프 생성 버튼 다시 활성화 (완료 또는 오류 시)
                self.graph_button.config(state="normal")
            
        except Exception as e:
            show_topmost_error("오류", f"그래프 준비 중 오류 발생: {str(e)}", parent=self.root)
            self.update_progress(0, "")
            # 오류 발생 시에도 그래프 버튼 활성화
            self.graph_button.config(state="normal")
        
    def delete_selected_apt(self):
        """선택된 아파트 목록에서 선택한 항목 삭제"""
        selection = self.selected_apt_listbox.curselection()
        if not selection:
            show_topmost_info("알림", "삭제할 아파트를 선택해주세요.", parent=self.root)
            return

        idx = selection[0]
        self.selected_apt_listbox.delete(idx)
        self.selected_apts.pop(idx)

        # 남은 아파트가 없거나 데이터가 수집되지 않은 경우 버튼 비활성화
        if len(self.selected_apts) == 0:
            self.graph_button.config(state="disabled")
            self.status_label.config(text="선택된 아파트가 없습니다.")
        else:
            # 남은 아파트 중 데이터가 수집된 것이 있는지 확인
            has_data = any(('trades_data' in apt and apt.get('trades_data')) or
                          ('jeonse_data' in apt and apt.get('jeonse_data'))
                          for apt in self.selected_apts)
            if not has_data:
                self.graph_button.config(state="disabled")
            self.status_label.config(text=f"선택된 아파트: {len(self.selected_apts)}개")
    
    def clear_all_apts(self):
        """선택된 아파트 목록 모두 삭제"""
        if not self.selected_apts:
            return

        self.selected_apt_listbox.delete(0, tk.END)
        self.selected_apts = []
        self.graph_button.config(state="disabled")
        self.status_label.config(text="선택된 아파트가 없습니다.")

    def clear_cache(self):
        """아파트 목록 캐시 초기화"""
        cache_count = len(self.apt_list_cache)
        if cache_count == 0:
            show_topmost_info("캐시 초기화", "초기화할 캐시가 없습니다.", parent=self.root)
            return

        response = ask_topmost_yesno(
            "캐시 초기화",
            f"현재 {cache_count}개 지역의 아파트 목록이 캐시되어 있습니다.\n\n"
            f"캐시를 초기화하시겠습니까?\n"
            f"(다음 조회 시 API에서 새로 불러옵니다)",
            parent=self.root
        )

        if response:
            self.apt_list_cache.clear()
            show_topmost_info("캐시 초기화", "캐시가 초기화되었습니다.", parent=self.root)
            print(f"🗑 캐시 초기화 완료 ({cache_count}개 지역)")

    def toggle_jeonse_options(self):
        """전세 데이터 수집 체크박스 상태 변경 시 호출"""
        if self.collect_jeonse_data.get():
            print("✅ 전세 데이터 수집 활성화")
        else:
            print("❌ 전세 데이터 수집 비활성화")

    def get_trade_cache_filename(self, sido, sigungu, dong, apt_name, area, data_type):
        """거래 데이터 캐시 파일 경로 생성"""
        # 파일명에 사용할 수 없는 문자 제거
        safe_sido = sido.replace('/', '_').replace('\\', '_')
        safe_sigungu = sigungu.replace('/', '_').replace('\\', '_')
        safe_dong = dong.replace('/', '_').replace('\\', '_')
        safe_apt = apt_name.replace('/', '_').replace('\\', '_').replace(' ', '_')

        # 지역 폴더 경로
        region_folder = os.path.join(self.trade_cache_path, f"{safe_sido}_{safe_sigungu}_{safe_dong}")
        os.makedirs(region_folder, exist_ok=True)

        # 파일명: 아파트명_면적_데이터타입.json
        filename = f"{safe_apt}_{area}_{data_type}.json"
        return os.path.join(region_folder, filename)

    def save_trade_cache(self, sido, sigungu, dong, apt_name, area, data_type, trades_data):
        """거래 데이터를 캐시 파일로 저장"""
        try:
            cache_file = self.get_trade_cache_filename(sido, sigungu, dong, apt_name, area, data_type)

            # 데이터를 JSON으로 직렬화 (datetime 변환)
            cache_data = {
                'apt_name': apt_name,
                'area': area,
                'sido': sido,
                'sigungu': sigungu,
                'dong': dong,
                'data_type': data_type,
                'cached_at': datetime.now().isoformat(),
                'trades': []
            }

            # 거래 데이터 변환
            for trade in trades_data:
                trade_copy = trade.copy()
                if 'date' in trade_copy and isinstance(trade_copy['date'], datetime):
                    trade_copy['date'] = trade_copy['date'].isoformat()
                cache_data['trades'].append(trade_copy)

            # JSON 파일로 저장
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)

            print(f"💾 캐시 저장: {apt_name} ({area}㎡) - {data_type} ({len(trades_data)}건)")
            return True
        except Exception as e:
            print(f"⚠️ 캐시 저장 실패: {str(e)}")
            return False

    def load_trade_cache(self, sido, sigungu, dong, apt_name, area, data_type):
        """캐시 파일에서 거래 데이터 로드"""
        try:
            cache_file = self.get_trade_cache_filename(sido, sigungu, dong, apt_name, area, data_type)

            if not os.path.exists(cache_file):
                return None

            with open(cache_file, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)

            # datetime 복원
            trades = []
            for trade in cache_data.get('trades', []):
                trade_copy = trade.copy()
                if 'date' in trade_copy and isinstance(trade_copy['date'], str):
                    trade_copy['date'] = datetime.fromisoformat(trade_copy['date'])
                trades.append(trade_copy)

            print(f"📂 캐시 로드: {apt_name} ({area}㎡) - {data_type} ({len(trades)}건)")
            return trades
        except Exception as e:
            print(f"⚠️ 캐시 로드 실패: {str(e)}")
            return None

    def clear_trade_cache(self):
        """거래 데이터 캐시 전체 삭제"""
        try:
            if not os.path.exists(self.trade_cache_path):
                show_topmost_info("캐시 삭제", "삭제할 캐시가 없습니다.", parent=self.root)
                return

            # 캐시 폴더 크기 계산
            total_size = 0
            file_count = 0
            for root, dirs, files in os.walk(self.trade_cache_path):
                for file in files:
                    if file.endswith('.json'):
                        file_path = os.path.join(root, file)
                        total_size += os.path.getsize(file_path)
                        file_count += 1

            if file_count == 0:
                show_topmost_info("캐시 삭제", "삭제할 캐시가 없습니다.", parent=self.root)
                return

            size_mb = total_size / (1024 * 1024)
            response = ask_topmost_yesno(
                "캐시 삭제",
                f"현재 {file_count}개의 캐시 파일이 있습니다.\n"
                f"총 용량: {size_mb:.2f} MB\n\n"
                f"모든 거래 데이터 캐시를 삭제하시겠습니까?\n"
                f"(다음 조회 시 API에서 새로 불러옵니다)",
                parent=self.root
            )

            if response:
                shutil.rmtree(self.trade_cache_path)
                os.makedirs(self.trade_cache_path, exist_ok=True)
                show_topmost_info("캐시 삭제", f"{file_count}개의 캐시 파일이 삭제되었습니다.", parent=self.root)
                print(f"🗑 거래 데이터 캐시 삭제 완료 ({file_count}개 파일, {size_mb:.2f} MB)")
        except Exception as e:
            show_topmost_error("오류", f"캐시 삭제 중 오류 발생: {str(e)}", parent=self.root)



    def collect_apt_data_background(self, apt_info, data_type):
        """단일 아파트의 거래 데이터를 백그라운드에서 수집 - 모든 월별 데이터 수집 (캐시 지원)"""
        try:
            apt_name = apt_info['apt_name']
            dong = apt_info['dong']
            sido = apt_info.get('sido', '')
            sigungu = apt_info.get('sigungu', '')
            sigungu_code = apt_info['sigungu_code']
            target_area = float(apt_info['area'])

            # 캐시에서 먼저 로드 시도
            cached_trades = self.load_trade_cache(sido, sigungu, dong, apt_name, target_area, data_type)

            # 3개월 기준 날짜 계산
            three_months_ago = datetime.now() - timedelta(days=90)

            if cached_trades is not None:
                # 캐시가 있는 경우: 3개월 이전 데이터는 캐시 사용, 최근 3개월은 API로 재조회
                print(f"📦 캐시 발견: {apt_name} ({target_area}㎡) - {data_type}")
                print(f"   → 3개월 이전 데이터는 캐시 사용, 최근 3개월 데이터는 API로 업데이트")

                # 3개월 이전 데이터만 필터링
                old_trades = [t for t in cached_trades if t['date'] < three_months_ago]
                print(f"   → 캐시에서 {len(old_trades)}건 로드")

                # 최근 3개월은 API로 조회 (부분 조회 모드)
                recent_months_only = True
            else:
                # 캐시가 없으면 전체 API 조회
                print(f"🔍 API 전체 조회: {apt_name} ({target_area}㎡) - {data_type}")
                old_trades = []
                recent_months_only = False

            # 데이터 유형에 따라 API 엔드포인트 설정
            if data_type == "purchase":
                api_endpoint = "http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
            else:  # jeonse
                api_endpoint = "http://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"

            # 데이터 수집 관련 설정
            trades = []
            current_date = datetime.now()
            
            # 병렬 처리 설정
            concurrent_requests = 36  # 동시 요청 수 증가 (성능 향상)
            
            # HTTP 세션 최적화
            session = requests.Session()
            adapter = requests.adapters.HTTPAdapter(
                pool_connections=concurrent_requests,
                pool_maxsize=concurrent_requests * 2,
                max_retries=1
            )
            session.mount('http://', adapter)
            session.mount('https://', adapter)
            
            # 모든 월별 데이터 수집을 위한 연도 범위 설정
            # 준공년도 확인하여 그 이후부터 데이터 수집
            build_year = 1990  # 기본값
            if 'build_year' in apt_info and apt_info['build_year']:
                try:
                    build_year = int(apt_info['build_year'])
                except:
                    pass  # 변환 실패시 기본값 사용
            
            # 모든 월별 요청 생성
            all_months = []

            if recent_months_only:
                # 최근 3개월만 조회
                for i in range(3):
                    search_date = current_date - timedelta(days=30 * i)
                    deal_ymd = search_date.strftime("%Y%m")
                    all_months.append((search_date, deal_ymd))
            else:
                # 전체 기간 조회 (준공년도 ~ 현재)
                # 현재 연도부터 준공연도까지 역순으로 처리 (최신 데이터 우선)
                for year in range(current_date.year, build_year - 1, -1):
                    # 현재 연도의 경우 현재 월까지만
                    max_month = 12 if year < current_date.year else current_date.month

                    for month in range(1, max_month + 1):
                        try:
                            search_date = datetime(year=year, month=month, day=1)
                            deal_ymd = search_date.strftime("%Y%m")
                            all_months.append((search_date, deal_ymd))
                        except ValueError:
                            continue
            
            # 필터링 함수 - 데이터 처리 최적화
            def process_item(item, apt_name, dong, target_area):
                # 빠른 필터링
                item_apt = item.findtext('aptNm', '').strip()
                if item_apt != apt_name:
                    return None
                
                item_dong = item.findtext('umdNm', '').strip()
                if item_dong != dong:
                    return None
                
                item_area = float(item.findtext('excluUseAr', '0'))
                if abs(item_area - target_area) > 1:
                    return None
                
                # 데이터 유형에 따라 처리
                if data_type == "purchase":
                    # 매매 데이터
                    try:
                        deal_amount = item.findtext('dealAmount', '0').replace(',', '')
                        trade = {
                            'date': datetime(
                                int(item.findtext('dealYear', '0')),
                                int(item.findtext('dealMonth', '0')),
                                int(item.findtext('dealDay', '1') or '1')
                            ),
                            'price': int(deal_amount),
                            'floor': int(item.findtext('floor', '0')),
                            'area': item_area
                        }
                        return trade
                    except (ValueError, TypeError):
                        return None
                else:
                    # 전세 데이터
                    try:
                        monthly_rent = item.findtext('monthlyRent', '0').replace(',', '')
                        is_monthly = monthly_rent and int(monthly_rent) > 0
                        
                        if not is_monthly:
                            deposit = item.findtext('deposit', '0').replace(',', '')
                            trade = {
                                'date': datetime(
                                    int(item.findtext('dealYear', '0')),
                                    int(item.findtext('dealMonth', '0')),
                                    int(item.findtext('dealDay', '1') or '1')
                                ),
                                'price': int(deposit),
                                'floor': int(item.findtext('floor', '0')),
                                'area': item_area,
                                'rent_type': '전세'
                            }
                            return trade
                        return None
                    except (ValueError, TypeError):
                        return None
            
            # 모든 월별 데이터를 배치 단위로 수집
            def collect_all_monthly_data():
                all_data = []
                batch_size = 24  # 한 번에 24개월 처리 (2년)
                batches = [all_months[i:i+batch_size] for i in range(0, len(all_months), batch_size)]
                
                # 연속 빈 배치 카운터 (연속으로 6개 배치가 비어있으면 중단)
                consecutive_empty_batches = 0
                max_consecutive_empty = 2  # 4년(48개월) 동안 데이터가 없으면 중단
                
                # 진행률 모니터링
                total_batches = len(batches)
                completed_batches = 0
                
                # 배치 단위로 처리
                for batch in batches:
                    completed_batches += 1
                    
                    # 연속으로 빈 배치가 max_consecutive_empty개 이상이면 중단
                    if consecutive_empty_batches >= max_consecutive_empty:
                        print(f"연속 {consecutive_empty_batches * batch_size}개월 동안 데이터가 없어 조기 종료")
                        break
                    
                    batch_data = []
                    
                    # 각 배치를 병렬 처리
                    with concurrent.futures.ThreadPoolExecutor(max_workers=concurrent_requests) as executor:
                        future_results = {}
                        
                        # 배치내 모든 월 요청 준비
                        for search_date, deal_ymd in batch:
                            url = (f"{api_endpoint}"
                                   f"?serviceKey={self.service_key}"
                                   f"&LAWD_CD={sigungu_code}"
                                   f"&DEAL_YMD={deal_ymd}"
                                   f"&numOfRows=5000")
                            
                            future = executor.submit(session.get, url, timeout=API_TIMEOUT)
                            future_results[future] = (search_date, deal_ymd)
                        
                        # 결과 수집
                        for future in concurrent.futures.as_completed(future_results.keys()):
                            search_date, deal_ymd = future_results[future]
                            
                            try:
                                response = future.result()
                                
                                if response.status_code == 200:
                                    try:
                                        root = ET.fromstring(response.text)
                                        items = root.findall('.//item')
                                        
                                        # 병렬 처리로 아이템 필터링
                                        monthly_trades = []
                                        for item in items:
                                            trade = process_item(item, apt_name, dong, target_area)
                                            if trade:
                                                monthly_trades.append(trade)
                                        
                                        if monthly_trades:
                                            batch_data.extend(monthly_trades)
                                    except ET.ParseError:
                                        pass
                                
                            except (requests.RequestException, requests.Timeout):
                                continue
                    
                    # 이번 배치에서 데이터가 있는지 확인
                    if batch_data:
                        all_data.extend(batch_data)
                        consecutive_empty_batches = 0  # 데이터가 있으면 카운터 리셋
                    else:
                        consecutive_empty_batches += 1  # 데이터가 없으면 카운터 증가
                    
                    # 진행 상태 로그
                    progress = (completed_batches / total_batches) * 100
                    year_range = f"{batch[-1][0].year}-{batch[0][0].year}"
                    print(f"진행: {progress:.1f}% 완료 - {year_range} 기간 {len(batch_data)}건 수집됨")
                
                return all_data
            
            # 모든 월별 데이터 수집 실행
            all_trades = collect_all_monthly_data()
            trades.extend(all_trades)

            # 캐시된 구 데이터와 병합
            if old_trades:
                print(f"   → API에서 {len(trades)}건 수집, 캐시 데이터 {len(old_trades)}건과 병합")
                trades.extend(old_trades)

            # 데이터 처리
            if trades:
                # 날짜순 정렬 및 중복 제거
                trades = sorted(trades, key=lambda x: x['date'])

                # 중복 제거 (같은 날짜, 가격, 층이면 중복으로 간주)
                unique_trades = []
                seen = set()
                for trade in trades:
                    key = (trade['date'], trade['price'], trade['floor'])
                    if key not in seen:
                        seen.add(key)
                        unique_trades.append(trade)
                trades = unique_trades

                print(f"   → 최종 {len(trades)}건 (중복 제거 완료)")

                # 전체 데이터를 캐시에 저장 (갱신)
                self.save_trade_cache(sido, sigungu, dong, apt_name, target_area, data_type, trades)

                # 데이터프레임 생성
                df = pd.DataFrame(trades)

                # 결과 반환
                return {'apt_info': apt_info, 'trades': trades, 'df': df}
            else:
                # 빈 데이터도 캐시에 저장 (불필요한 재조회 방지)
                self.save_trade_cache(sido, sigungu, dong, apt_name, target_area, data_type, [])
                return {'apt_info': apt_info, 'trades': []}

        except Exception as e:
            print(f"백그라운드 데이터 수집 중 오류: {str(e)}")
            return {'apt_info': apt_info, 'error': str(e), 'trades': []}
    
    def add_apt_to_selection(self, apt_info):
        """선택된 아파트 목록에 아파트 추가 및 자동 데이터 수집"""
        print(f"\n[디버그] ===== add_apt_to_selection 호출 =====")
        print(f"[디버그] 받은 apt_info: {apt_info}")
        print(f"[디버그] =========================================\n")

        # 최대 3개까지만 추가 가능
        if len(self.selected_apts) >= 3:
            show_topmost_info("알림", "최대 3개의 아파트만 선택할 수 있습니다.", parent=self.root)
            return False

        # 중복 확인
        for existing in self.selected_apts:
            if existing['apt_name'] == apt_info['apt_name'] and existing['area'] == apt_info['area']:
                show_topmost_info("알림", "이미 선택된 아파트입니다.", parent=self.root)
                return False

        # API 조회에 필요한 필수 정보 확인
        required_fields = ['apt_name', 'area', 'sido', 'sigungu', 'dong', 'sigungu_code']
        missing_fields = [field for field in required_fields if field not in apt_info]
        if missing_fields:
            print(f"[디버그] ❌ 누락된 필수 필드: {missing_fields}")
            show_topmost_error("오류", f"아파트 정보에 필수 필드가 없습니다: {', '.join(missing_fields)}", parent=self.root)
            return False

        print(f"[디버그] ✅ 모든 필수 필드 확인 완료")
        print(f"[디버그] sigungu_code: {apt_info['sigungu_code']}")
                
        # 목록에 추가
        self.selected_apts.append(apt_info)
        display_text = f"{apt_info['apt_name']} ({apt_info['area']}㎡) - {apt_info['sido']} {apt_info['sigungu']} {apt_info['dong']}"
        self.selected_apt_listbox.insert(tk.END, display_text)
        
        # 데이터 수집 상태 메시지 및 프로그레스 초기화
        self.status_label.config(text=f"🔄 {apt_info['apt_name']} ({apt_info['area']}㎡) 데이터 수집 준비 중...")
        self.update_progress(0, "")

        # 그래프 생성 버튼 비활성화 (수집 시작) - 회색으로 표시됨
        self.graph_button.config(state="disabled", text="📊 데이터 수집 중...")

        # tkinter 변수들을 미리 읽어두기 (스레드에서 직접 접근 방지)
        should_collect_jeonse = self.collect_jeonse_data.get()

        # 백그라운드에서 매매 및 전세 데이터 수집 시작
        def background_collection():
            # 필요한 정보 준비
            apt_info_copy = apt_info.copy()

            # 1단계: 매매 데이터 수집 (0-50%)
            self.root.after(0, lambda: self.update_progress(10, "📊 매매 데이터 수집 중..."))
            purchase_result = self.collect_apt_data_background(apt_info_copy, "purchase")

            purchase_count = 0
            if purchase_result and 'trades' in purchase_result and purchase_result['trades']:
                purchase_count = len(purchase_result['trades'])
                # 매매 데이터 저장
                for i, selected_apt in enumerate(self.selected_apts):
                    if selected_apt['apt_name'] == apt_info['apt_name'] and str(selected_apt['area']) == str(apt_info['area']):
                        self.selected_apts[i]['trades_data'] = purchase_result['trades']

                        # 엑셀 파일 저장
                        if 'df' in purchase_result and purchase_result['df'] is not None and not purchase_result['df'].empty:
                            try:
                                excel_path = self.save_apt_data_to_excel(purchase_result['df'], apt_info, "purchase")
                                self.selected_apts[i]['excel_path'] = excel_path
                            except Exception as e:
                                print(f"매매 데이터 엑셀 저장 중 오류: {str(e)}")
                        break

            self.root.after(0, lambda: self.update_progress(50, f"✅ 매매 {purchase_count}건 수집 완료"))

            # 2단계: 전세 데이터 수집 (50-100%) - 선택적
            jeonse_count = 0
            if should_collect_jeonse:
                self.root.after(0, lambda: self.update_progress(60, "🏠 전세 데이터 수집 중..."))
                try:
                    jeonse_result = self.collect_apt_data_background(apt_info_copy, "jeonse")

                    if jeonse_result and 'trades' in jeonse_result and jeonse_result['trades']:
                        jeonse_count = len(jeonse_result['trades'])
                        # 전세 데이터 저장
                        for i, selected_apt in enumerate(self.selected_apts):
                            if selected_apt['apt_name'] == apt_info['apt_name'] and str(selected_apt['area']) == str(apt_info['area']):
                                self.selected_apts[i]['jeonse_data'] = jeonse_result['trades']

                                # 엑셀 파일 저장
                                if 'df' in jeonse_result and jeonse_result['df'] is not None and not jeonse_result['df'].empty:
                                    try:
                                        excel_path = self.save_apt_data_to_excel(jeonse_result['df'], apt_info, "jeonse")
                                        self.selected_apts[i]['jeonse_excel_path'] = excel_path
                                    except Exception as e:
                                        print(f"전세 데이터 엑셀 저장 중 오류: {str(e)}")
                                break
                    elif jeonse_result and 'error' in jeonse_result:
                        # 전세 데이터 조회 실패 시 에러 메시지 출력하고 계속 진행
                        error_msg = jeonse_result['error']
                        print(f"⚠ 전세 데이터 조회 실패: {error_msg}")
                        self.root.after(0, lambda: self.update_progress(50, f"⚠ 전세 데이터 조회 실패 (매매 데이터만 사용)"))
                except Exception as e:
                    # 예외 발생 시에도 계속 진행
                    print(f"⚠ 전세 데이터 수집 중 예외 발생: {str(e)}")
                    self.root.after(0, lambda: self.update_progress(50, f"⚠ 전세 데이터 수집 실패 (매매 데이터만 사용)"))
            else:
                print("전세 데이터 수집 건너뜀 (옵션 비활성화)")
                self.root.after(0, lambda: self.update_progress(50, "전세 데이터 수집 건너뜀"))

            self.safe_after(0, lambda: self.update_progress(100, f"✅ 전세 {jeonse_count}건 수집 완료"))

            # 수집 완료 후 UI 업데이트
            update_message = f"✅ {apt_info['apt_name']} ({apt_info['area']}㎡) 수집 완료 - 매매: {purchase_count}건 / 전세: {jeonse_count}건"
            self.safe_after(100, lambda: self.status_label.config(text=update_message))

            # 데이터 수집 완료 확인 - 매매 또는 전세 데이터가 있으면 그래프 버튼 활성화
            has_data = any(('trades_data' in apt and apt.get('trades_data')) or
                          ('jeonse_data' in apt and apt.get('jeonse_data'))
                          for apt in self.selected_apts)

            # 완료 메시지와 함께 그래프 버튼 활성화 안내
            if has_data:
                final_message = f"🎉 데이터 수집 완료! 이제 '📊 그래프 생성하기' 버튼을 눌러주세요."
                self.safe_after(200, lambda: self.status_label.config(text=final_message))
                self.safe_after(200, lambda: self.graph_button.config(state="normal", text="📊 그래프 생성하기"))
                # 3초 후 프로그레스바 초기화
                self.safe_after(3000, lambda: self.update_progress(0, final_message))
            else:
                # 데이터가 없는 경우에도 그래프 버튼 활성화 (사용자가 다시 시도할 수 있도록)
                warning_message = "⚠️ 수집된 데이터가 없습니다. 다른 아파트를 선택해주세요."
                self.safe_after(200, lambda: self.status_label.config(text=warning_message))
                self.safe_after(200, lambda: self.graph_button.config(state="normal", text="📊 그래프 생성하기"))
                self.safe_after(200, lambda: self.update_progress(0, ""))
        
        # 백그라운드 스레드 시작
        bg_thread = threading.Thread(target=background_collection, daemon=True)
        bg_thread.start()
        
        return True
    
    def update_progress(self, value, message=""):
        """프로그레스 바 업데이트"""
        self.progress["value"] = value
        if message:
            self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def on_sido_selected(self, event):
        """시/도 선택 시 처리 (간결한 시군구 표시)"""
        sido = self.sido_combobox.get()
        if sido in self.sigungu_dict:
            self.sigungu_combobox['values'] = sorted(self.sigungu_dict[sido])
            self.sigungu_combobox.set("시/군/구 선택")
            self.dong_combobox.set("읍/면/동 선택")
    
    def on_sigungu_selected(self, event):
        """시/군/구 선택 시 처리 (구 아래 동 계층 구조 표시)"""
        sigungu = self.sigungu_combobox.get()

        if sigungu in self.dong_dict:
            dong_list = self.dong_dict[sigungu]
            expanded_list = []

            for item in sorted(dong_list):
                expanded_list.append(item)

                # 구인 경우 하위 동 목록 추가
                if item.endswith('구'):
                    gu_key = f"{sigungu}_{item}"
                    if gu_key in self.dong_dict:
                        sub_dongs = sorted(self.dong_dict[gu_key])
                        for sub_dong in sub_dongs:
                            expanded_list.append(f"  └ {sub_dong}")

            self.dong_combobox['values'] = expanded_list
            self.dong_combobox.set("읍/면/동 선택")
        else:
            self.dong_combobox['values'] = []
            self.dong_combobox.set("읍/면/동 선택")
    
    def on_dong_selected(self, event):
        """읍/면/동 선택 시 처리 - 메인 UI 프로그레스바 사용"""
        sido = self.sido_combobox.get()
        sigungu = self.sigungu_combobox.get()
        dong = self.dong_combobox.get()

        if "선택" not in [sido, sigungu, dong]:
            # 중복 이름 시군구 처리 (강서구(서), 강서구(부) 등)
            if sigungu in self.special_sigungu_names:
                # 특별 시군구에서 원래 시도명과 시군구명 가져오기
                original_sido, original_sigungu = self.special_sigungu_names[sigungu]
                region_code = self.region_codes.get((original_sido, sigungu, dong))
            else:
                region_code = self.region_codes.get((sido, sigungu, dong))

            if region_code:
                _, sigungu_code = region_code

                # 프로그레스바 시작
                self.update_progress(5, f"🔄 {dong} 아파트 목록을 불러오는 중...")

                # 진행 콜백 함수
                def progress_callback(progress, msg):
                    self.update_progress(5 + int(progress * 0.9), msg)

                try:
                    apt_list = self.get_apt_list_from_api(sigungu_code, dong, progress_callback=progress_callback)

                    if apt_list:
                        self.update_progress(100, f"✅ {dong}의 아파트 {len(apt_list)}개 검색 완료")
                        self.status_label.config(text=f"{dong}의 아파트 {len(apt_list)}개가 검색되었습니다.")
                    else:
                        self.update_progress(0, "")
                        self.status_label.config(text=f"{dong}의 검색 가능한 아파트가 없습니다.")
                except Exception as e:
                    self.update_progress(0, "")
                    self.status_label.config(text=f"오류 발생: {str(e)}")

                # 잠시 후 프로그레스바 초기화
                self.root.after(2000, lambda: self.update_progress(0, ""))
            else:
                self.status_label.config(text=f"지역 코드를 찾을 수 없습니다: {sido} {sigungu} {dong}")
                        
    def show_settings_dialog(self):
        """설정 대화상자 표시 - 캐시 경로 추가"""
        settings = tk.Toplevel(self.root)
        settings.title("설정")
        settings.geometry("550x450")  # 창 크기 더 크게 조정 (캐시 설정 추가)
        settings.resizable(False, False)
        settings.transient(self.root)
        settings.grab_set()
        
        # 다운로드 경로 설정
        ttk.Label(settings, text="다운로드 경로:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        download_path_var = tk.StringVar(value=self.download_path)
        download_entry = ttk.Entry(settings, textvariable=download_path_var, width=40)
        download_entry.grid(row=0, column=1, padx=5, pady=10)
        
        def select_download_path():
            path = filedialog.askdirectory(initialdir=self.download_path)
            if path:
                download_path_var.set(path)
        
        ttk.Button(settings, text="찾아보기", command=select_download_path).grid(row=0, column=2, padx=5, pady=10)
        
        # 히스토리 경로 설정
        ttk.Label(settings, text="히스토리 경로:").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        history_path_var = tk.StringVar(value=self.history_path)
        history_entry = ttk.Entry(settings, textvariable=history_path_var, width=40)
        history_entry.grid(row=1, column=1, padx=5, pady=10)
        
        def select_history_path():
            path = filedialog.askdirectory(initialdir=self.history_path)
            if path:
                history_path_var.set(path)
        
        ttk.Button(settings, text="찾아보기", command=select_history_path).grid(row=1, column=2, padx=5, pady=10)
        
        # 법정동 코드 파일 경로 설정
        ttk.Label(settings, text="법정동 파일 경로:").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        lawdong_path_var = tk.StringVar(value=self.lawdong_path)
        lawdong_entry = ttk.Entry(settings, textvariable=lawdong_path_var, width=40)
        lawdong_entry.grid(row=2, column=1, padx=5, pady=10)
        
        def select_lawdong_path():
            path = filedialog.askopenfilename(
                initialdir=os.path.dirname(self.lawdong_path),
                title="법정동 코드 파일 선택",
                filetypes=[("Text files", "*.txt")]
            )
            if path:
                lawdong_path_var.set(path)
        
        ttk.Button(settings, text="찾아보기", command=select_lawdong_path).grid(row=2, column=2, padx=5, pady=10)
        
        # 단지정보 파일 경로 설정 (새로 추가)
        ttk.Label(settings, text="단지정보 파일 경로:").grid(row=3, column=0, sticky="w", padx=10, pady=10)
        
        # 클래스에 단지정보 경로 변수가 없다면 기본값으로 초기화
        if not hasattr(self, 'complex_info_path'):
            self.complex_info_path = os.path.join(os.getcwd(), "data", "complex_info.xlsx")
        
        complex_info_path_var = tk.StringVar(value=self.complex_info_path)
        complex_info_entry = ttk.Entry(settings, textvariable=complex_info_path_var, width=40)
        complex_info_entry.grid(row=3, column=1, padx=5, pady=10)
        
        def select_complex_info_path():
            path = filedialog.askopenfilename(
                initialdir=os.path.dirname(self.complex_info_path) if hasattr(self, 'complex_info_path') else os.getcwd(),
                title="단지정보 파일 선택",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            if path:
                complex_info_path_var.set(path)
        
        ttk.Button(settings, text="찾아보기", command=select_complex_info_path).grid(row=3, column=2, padx=5, pady=10)

        # 거래 데이터 캐시 경로 설정 (새로 추가)
        ttk.Label(settings, text="거래 데이터 캐시 경로:").grid(row=4, column=0, sticky="w", padx=10, pady=10)
        trade_cache_path_var = tk.StringVar(value=self.trade_cache_path)
        trade_cache_entry = ttk.Entry(settings, textvariable=trade_cache_path_var, width=40)
        trade_cache_entry.grid(row=4, column=1, padx=5, pady=10)

        def select_trade_cache_path():
            path = filedialog.askdirectory(initialdir=self.trade_cache_path)
            if path:
                trade_cache_path_var.set(path)

        ttk.Button(settings, text="찾아보기", command=select_trade_cache_path).grid(row=4, column=2, padx=5, pady=10)

        # 캐시 관리 버튼 프레임
        cache_frame = ttk.LabelFrame(settings, text="📦 캐시 관리", padding=10)
        cache_frame.grid(row=5, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        # 캐시 정보 표시
        cache_info_label = ttk.Label(cache_frame, text="거래 데이터 캐시를 관리합니다.")
        cache_info_label.pack(pady=5)

        # 캐시 삭제 버튼
        def delete_trade_cache():
            self.clear_trade_cache()

        ttk.Button(cache_frame, text="🗑 거래 데이터 캐시 삭제", command=delete_trade_cache).pack(side="left", padx=5)

        # 캐시 폴더 열기 버튼
        def open_cache_folder():
            if os.path.exists(self.trade_cache_path):
                os.startfile(self.trade_cache_path)
            else:
                show_topmost_info("알림", "캐시 폴더가 존재하지 않습니다.", parent=settings)

        ttk.Button(cache_frame, text="📁 캐시 폴더 열기", command=open_cache_folder).pack(side="left", padx=5)

        # 설정 정보 프레임
        info_frame = ttk.Frame(settings, padding=5)
        info_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=10, pady=5)

        # 설정 정보 텍스트
        info_text = "※ 설정은 이 프로그램에서만 적용되며, 다른 프로그램에 영향을 주지 않습니다."
        ttk.Label(info_frame, text=info_text, foreground="blue").pack(side="left")
    
        # 저장 버튼
        def save_settings():
            # 다운로드 경로 생성
            new_download_path = download_path_var.get()
            if new_download_path:
                os.makedirs(new_download_path, exist_ok=True)
                self.download_path = new_download_path
                
            # 히스토리 경로 설정 및 생성
            new_history_path = history_path_var.get()
            if new_history_path:
                os.makedirs(new_history_path, exist_ok=True)
                self.history_path = new_history_path
            else:
                # 히스토리 경로가 지정되지 않은 경우 다운로드 경로 내에 history 폴더 생성
                self.history_path = os.path.join(self.download_path, "history")
                os.makedirs(self.history_path, exist_ok=True)
                
            # 법정동 경로 설정
            new_lawdong_path = lawdong_path_var.get()
            if new_lawdong_path and os.path.exists(new_lawdong_path):
                self.lawdong_path = new_lawdong_path
            
            # 단지정보 경로 설정 (새로 추가)
            new_complex_info_path = complex_info_path_var.get()
            if new_complex_info_path and os.path.exists(new_complex_info_path):
                self.complex_info_path = new_complex_info_path

            # 거래 데이터 캐시 경로 설정 및 생성
            new_trade_cache_path = trade_cache_path_var.get()
            if new_trade_cache_path:
                os.makedirs(new_trade_cache_path, exist_ok=True)
                self.trade_cache_path = new_trade_cache_path

            # 설정 저장 - 단지정보 경로 포함
            # 설정 저장 - 세부정보 옵션 포함
            settings_data = {
                'download_path': self.download_path,
                'history_path': self.history_path,
                'lawdong_path': self.lawdong_path,
                'complex_info_path': self.complex_info_path,
                'trade_cache_path': self.trade_cache_path,
                'graph_options': {
                    'show_monthly_avg': self.show_monthly_avg.get(),
                    'show_monthly_max': self.show_monthly_max.get(),
                    'show_scatter_plot': self.show_scatter_plot.get(),
                    'show_jeonse': self.show_jeonse.get(),
                    'show_jeonse_monthly_avg': self.show_jeonse_monthly_avg.get(),
                    'show_jeonse_monthly_max': self.show_jeonse_monthly_max.get(),
                    'show_jeonse_scatter_plot': self.show_jeonse_scatter_plot.get(),
                    # 세부정보 옵션 추가
                    'show_complex_info': self.show_complex_info.get(),
                    # 전세 데이터 수집 여부 옵션 추가
                    'collect_jeonse_data': self.collect_jeonse_data.get()
                }
            }

            settings_file = os.path.join(os.getcwd(), 'real_estate_analyzer_settings.json')
            
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)
                
            # 법정동 파일 다시 로드
            self.load_lawdong_file()
            
            # 단지정보 파일 로드 (필요한 경우)
            if hasattr(self, 'load_complex_info'):
                self.load_complex_info()
            
            # 히스토리 리로드
            self.history_list = self.load_history()
            self.update_history_display()
            
            show_topmost_info("알림", "설정이 저장되었습니다.", parent=self.root)
            settings.destroy()
        
        button_frame = ttk.Frame(settings, padding=5)
        button_frame.grid(row=5, column=0, columnspan=3, sticky="e", padx=10, pady=10)
        
        ttk.Button(button_frame, text="저장", command=save_settings).pack(side="right", padx=5)
        ttk.Button(button_frame, text="취소", command=settings.destroy).pack(side="right", padx=5)
    
    def open_download_folder(self):
        """다운로드 폴더 열기"""
        try:
            import platform
            import subprocess
            
            if platform.system() == "Windows":
                os.startfile(self.download_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", self.download_path])
            else:  # Linux
                subprocess.Popen(["xdg-open", self.download_path])
        except Exception as e:
            show_topmost_error("오류", f"폴더를 열 수 없습니다: {str(e)}", parent=self.root)
    
    def delete_selected_history(self):
        """선택된 히스토리 항목 삭제"""
        selection = self.history_tree.selection()
        if not selection:
            show_topmost_info("알림", "삭제할 항목을 선택해주세요.", parent=self.root)
            return
        
        if ask_topmost_yesno("확인", "선택한 항목을 삭제하시겠습니까?", parent=self.root):
            for item_id in selection:
                idx = self.history_tree.index(item_id)
                file_path = self.history_list[idx]['file_path']
                
                # 파일 삭제
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        show_topmost_error("오류", f"파일 삭제 중 오류 발생: {str(e)}", parent=self.root)
                        continue
                
                # 이미지 파일 삭제
                try:
                    apt_name = self.history_list[idx]['apt_name']
                    area = self.history_list[idx]['area']
                    
                    apt_name_clean = ''.join(char for char in apt_name if char.isalnum() or char.isspace())
                    apt_name_clean = apt_name_clean.replace(' ', '_')
                    
                    img_filename = f"{apt_name_clean}_{area}m2.jpg"
                    img_path = os.path.join(self.download_path, img_filename)
                    
                    if os.path.exists(img_path):
                        os.remove(img_path)
                except:
                    pass
            
            # 히스토리 리스트 갱신
            self.history_list = self.load_history()
            self.update_history_display()
            
            show_topmost_info("알림", "선택한 항목이 삭제되었습니다.", parent=self.root)
    
    def delete_all_history(self):
        """모든 히스토리 항목 삭제"""
        if not self.history_list:
            show_topmost_info("알림", "삭제할 히스토리가 없습니다.", parent=self.root)
            return
        
        if ask_topmost_yesno("확인", "모든 히스토리를 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다.", parent=self.root):
            # 히스토리 폴더 내 모든 파일 삭제
            try:
                for file in os.listdir(self.history_path):
                    file_path = os.path.join(self.history_path, file)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                
                # 히스토리 리스트 초기화
                self.history_list = []
                self.update_history_display()
                
                show_topmost_info("알림", "모든 히스토리가 삭제되었습니다.", parent=self.root)
            except Exception as e:
                show_topmost_error("오류", f"히스토리 삭제 중 오류 발생: {str(e)}", parent=self.root)

    
    def find_and_show_history_image(self, history_item):
        """히스토리 항목에 대응하는 이미지 파일을 찾아서 표시"""
        try:
            # 가능한 이미지 파일명 패턴들 생성
            possible_filenames = []
            
            apt_name = history_item.get('apt_name', '')
            area_info = history_item.get('area', '')
            
            if apt_name:
                # 아파트명 정리
                apt_name_clean = ''.join(char for char in apt_name if char.isalnum() or char.isspace())
                apt_name_clean = apt_name_clean.replace(' ', '_')
                
                # 면적 정보에서 숫자만 추출
                if '(' in area_info and ')' in area_info:
                    # "아파트명 (84㎡)" 형태에서 84 추출
                    area_match = re.search(r'\((\d+)㎡\)', area_info)
                    if area_match:
                        area_num = area_match.group(1)
                        possible_filenames.append(f"{apt_name_clean}_{area_num}m2.jpg")
                        possible_filenames.append(f"{apt_name_clean}_{area_num}m2.png")
                
                # vs가 포함된 경우 (다중 아파트 비교)
                if ' vs ' in apt_name:
                    # 다중 아파트 비교 파일명 패턴
                    multi_name = apt_name.replace(' vs ', '_vs_').replace(' ', '_').replace('(', '_').replace(')', '_').replace('㎡', 'm2')
                    possible_filenames.append(f"multi_apt_comparison_{multi_name}.jpg")
                    possible_filenames.append(f"multi_apt_comparison_{multi_name}.png")
            
            # 다운로드 폴더에서 파일 검색
            print(f"[디버그] 검색할 파일명들: {possible_filenames}")
            
            for filename in possible_filenames:
                image_path = os.path.join(self.download_path, filename)
                print(f"[디버그] 이미지 파일 검색: {image_path}")
                
                if os.path.exists(image_path):
                    print(f"[디버그] 이미지 파일 발견: {image_path}")
                    # 히스토리 항목에 이미지 경로 업데이트
                    history_item['image_path'] = image_path
                    self.show_graph_popup(history_item)
                    return True
            
            # 패턴 매칭으로도 찾지 못한 경우, 다운로드 폴더의 모든 이미지 파일을 검사
            print(f"[디버그] 패턴 매칭 실패, 전체 이미지 파일 검색")
            
            if os.path.exists(self.download_path):
                for file in os.listdir(self.download_path):
                    if file.lower().endswith(('.jpg', '.jpeg', '.png')) and file != "graph.jpg":
                        # 파일명에 아파트명이 포함되어 있는지 확인
                        if apt_name and any(word in file for word in apt_name.split() if len(word) > 1):
                            image_path = os.path.join(self.download_path, file)
                            print(f"[디버그] 유사한 이미지 파일 발견: {image_path}")
                            
                            # 사용자에게 확인
                            if ask_topmost_yesno("이미지 파일 확인", f"다음 이미지 파일이 해당 히스토리와 관련된 것 같습니다.\n\n{file}\n\n이 파일을 표시하시겠습니까?", parent=self.root):
                                history_item['image_path'] = image_path
                                self.show_graph_popup(history_item)
                                return True
            
            return False
            
        except Exception as e:
            print(f"[디버그] 이미지 파일 검색 중 오류: {str(e)}")
            return False    
                
    
    # 3. 히스토리 선택 시 처리 함수 수정
    def on_history_select(self, event):
        """히스토리 항목 선택 시 처리 - 그래프 이미지 표시 (수정된 버전)"""
        selection = self.history_tree.selection()
        if not selection:
            return
        
        try:
            idx = self.history_tree.index(selection[0])
            if idx >= len(self.history_list):
                show_topmost_info("알림", "히스토리 항목을 찾을 수 없습니다.", parent=self.root)
                return
                
            history_item = self.history_list[idx]
            print(f"[디버그] 히스토리 선택: {history_item}")
            
            # 그래프 이미지 표시
            if 'image_path' in history_item and os.path.exists(history_item['image_path']):
                print(f"[디버그] 이미지 파일 존재: {history_item['image_path']}")
                self.show_graph_popup(history_item)
            else:
                # 이미지 파일이 없는 경우, 다운로드 폴더에서 찾아보기
                print(f"[디버그] 이미지 파일 없음, 다운로드 폴더에서 검색")
                image_found = self.find_and_show_history_image(history_item)
                
                if not image_found:
                    # 엑셀 파일이라도 있으면 열기
                    if 'excel_path' in history_item and history_item['excel_path'] and os.path.exists(history_item['excel_path']):
                        print(f"[디버그] 엑셀 파일 열기: {history_item['excel_path']}")
                        os.startfile(history_item['excel_path'])
                    else:
                        show_topmost_info("알림", "그래프 이미지 및 엑셀 파일을 찾을 수 없습니다.", parent=self.root)
        except Exception as e:
            print(f"[디버그] 히스토리 선택 중 오류: {str(e)}")
            show_topmost_error("오류", f"히스토리 항목 처리 중 오류가 발생했습니다: {str(e)}", parent=self.root)
    # 5. 그래프 이미지 저장 함수
    def save_graph_image(self, image_path):
        """그래프 이미지를 다른 이름으로 저장"""
        try:
            # 원본 파일명 추출
            file_name = os.path.basename(image_path)
            
            # 저장 대화상자 표시
            save_path = filedialog.asksaveasfilename(
                initialdir=os.path.dirname(image_path),
                initialfile=file_name,
                defaultextension=".jpg",
                filetypes=[("JPEG 이미지", "*.jpg"), ("PNG 이미지", "*.png"), ("모든 파일", "*.*")]
            )
            
            if save_path:
                # 원본 이미지 복사
                shutil.copy2(image_path, save_path)
                show_topmost_info("알림", "이미지가 저장되었습니다.", parent=self.root)
        except Exception as e:
            show_topmost_error("오류", f"이미지 저장 중 오류 발생: {str(e)}", parent=self.root)
    
    # 6. 히스토리 트리뷰 컬럼 설정 변경 (setup_gui 함수 내 히스토리 트리뷰 설정 부분)
    # 아래 코드를 setup_gui 함수 내 히스토리 트리뷰 설정 부분에 적용
    """
    # 히스토리 트리뷰
    self.history_tree = ttk.Treeview(history_frame, 
                                  columns=("date", "chart_title", "apt", "area", "max_trade"),
                                  show="headings",
                                  height=24)
    
    # 컬럼 헤더 설정
    self.history_tree.heading("date", text="검색일시")
    self.history_tree.heading("chart_title", text="차트 제목")
    self.history_tree.heading("apt", text="아파트")
    self.history_tree.heading("area", text="면적")
    self.history_tree.heading("max_trade", text="최근 최고거래가")
    
    # 컬럼 너비 조정
    self.history_tree.column("date", width=100)
    self.history_tree.column("chart_title", width=120)
    self.history_tree.column("apt", width=120)
    self.history_tree.column("area", width=50)
    self.history_tree.column("max_trade", width=100)
    """
    # 4. 그래프 팝업 표시 함수
    def show_graph_popup(self, history_item):
        """그래프 이미지를 팝업 창으로 표시 (개선된 버전)"""
        try:
            image_path = history_item['image_path']
            
            if not os.path.exists(image_path):
                show_topmost_error("오류", f"이미지 파일을 찾을 수 없습니다:\n{image_path}", parent=self.root)
                return
            
            print(f"[디버그] 이미지 팝업 표시: {image_path}")
            
            # 팝업 창 생성
            popup = tk.Toplevel(self.root)
            
            # 제목 설정 (차트 제목과 아파트 정보)
            if history_item.get('type') == 'multi':
                popup.title(f"다중 아파트 비교 분석: {history_item.get('apt_name', '')}")
            else:
                apt_name = history_item.get('apt_name', '')
                area = history_item.get('area', '')
                popup.title(f"아파트 실거래가 분석: {apt_name} {area}")
            
            # 창 크기 설정
            width = 1000
            height = 700
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            popup.geometry(f"{width}x{height}+{x}+{y}")
            
            # 이미지 로드 및 크기 조정
            try:
                img = Image.open(image_path)
                img_width, img_height = img.size
                
                # 화면 크기에 맞게 스케일링
                scale_factor = min(width / img_width, height / img_height * 0.9)  # 버튼 영역 고려하여 90%로 제한
                new_width = int(img_width * scale_factor)
                new_height = int(img_height * scale_factor)
                
                img = img.resize((new_width, new_height), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
            except Exception as e:
                show_topmost_error("오류", f"이미지 로드 중 오류가 발생했습니다:\n{str(e)}", parent=self.root)
                popup.destroy()
                return
            
            # 메인 프레임
            main_frame = ttk.Frame(popup)
            main_frame.pack(fill="both", expand=True)
            
            # 캔버스에 이미지 표시 (스크롤바 추가)
            canvas_frame = ttk.Frame(main_frame)
            canvas_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            canvas = tk.Canvas(canvas_frame, width=new_width, height=new_height)
            canvas.pack(side="left", fill="both", expand=True)
            
            # 스크롤바 추가
            v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
            h_scrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
            
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            v_scrollbar.pack(side="right", fill="y")
            h_scrollbar.pack(side="bottom", fill="x")
            
            # 이미지 표시
            canvas.create_image(0, 0, anchor="nw", image=photo)
            canvas.image = photo  # 참조를 유지해야 이미지가 표시됨
            
            # 캔버스 스크롤 영역 설정
            canvas.configure(scrollregion=canvas.bbox("all"))
            
            # 버튼 프레임
            btn_frame = ttk.Frame(popup, padding=10)
            btn_frame.pack(side="bottom", fill="x")
            
            # 엑셀 버튼 - 엑셀 파일 열기
            if 'excel_path' in history_item and history_item['excel_path'] and os.path.exists(history_item['excel_path']):
                excel_btn = ttk.Button(btn_frame, text="엑셀 파일 열기", 
                                    command=lambda: os.startfile(history_item['excel_path']))
                excel_btn.pack(side="left", padx=5, pady=5)
            
            # 저장 버튼 - 다른 이름으로 이미지 저장
            save_btn = ttk.Button(btn_frame, text="이미지 저장", 
                                 command=lambda: self.save_graph_image(image_path))
            save_btn.pack(side="left", padx=5, pady=5)
            
            # 닫기 버튼
            close_btn = ttk.Button(btn_frame, text="닫기", command=popup.destroy)
            close_btn.pack(side="right", padx=5, pady=5)
            
            # 팝업 창을 최상위로
            popup.lift()
            popup.focus_force()
            
        except Exception as e:
            print(f"[디버그] 그래프 팝업 표시 중 오류: {str(e)}")
            show_topmost_error("오류", f"그래프 표시 중 오류가 발생했습니다:\n{str(e)}", parent=self.root)

         

    def open_history_excel(self, file_path):
        """히스토리 엑셀 파일 열기"""
        try:
            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                show_topmost_info("알림", "엑셀 파일을 찾을 수 없습니다.", parent=self.root)
        except Exception as e:
            show_topmost_error("오류", f"엑셀 파일을 열 수 없습니다: {str(e)}", parent=self.root)
    

        
    # 1. 먼저 get_apt_list_from_api 함수를 수정하여 준공연도 정보 가져오기 (신고가 프로그램 방식으로 개선)
    def get_apt_list_from_api(self, sigungu_code, dong, progress_callback=None):
        """국토부 API에서 아파트 목록 가져오기 (기축 + 신축 통합, 신고가 프로그램 방식)"""
        print(f"\n=== get_apt_list_from_api 호출 ===")
        print(f"시군구 코드: {sigungu_code}")
        print(f"동: {dong}")

        apt_info = {}
        current_date = datetime.now()

        # 최근 3개월만 검색
        total_steps = 3 * 2  # 3개월 × 2가지 API (기축, 신축)
        current_step = 0

        for i in range(3):
            search_date = current_date - timedelta(days=30*i)
            deal_ymd = search_date.strftime("%Y%m")

            # 기축 아파트 API
            url_existing = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                           f"?serviceKey={self.service_key}"
                           f"&LAWD_CD={sigungu_code}"
                           f"&DEAL_YMD={deal_ymd}"
                           f"&numOfRows=1000")

            # 신축 아파트(분양권) API
            url_new = (f"http://apis.data.go.kr/1613000/RTMSDataSvcSilvTrade/getRTMSDataSvcSilvTrade"
                       f"?serviceKey={self.service_key}"
                       f"&LAWD_CD={sigungu_code}"
                       f"&DEAL_YMD={deal_ymd}"
                       f"&numOfRows=1000")

            print(f"\n월 {deal_ymd} API 호출:")
            print(f"URL: {url_existing}")

            for url, apt_type in [(url_existing, "기축"), (url_new, "신축")]:
                current_step += 1
                if progress_callback:
                    progress = int((current_step / total_steps) * 100)
                    progress_callback(progress, f"📡 {dong} 아파트 목록 조회 중... ({deal_ymd[:4]}.{deal_ymd[4:]} {apt_type})")
                try:
                    response = requests.get(url, timeout=API_TIMEOUT)
                    time.sleep(0.2)  # API 호출 간격 조절
                    print(f"{apt_type} API 응답 상태: {response.status_code}")

                    if response.status_code == 200:
                        root = ET.fromstring(response.text)
                        items = root.findall('.//item')
                        print(f"{apt_type} 조회된 전체 항목 수: {len(items)}")

                        dong_count = 0
                        for item in items:
                            item_dong = item.findtext('umdNm', '').strip()
                            if item_dong == dong:
                                dong_count += 1
                                apt_name = item.findtext('aptNm', '').strip()

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

        # 결과 리스트 생성
        apt_list = []
        new_apts = []

        # 신축 아파트 먼저 추가
        for apt_name, info in sorted(apt_info.items()):
            if info['type'] == '신축':
                if info['build_year'] and info['build_year'] != "분양":
                    apt_str = f"[신축] {apt_name} [{info['road_addr']} / {info['jibun_addr']}] (준공: {info['build_year']}년)"
                else:
                    apt_str = f"[신축] {apt_name} [{info['road_addr']} / {info['jibun_addr']}] (분양중)"
                new_apts.append(apt_str)

        # 기축 아파트 - 준공연도 있는 것
        existing_apts = [f"{apt_name} [{info['road_addr']} / {info['jibun_addr']}] (준공: {info['build_year']}년)"
                        for apt_name, info in sorted(apt_info.items())
                        if info['type'] == '기축' and info['build_year']]

        # 기축 아파트 - 준공연도 없는 것
        no_year_apts = [f"{apt_name} [{info['road_addr']} / {info['jibun_addr']}]"
                       for apt_name, info in sorted(apt_info.items())
                       if info['type'] == '기축' and not info['build_year']]

        logging.info(f"검색 결과 - 신축: {len(new_apts)}개, 기축: {len(existing_apts)}개")
        print(f"최종 결과 - 신축: {len(new_apts)}개, 기축: {len(existing_apts)}개, 연도없음: {len(no_year_apts)}개")

        apt_list_result = new_apts + existing_apts + no_year_apts

        # 캐시에 저장
        cache_key = (sigungu_code, dong)
        self.apt_list_cache[cache_key] = apt_list_result
        print(f"✅ 캐시에 저장됨: {cache_key} ({len(apt_list_result)}개)")

        return apt_list_result

        
    # 2. 반응형 UI 개선 (사용자 피드백 추가)
    def create_graph_from_selected(self):
        """선택된 아파트 정보로 그래프 생성 (데이터 저장과 그래프 생성 분리)"""
        if not self.selected_apts:
            show_topmost_error("오류", "선택된 아파트가 없습니다.", parent=self.root)
            return
            
        try:
            # 그래프 생성 버튼 비활성화
            self.graph_button.config(state="disabled")
            
            # 진행 상태 표시
            self.update_progress(5, "분석 준비 중...")
            
            # 데이터 수집 진행 관리
            total_apts = len(self.selected_apts)
            progress_per_apt = 60 / total_apts  # 데이터 수집 60%, 저장 20%, 그래프 생성 20%
            
            # 각 아파트별 거래 데이터 확인 (이미 수집된 데이터 사용)
            for i, apt_info in enumerate(self.selected_apts):
                apt_name = apt_info['apt_name']
                area = apt_info['area']

                self.update_progress(5 + (i * progress_per_apt),
                                    f"{apt_name} ({area}㎡) 데이터 확인 중...")

                # 이미 수집된 데이터가 있는지 확인
                if 'trades_data' not in apt_info or not apt_info.get('trades_data'):
                    # 데이터가 없으면 새로 수집
                    self.update_progress(5 + (i * progress_per_apt),
                                        f"{apt_name} ({area}㎡) 매매 데이터 수집 중...")
                    trades = self.get_trade_data(apt_info)
                    self.selected_apts[i]['trades_data'] = trades
                else:
                    trades = apt_info['trades_data']

                self.update_progress(5 + ((i+1) * progress_per_apt),
                                    f"{apt_name} ({area}㎡) 매매 데이터 확인 완료 ({len(trades)}건)")

                # 거래 데이터가 없는 경우 경고
                if not trades:
                    print(f"⚠️ {apt_name} ({area}㎡)의 매매 거래 데이터가 없습니다.")
            
            # 1. 먼저 엑셀 파일로 데이터 저장
            self.update_progress(70, "데이터 파일 저장 중...")

            # 처리할 아파트 목록과 데이터프레임 생성 (매매 + 전세)
            apt_dfs = []
            jeonse_dfs = []

            for apt_info in self.selected_apts:
                # 매매 데이터 처리
                trades = apt_info.get('trades_data', [])
                if trades:
                    # 데이터 정렬
                    trades = sorted(trades, key=lambda x: x['date'])
                    # 데이터프레임 생성
                    df = pd.DataFrame(trades)
                    # 엑셀 파일 저장
                    try:
                        self.save_apt_data_to_excel(df, apt_info, "purchase")
                        apt_dfs.append(df)
                    except Exception as e:
                        print(f"⚠️ {apt_info['apt_name']} 매매 데이터 저장 중 오류: {str(e)}")
                else:
                    # 매매 데이터가 없어도 전세 데이터를 위해 빈 df 추가
                    apt_dfs.append(pd.DataFrame())

                # 전세 데이터 처리
                jeonse_trades = apt_info.get('jeonse_data', [])
                if jeonse_trades:
                    # 데이터 정렬
                    jeonse_trades = sorted(jeonse_trades, key=lambda x: x['date'])
                    # 데이터프레임 생성
                    jeonse_df = pd.DataFrame(jeonse_trades)
                    # 엑셀 파일 저장
                    try:
                        self.save_apt_data_to_excel(jeonse_df, apt_info, "jeonse")
                        jeonse_dfs.append(jeonse_df)
                    except Exception as e:
                        print(f"⚠️ {apt_info['apt_name']} 전세 데이터 저장 중 오류: {str(e)}")
                else:
                    # 전세 데이터가 없으면 빈 df 추가
                    jeonse_dfs.append(pd.DataFrame())

            # 2. 그래프 생성은 저장 성공 후 별도로 처리
            self.update_progress(85, "그래프 생성 중...")

            # 매매 또는 전세 데이터가 하나라도 있는지 확인
            has_purchase_data = any(not df.empty for df in apt_dfs)
            has_jeonse_data = any(not df.empty for df in jeonse_dfs)

            if not has_purchase_data and not has_jeonse_data:
                show_topmost_info("알림", "선택한 모든 아파트에 매매/전세 거래 데이터가 없습니다.", parent=self.root)
                self.update_progress(0, "")
                self.graph_button.config(state="normal")
                return

            try:
                # 그래프 생성 전 데이터 준비
                monthly_dfs = []
                monthly_jeonse_dfs = []

                # 매매 데이터 처리
                for i, df in enumerate(apt_dfs):
                    if not df.empty:
                        # 아파트 이름과 면적 추가
                        df['apt_name'] = self.selected_apts[i]['apt_name']
                        df['area'] = self.selected_apts[i]['area']

                        # 월별 평균가 계산
                        df_monthly = df.groupby(pd.Grouper(key='date', freq='ME')).agg({'price': 'mean'}).reset_index()
                        df_monthly['apt_name'] = self.selected_apts[i]['apt_name']
                        df_monthly['area'] = self.selected_apts[i]['area']
                        monthly_dfs.append(df_monthly)
                    else:
                        monthly_dfs.append(pd.DataFrame())

                # 전세 데이터 처리
                for i, jeonse_df in enumerate(jeonse_dfs):
                    if not jeonse_df.empty:
                        # 아파트 이름과 면적 추가
                        jeonse_df['apt_name'] = self.selected_apts[i]['apt_name']
                        jeonse_df['area'] = self.selected_apts[i]['area']

                        # 월별 평균가 계산
                        jeonse_monthly = jeonse_df.groupby(pd.Grouper(key='date', freq='ME')).agg({'price': 'mean'}).reset_index()
                        jeonse_monthly['apt_name'] = self.selected_apts[i]['apt_name']
                        jeonse_monthly['area'] = self.selected_apts[i]['area']
                        monthly_jeonse_dfs.append(jeonse_monthly)
                    else:
                        monthly_jeonse_dfs.append(pd.DataFrame())

                # 그래프 생성 (전세 데이터도 포함)
                self.create_multi_chart(apt_dfs, monthly_dfs, jeonse_dfs, monthly_jeonse_dfs)
                
                # 히스토리 업데이트
                self.history_list = self.load_history()
                self.update_history_display()
                
                self.update_progress(100, "분석 완료!")
                
                # 그래프 이미지 표시
                if os.path.exists(self.image_path):
                    os.startfile(self.image_path)
                    
            except Exception as e:
                show_topmost_error("오류", f"그래프 생성 중 오류 발생: {str(e)}", parent=self.root)
                import traceback
                traceback.print_exc()  # 콘솔에 상세 오류 출력
            
        except Exception as e:
            show_topmost_error("오류", f"분석 중 오류 발생: {str(e)}", parent=self.root)
            import traceback
            traceback.print_exc()  # 콘솔에 상세 오류 출력
        finally:
            # 그래프 생성 버튼 다시 활성화
            self.graph_button.config(state="normal")

    def save_apt_data_to_excel(self, df, apt_info, data_type="purchase"):
        """아파트 거래 데이터를 엑셀로 저장하는 함수"""
        try:
            apt_name = apt_info['apt_name']
            area = apt_info['area']
            
            apt_name_clean = ''.join(char for char in apt_name if char.isalnum() or char.isspace())
            apt_name_clean = apt_name_clean.replace(' ', '_')
            
            # 데이터 유형에 따른 파일명 설정
            type_suffix = "매매" if data_type == "purchase" else "전세"
            filename = f"{apt_name_clean}_{area}m2_{type_suffix}.xlsx"
            excel_path = os.path.join(self.download_path, filename)
            
            # 데이터프레임을 엑셀로 저장
            writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
            
            # 기본 정보 시트 생성
            info_df = pd.DataFrame([
                ["단지 정보", ""],
                ["단지명", apt_name],
                ["지역", f"{apt_info['sido']} {apt_info['sigungu']} {apt_info['dong']}"],
                ["전용면적", f"{area}㎡"],
                ["거래유형", f"{type_suffix}"],
                ["거래건수", len(df)],
                ["최고거래가", f"{df['price'].max():,.0f}만원" if not df.empty else "정보없음"],
                ["최저거래가", f"{df['price'].min():,.0f}만원" if not df.empty else "정보없음"]
            ])
            
            info_df.to_excel(writer, sheet_name='기본정보', header=False, index=False)
            
            # 거래 데이터 시트 생성
            if not df.empty:
                trade_df = df.copy()
                trade_df['date'] = trade_df['date'].dt.strftime('%Y-%m-%d')
                
                # 거래 유형에 따른 컬럼명 설정
                if data_type == "purchase":
                    trade_df.columns = ['거래일자', '가격(만원)', '층', '면적(㎡)']
                else:
                    # 전세인 경우 rent_type 컬럼 처리
                    if 'rent_type' in trade_df.columns:
                        trade_df = trade_df.drop(columns=['rent_type'])
                    trade_df.columns = ['거래일자', '전세가(만원)', '층', '면적(㎡)']
                
                trade_df = trade_df.sort_values('거래일자', ascending=False)
                
                trade_df.to_excel(writer, sheet_name='거래내역', index=False)
            else:
                # 빈 데이터프레임 저장
                if data_type == "purchase":
                    empty_df = pd.DataFrame(columns=['거래일자', '가격(만원)', '층', '면적(㎡)'])
                else:
                    empty_df = pd.DataFrame(columns=['거래일자', '전세가(만원)', '층', '면적(㎡)'])
                empty_df.to_excel(writer, sheet_name='거래내역', index=False)
            
            # 엑셀 파일 저장
            writer.close()
            
            # 히스토리에 추가
            history_filename = f"history_{apt_name.replace(' ', '_')}_{area}_{type_suffix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            history_path = os.path.join(self.history_path, history_filename)
            
            # 히스토리용 파일 복사
            if os.path.exists(excel_path):
                shutil.copy2(excel_path, history_path)
            
            return excel_path
            
        except Exception as e:
            print(f"엑셀 저장 중 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    
    def on_closing(self):
        """프로그램 종료 시 설정 자동 저장 및 안전한 종료"""
        # 중복 호출 방지
        if self.is_closing:
            return
        self.is_closing = True

        try:
            # 현재 설정을 JSON 파일로 저장
            settings_file = os.path.join(os.getcwd(), 'real_estate_analyzer_settings.json')

            settings_data = {
                'download_path': self.download_path,
                'history_path': self.history_path,
                'lawdong_path': self.lawdong_path,
                'complex_info_path': self.complex_info_path,
                'trade_cache_path': self.trade_cache_path,
                'graph_options': {
                    'show_monthly_avg': self.show_monthly_avg.get(),
                    'show_monthly_max': self.show_monthly_max.get(),
                    'show_scatter_plot': self.show_scatter_plot.get(),
                    'show_jeonse': self.show_jeonse.get(),
                    'show_jeonse_monthly_avg': self.show_jeonse_monthly_avg.get(),
                    'show_jeonse_monthly_max': self.show_jeonse_monthly_max.get(),
                    'show_jeonse_scatter_plot': self.show_jeonse_scatter_plot.get(),
                    'show_complex_info': self.show_complex_info.get(),
                    'collect_jeonse_data': self.collect_jeonse_data.get()
                }
            }

            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)

            print(f"✅ 설정이 자동 저장되었습니다: {settings_file}")
        except Exception as e:
            print(f"⚠️ 설정 저장 중 오류 발생: {str(e)}")

        # 프로그램 종료
        self.root.destroy()

    def get_trade_data(self, apt_info):
        """선택한 아파트의 실거래 데이터 가져오기 (최적화된 버전)"""
        apt_name = apt_info['apt_name']
        target_area = apt_info['area']
        sigungu_code = apt_info['sigungu_code']
        dong = apt_info['dong']
        
        trades = []
        current_date = datetime.now()
        
        # 데이터 수집 설정
        max_months = 360  # 최대 30년(360개월)
        consecutive_empty_months_limit = 12  # 연속 12개월(1년) 동안 데이터가 없으면 중단
        
        # 진행 상황을 표시할 창 생성
        progress_window = tk.Toplevel(self.root)
        progress_window.title("거래 데이터 수집 중...")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        ttk.Label(progress_window, 
                  text=f"{apt_name} ({target_area}㎡) 실거래 데이터를 수집 중입니다...",
                  wraplength=350).pack(pady=10)
        
        progress_label = ttk.Label(progress_window, text="0% 완료")
        progress_label.pack(pady=5)
        
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=350, mode="determinate")
        progress_bar.pack(fill="x", padx=20, pady=10)
        
        cancel_button = ttk.Button(progress_window, text="중단", command=progress_window.destroy)
        cancel_button.pack(pady=5)
        
        # 취소 플래그
        cancel_flag = [False]
        progress_window.protocol("WM_DELETE_WINDOW", lambda: setattr(cancel_flag, 0, True) or progress_window.destroy())
        
        # 데이터 수집 함수
        def collect_data():
            nonlocal trades
            consecutive_empty_months = 0  # 연속으로 데이터가 없는 월 수
            
            try:
                for month in range(max_months):
                    if cancel_flag[0]:
                        break
                        
                    # 진행 상태 업데이트
                    progress = min(100, (month / max_months) * 100)
                    progress_bar['value'] = progress
                    
                    # 수집한 데이터 개수에 따라 진행 상태 메시지 업데이트
                    if trades:
                        progress_label.config(text=f"{progress:.1f}% 완료 - {len(trades)}건 수집됨 ({month+1}/{max_months}개월)")
                    else:
                        progress_label.config(text=f"{progress:.1f}% 완료 ({month+1}/{max_months}개월)")
                    
                    progress_window.update_idletasks()
                    
                    # 현재 조회할 월 계산
                    search_date = current_date - timedelta(days=30 * month)
                    deal_ymd = search_date.strftime("%Y%m")
                    
                    # API 호출
                    url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                           f"?serviceKey={self.service_key}"
                           f"&LAWD_CD={sigungu_code}"
                           f"&DEAL_YMD={deal_ymd}"
                           f"&numOfRows=1000")
                    
                    try:
                        response = requests.get(url, timeout=API_TIMEOUT)  # 타임아웃 설정
                        monthly_trades = []  # 이번 달 거래 데이터
                        
                        if response.status_code == 200:
                            root = ET.fromstring(response.text)
                            items = root.findall('.//item')
                            
                            for item in items:
                                item_apt = item.findtext('aptNm', '').strip()
                                item_dong = item.findtext('umdNm', '').strip()
                                
                                if item_apt == apt_name and item_dong == dong:
                                    area = float(item.findtext('excluUseAr', '0'))
                                    # 지정된 전용면적과 일치하는지 확인 (±1㎡ 오차 허용)
                                    if abs(area - float(target_area)) <= 1:
                                        trade = {
                                            'date': datetime(
                                                int(item.findtext('dealYear')),
                                                int(item.findtext('dealMonth')),
                                                int(item.findtext('dealDay', '1'))
                                            ),
                                            'price': int(item.findtext('dealAmount').replace(',', '')),
                                            'floor': int(item.findtext('floor', '0')),
                                            'area': area
                                        }
                                        monthly_trades.append(trade)
                        
                        # 이번 달 데이터 개수 확인
                        if monthly_trades:
                            consecutive_empty_months = 0  # 데이터가 있으면 카운터 리셋
                            trades.extend(monthly_trades)  # 전체 거래 목록에 추가
                            
                            # 진행 상태 업데이트
                            progress_label.config(text=f"{progress:.1f}% 완료 - {len(trades)}건 수집됨 ({month+1}/{max_months}개월)")
                            progress_window.update_idletasks()
                        else:
                            consecutive_empty_months += 1  # 데이터가 없으면 카운터 증가
                        
                        # 일정 기간 연속으로 데이터가 없으면 조기 종료
                        if consecutive_empty_months >= consecutive_empty_months_limit:
                            progress_label.config(text=f"데이터 수집 완료 - {consecutive_empty_months}개월 연속 거래 내역 없음")
                            break
                            
                        # 잠시 대기하여 API 서버 부하 방지
                        time.sleep(0.1)
                        
                    except Exception as e:
                        print(f"API 호출 중 오류: {str(e)}")
                        # 오류가 발생해도 계속 진행
                        continue
                
                # 진행 상태 100%로 설정
                progress_bar['value'] = 100
                progress_label.config(text=f"100% 완료 - 총 {len(trades)}건 수집됨")
                progress_window.update_idletasks()
                
                # 잠시 후 창 닫기
                time.sleep(0.5)
                if not cancel_flag[0]:
                    progress_window.destroy()
                    
            except Exception as e:
                show_topmost_error("오류", f"데이터 수집 중 오류 발생: {str(e)}", parent=self.root)
                progress_window.destroy()
        
        # 별도 스레드로 데이터 수집 실행
        import threading
        thread = threading.Thread(target=collect_data)
        thread.daemon = True
        thread.start()
        
        # 창이 닫힐 때까지 대기
        self.root.wait_window(progress_window)
        
        # 결과 반환
        if cancel_flag[0]:
            return []  # 취소된 경우 빈 목록 반환
        
        # 결과를 날짜순으로 정렬하여 반환
        return sorted(trades, key=lambda x: x['date'])

    
    # 2. 호출 부분 수정 - analyze_and_visualize_multi 함수 수정
    def analyze_and_visualize_multi(self):
        """여러 아파트 실거래 데이터 분석 및 시각화"""
        try:
            # 각 아파트별 데이터프레임 생성
            apt_dfs = []
            for apt_info in self.selected_apts:
                trades = apt_info['trades_data']
                if not trades:
                    continue
                    
                # 데이터 정렬
                trades = sorted(trades, key=lambda x: x['date'])
                
                # 데이터프레임 생성
                df = pd.DataFrame(trades)
                df['apt_name'] = apt_info['apt_name']
                df['area'] = apt_info['area']
                
                apt_dfs.append(df)
            
            if not apt_dfs:
                show_topmost_info("알림", "선택한 아파트의 실거래 데이터가 없습니다.", parent=self.root)
                return
                
            # 모든 데이터프레임 결합
            combined_df = pd.concat(apt_dfs)
            
            # 아파트별 월별 평균가 계산
            monthly_dfs = []
            for i, df in enumerate(apt_dfs):
                # 평균가 추가 (ME: Month End 사용)
                df_monthly = df.groupby(pd.Grouper(key='date', freq='ME')).agg({'price': 'mean'}).reset_index()
                df_monthly['apt_name'] = self.selected_apts[i]['apt_name']
                df_monthly['area'] = self.selected_apts[i]['area']
                monthly_dfs.append(df_monthly)
                
            combined_monthly = pd.concat(monthly_dfs) if monthly_dfs else None
            
            # 차트 생성
            self.create_multi_chart(apt_dfs, monthly_dfs)
            
            # 분석 데이터 저장
            for i, df in enumerate(apt_dfs):
                # self.save_analysis_result가 2개의 인자만 받도록 수정
                self.save_analysis_result(df, self.selected_apts[i])
            
            # 히스토리 리스트 갱신 - 여기서 한 번만 수행
            self.history_list = self.load_history()
            self.update_history_display()
            
        except Exception as e:
            show_topmost_error("오류", f"분석 중 오류 발생: {str(e)}", parent=self.root)
            import traceback
            traceback.print_exc()
            raise
    
        # create_multi_chart 함수에서 time 관련 코드 수정
    # create_multi_chart 함수의 시계열 그래프 부분 수정
        # create_multi_chart 함수에서 표 부분 수정

    def create_multi_chart(self, apt_dfs, monthly_dfs, jeonse_dfs=None, monthly_jeonse_dfs=None):
        """여러 아파트 거래 데이터 차트 생성 (seaborn 스타일, 개선된 레이블 배치)"""
        # seaborn 스타일 설정
        sns.set_style("whitegrid")
        sns.set_context("notebook", font_scale=1.1)

        # seaborn 설정 후 한글 폰트 재설정 (seaborn이 폰트를 덮어쓰므로)
        plt.rcParams['font.family'] = 'Malgun Gothic'
        plt.rcParams['axes.unicode_minus'] = False  # 마이너스 기호 깨짐 방지

        # 세부정보 가져오기가 선택된 경우만 단지정보 로드
        complex_df = None
        if self.show_complex_info.get():
            complex_df = self.load_complex_info()

        # 매매가와 전세가 데이터 모두 있는 경우 레이아웃 조정
        has_jeonse = jeonse_dfs and any(not df.empty for df in jeonse_dfs) and self.show_jeonse.get()

        # 단지세부정보 선택 여부에 따라 그래프 영역 크기 조정
        from matplotlib.gridspec import GridSpec

        if self.show_complex_info.get():
            # 하단에 단지세부정보 테이블이 있을 경우
            fig = plt.figure(figsize=(18, 13))  # 크기 증가 (수익률 테이블 공간 추가)
            gs = GridSpec(18, 1, figure=fig, hspace=0.6)  # 행 증가 및 간격 증가
            ax = fig.add_subplot(gs[0:9, 0])  # 그래프 영역: 0-8행 (9행)
            ax_yield = fig.add_subplot(gs[10:12, 0])  # 수익률 테이블 영역: 10-11행 (2행) - 1행 간격 추가
            # 단지세부정보 테이블은 나중에 gs[13:18, 0] 영역에 배치
        else:
            # 단지세부정보 없이 수익률 테이블만 표시
            fig = plt.figure(figsize=(18, 11))  # 크기 증가
            gs = GridSpec(13, 1, figure=fig, hspace=0.6)  # 행 증가 및 간격 증가
            ax = fig.add_subplot(gs[0:9, 0])  # 그래프 영역: 0-8행 (9행)
            ax_yield = fig.add_subplot(gs[10:12, 0])  # 수익률 테이블 영역: 10-11행 (2행) - 1행 간격 추가

        # 매매가 색상 - 더 진하고 가시성 좋은 색상
        colors = [
            '#E74C3C',  # 밝은 빨강 (매우 눈에 잘 띔)
            '#3498DB',  # 밝은 파랑
            '#2ECC71',  # 밝은 초록
            '#F39C12',  # 밝은 주황
            '#9B59B6',  # 보라
            '#1ABC9C',  # 청록
        ]
        markers = ['o', 's', '^', 'D', 'X', 'P']

        # 전세가 색상 (매매가보다 연한 색상)
        jeonse_colors = [
            '#F1948A',  # 연한 빨강
            '#85C1E2',  # 연한 파랑
            '#82E0AA',  # 연한 초록
            '#F8C471',  # 연한 주황
            '#BB8FCE',  # 연한 보라
            '#76D7C4',  # 연한 청록
        ]
        jeonse_markers = ['v', 'd', '*', 'p', '8', 'h']
        
        max_price = 0
        min_price = float('inf')
        first_date = datetime.now()
        last_date = datetime(1990, 1, 1)

        # 연복리 계산 및 매전갭 정보를 저장할 변수
        cagr_info = []

        # 어노테이션 저장 리스트 (adjustText를 위한)
        annotations = []
        texts_for_adjust = []
        
        # 1. 매매가 데이터 플롯
        for i, df in enumerate(apt_dfs):
            if df.empty:
                continue
                
            apt_info = None
            for apt in self.selected_apts:
                if apt['apt_name'] == df['apt_name'].iloc[0] and str(apt['area']) == str(df['area'].iloc[0]):
                    apt_info = apt
                    break
                    
            if not apt_info:
                continue
                
            apt_name = apt_info['apt_name']
            area = apt_info['area']
            
            color = colors[i % len(colors)]
            marker = markers[i % len(markers)]
            
            # 날짜 범위 업데이트
            if df['date'].min() < first_date:
                first_date = df['date'].min()
            if df['date'].max() > last_date:
                last_date = df['date'].max()
            
            # 점도표 (실거래가)
            if self.show_scatter_plot.get():
                alpha = 0.5 if len(apt_dfs) > 1 else 0.7
                ax.scatter(df['date'], df['price'], 
                           color=color, alpha=alpha, s=40, 
                           marker=marker,
                           label=f'{apt_name} ({area}㎡) 실거래', zorder=3)
            
            # 월별 데이터 처리 개선
            # 월별 고유 값 추출(년-월 형식)
            df_sorted = df.sort_values('date')
            df_sorted['year_month'] = df_sorted['date'].dt.strftime('%Y-%m')
            
            # 월별 데이터 계산
            monthly_data = df_sorted.groupby('year_month').agg({
                'date': 'first',  # 각 월의 첫 날짜
                'price': ['mean', 'max']  # 평균가와 최고가
            }).reset_index()
            
            # 컬럼 이름 재설정
            monthly_data.columns = ['year_month', 'date', 'avg_price', 'max_price']
            
            # 날짜 순서로 정렬
            monthly_data = monthly_data.sort_values('date')
            
            # 월 평균가격 그래프
            if self.show_monthly_avg.get():
                if len(monthly_data) > 0:
                    ax.plot(monthly_data['date'], monthly_data['avg_price'], 
                            color=color, linewidth=2.5, 
                            label=f'{apt_name} ({area}㎡) 월평균', zorder=4)
            
            # 월 최고가 그래프
            # 매매가 월 최고가 그래프에서 어노테이션 위치 조정
            # 매매가 월 최고가 그래프에서 어노테이션 위치 조정
            # 매매가 월 최고가 어노테이션 위치 조정 부분 수정
            if self.show_monthly_max.get():
                if len(monthly_data) > 0:
                    # 월 최고가 선 그래프
                    ax.plot(monthly_data['date'], monthly_data['max_price'], 
                            color=color, linewidth=1.5, linestyle='--', dashes=(2, 2),
                            label=f'{apt_name} ({area}㎡) 월최고가', zorder=4)
                    
                    # 월 최고가 중 가장 높은 값 (최고가)
                    monthly_max_idx = monthly_data['max_price'].idxmax()
                    monthly_max = monthly_data.iloc[monthly_max_idx]
                    
                    # 가장 최근 거래가
                    monthly_latest = monthly_data.iloc[-1]
                    
                    # 최고가와 최근가의 날짜 차이 계산 (6개월 이내인지 확인)
                    date_diff = abs((monthly_latest['date'] - monthly_max['date']).days)
                    close_dates = date_diff <= 180  # 6개월(180일) 이내
                    
                    # 상대적 가격 위치에 따른 어노테이션 위치 조정
                    # 현재 단지의 최고가
                    current_max_price = monthly_max['max_price']

                    # 요구사항 1, 2: 단지 수에 따른 위치 조정
                    # x 좌표를 기존 -40에서 -80으로 수정하여 더 왼쪽으로 이동
                    if len(apt_dfs) == 2:
                        # 2개 단지 비교일 경우: 큰 값은 위로, 작은 값은 아래로
                        # 다른 단지의 인덱스 구하기
                        other_idx = 1 if i == 0 else 0
                        if not apt_dfs[other_idx].empty:
                            other_max_price = apt_dfs[other_idx]['price'].max()
                            if current_max_price > other_max_price:
                                max_xytext = (-100, 30)  # 더 높은 가격은 위로
                            else:
                                max_xytext = (-100, -30)  # 더 낮은 가격은 아래로
                        else:
                            max_xytext = (-100, 30)  # 다른 단지 데이터 없으면 위로
                    elif len(apt_dfs) >= 3:
                        # 3개 이상 단지 비교: 가장 큰 값은 위로, 가장 작은 값은 아래로, 중간값은 그대로
                        all_max_prices = [df['price'].max() for df in apt_dfs if not df.empty]
                        sorted_prices = sorted(all_max_prices)

                        if len(sorted_prices) >= 3:
                            if current_max_price == sorted_prices[-1]:  # 가장 큰 값
                                max_xytext = (-100, 50)  # 최상단
                            elif current_max_price == sorted_prices[0]:  # 가장 작은 값
                                max_xytext = (-100, -50)  # 최하단
                            else:  # 중간값
                                max_xytext = (-100, 0)  # 중앙
                        else:
                            # 정렬된 가격이 3개보다 적은 경우
                            if current_max_price == max(sorted_prices):
                                max_xytext = (-100, 30)  # 위로
                            else:
                                max_xytext = (-100, -30)  # 아래로
                    else:
                        # 단일 단지인 경우
                        max_xytext = (-80, 30)  # 좌측 위로
                    
                    # 최고가 점 표시
                    ax.scatter([monthly_max['date']], [monthly_max['max_price']],
                              color='white', edgecolor=color, s=100, linewidth=2.5, zorder=5)

                    # 최고가 어노테이션 추가 (더 세련된 스타일)
                    ann = ax.annotate(f"{monthly_max['max_price']:,.0f}만\n(최고가)",
                              xy=(monthly_max['date'], monthly_max['max_price']),
                              xytext=max_xytext,
                              textcoords='offset points',
                              bbox=dict(boxstyle='round,pad=0.6', fc='white', alpha=0.95,
                                       edgecolor=color, linewidth=2),
                              fontsize=10,
                              color='red',
                              fontweight='bold',
                              zorder=6,
                              arrowprops=dict(
                                  arrowstyle='->',
                                  color=color,
                                  lw=1.5,
                                  alpha=0.8,
                                  connectionstyle="arc3,rad=0.2"
                              ))
                    texts_for_adjust.append(ann)
                    
                    # 현재가 어노테이션 표시 조건 - 최고가와 다르고, 날짜가 충분히 떨어진 경우만
                    if monthly_latest['date'] != monthly_max['date'] and not close_dates:
                        
                        # 요구사항 1, 2: 단지 수에 따른 위치 조정 (현재가)
                        # 2개 단지 비교일 경우 코드 수정
                        if len(apt_dfs) == 2:
                            # 다른 단지의 인덱스 구하기
                            other_idx = 1 if i == 0 else 0
                            
                            # 다른 단지의 데이터가 있는지 확인
                            if not apt_dfs[other_idx].empty and len(monthly_dfs) > other_idx:
                                other_monthly_data = monthly_dfs[other_idx]
                                
                                if not other_monthly_data.empty:
                                    # 다른 단지의 최근 월별 데이터
                                    other_latest = other_monthly_data.iloc[-1]
                                    
                                    # 현재 가격이 다른 단지보다 높으면 위로, 낮으면 아래로
                                    if monthly_latest['max_price'] > other_latest['price']:
                                        latest_xytext = (30, 30)  # 우측 위로
                                    else:
                                        latest_xytext = (30, -30)  # 우측 아래로
                                else:
                                    # 다른 단지의 월별 데이터가 없으면 기본값
                                    latest_xytext = (30, 0)
                            else:
                                # 다른 단지의 데이터가 없으면 기본값
                                latest_xytext = (30, 0)
                        # 3개 이상 단지 비교: 가장 큰 값은 위로, 가장 작은 값은 아래로, 중간값은 그대로
                        elif len(apt_dfs) >= 3:
                            # 모든 단지의 최근 가격을 비교
                            latest_prices = []
                            
                            for j, monthly_df in enumerate(monthly_dfs):
                                if j != i and not monthly_df.empty:  # 현재 단지가 아니고, 데이터가 있는 경우
                                    latest_prices.append(monthly_df.iloc[-1]['price'])
                            
                            if latest_prices:  # 다른 단지의, 비교할 데이터가 있는 경우
                                current_price = monthly_latest['max_price']
                                
                                # 모든 가격 중 현재 가격의 상대적 위치 확인
                                if current_price >= max(latest_prices):  # 가장 큰 값
                                    latest_xytext = (30, 50)  # 우측 위로
                                elif current_price <= min(latest_prices):  # 가장 작은 값
                                    latest_xytext = (30, -50)  # 우측 아래로
                                else:  # 중간값
                                    latest_xytext = (30, 0)  # 위치 변경 없음
                            else:
                                # 비교할 다른 단지 데이터가 없는 경우
                                latest_xytext = (30, 0)
                        else:
                            # 단일 단지인 경우
                            latest_xytext = (30, -30)  # 우측 아래로

                        
                        # 최근 거래가 점 표시
                        ax.scatter([monthly_latest['date']], [monthly_latest['max_price']],
                                  color='white', edgecolor=color, s=100, linewidth=2.5, zorder=5)

                        # 현재가 어노테이션 추가 (더 세련된 스타일)
                        ann = ax.annotate(f"{monthly_latest['max_price']:,.0f}만\n({monthly_latest['date'].strftime('%y.%m')})",
                                  xy=(monthly_latest['date'], monthly_latest['max_price']),
                                  xytext=latest_xytext,
                                  textcoords='offset points',
                                  bbox=dict(boxstyle='round,pad=0.6', fc='white', alpha=0.95,
                                           edgecolor=color, linewidth=2),
                                  fontsize=10,
                                  color='green',
                                  fontweight='bold',
                                  zorder=6,
                                  arrowprops=dict(
                                      arrowstyle='->',
                                      color=color,
                                      lw=1.5,
                                      alpha=0.8,
                                      connectionstyle="arc3,rad=-0.2"
                                  ))
                        texts_for_adjust.append(ann)
                    

            
            # 최대/최소 가격 갱신
            if not df.empty:
                if df['price'].max() > max_price:
                    max_price = df['price'].max()
                if df['price'].min() < min_price:
                    min_price = df['price'].min()
                    
            # 연복리 계산 추가
            if len(df) >= 2:
                first_trade = df.iloc[0]
                last_trade = df.iloc[-1]
                max_trade = df.loc[df['price'].idxmax()]  # 최고거래가
                price_change = last_trade['price'] - first_trade['price']
                years = (last_trade['date'] - first_trade['date']).days / 365.25
                
                if years > 0 and first_trade['price'] > 0:
                    cagr = (((last_trade['price'] / first_trade['price'])) ** (1/years) - 1) * 100

                    # 전고점 대비 하락률 계산
                    decline_rate = 0.0
                    if max_trade['price'] > 0:
                        decline_rate = ((max_trade['price'] - last_trade['price']) / max_trade['price']) * 100

                    # 최근 매매가 추가
                    cagr_info.append({
                        'apt_name': apt_name,
                        'area': area,
                        'color': color,
                        'first_date': first_trade['date'],
                        'last_date': last_trade['date'],
                        'years': years,
                        'first_price': first_trade['price'],
                        'last_price': last_trade['price'],
                        'max_price': max_trade['price'],  # 최고거래가 추가
                        'max_date': max_trade['date'],    # 최고거래가 날짜 추가
                        'decline_rate': decline_rate,      # 전고점 대비 하락률 추가
                        'change': price_change,
                        'cagr': cagr,
                        'type': '매매',
                        'jeonse_price': 0,  # 전세가 (기본값)
                        'gap': 0,  # 매전갭 (기본값)
                        'min_gap_2020': None,  # 2020년 이후 최소 갭 (기본값)
                        'min_gap_date_2020': None  # 2020년 이후 최소 갭 날짜 (기본값)
                    })
        
        # 2. 전세가 데이터 플롯 (매매가와 동일한 옵션 적용)
        if jeonse_dfs and self.show_jeonse.get():
            for i, df in enumerate(jeonse_dfs):
                if df.empty:
                    continue
                    
                apt_info = None
                for apt in self.selected_apts:
                    if apt['apt_name'] == df['apt_name'].iloc[0] and str(apt['area']) == str(df['area'].iloc[0]):
                        apt_info = apt
                        break
                        
                if not apt_info:
                    continue
                    
                apt_name = apt_info['apt_name']
                area = apt_info['area']
                
                color = jeonse_colors[i % len(jeonse_colors)]
                marker = jeonse_markers[i % len(jeonse_markers)]
                
                # 날짜 범위 업데이트
                if df['date'].min() < first_date:
                    first_date = df['date'].min()
                if df['date'].max() > last_date:
                    last_date = df['date'].max()
                
                # 전세 점도표 (매매가와 동일한 옵션 적용)
                if self.show_jeonse_scatter_plot.get():
                    alpha = 0.5 if len(jeonse_dfs) > 1 else 0.7
                    ax.scatter(df['date'], df['price'], 
                              color=color, alpha=alpha, s=35, 
                              marker=marker,
                              label=f'{apt_name} ({area}㎡) 전세 실거래', zorder=2)
                
                # 월별 데이터 처리 - 매매가와 동일한 방식
                df_sorted = df.sort_values('date')
                df_sorted['year_month'] = df_sorted['date'].dt.strftime('%Y-%m')
                
                # 월별 전세 데이터 계산 (평균, 최고가)
                monthly_data = df_sorted.groupby('year_month').agg({
                    'date': 'first',
                    'price': ['mean', 'max']  # 평균가와 최고가 모두 계산
                }).reset_index()
                
                # 컬럼명 설정
                monthly_data.columns = ['year_month', 'date', 'avg_price', 'max_price']
                monthly_data = monthly_data.sort_values('date')
                
                # 전세 월평균가 추세선 (매매가와 동일한 옵션 적용)
                if self.show_jeonse_monthly_avg.get() and len(monthly_data) > 0:
                    ax.plot(monthly_data['date'], monthly_data['avg_price'], 
                           color=color, linewidth=2, 
                           label=f'{apt_name} ({area}㎡) 전세 월평균', zorder=3)
                
                # 전세 월최고가 추세선 (매매가와 동일한 옵션 적용)
                # 전세 월최고가 추세선에도 동일한 로직 적용
                # 전세가 어노테이션 위치 조정 부분도 유사하게 수정
                # 전세가 어노테이션 위치 조정 부분도 유사하게 수정
                if self.show_jeonse_monthly_max.get() and len(monthly_data) > 0:
                    # 월 최고가 선 그래프
                    ax.plot(monthly_data['date'], monthly_data['max_price'], 
                           color=color, linewidth=1.5, linestyle='--', dashes=(2, 2),
                           label=f'{apt_name} ({area}㎡) 전세 월최고가', zorder=3)
                    
                    # 전세 최고가 표시
                    monthly_max_idx = monthly_data['max_price'].idxmax()
                    monthly_max = monthly_data.iloc[monthly_max_idx]
                    
                    # 최근 전세가
                    monthly_latest = monthly_data.iloc[-1]
                    
                    # 최고가와 최근가의 날짜 차이 계산 (6개월 이내인지 확인)
                    date_diff = abs((monthly_latest['date'] - monthly_max['date']).days)
                    close_dates = date_diff <= 180  # 6개월(180일) 이내
                    
                    # 요구사항 1, 2: 단지 수에 따른 위치 조정 (전세 최고가)
                    # x 좌표를 기존 -40에서 -80으로 수정하여 더 왼쪽으로 이동
                    if len(jeonse_dfs) == 2:
                        # 2개 단지 비교일 경우: 큰 값은 위로, 작은 값은 아래로
                        if i == 0:  # 첫 번째 단지
                            jeonse_max_xytext = (-100, 30)  # 좌측 위로 (x 값 -80)
                        else:  # 두 번째 단지
                            jeonse_max_xytext = (-100, -30)  # 좌측 아래로 (x 값 -80)
                    elif len(jeonse_dfs) >= 3:
                        # 3개 이상 단지 비교: 가장 큰 값은 위로, 가장 작은 값은 아래로, 중간값은 그대로
                        sorted_prices = sorted([df['price'].max() for df in jeonse_dfs if not df.empty])
                        if len(sorted_prices) >= 3:
                            if df['price'].max() == sorted_prices[-1]:  # 가장 큰 값
                                jeonse_max_xytext = (-100, 30)  # 좌측 위로 (x 값 -80)
                            elif df['price'].max() == sorted_prices[0]:  # 가장 작은 값
                                jeonse_max_xytext = (-100, -30)  # 좌측 아래로 (x 값 -80)
                            else:  # 중간값
                                jeonse_max_xytext = (-100, 0)  # 좌측으로만 이동 (x 값 -80)
                        else:
                            # 정렬된 가격이 3개보다 적은 경우
                            if df['price'].max() == max(sorted_prices):
                                jeonse_max_xytext = (-100, 30)  # 좌측 위로 (x 값 -80)
                            else:
                                jeonse_max_xytext = (-100, -30)  # 좌측 아래로 (x 값 -80)
                    else:
                        # 단일 단지인 경우
                        jeonse_max_xytext = (-100, 30)  # 좌측 위로 (x 값 -80)
                    
                    # 전세 최고가 점 표시
                    ax.scatter([monthly_max['date']], [monthly_max['max_price']], 
                              color='white', edgecolor=color, s=70, linewidth=2, zorder=5)
                    
                    # 요구사항 3: 최고가는 항상 왼쪽에 위치 (x 값 -40)
                    ax.annotate(f"전세 {monthly_max['max_price']:,.0f}만원 (최고가)\n"
                               f"({monthly_max['date'].strftime('%Y-%m')})",
                               xy=(monthly_max['date'], monthly_max['max_price']),
                               xytext=jeonse_max_xytext, 
                               textcoords='offset points',
                               bbox=dict(boxstyle='round,pad=0.5', fc='white', alpha=0.8, edgecolor=color),
                               fontsize=9,
                               color='red',
                               fontweight='bold',
                               zorder=6,
                               arrowprops=dict(
                                   arrowstyle='->',
                                   color=color,
                                   lw=1.2,
                                   alpha=0.7
                               ))
                    
                    # 현재가 어노테이션 표시 조건 - 최고가와 다르고, 날짜가 충분히 떨어진 경우만
                    if monthly_latest['date'] != monthly_max['date'] and not close_dates:
                        
                        # 요구사항 1, 2: 단지 수에 따른 위치 조정 (전세 현재가)
                        # 2개 단지 비교일 경우 코드 수정
                        # 2개 단지 비교일 경우 코드 수정
                        if len(jeonse_dfs) == 2:
                            # 다른 단지의 인덱스 구하기
                            other_idx = 1 if i == 0 else 0
                            
                            # 다른 단지의 데이터가 있는지 확인
                            if not jeonse_dfs[other_idx].empty and len(monthly_jeonse_dfs) > other_idx:
                                other_monthly_data = monthly_jeonse_dfs[other_idx]
                                
                                if not other_monthly_data.empty:
                                    # 다른 단지의 최근 월별 데이터
                                    other_latest = other_monthly_data.iloc[-1]
                                    
                                    # 현재 가격이 다른 단지보다 높으면 위로, 낮으면 아래로
                                    if monthly_latest['max_price'] > other_latest['price']:
                                        jeonse_latest_xytext = (30, 30)  # 우측 위로
                                    else:
                                        jeonse_latest_xytext = (30, -30)  # 우측 아래로
                                else:
                                    # 다른 단지의 월별 데이터가 없으면 기본값
                                    jeonse_latest_xytext = (30, 0)
                            else:
                                # 다른 단지의 데이터가 없으면 기본값
                                jeonse_latest_xytext = (30, 0)
                        # 3개 이상 단지 비교: 가장 큰 값은 위로, 가장 작은 값은 아래로, 중간값은 그대로
                        elif len(jeonse_dfs) >= 3:
                            # 모든 단지의 최근 가격을 비교
                            latest_prices = []
                            
                            for j, monthly_df in enumerate(monthly_jeonse_dfs):
                                if j != i and not monthly_df.empty:  # 현재 단지가 아니고, 데이터가 있는 경우
                                    latest_prices.append(monthly_df.iloc[-1]['price'])
                            
                            if latest_prices:  # 다른 단지의, 비교할 데이터가 있는 경우
                                current_price = monthly_latest['max_price']
                                
                                # 모든 가격 중 현재 가격의 상대적 위치 확인
                                if current_price >= max(latest_prices):  # 가장 큰 값
                                    jeonse_latest_xytext = (30, 50)  # 우측 위로
                                elif current_price <= min(latest_prices):  # 가장 작은 값
                                    jeonse_latest_xytext = (30, -50)  # 우측 아래로
                                else:  # 중간값
                                    jeonse_latest_xytext = (30, 0)  # 위치 변경 없음
                            else:
                                # 비교할 다른 단지 데이터가 없는 경우
                                jeonse_latest_xytext = (30, 0)
                        else:
                            # 단일 단지인 경우
                            jeonse_latest_xytext = (30, -30)  # 우측 아래로
                        # 최근 전세가 점 표시
                        ax.scatter([monthly_latest['date']], [monthly_latest['max_price']], 
                                  color='white', edgecolor=color, s=70, linewidth=2, zorder=5)
                        
                        # 요구사항 3: 현재가는 항상 오른쪽에 위치
                        ax.annotate(f"전세 {monthly_latest['max_price']:,.0f}만원\n"
                                   f"({monthly_latest['date'].strftime('%Y-%m')})",
                                   xy=(monthly_latest['date'], monthly_latest['max_price']),
                                   xytext=jeonse_latest_xytext, 
                                   textcoords='offset points',
                                   bbox=dict(boxstyle='round,pad=0.5', fc='white', alpha=0.8, edgecolor=color),
                                   fontsize=9, 
                                   zorder=6,
                                   arrowprops=dict(
                                       arrowstyle='->',
                                       color=color,
                                       lw=1.2,
                                       alpha=0.7
                                   ))
                    
                    # 최근 전세가와 최고가가 다르고, 6개월 이상 차이나는 경우에만
  
                
                # 전세 최대/최소 가격도 전체 그래프 범위에 포함
                if not df.empty:
                    if df['price'].max() > max_price:
                        max_price = df['price'].max()
                    if df['price'].min() < min_price:
                        min_price = df['price'].min()
                        
                # 전세 연복리 계산 추가
                if len(df) >= 2:
                    first_trade = df.iloc[0]
                    last_trade = df.iloc[-1]
                    max_trade = df.loc[df['price'].idxmax()]  # 최고거래가
                    price_change = last_trade['price'] - first_trade['price']
                    years = (last_trade['date'] - first_trade['date']).days / 365.25
                    
                    if years > 0 and first_trade['price'] > 0:
                        cagr = (((last_trade['price'] / first_trade['price'])) ** (1/years) - 1) * 100
                        
                        # 전고점 대비 하락률 계산 (전세)
                        decline_rate = 0.0
                        if max_trade['price'] > 0:
                            decline_rate = ((max_trade['price'] - last_trade['price']) / max_trade['price']) * 100

                        # 연복리 정보 저장
                        jeonse_info = {
                            'apt_name': apt_name,
                            'area': area,
                            'color': color,
                            'first_date': first_trade['date'],
                            'last_date': last_trade['date'],
                            'years': years,
                            'first_price': first_trade['price'],
                            'last_price': last_trade['price'],
                            'max_price': max_trade['price'],  # 최고거래가 추가
                            'max_date': max_trade['date'],    # 최고거래가 날짜 추가
                            'decline_rate': decline_rate,      # 전고점 대비 하락률 추가
                            'change': price_change,
                            'cagr': cagr,
                            'type': '전세'
                        }
                        
                        # 매매 정보와 연결하여 매전갭 계산
                        for j, info in enumerate(cagr_info):
                            if info['apt_name'] == apt_name and info['area'] == area and info['type'] == '매매':
                                cagr_info[j]['jeonse_price'] = last_trade['price']
                                cagr_info[j]['gap'] = info['last_price'] - last_trade['price']

                                # 2020년 이후 최소 갭 계산
                                # 매매 데이터와 전세 데이터를 날짜별로 매칭
                                maemae_df = apt_dfs[i]
                                jeonse_df = df

                                # 2020년 1월 1일 이후 데이터만 필터링
                                cutoff_date = datetime(2020, 1, 1)
                                maemae_2020 = maemae_df[maemae_df['date'] >= cutoff_date].copy()
                                jeonse_2020 = jeonse_df[jeonse_df['date'] >= cutoff_date].copy()

                                if not maemae_2020.empty and not jeonse_2020.empty:
                                    # 각 매매 날짜에 대해 가장 가까운 전세가 찾기
                                    min_gap = float('inf')
                                    min_gap_date = None

                                    for _, maemae_row in maemae_2020.iterrows():
                                        # 해당 매매 날짜 이전의 가장 최근 전세가 찾기
                                        recent_jeonse = jeonse_2020[jeonse_2020['date'] <= maemae_row['date']]
                                        if not recent_jeonse.empty:
                                            closest_jeonse = recent_jeonse.iloc[-1]  # 가장 최근 전세가
                                            gap = maemae_row['price'] - closest_jeonse['price']
                                            if gap < min_gap:
                                                min_gap = gap
                                                min_gap_date = maemae_row['date']

                                    if min_gap != float('inf'):
                                        cagr_info[j]['min_gap_2020'] = min_gap
                                        cagr_info[j]['min_gap_date_2020'] = min_gap_date

                                break

                        # 전세 연복리 정보 추가
                        cagr_info.append(jeonse_info)
        
        # 날짜 범위 설정 (적절한 여백 추가)
        date_range = (last_date - first_date).days
        date_padding = timedelta(days=date_range * 0.05)  # 5% 여백
        ax.set_xlim(first_date - date_padding, last_date + date_padding)
        
        # 가격 범위 설정 (적절한 여백 추가)
        if min_price < float('inf') and max_price > 0:
            price_range = max_price - min_price
            price_padding = price_range * 0.1  # 10% 여백
            ax.set_ylim(max(0, min_price - price_padding), max_price + price_padding)
        
        # 차트 제목 변경 - '부태리의 실거래가 비교 분석'으로 변경
        title_text = "부태리의 실거래가 비교 분석\n"
        
        # 매매가 데이터의 아파트 정보 먼저 표시
        for i, df in enumerate(apt_dfs):
            if not df.empty:
                # 해당 아파트 정보 찾기
                apt_info = None
                for apt in self.selected_apts:
                    if apt['apt_name'] == df['apt_name'].iloc[0] and str(apt['area']) == str(df['area'].iloc[0]):
                        apt_info = apt
                        break
                
                if apt_info:
                    apt_name = apt_info['apt_name']
                    area = apt_info['area']
                    
                    # 준공연도 정보가 있으면 표시
                    build_year_text = ""
                    if 'build_year' in apt_info and apt_info['build_year']:
                        current_year = datetime.now().year
                        age = current_year - int(apt_info['build_year'])
                        build_year_text = f", {apt_info['build_year']}년 준공({age}년차)"
                    
                    title_text += f"{apt_name} ({area}㎡{build_year_text})"
                    
                    if i < len(apt_dfs) - 1 and i < len(apt_dfs) - 1 and not apt_dfs[i+1].empty:
                        title_text += " vs "
        
        ax.set_title(title_text, pad=20, fontsize=18, fontweight='bold')

        ax.set_xlabel('', fontsize=12, labelpad=15)  # labelpad 추가로 여백 확보
        ax.set_ylabel('가격(만원)', fontsize=12)
        
        # x축 날짜 형식 설정
        if date_range > 365 * 5:  # 5년 이상
            ax.xaxis.set_major_locator(mdates.YearLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
            ax.xaxis.set_minor_locator(mdates.MonthLocator([1, 7]))  # 1월, 7월에 minor tick
        else:
            ax.xaxis.set_major_locator(mdates.MonthLocator([1, 4, 7, 10]))  # 분기별
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
                    
        plt.setp(ax.get_xticklabels(), rotation=-45, ha='center')
        
        # y축 숫자 형식 설정 (천 단위 콤마)
        import matplotlib.ticker as ticker
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
        
        # adjustText를 사용하여 레이블 자동 조정
        if ADJUSTTEXT_AVAILABLE and texts_for_adjust:
            try:
                adjust_text(
                    texts_for_adjust,
                    ax=ax,
                    arrowprops=dict(
                        arrowstyle='->',
                        color='gray',
                        lw=0.5,
                        alpha=0.5,
                        shrinkA=5,  # 화살표 시작점을 텍스트에서 5pt 떨어뜨림
                        shrinkB=5   # 화살표 끝점을 포인트에서 5pt 떨어뜨림
                    ),
                    expand_points=(1.2, 1.3),
                    expand_text=(1.2, 1.3),
                    force_points=(0.5, 0.5),
                    force_text=(0.5, 0.5),
                    avoid_self=True,
                    only_move={'points':'y', 'text':'xy'},
                    lim=500  # 최대 반복 횟수 제한
                )
            except Exception as e:
                print(f"⚠️ adjustText 적용 중 경고 (무시 가능): {str(e)}")

        # seaborn 스타일 그리드 설정
        ax.grid(True, linestyle='--', alpha=0.3, linewidth=0.8, color='gray')
        ax.set_axisbelow(True)  # 그리드를 데이터 아래로

        # x축 tick label 패딩 추가 (테이블과 간격 확보)
        ax.tick_params(axis='x', pad=10)

        # 범례 설정 (seaborn 스타일)
        legend = ax.legend(loc='upper left', bbox_to_anchor=(0.01, 0.99),
                          fontsize=10, framealpha=0.95, frameon=True,
                          fancybox=True, shadow=True,
                          edgecolor='gray', facecolor='white')

        # 2개 단지 비교 시 2020년 이후 두 단지 매매가 갭이 가장 작았던 시점 표시
        if len(apt_dfs) == 2 and not apt_dfs[0].empty and not apt_dfs[1].empty:
            # 2020년 이후 데이터만 필터링
            cutoff_date = datetime(2020, 1, 1)
            apt1_2020 = apt_dfs[0][apt_dfs[0]['date'] >= cutoff_date].copy()
            apt2_2020 = apt_dfs[1][apt_dfs[1]['date'] >= cutoff_date].copy()

            if not apt1_2020.empty and not apt2_2020.empty:
                # 월별 평균가 계산
                apt1_2020['year_month'] = apt1_2020['date'].dt.to_period('M')
                apt2_2020['year_month'] = apt2_2020['date'].dt.to_period('M')

                apt1_monthly = apt1_2020.groupby('year_month')['price'].mean().reset_index()
                apt2_monthly = apt2_2020.groupby('year_month')['price'].mean().reset_index()

                # 두 단지 모두 데이터가 있는 월만 비교
                common_months = set(apt1_monthly['year_month']) & set(apt2_monthly['year_month'])

                min_gap = float('inf')
                min_gap_date = None
                min_gap_price1 = 0
                min_gap_price2 = 0

                for month in common_months:
                    apt1_price = apt1_monthly[apt1_monthly['year_month'] == month]['price'].iloc[0]
                    apt2_price = apt2_monthly[apt2_monthly['year_month'] == month]['price'].iloc[0]
                    gap = abs(apt1_price - apt2_price)

                    if gap < min_gap:
                        min_gap = gap
                        min_gap_date = month.to_timestamp()  # Period를 Timestamp로 변환
                        min_gap_price1 = apt1_price
                        min_gap_price2 = apt2_price

                if min_gap_date is not None:
                    # 그래프 하단에 표시 (그래프를 가리지 않도록)
                    avg_price = (min_gap_price1 + min_gap_price2) / 2
                    apt1_name = apt_dfs[0]['apt_name'].iloc[0]
                    apt2_name = apt_dfs[1]['apt_name'].iloc[0]

                    # y축 범위 가져오기
                    y_min, y_max = ax.get_ylim()
                    y_range = y_max - y_min

                    # 하단 20% 위치에 배치
                    label_y_position = y_min + (y_range * 0.2)

                    # 어노테이션 추가 (아래쪽으로 화살표)
                    ax.annotate(
                        f'📍 두 단지 최소 갭\n{min_gap_date.strftime("%Y.%m")}\n갭: {min_gap:,.0f}만원',
                        xy=(min_gap_date, avg_price),
                        xytext=(0, -120),  # 아래쪽으로 더 멀리 배치
                        textcoords='offset points',
                        fontsize=10,
                        fontweight='bold',
                        color='#E74C3C',
                        bbox=dict(boxstyle='round,pad=0.8', facecolor='#FFF3CD',
                                 edgecolor='#E74C3C', linewidth=2, alpha=0.95),
                        arrowprops=dict(arrowstyle='->', color='#E74C3C', lw=2,
                                      connectionstyle="arc3,rad=-0.3"),
                        zorder=15,
                        horizontalalignment='center'
                    )

        # 매매·전세가 및 수익률 정보 테이블을 그래프 아래 별도 영역에 표시
        # ax_yield 영역 설정
        ax_yield.axis('off')  # 축 숨기기

        if cagr_info:
            # 테이블 데이터 준비
            headers = ['아파트', '종류', '최고거래가', '최근가격', '매전갭', '하락률', '연복리']
            table_data = []

            for info in cagr_info:
                apt_short = info['apt_name']
                if len(apt_short) > 12:
                    apt_short = apt_short[:12] + ".."

                # 1줄로 표기: 아파트명(면적)
                apt_label = f"{apt_short}({info['area']}㎡)"
                trade_type = info['type']
                # 1줄로 표기: 99,000만원
                max_price_text = f"{info['max_price']:,.0f}만원"
                last_price = f"{info['last_price']:,.0f}만원"

                gap = ''
                if trade_type == '매매' and info['jeonse_price'] > 0:
                    gap = f"{info['gap']:,.0f}만원"

                # 하락률에 마이너스 기호 추가
                decline_text = f"-{info['decline_rate']:.1f}%"
                cagr_value = f"{info['cagr']:.1f}%"

                table_data.append([apt_label, trade_type, max_price_text, last_price, gap, decline_text, cagr_value])

            # matplotlib 테이블 생성
            table = ax_yield.table(cellText=table_data,
                                   colLabels=headers,
                                   cellLoc='center',
                                   loc='center',
                                   bbox=[0, 0, 1, 1])

            # 테이블 스타일 설정
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 1.8)  # 행 높이 조정 (1줄 표기이므로 낮춤)

            # 헤더 스타일
            for i in range(len(headers)):
                cell = table[(0, i)]
                cell.set_facecolor('#4472C4')
                cell.set_text_props(weight='bold', color='white', fontsize=11)
                cell.set_edgecolor('white')
                cell.set_linewidth(1.5)

            # 데이터 셀 스타일
            for i in range(len(table_data)):
                # 배경색 교대로 설정
                bg_color = '#f5f5f5' if i % 2 == 0 else 'white'

                for j in range(len(headers)):
                    cell = table[(i+1, j)]

                    # 첫 번째 열(아파트명)은 아파트 색상으로 표시
                    if j == 0:
                        cell.set_facecolor(cagr_info[i]['color'])
                        cell.set_text_props(weight='bold', color='white', fontsize=10)
                    # 두 번째 열(종류)은 매매/전세 구분하여 눈에 띄는 색상
                    elif j == 1:
                        trade_type = table_data[i][1]
                        if trade_type == '매매':
                            cell.set_facecolor('#E74C3C')  # 빨간색 (매매)
                            cell.set_text_props(weight='bold', color='white', fontsize=11)
                        else:  # 전세
                            cell.set_facecolor('#3498DB')  # 파란색 (전세)
                            cell.set_text_props(weight='bold', color='white', fontsize=11)
                    else:
                        cell.set_facecolor(bg_color)
                        cell.set_text_props(color='black', fontsize=9.5)

                    cell.set_edgecolor('#cccccc')
                    cell.set_linewidth(0.5)

            # 테이블 제목
            ax_yield.text(0.5, 1.05, '매매·전세가 및 수익률 정보',
                         transform=ax_yield.transAxes,
                         fontsize=12, fontweight='bold',
                         horizontalalignment='center',
                         verticalalignment='bottom')

        # 단지세부정보가 선택된 경우에만 하단 테이블 표시
        if self.show_complex_info.get():
            # 단지정보 테이블 생성 함수 - 가격 정보 제외
            # create_multi_chart 함수 내에서 단지정보 테이블 생성 함수 수정
            def create_complex_info_table():
                """단지정보 테이블 데이터 생성 - 가격 정보 제외"""
                table_data = []
                # 기본 헤더에서 가격 정보 제외
                headers = ['아파트명']
                
                # 세부정보 가져오기가 선택된 경우만 추가 헤더
                if self.show_complex_info.get() and complex_df is not None:
                    headers.extend(['단지분류', '동수/세대수/최고층', '임대세대', '난방방식', '시공사', '주차대수'])
                
                for i, df in enumerate(apt_dfs):
                    if df.empty:
                        continue
                        
                    apt_name = df['apt_name'].iloc[0]
                    area = df['area'].iloc[0]
                    
                    # 기본 정보만 담은 데이터 구성 (가격 정보 제외)
                    row_data = [
                        f"{apt_name} ({area}㎡)"
                    ]
                    
                    # 세부정보 가져오기가 선택된 경우만 추가 정보 조회
                    if self.show_complex_info.get() and complex_df is not None:
                        # 아파트 정보 전체 전달
                        apt_info = None
                        for apt in self.selected_apts:
                            if apt['apt_name'] == apt_name and str(apt['area']) == str(area):
                                apt_info = apt
                                break
                        
                        if apt_info:
                            complex_info = self.get_complex_info_by_address(complex_df, apt_info)
                            
                            if complex_info:
                                # 동수/세대수/최고층 정보 결합
                                building_info = f"{complex_info.get('n', 0)}동/"
                                building_info += f"{complex_info.get('o', 0)}세대/"
                                building_info += f"{complex_info.get('bo', 0)}층"
                                
                                # 단지분류 정보
                                complex_type = complex_info.get('g', '-')
                                
                                # 임대세대수
                                rental_units = complex_info.get('q', '-')
                                
                                # 난방방식
                                heating_system = complex_info.get('u', '-')
                                
                                # 시공사
                                constructor = complex_info.get('w', '-')
                                
                                # 주차대수 (세대당 주차대수 비율 추가)
                                parking_spots = complex_info.get('ba', '-')
                                total_households = complex_info.get('o', 0)
                                
                                # 주차대수와 세대당 주차대수 비율 계산
                                parking_ratio = "-"
                                if parking_spots != '-' and total_households and int(total_households) > 0:
                                    try:
                                        # 문자열 타입인 경우를 대비해 정수로 변환
                                        parking_spots_num = int(parking_spots)
                                        total_households_num = int(total_households)
                                        if total_households_num > 0:
                                            ratio = parking_spots_num / total_households_num
                                            parking_ratio = f"{parking_spots}(세대당 주차대수 {ratio:.1f})"
                                        else:
                                            parking_ratio = f"{parking_spots}"
                                    except (ValueError, TypeError):
                                        parking_ratio = f"{parking_spots}"
                                else:
                                    parking_ratio = f"{parking_spots}"
                            else:
                                # 단지정보를 찾을 수 없는 경우 기본값
                                building_info = "-/-/-"
                                complex_type = "-"
                                rental_units = "-"
                                heating_system = "-"
                                constructor = "-"
                                parking_ratio = "-"
                        else:
                            # 아파트 정보를 찾을 수 없는 경우 기본값
                            building_info = "-/-/-"
                            complex_type = "-"
                            rental_units = "-"
                            heating_system = "-"
                            constructor = "-"
                            parking_ratio = "-"
                        
                        # 세부정보 추가
                        row_data.extend([
                            complex_type,          # 단지분류
                            building_info,         # 동수/세대수/최고층
                            str(rental_units),     # 임대세대수
                            heating_system,        # 난방방식
                            constructor,           # 시공사
                            parking_ratio          # 주차대수(세대당 주차대수 x.x)
                        ])
                    
                    table_data.append(row_data)
                
                return table_data, headers
    
            # 상세 정보 및 통계 테이블 (단지세부정보가 선택된 경우에만)
            stats_ax = fig.add_subplot(gs[13:18, 0])  # GridSpec의 13-17행 (5행) 사용
            stats_ax.axis('off')
            
            # 단지정보 테이블 데이터 생성
            table_data, headers = create_complex_info_table()
            
            # 테이블 생성
            if table_data:
                table = stats_ax.table(
                    cellText=table_data,
                    colLabels=headers,
                    loc='center',
                    cellLoc='center',
                    colColours=['#f2f2f2'] * len(headers),
                    cellColours=[['#ffffff' if i % 2 == 0 else '#f9f9f9' for j in range(len(headers))] for i in range(len(table_data))]
                )
                
                # 테이블 스타일 조정
                table.auto_set_font_size(False)
                table.set_fontsize(9)
                table.scale(1, 1.5)  # 행 높이 조정
                
                # 열 너비 자동 조정
                for key, cell in table.get_celld().items():
                    cell.set_linewidth(0.5)
                    cell.set_edgecolor('#cccccc')
                    
                    # 열별 너비 설정
                    if key[1] == 0:  # 아파트명
                        cell.set_width(0.18)
                    elif key[1] == 1:  # 단지분류
                        cell.set_width(0.10)
                    elif key[1] == 2:  # 동수/세대수/최고층
                        cell.set_width(0.15)
                    elif key[1] == 3:  # 임대세대
                        cell.set_width(0.10)
                    elif key[1] == 4:  # 난방방식
                        cell.set_width(0.12)
                    elif key[1] == 5:  # 시공사
                        cell.set_width(0.15)
                    elif key[1] == 6:  # 주차대수
                        cell.set_width(0.20)
        
        # 워터마크 추가 (그래프 우측 하단에 위치)
        ax.text(0.99, 0.01,
                f'만든이 부태리 : https://blog.naver.com/landlover333',
                fontsize=12,
                color='#3366CC',
                alpha=0.7,
                transform=ax.transAxes,
                verticalalignment='bottom',
                horizontalalignment='right',
                fontweight='bold',
                bbox=dict(facecolor='white', alpha=0.6, pad=3, boxstyle='round,pad=0.5'),
                zorder=10
               )

        # GridSpec 여백 조정은 이미 hspace=0.3으로 설정됨
        # plt.tight_layout()은 GridSpec과 충돌할 수 있으므로 사용하지 않음
        
        # 파일명 생성 및 이미지 저장
        filename = "multi_apt_comparison_"
        for i, df in enumerate(apt_dfs):
            if not df.empty:
                apt_name = df['apt_name'].iloc[0]
                area = df['area'].iloc[0]
                apt_name_clean = ''.join(char for char in apt_name if char.isalnum() or char.isspace())
                apt_name_clean = apt_name_clean.replace(' ', '_')
                filename += f"{apt_name_clean}_{area}m2"
                
                # 다음 아파트가 있는 경우에만 vs 추가
                if i < len(apt_dfs) - 1 and not apt_dfs[i+1].empty:
                    filename += "_vs_"
        
        # 그래프 옵션을 파일명에 추가
        options = []
        if self.show_monthly_avg.get():
            options.append("avg")
        if self.show_monthly_max.get():
            options.append("max")
        if self.show_scatter_plot.get():
            options.append("scatter")
        if self.show_jeonse.get():
            options.append("jeonse")
            
        if options:
            filename += "_" + "_".join(options)
        
        filename += ".jpg"
        
        self.image_path = os.path.join(self.download_path, filename)
        
        # 이미지 저장
        plt.savefig(self.image_path, bbox_inches='tight', dpi=600, pad_inches=0.3)
        plt.close('all')  # 모든 figure 완전히 닫기
        gc.collect()  # 가비지 컬렉션으로 메모리 정리

        return self.image_path
                    
    def _extract_complex_info(self, row):
        """단지정보 추출 헬퍼 함수"""
        complex_info = {
            'g': row.get('g', '-'),    # 단지분류
            'n': row.get('n', 0),      # 동수
            'o': row.get('o', 0),      # 세대수
            'bo': row.get('bo', 0),    # 최고층수
            'q': row.get('q', '-'),    # 임대세대수
            'u': row.get('u', '-'),    # 난방방식
            'w': row.get('w', '-'),    # 시공사
            'ba': row.get('ba', '-')   # 총 주차대수
        }
        return complex_info

    def _show_complex_selection_dialog(self, matching_df, apt_name, dong):
        """매칭되는 단지가 여러 개일 때 사용자가 선택할 수 있는 다이얼로그 표시"""
        if matching_df.empty:
            return None
        
        # 다이얼로그 생성
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{apt_name} - 단지정보 선택")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 안내 메시지
        ttk.Label(dialog, 
                 text=f"'{apt_name}'의 정확한 단지정보를 찾을 수 없습니다.\n{dong} 지역에서 매칭되는 단지를 선택해주세요.",
                 wraplength=550,
                 font=self.font_normal).pack(pady=10, padx=10)
        
        # 단지 목록 표시
        frame = ttk.Frame(dialog)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 트리뷰 생성
        columns = ("단지명", "주소", "세대수", "동수", "최고층")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)
        
        # 컬럼 설정
        for col in columns:
            tree.heading(col, text=col)
        
        tree.column("단지명", width=150)
        tree.column("주소", width=250)
        tree.column("세대수", width=60, anchor="center")
        tree.column("동수", width=50, anchor="center")
        tree.column("최고층", width=60, anchor="center")
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        
        # 데이터 추가
        for _, row in matching_df.iterrows():
            # 단지명, 주소, 세대수, 동수, 최고층 정보
            complex_name = row.get('a', '정보없음')
            address = row.get('h', '정보없음')
            households = row.get('o', '0')
            buildings = row.get('n', '0')
            max_floor = row.get('bo', '0')
            
            tree.insert("", "end", values=(complex_name, address, households, buildings, max_floor))
        
        # 결과 저장용 변수
        result = [None]  # 리스트로 감싸서 nonlocal 없이 수정 가능하게 함
        
        # 선택 완료 함수
        def on_select():
            selection = tree.selection()
            if selection:
                idx = tree.index(selection[0])
                selected_row = matching_df.iloc[idx]
                selected_row_idx = matching_df.index[idx]
                
                print(f"[디버그] 사용자 선택: 엑셀 {selected_row_idx + 2}행")
                print(f"[디버그] 선택된 단지: {selected_row.get('a', 'N/A')}")
                
                result[0] = selected_row
                dialog.destroy()
            else:
                show_topmost_info("알림", "단지를 선택해주세요.", parent=self.root)
        
        # 취소 함수
        def on_cancel():
            print(f"[디버그] 사용자가 단지 선택 취소")
            result[0] = None
            dialog.destroy()
        
        # 버튼 프레임
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill="x", pady=10)
        
        # 버튼 생성
        ttk.Button(btn_frame, text="선택", command=on_select).pack(side="right", padx=10)
        ttk.Button(btn_frame, text="취소", command=on_cancel).pack(side="right", padx=10)
        
        # 더블 클릭 이벤트
        tree.bind("<Double-1>", lambda e: on_select())
        
        # 다이얼로그가 닫힐 때까지 대기
        self.root.wait_window(dialog)
        
        return result[0]
    
    
    def get_complex_info_by_address(self, complex_df, apt_info):
        """법정동+번지를 통해 단지정보 찾기 - 디버그 정보 추가"""
        if complex_df is None or complex_df.empty:
            print(f"[디버그] 단지정보 엑셀 파일이 없거나 비어있음")
            return None
        
        try:
            # apt_info에서 필요한 정보 추출
            apt_name = apt_info['apt_name']
            dong = apt_info['dong']      # 법정동
            jibun_addr = apt_info.get('jibun_addr', '')  # 지번주소 (없을 경우 빈 문자열)
            
            print(f"\n[디버그] ===== 단지정보 검색 시작 =====")
            print(f"[디버그] 검색 대상 아파트: {apt_name}")
            print(f"[디버그] 법정동: {dong}")
            print(f"[디버그] 지번주소: {jibun_addr}")
            print(f"[디버그] 전체 엑셀 데이터 행 수: {len(complex_df)}")
            
            # 법정동+번지로 단지 찾기
            matching_address = f"{dong} {jibun_addr}" if jibun_addr else dong
            print(f"[디버그] 매칭 시도할 주소: '{matching_address}'")
            
            # 정확한 매칭 먼저 시도 (h열에 법정동+번지가 포함되는 경우)
            exact_matches = complex_df[complex_df['h'].str.contains(matching_address, na=False)]
            print(f"[디버그] 정확한 매칭 결과 행 수: {len(exact_matches)}")
            
            if not exact_matches.empty:
                # 매칭된 행들의 인덱스와 내용 출력
                for idx, (row_idx, row) in enumerate(exact_matches.iterrows()):
                    print(f"[디버그] 정확 매칭 {idx+1}번째 - 엑셀 행: {row_idx + 2} (헤더 포함)")  # +2는 0-based 인덱스 + 헤더행
                    print(f"[디버그]   - 단지명(a열): {row.get('a', 'N/A')}")
                    print(f"[디버그]   - 주소(h열): {row.get('h', 'N/A')}")
                    print(f"[디버그]   - 세대수(o열): {row.get('o', 'N/A')}")
                    print(f"[디버그]   - 동수(n열): {row.get('n', 'N/A')}")
                
                # 정확히 매칭되는 경우 첫 번째 결과 사용
                selected_row = exact_matches.iloc[0]
                selected_row_idx = exact_matches.index[0]
                print(f"[디버그] 선택된 행: {selected_row_idx + 2} (헤더 포함)")
                
                complex_info = self._extract_complex_info(selected_row)
                print(f"[디버그] 추출된 단지정보: {complex_info}")
                return complex_info
            
            # 정확한 매칭이 안되면 법정동만으로 매칭되는 모든 단지 찾기
            print(f"[디버그] 정확한 매칭 실패, 법정동 '{dong}'로 재검색")
            dong_matches = complex_df[complex_df['h'].str.contains(dong, na=False)]
            print(f"[디버그] 법정동 매칭 결과 행 수: {len(dong_matches)}")
            
            if not dong_matches.empty:
                # 법정동으로 매칭되는 단지들 정보 출력
                for idx, (row_idx, row) in enumerate(dong_matches.iterrows()):
                    print(f"[디버그] 법정동 매칭 {idx+1}번째 - 엑셀 행: {row_idx + 2} (헤더 포함)")
                    print(f"[디버그]   - 단지명(a열): {row.get('a', 'N/A')}")
                    print(f"[디버그]   - 주소(h열): {row.get('h', 'N/A')}")
                    print(f"[디버그]   - 세대수(o열): {row.get('o', 'N/A')}")
                    print(f"[디버그]   - 동수(n열): {row.get('n', 'N/A')}")
                
                # 법정동으로 매칭되는 단지가 여러 개 있으면 사용자가 선택하게 함
                selected_row = self._show_complex_selection_dialog(dong_matches, apt_name, dong)
                if selected_row is not None:
                    # 선택된 행의 인덱스 찾기
                    selected_row_idx = dong_matches[dong_matches.eq(selected_row).all(axis=1)].index[0]
                    print(f"[디버그] 사용자가 선택한 행: {selected_row_idx + 2} (헤더 포함)")
                    print(f"[디버그] 선택된 단지명: {selected_row.get('a', 'N/A')}")
                    
                    complex_info = self._extract_complex_info(selected_row)
                    print(f"[디버그] 추출된 단지정보: {complex_info}")
                    return complex_info
                else:
                    print(f"[디버그] 사용자가 단지 선택을 취소함")
            
            # 매칭되는 정보가 없는 경우
            print(f"[디버그] 매칭되는 단지정보 없음")
            print(f"[디버그] ===== 단지정보 검색 종료 =====\n")
            return None
            
        except Exception as e:
            print(f"[디버그] 단지정보 찾기 중 오류: {str(e)}")
            import traceback
            print(f"[디버그] 상세 오류:\n{traceback.format_exc()}")
            return None



    def show_apt_list(self):
        """아파트 목록 조회 및 표시 (구 선택 시 모든 동 조회 기능 추가)"""
        # 버튼 비활성화 및 텍스트 변경
        self.apt_list_button.config(state='disabled', text='🔄 아파트 목록 조회 중...')

        # 즉시 로딩 표시
        self.update_progress(5, "🔄 아파트 목록 조회를 준비하고 있습니다...")
        self.root.update_idletasks()  # UI 즉시 업데이트

        sido = self.sido_combobox.get()
        sigungu = self.sigungu_combobox.get()
        dong = self.dong_combobox.get()
        original_dong = dong

        print(f"\n{'='*80}")
        print(f"=== show_apt_list 호출 ===")
        print(f"sido: '{sido}'")
        print(f"sigungu: '{sigungu}'")
        print(f"dong: '{dong}'")
        print(f"{'='*80}\n")

        # 하위 동 처리 (  └ 로 시작하는 경우)
        parent_gu = None
        if dong.startswith("  └ "):
            dong = dong.replace("  └ ", "").strip()
            # 상위 구 찾기
            dong_list = self.dong_combobox['values']
            for i, item in enumerate(dong_list):
                if item == original_dong:
                    for j in range(i-1, -1, -1):
                        if not dong_list[j].startswith("  └ ") and dong_list[j].endswith('구'):
                            parent_gu = dong_list[j]
                            break
                    break

        if "선택" in [sido, sigungu, dong]:
            self.update_progress(0, "")
            # 버튼 다시 활성화
            self.apt_list_button.config(state='normal', text='🏢 아파트 목록 조회')
            show_topmost_error("오류", "지역을 모두 선택해주세요.", parent=self.root)
            return

        # 구 선택 시 처리 - 하위 모든 동의 아파트 조회
        if dong.endswith('구'):
            print(f">>> ✅ '{dong}' 구 선택됨 - 하위 모든 동 조회 시작")
            self.update_progress(10, f"📍 {dong} 하위 동 목록을 확인하는 중...")

            # 구 코드 찾기
            sigungu_code_to_use = None
            if hasattr(self, 'gu_info'):
                if sigungu in self.sigungu_to_full_info:
                    _, original_si, _ = self.sigungu_to_full_info[sigungu]
                else:
                    original_si = sigungu.replace('(경)', '').replace('(충)', '').replace('(전)', '').strip()
                gu_key = f"{sido}_{original_si}_{dong}"
                print(f"gu_key 조회: {gu_key}")
                if gu_key in self.gu_info:
                    sigungu_code_to_use = self.gu_info[gu_key]
                    print(f"구 코드 찾음: {sigungu_code_to_use}")

            if not sigungu_code_to_use:
                self.update_progress(0, "")
                # 버튼 다시 활성화
                self.apt_list_button.config(state='normal', text='🏢 아파트 목록 조회')
                show_topmost_error("오류", f"{dong}의 코드를 찾을 수 없습니다.", parent=self.root)
                return

            # 해당 구 하위의 모든 동 목록 가져오기
            gu_key = f"{sigungu}_{dong}"
            if gu_key in self.dong_dict:
                sub_dongs = [d.replace("  └ ", "").strip() for d in self.dong_dict[gu_key]]
                print(f"'{dong}' 구 하위 동 목록: {sub_dongs}")

                if not sub_dongs:
                    self.update_progress(0, "")
                    # 버튼 다시 활성화
                    self.apt_list_button.config(state='normal', text='🏢 아파트 목록 조회')
                    show_topmost_info("알림", f"{dong} 하위에 동이 없습니다.", parent=self.root)
                    return

                # 모든 하위 동의 아파트 목록 수집
                all_apts = {}

                for idx, sub_dong in enumerate(sub_dongs):
                    # 각 동별 진행률 계산
                    base_progress = int((idx / len(sub_dongs)) * 100)

                    # 캐시 확인
                    cache_key = (sigungu_code_to_use, sub_dong)
                    if cache_key in self.apt_list_cache:
                        apt_list = self.apt_list_cache[cache_key]
                        print(f"💾 캐시에서 로드됨: {sub_dong} ({len(apt_list)}개)")
                        self.update_progress(base_progress, f"💾 {sub_dong} 캐시에서 로드 중... ({idx+1}/{len(sub_dongs)})")
                    else:
                        def sub_dong_progress(progress, msg):
                            # 동 내부의 API 호출 진행률을 전체 진행률에 반영
                            total_progress = base_progress + int((progress / 100) * (100 / len(sub_dongs)))
                            self.update_progress(total_progress, msg)

                        try:
                            apt_list = self.get_apt_list_from_api(sigungu_code_to_use, sub_dong, progress_callback=sub_dong_progress)
                        except Exception as e:
                            print(f"{sub_dong} 조회 중 오류: {str(e)}")
                            continue

                    # 아파트 목록 추가
                    if apt_list:
                        for apt in apt_list:
                            # 동 정보를 포함하여 중복 제거
                            all_apts[f"{sub_dong}|{apt}"] = (sub_dong, apt)

                    if cache_key not in self.apt_list_cache:
                        time.sleep(0.3)  # API 호출 시에만 간격 조절

                self.update_progress(100, "✅ 검색 완료")

                if all_apts:
                    # 아파트 목록을 동별로 정리하여 표시
                    formatted_list = [f"[{sub_dong}] {apt}" for sub_dong, apt in all_apts.values()]

                    dialog = AptSelectDialog(
                        self,  # self.root 대신 self(앱 인스턴스) 전달
                        formatted_list,
                        self.service_key,
                        sigungu_code_to_use,
                        dong,  # 구 이름 전달
                        sido,
                        sigungu,
                        title=f"{dong} 전체 아파트 목록 ({len(formatted_list)}개)"
                    )
                    self.root.wait_window(dialog.top)

                    # 다이얼로그 내부에서 이미 add_apt_to_selection 호출됨
                    # 상태 메시지만 업데이트
                    if self.selected_apts:
                        self.status_label.config(text="아파트가 선택되었습니다. 추가로 아파트를 선택하거나 그래프 생성 버튼을 눌러 분석을 시작하세요.")
                else:
                    self.update_progress(0, "")
                    show_topmost_info("알림", f"{dong} 하위 동들에 거래 내역이 있는 아파트가 없습니다.", parent=self.root)
            else:
                self.update_progress(0, "")
                show_topmost_error("오류", f"{dong} 하위 동 목록을 찾을 수 없습니다.", parent=self.root)

            self.update_progress(0, "")
            # 버튼 다시 활성화
            self.apt_list_button.config(state='normal', text='🏢 아파트 목록 조회')
            return

        # 일반 동 선택 처리
        sigungu_code_to_use = None

        # parent_gu가 있는 경우 (하위 동)
        if parent_gu:
            if hasattr(self, 'gu_info'):
                # 원래 시 이름 추출 (괄호 제거)
                if sigungu in self.sigungu_to_full_info:
                    _, original_si, _ = self.sigungu_to_full_info[sigungu]
                else:
                    original_si = sigungu.replace('(경)', '').replace('(충)', '').replace('(전)', '').strip()

                # gu_info에서 코드 찾기
                gu_key = f"{sido}_{original_si}_{parent_gu}"
                if gu_key in self.gu_info:
                    sigungu_code_to_use = self.gu_info[gu_key]

            # gu_info에서 못 찾았으면 region_codes에서 찾기
            if not sigungu_code_to_use:
                # 원래 시 이름으로 region_key 생성 (괄호 제거된 버전)
                if sigungu in self.sigungu_to_full_info:
                    _, original_si, _ = self.sigungu_to_full_info[sigungu]
                    region_key = f"{original_si}_{parent_gu}_{dong}"
                else:
                    original_si = sigungu.replace('(경)', '').replace('(충)', '').replace('(전)', '').strip()
                    region_key = f"{original_si}_{parent_gu}_{dong}"

                if region_key in self.region_codes:
                    _, sigungu_code_to_use = self.region_codes[region_key]
        else:
            # 일반 동인 경우
            print(f"[디버그] 일반 동 코드 조회: sido='{sido}', sigungu='{sigungu}', dong='{dong}'")
            region_code = self.region_codes.get((sido, sigungu, dong))
            if region_code:
                sigungu_code_to_use = region_code[1]
                print(f"[디버그] region_codes에서 찾음: {sigungu_code_to_use}")
            else:
                print(f"[디버그] region_codes에서 못 찾음. 사용 가능한 키 확인:")
                # 해당 시도의 코드만 출력 (너무 많으면 필터링)
                matching_keys = [k for k in self.region_codes.keys() if isinstance(k, tuple) and k[0] == sido and k[2] == dong]
                for key in matching_keys[:5]:  # 최대 5개만 출력
                    print(f"  - {key}")

        if not sigungu_code_to_use:
            print(f"[디버그] sigungu_to_full_info에서 조회 시도: sigungu='{sigungu}'")
            if sigungu in self.sigungu_to_full_info:
                _, _, sigungu_code_to_use = self.sigungu_to_full_info[sigungu]
                print(f"[디버그] sigungu_to_full_info에서 찾음: {sigungu_code_to_use}")
            else:
                print(f"[디버그] 코드를 찾을 수 없음!")
                print(f"[디버그] 사용 가능한 sigungu_to_full_info 키:")
                for key in list(self.sigungu_to_full_info.keys())[:10]:  # 최대 10개만
                    print(f"  - '{key}'")
                self.update_progress(0, "")
                # 버튼 다시 활성화
                self.apt_list_button.config(state='normal', text='🏢 아파트 목록 조회')
                show_topmost_error("오류", "해당 지역의 코드를 찾을 수 없습니다.", parent=self.root)
                return

        try:
            # 캐시 확인
            cache_key = (sigungu_code_to_use, dong)
            if cache_key in self.apt_list_cache:
                apt_list = self.apt_list_cache[cache_key]
                print(f"💾 캐시에서 로드됨: {cache_key} ({len(apt_list)}개)")
                self.update_progress(100, f"💾 캐시에서 로드 완료 ({len(apt_list)}개)")
            else:
                self.update_progress(10, "🔍 아파트 목록 검색 준비 중...")

                # 진행 상황 콜백 함수
                def progress_callback(progress, msg):
                    self.update_progress(10 + int(progress * 0.8), msg)

                apt_list = self.get_apt_list_from_api(sigungu_code_to_use, dong, progress_callback=progress_callback)

                self.update_progress(100, "✅ 검색 완료")

            if apt_list:
                # 여기서 sido와 sigungu 값을 함께 전달
                dialog = AptSelectDialog(
                    self,  # self.root 대신 self(앱 인스턴스) 전달
                    apt_list,
                    self.service_key,
                    sigungu_code_to_use,
                    dong,
                    sido,    # sido 값 전달
                    sigungu, # sigungu 값 전달
                    title=f"{dong} 아파트 목록"
                )
                self.root.wait_window(dialog.top)

                # 다이얼로그 내부에서 이미 add_apt_to_selection 호출됨
                # 상태 메시지만 업데이트
                if self.selected_apts:
                    self.status_label.config(text="아파트가 선택되었습니다. 추가로 아파트를 선택하거나 그래프 생성 버튼을 눌러 분석을 시작하세요.")
                    
            else:
                self.update_progress(0, "")
                show_topmost_info("알림", f"{dong}에 거래 내역이 있는 아파트가 없습니다.", parent=self.root)
        except Exception as e:
            self.update_progress(0, "")
            show_topmost_error("오류", f"아파트 목록 조회 중 오류가 발생했습니다: {str(e)}", parent=self.root)
        finally:
            self.update_progress(0, "")
            # 버튼 다시 활성화 및 텍스트 복원
            self.apt_list_button.config(state='normal', text='🏢 아파트 목록 조회')
    
    def analyze_and_visualize(self, trades):
        """실거래 데이터 분석 및 시각화"""
        try:
            self.update_progress(50, "실거래 데이터 분석 중...")
            
            if not trades:
                show_topmost_info("알림", "해당 아파트의 실거래 데이터가 없습니다.", parent=self.root)
                return
            
            # 거래 데이터 저장 (체크박스 업데이트 함수에서 사용)
            self.trades_data = trades
            
            # 데이터 정렬
            trades = sorted(trades, key=lambda x: x['date'])
            
            # 데이터프레임 생성
            df = pd.DataFrame(trades)
            
            # 평균가 추가 (ME: Month End 사용)
            df_monthly = df.groupby(pd.Grouper(key='date', freq='ME')).agg({'price': 'mean'}).reset_index()
            
            # 차트 생성
            self.create_chart(df, df_monthly)
            
            # 분석 데이터 저장
            self.save_analysis_result(df)
            
            self.update_progress(100, "분석 완료!")
            
            # 차트 표시
            if os.path.exists(self.image_path):
                os.startfile(self.image_path)
            
            # 히스토리 갱신
            self.history_list = self.load_history()
            self.update_history_display()
            
        except Exception as e:
            show_topmost_error("오류", f"분석 중 오류 발생: {str(e)}", parent=self.root)
            self.update_progress(0, "")
        
    def create_chart(self, df, df_monthly):
        """거래 데이터 차트 생성 - 사용자 선택에 따른 그래프 표시"""
        plt.figure(figsize=(16, 9))
        
        # 메인 그래프 영역
        ax = plt.subplot2grid((10, 1), (0, 0), rowspan=7)
        
        # 월 단위로 데이터 그룹화하여 평균 및 최대값 계산
        # 날짜 기준으로 정렬
        df_sorted = df.sort_values('date')
        
        # 월별 고유 값 추출(년-월 형식)
        df_sorted['year_month'] = df_sorted['date'].dt.strftime('%Y-%m')
        
        # 월별 데이터 계산
        monthly_data = df_sorted.groupby('year_month').agg({
            'date': 'first',  # 각 월의 첫 날짜
            'price': ['mean', 'max']  # 평균가와 최고가
        }).reset_index()
        
        # 컬럼 이름 재설정
        monthly_data.columns = ['year_month', 'date', 'avg_price', 'max_price']
        
        # 날짜 순서로 정렬
        monthly_data = monthly_data.sort_values('date')

        # 차트 설정 - 준공연도 정보 추가
        apt_info = self.selected_apt_info
        
        # 준공연도 정보가 있으면 표시
        build_year_text = ""
        if 'build_year' in apt_info and apt_info['build_year']:
            current_year = datetime.now().year
            age = current_year - int(apt_info['build_year'])
            build_year_text = f", {apt_info['build_year']}년 준공({age}년차)"
        
        chart_title = f"{apt_info['apt_name']} ({apt_info['area']}㎡{build_year_text}) 실거래가 분석\n{apt_info['sido']} {apt_info['sigungu']} {apt_info['dong']}"
        ax.set_title(chart_title, pad=20, fontsize=18, fontweight='bold')
        
        
        # 점도표 (실거래가) - 점도표 체크박스가 선택된 경우만 표시
        if self.show_scatter_plot.get():
            print("점도표 체크박스 선택됨 - 점도표 그리기")
            high_floor_color = '#69B1FF'  # 5층 이상 (파란색)
            low_floor_color = '#C0C0C0'   # 5층 미만 (회색)
            
            high_trades = df[df['floor'] > 4]
            low_trades = df[df['floor'] <= 4]
            
            if not high_trades.empty:
                ax.scatter(high_trades['date'], high_trades['price'], 
                           color=high_floor_color, alpha=0.6, s=50, 
                           label='실거래(5층↑)', zorder=3)
            
            if not low_trades.empty:
                ax.scatter(low_trades['date'], low_trades['price'], 
                           color=low_floor_color, alpha=0.6, s=50, 
                           label='실거래(4층↓)', zorder=3)
        else:
            print("점도표 체크박스 미선택 - 점도표 그리지 않음")
        
        # 월 평균가격 체크박스가 선택된 경우
        if self.show_monthly_avg.get():
            print("월 평균가격 체크박스 선택됨 - 평균가 그래프 그리기")
            if len(monthly_data) > 0:
                ax.plot(monthly_data['date'], monthly_data['avg_price'], 
                        color='#0066CC', linewidth=3, label='월별 평균가', zorder=4)
        else:
            print("월 평균가격 체크박스 미선택 - 평균가 그래프 그리지 않음")
        
        # 월 최고가 체크박스가 선택된 경우
        if self.show_monthly_max.get():
            print("월 최고가 체크박스 선택됨 - 최고가 그래프 그리기") 
            if len(monthly_data) > 0:
                ax.plot(monthly_data['date'], monthly_data['max_price'], 
                        color='#FF5050', linewidth=2.5, linestyle='--', 
                        label='월별 최고가', zorder=4)
        else:
            print("월 최고가 체크박스 미선택 - 최고가 그래프 그리지 않음")
        
        # 나머지 코드는 동일...
        
        # 절대 최고가, 최저가, 현재가 표시
        max_trade = df.loc[df['price'].idxmax()]
        min_trade = df.loc[df['price'].idxmin()]
        latest_trade = df.iloc[-1]
        
        # 절대 최고가 표시
        ax.scatter([max_trade['date']], [max_trade['price']], 
                   color='white', edgecolor='red', s=100, linewidth=2, zorder=5)
        ax.annotate(f"{max_trade['price']:,}만원\n(최고가)",
                    xy=(max_trade['date'], max_trade['price']),
                    xytext=(10, 10), textcoords='offset points',
                    bbox=dict(boxstyle='round,pad=0.5', fc='white', alpha=0.8),
                    fontsize=10, zorder=6)
        
        # 절대 최저가 표시
        ax.scatter([min_trade['date']], [min_trade['price']], 
                   color='white', edgecolor='blue', s=100, linewidth=2, zorder=5)
        ax.annotate(f"{min_trade['price']:,}만원\n(최저가)",
                    xy=(min_trade['date'], min_trade['price']),
                    xytext=(-70, -30), textcoords='offset points',
                    bbox=dict(boxstyle='round,pad=0.5', fc='white', alpha=0.8),
                    fontsize=10, zorder=6)
        
        # 최근 거래가 표시
        ax.scatter([latest_trade['date']], [latest_trade['price']], 
                   color='white', edgecolor='green', s=100, linewidth=2, zorder=5)
        ax.annotate(f"{latest_trade['price']:,}만원\n(최근거래)",
                    xy=(latest_trade['date'], latest_trade['price']),
                    xytext=(10, -30), textcoords='offset points',
                    bbox=dict(boxstyle='round,pad=0.5', fc='white', alpha=0.8),
                    fontsize=10, zorder=6)
        
        # 차트 설정
        apt_info = self.selected_apt_info
        chart_title = f"{apt_info['apt_name']} ({apt_info['area']}㎡) 실거래가 분석\n{apt_info['sido']} {apt_info['sigungu']} {apt_info['dong']}"
        ax.set_title(chart_title, pad=20, fontsize=18, fontweight='bold')
        
        ax.set_xlabel('거래 일자', fontsize=12)
        ax.set_ylabel('가격(만원)', fontsize=12)
        
        # x축 날짜 형식 설정
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.setp(ax.get_xticklabels(), rotation=-45, ha='center')
        
        # y축 숫자 형식 설정 (천 단위 콤마)
        import matplotlib.ticker as ticker
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
        
        # 그리드 설정
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # 범례 설정
        ax.legend(loc='upper left', bbox_to_anchor=(0.02, 0.98), fontsize=10)
        
        # 정보 박스 영역 (하단)
        info_ax = plt.subplot2grid((10, 1), (7, 0), rowspan=3)
        info_ax.axis('off')
        
        # 통계 정보 계산
        avg_price = df['price'].mean()
        med_price = df['price'].median()
        price_range = max_trade['price'] - min_trade['price']
        price_range_percent = (price_range / min_trade['price']) * 100
        
        # 첫 거래와 마지막 거래 사이 가격 변화율
        first_trade = df.iloc[0]
        last_trade = df.iloc[-1]
        price_change = last_trade['price'] - first_trade['price']
        price_change_percent = (price_change / first_trade['price']) * 100
        
        # 연복리 계산
        years = (last_trade['date'] - first_trade['date']).days / 365.25
        cagr = (((last_trade['price'] / first_trade['price'])) ** (1/years) - 1) * 100 if years > 0 else 0
        
        # 월별 최고가 상승률 계산
        monthly_max_change = 0
        monthly_max_change_percent = 0
        
        if len(monthly_data) > 1:
            first_max = monthly_data['max_price'].iloc[0]
            last_max = monthly_data['max_price'].iloc[-1]
            monthly_max_change = last_max - first_max
            monthly_max_change_percent = (monthly_max_change / first_max) * 100 if first_max > 0 else 0
        
        # 정보 텍스트 생성 (월별 최고가 정보 추가)
        info_text = (
            f"거래 통계 정보\n\n"
            f"총 거래건수: {len(df)}건\n"
            f"거래기간: {first_trade['date'].strftime('%Y-%m-%d')} ~ {last_trade['date'].strftime('%Y-%m-%d')}\n\n"
            f"평균가: {avg_price:,.0f}만원\n"
            f"중간값: {med_price:,.0f}만원\n"
            f"최고가: {max_trade['price']:,.0f}만원 ({max_trade['date'].strftime('%Y-%m-%d')})\n"
            f"최저가: {min_trade['price']:,.0f}만원 ({min_trade['date'].strftime('%Y-%m-%d')})\n\n"
            f"가격변동폭: {price_range:,.0f}만원 ({price_range_percent:.1f}%)\n"
            f"전체기간 가격변화: {price_change:,.0f}만원 ({price_change_percent:.1f}%)\n"
            f"연복리 수익률: {cagr:.2f}%\n\n"
            f"월별 최고가 변화: {monthly_max_change:,.0f}만원 ({monthly_max_change_percent:.1f}%)"
        )
        
        # 정보 텍스트 표시
        text_style = dict(
            ha='center', 
            va='center',
            fontsize=11,
            fontweight='bold',
            linespacing=1.5,
            bbox=dict(
                facecolor='white',
                edgecolor='#6366F1',
                linewidth=2,
                alpha=0.9,
                boxstyle='round4,pad=0.6,rounding_size=0.2'
            )
        )
        
        info_ax.text(0.5, 0.5, info_text, transform=info_ax.transAxes, **text_style)
        
        # 워터마크 추가
        ax.text(0.02, 0.02,
                '국토부 실거래가 분석 시각화 도구',
                fontsize=8,
                color='gray',
                alpha=0.6,
                transform=ax.transAxes,
                verticalalignment='bottom')
        
        # 여백 조정 - 타이트 레이아웃 경고 해결을 위해 수정
        # 여백 조정 - 더 넓은 여백 확보
        plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.2, hspace=0.5)  # bottom과 hspace 값 증가
        
        # 그래프와 x축 사이 여백 추가
        ax.tick_params(axis='x', pad=10)  # x축 라벨과 축 사이 간격 키우기
        
        # x축 레이블 위치 조정
        ax.xaxis.set_label_coords(0.5, -0.15)  # x축 레이블 위치 아래로 이동
        
        # 파일명 생성
        apt_name_clean = ''.join(char for char in apt_info['apt_name'] if char.isalnum() or char.isspace())
        apt_name_clean = apt_name_clean.replace(' ', '_')
        
        filename = f"{apt_name_clean}_{apt_info['area']}m2.jpg"
        self.image_path = os.path.join(self.download_path, filename)
        
        # 이미지 저장
        plt.savefig(self.image_path, bbox_inches='tight', dpi=600, pad_inches=0.3)
        plt.close('all')  # 모든 figure 완전히 닫기
        gc.collect()  # 가비지 컬렉션으로 메모리 정리

        return self.image_path
    
    def save_analysis_result(self, df):
        """분석 결과 저장"""
        try:
            # 기본 정보 및 결과 텍스트 생성
            apt_info = self.selected_apt_info
            max_trade = df.loc[df['price'].idxmax()]
            
            # 엑셀 파일에 저장
            self.save_analysis_result_to_excel(df)
            
            # 히스토리에 추가
            history_filename = f"history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            history_path = os.path.join(self.history_path, history_filename)
            
            # 히스토리용 파일 복사
            if os.path.exists(self.excel_path):
                shutil.copy2(self.excel_path, history_path)
            
            # 히스토리 리스트 갱신
            self.history_list = self.load_history()
            self.update_history_display()
            
        except Exception as e:
            print(f"분석 결과 저장 중 오류: {str(e)}")
    
    # 6. 엑셀 저장 함수에 준공연도 정보 추가
    def save_analysis_result_to_excel(self, df, apt_info):
        """분석 결과를 엑셀 파일로 저장"""
        try:
            apt_name = apt_info['apt_name']
            area = apt_info['area']
            
            apt_name_clean = ''.join(char for char in apt_name if char.isalnum() or char.isspace())
            apt_name_clean = apt_name_clean.replace(' ', '_')
            
            filename = f"{apt_name_clean}_{area}m2.xlsx"
            excel_path = os.path.join(self.download_path, filename)
            
            # 데이터프레임을 엑셀로 저장
            writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
            
            # 기본 정보 시트 생성 - 준공연도 추가
            build_year_info = "정보없음"
            if 'build_year' in apt_info and apt_info['build_year']:
                current_year = datetime.now().year
                age = current_year - int(apt_info['build_year'])
                build_year_info = f"{apt_info['build_year']}년 준공 ({age}년차)"
                
            info_df = pd.DataFrame([
                ["단지 정보", ""],
                ["단지명", apt_name],
                ["지역", f"{apt_info['sido']} {apt_info['sigungu']} {apt_info['dong']}"],
                ["전용면적", f"{area}㎡"],
                ["준공연도", build_year_info],  # 준공연도 정보 추가
                ["거래건수", len(df)],
                ["최고거래가", f"{df['price'].max():,.0f}만원" if not df.empty else "정보없음"],
                ["최저거래가", f"{df['price'].min():,.0f}만원" if not df.empty else "정보없음"]
            ])
            
            info_df.to_excel(writer, sheet_name='기본정보', header=False, index=False)
            
            
            # 거래 데이터 시트 생성
            if not df.empty:
                trade_df = df.copy()
                trade_df['date'] = trade_df['date'].dt.strftime('%Y-%m-%d')
                trade_df.columns = ['거래일자', '가격(만원)', '층', '면적(㎡)']
                trade_df = trade_df.sort_values('거래일자', ascending=False)
                
                trade_df.to_excel(writer, sheet_name='거래내역', index=False)
            else:
                # 빈 데이터프레임 저장
                empty_df = pd.DataFrame(columns=['거래일자', '가격(만원)', '층', '면적(㎡)'])
                empty_df.to_excel(writer, sheet_name='거래내역', index=False)
            
            # 엑셀 파일 저장
            writer.close()
            
            return excel_path
            
        except Exception as e:
            print(f"엑셀 저장 중 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return None


class AptSelectDialog:
    def __init__(self, parent, apt_list, service_key, sigungu_code, dong, sido, sigungu, title="아파트 선택"):
        # parent가 RealEstateAnalyzerApp 인스턴스인 경우
        if hasattr(parent, 'root'):
            self.app = parent  # RealEstateAnalyzerApp 인스턴스
            self.parent = parent.root  # Tk 객체
        else:
            self.parent = parent  # Tk 객체
            self.app = parent

        self.service_key = service_key
        self.sigungu_code = sigungu_code
        self.dong = dong
        self.sido = sido  # 시도 값 직접 전달
        self.sigungu = sigungu  # 시군구 값 직접 전달
        self.apt_list = apt_list
        self.result = None
        self.selected_apt = None

        self.top = tk.Toplevel(self.parent)
        self.top.title(title)
        self.top.attributes('-topmost', True)
        
        # 창 크기와 위치 설정
        width = 800
        height = 500
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.top.geometry(f"{width}x{height}+{x}+{y}")
        
        # 부모의 폰트 가져오기
        if hasattr(self.app, 'font_normal'):
            self.font_normal = self.app.font_normal
            self.font_large = self.app.font_large
            self.font_button = self.app.font_button
        else:
            self.font_normal = ('Malgun Gothic', 9)
            self.font_large = ('Malgun Gothic', 11)
            self.font_button = ('Malgun Gothic', 9)
        
        # 검색창 프레임
        search_frame = ttk.Frame(self.top, padding="5")
        search_frame.pack(fill='x', padx=5, pady=5)
        
        # 폰트 적용
        ttk.Label(search_frame, text=f"{self.dong} 아파트 목록", 
                 font=self.font_large).pack(side='left')
        
        # 검색 입력 필드
        ttk.Label(search_frame, text="검색:", font=self.font_normal).pack(side='left', padx=(20, 0))
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_apartments)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40,
                               font=self.font_normal)
        search_entry.pack(side='left', fill='x', expand=True, padx=5)
        
        # 리스트박스 프레임
        list_frame = ttk.Frame(self.top, padding="5")
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        # 리스트박스에 폰트 적용
        self.listbox = tk.Listbox(list_frame, 
                                yscrollcommand=scrollbar.set,
                                font=self.font_normal)
        self.listbox.pack(fill='both', expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # 아파트 목록 초기화
        self.update_listbox(apt_list)
        
        # 선택 버튼 프레임
        button_frame = ttk.Frame(self.top, padding="5")
        button_frame.pack(fill='x', padx=5, pady=5)
        
        # 선택 버튼
        select_button = ttk.Button(button_frame, text="선택", command=self.on_button_select)
        select_button.pack(side='right', padx=5)
        
        cancel_button = ttk.Button(button_frame, text="취소", command=self.top.destroy)
        cancel_button.pack(side='right', padx=5)
        
        # 더블 클릭 이벤트
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
            show_topmost_info("알림", "아파트를 선택해주세요.", parent=self.parent)
            return
        self.on_select(None)
        
    # 2. AptSelectDialog 클래스의 on_select 메서드 수정
    def on_select(self, event):
        if self.listbox.curselection():
            full_text = self.listbox.get(self.listbox.curselection())

            # [동명] 형식으로 시작하는 경우 처리 (구 선택 시)
            actual_dong = None
            if full_text.startswith('[') and ']' in full_text:
                # 첫 번째 [...] 는 동 이름
                first_bracket_end = full_text.find(']')
                actual_dong = full_text[1:first_bracket_end].strip()
                # 동 이름 제거
                remaining_text = full_text[first_bracket_end+1:].strip()
            else:
                remaining_text = full_text

            # 이제 remaining_text에서 아파트명과 주소 정보 추출
            # '[' 와 ']' 사이의 주소 정보 추출
            if '[' in remaining_text and ']' in remaining_text:
                address_start = remaining_text.find('[')
                address_end = remaining_text.find(']')
                address_info = remaining_text[address_start+1:address_end]
                apt_name = remaining_text[:address_start].strip()
            else:
                address_info = ""
                apt_name = remaining_text.split('[')[0].strip()

            # 도로명 주소와 지번 주소 분리
            parts = address_info.split(' / ')
            road_addr = parts[0] if len(parts) > 0 else ""
            jibun_addr = parts[1] if len(parts) > 1 else ""  # 지번 주소

            # 지번 주소에서 번지만 추출 (법정동 이름은 이미 dong에 있음)
            jibun_number = jibun_addr.split()[-1] if jibun_addr and len(jibun_addr.split()) > 0 else ""

            # 준공연도 추출
            build_year = ""
            if "(준공:" in full_text:
                build_year_part = full_text.split("(준공:")[1].strip()
                build_year = build_year_part.split("년")[0].strip()

            # 선택한 아파트 정보 저장
            self.selected_apt = apt_name
            self.simple_addr = jibun_number  # 법정동을 제외한 번지만 저장
            self.jibun_addr = jibun_number   # 번지 정보 별도 저장
            self.build_year = build_year     # 준공연도 저장

            # actual_dong이 있으면 dong을 업데이트 (구 선택 시)
            if actual_dong:
                self.dong = actual_dong
                print(f"[디버그] 구 선택 시 실제 동 이름 업데이트: {actual_dong}")

            # 전용면적 목록 다이얼로그 표시
            self.show_area_dialog()
    
    def show_area_dialog(self):
        """전용면적 목록 다이얼로그 표시 (수정된 버전)"""
        # 전용면적 목록 가져오기
        area_list = self.get_areas_for_apt(self.selected_apt)
        
        if not area_list:
            show_topmost_info("알림", "해당 아파트의 전용면적 정보를 찾을 수 없습니다.", parent=self.parent)
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
                show_topmost_info("알림", "전용면적을 선택해주세요.", parent=self.parent)
                return

            selected_area = listbox.get(listbox.curselection())
            area_value = selected_area.replace('㎡', '').strip()

            # 아파트 정보 구성 전 디버그 로그
            print(f"\n[디버그] ===== 아파트 정보 구성 =====")
            print(f"[디버그] apt_name: {self.selected_apt}")
            print(f"[디버그] area: {area_value}")
            print(f"[디버그] sido: {self.sido}")
            print(f"[디버그] sigungu: {self.sigungu}")
            print(f"[디버그] dong: {self.dong}")
            print(f"[디버그] sigungu_code: {self.sigungu_code}")
            print(f"[디버그] build_year: {self.build_year}")
            print(f"[디버그] ==============================\n")

            # 아파트 정보만 구성 (거래 데이터는 나중에 그래프 생성 시 조회)
            apt_info = {
                'apt_name': self.selected_apt,
                'jibun_addr': self.jibun_addr,   # 지번 주소 추가
                'area': area_value,
                'sido': self.sido,
                'sigungu': self.sigungu,
                'dong': self.dong,
                'sigungu_code': self.sigungu_code,
                'build_year': self.build_year    # 준공연도 추가
            }

            # 면적 선택 다이얼로그 닫기
            area_dialog.destroy()

            # 메인 앱에 아파트 추가
            if hasattr(self.app, 'add_apt_to_selection'):
                self.app.add_apt_to_selection(apt_info)

            # 사용자에게 추가 선택 여부 묻기
            response = ask_topmost_yesno(
                "추가 선택",
                f"'{self.selected_apt} ({area_value}㎡)'이(가) 선택되었습니다.\n\n같은 지역에서 다른 아파트를 추가로 선택하시겠습니까?",
                parent=self.top
            )

            if not response:
                # 아니오 선택 시 다이얼로그 닫기
                self.result = apt_info  # 마지막 선택 정보 저장
                self.top.destroy()
            else:
                # 예 선택 시 다이얼로그를 열어둠 (사용자가 다시 선택 가능)
                # 검색창 초기화 및 포커스
                self.search_var.set("")
                self.listbox.focus_set()
        
        # 이벤트 바인딩 - 더블 클릭 수정
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
        """해당 아파트의 전용면적 목록 가져오기 (최적화된 고속 버전)"""
        areas = set()
        current_date = datetime.now()
        
        # 진행 상황을 표시할 창
        progress_window = tk.Toplevel(self.top)
        progress_window.title("데이터 수집 중...")
        progress_window.geometry("300x100")
        progress_window.transient(self.top)
        progress_window.grab_set()
        
        ttk.Label(progress_window, text=f"{apt_name} 전용면적 정보를 수집 중입니다...").pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        progress_bar.pack(fill="x", padx=20, pady=10)
        
        # 취소 플래그
        cancel_flag = [False]
        progress_window.protocol("WM_DELETE_WINDOW", lambda: setattr(cancel_flag, 0, True) or progress_window.destroy())
        
        # 캐시 초기화 - 아파트별 전용면적 캐싱
        if not hasattr(self, 'apt_area_cache'):
            self.apt_area_cache = {}
        
        # 캐시에 데이터가 있으면 바로 반환
        if apt_name in self.apt_area_cache:
            # 곧바로 창 닫기
            progress_bar['value'] = 100
            progress_window.update_idletasks()
            time.sleep(0.3)
            progress_window.destroy()
            return self.apt_area_cache[apt_name]
        
        # 병렬 요청 최적화 설정
        concurrent_requests = 12  # 동시 요청 수 증가
        
        # HTTP 세션 최적화
        session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=concurrent_requests,
            pool_maxsize=concurrent_requests * 2,
            max_retries=1
        )
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        
        # 최근 6개월만 조회하여 속도 향상 (전체 조회할 필요 없음)
        max_months = 6  # 조회 기간 크게 단축 (6개월 데이터면 충분)
        consecutive_empty_months_limit = 2  # 연속 2개월만 비어있어도 조기 종료
        
        # 데이터 수집 함수
        def collect_areas():
            nonlocal areas
            consecutive_empty_months = 0
            
            try:
                # 전체 월 요청 준비 - 병렬처리로 한 번에 요청
                all_requests = []
                for month in range(max_months):
                    search_date = current_date - timedelta(days=30 * month)
                    deal_ymd = search_date.strftime("%Y%m")
                    
                    # 매매 API와 전세 API 모두 요청 (전용면적을 모두 찾기 위해)
                    trade_url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                               f"?serviceKey={self.service_key}"
                               f"&LAWD_CD={self.sigungu_code}"
                               f"&DEAL_YMD={deal_ymd}"
                               f"&numOfRows=1000")
                    
                    rent_url = (f"http://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"
                               f"?serviceKey={self.service_key}"
                               f"&LAWD_CD={self.sigungu_code}"
                               f"&DEAL_YMD={deal_ymd}"
                               f"&numOfRows=1000")
                    
                    all_requests.append((trade_url, 'trade', month))
                    all_requests.append((rent_url, 'rent', month))
                
                # 모든 요청을 병렬로 처리
                with concurrent.futures.ThreadPoolExecutor(max_workers=concurrent_requests) as executor:
                    # 모든 요청 실행
                    future_to_request = {
                        executor.submit(session.get, url, timeout=API_TIMEOUT): (req_type, month_idx)
                        for url, req_type, month_idx in all_requests
                    }
                    
                    # 진행 상태 업데이트
                    total_requests = len(all_requests)
                    processed = 0
                    
                    # 결과 처리
                    for future in concurrent.futures.as_completed(future_to_request):
                        processed += 1
                        req_type, month_idx = future_to_request[future]
                        
                        # 진행 상태 업데이트
                        progress = min(100, (processed / total_requests) * 100)
                        progress_bar['value'] = progress
                        progress_window.update_idletasks()
                        
                        # 이미 충분한 데이터가 모였다면 나머지 요청은 건너뛰기
                        if len(areas) >= 5:
                            continue
                            
                        # 취소 또는 연속 빈 데이터가 많으면 중단
                        if cancel_flag[0] or consecutive_empty_months >= consecutive_empty_months_limit:
                            continue
                        
                        # 응답 처리
                        try:
                            response = future.result()
                            month_areas = set()  # 이번 조회에서 찾은 전용면적
                            
                            if response.status_code == 200:
                                try:
                                    # XML 파싱 최적화 - 빠른 검색 로직 적용
                                    root = ET.fromstring(response.text)
                                    
                                    # 빠른 필터링 - 아파트 이름이 일치하는 항목만 선택
                                    for item in root.findall('.//item'):
                                        item_apt = item.findtext('aptNm', '').strip()
                                        
                                        # 정확한 아파트명 매칭
                                        if item_apt == apt_name:
                                            item_area = float(item.findtext('excluUseAr', '0'))
                                            if item_area > 0:
                                                # 소수점 없는 정수로 변환 (정확한 면적)
                                                area_int = str(int(item_area))
                                                areas.add(area_int)
                                                month_areas.add(area_int)
                                except ET.ParseError:
                                    pass  # XML 파싱 오류 무시
                            
                            # 이번 응답에서 데이터가 있었는지 확인
                            if month_areas:
                                consecutive_empty_months = 0  # 데이터가 있으면 카운터 리셋
                            elif req_type == 'trade' and month_idx not in [m for _, _, m in all_requests if _ == 'rent']:
                                # 매매 데이터가 없고, 해당 월의 전세 데이터도 아직 확인하지 않은 경우
                                consecutive_empty_months += 0.5  # 절반만 카운트
                            elif req_type == 'rent':
                                # 전세 데이터가 없는 경우
                                consecutive_empty_months += 0.5  # 절반만 카운트
                        
                        except Exception as e:
                            # 요청 처리 중 오류 발생 - 무시하고 계속
                            continue
                
                # 데이터 처리 완료
                progress_bar['value'] = 100
                progress_window.update_idletasks()
                
                # 조회 결과를 캐시에 저장
                self.apt_area_cache[apt_name] = sorted(list(areas), key=float)
                
            except Exception as e:
                print(f"전용면적 정보 수집 중 오류: {str(e)}")
            
            # 창 닫기
            time.sleep(0.3)  # 약간의 지연
            if not cancel_flag[0]:
                progress_window.destroy()
        
        # 별도 스레드로 데이터 수집 실행
        import threading
        thread = threading.Thread(target=collect_areas)
        thread.daemon = True
        thread.start()
        
        # 창이 닫힐 때까지 대기
        self.top.wait_window(progress_window)
        
        # 취소되었거나 데이터가 없는 경우 처리
        if cancel_flag[0] or not areas:
            return []
        
        # 캐시에 저장
        if apt_name not in self.apt_area_cache:
            self.apt_area_cache[apt_name] = sorted(list(areas), key=float)
        
        return sorted(list(areas), key=float)


def main():
    app = RealEstateAnalyzerApp()
    app.root.mainloop()


if __name__ == "__main__":
    main()