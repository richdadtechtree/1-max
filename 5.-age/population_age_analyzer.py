import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm
import seaborn as sns
import json
import os
import numpy as np

# 한글 폰트 설정
import matplotlib
matplotlib.rcParams['font.family'] = 'Malgun Gothic'
matplotlib.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# seaborn 한글 폰트 설정
sns.set_theme(font='Malgun Gothic', rc={'axes.unicode_minus': False}, style="whitegrid")
sns.set_palette("husl")

class PopulationAgeAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("연령 및 인구 분석 프로그램")
        self.root.geometry("1920x1000")  # 가로 길이 확대

        # 설정 파일 경로
        self.config_file = "population_age_config.json"

        # 데이터 변수
        self.data_file = None
        self.df = None

        # 지역 선택 변수
        self.selected_sido = None
        self.selected_sigungu = None
        self.selected_dong = None
        self.selected_detail = None  # d열 (상세 동)

        # 비교용 지역 목록 (최대 3개)
        self.comparison_regions = []

        # 설정 로드
        self.load_config()

        # GUI 생성
        self.create_widgets()

        # 데이터 로드 시도
        if self.data_file:
            self.load_data()

    def create_widgets(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 설정 프레임
        setting_frame = ttk.LabelFrame(main_frame, text="설정", padding="10")
        setting_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # 파일 선택 버튼
        ttk.Label(setting_frame, text="데이터 파일:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.file_label = ttk.Label(setting_frame, text=self.data_file or "파일 미선택",
                                    foreground="blue" if self.data_file else "red")
        self.file_label.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Button(setting_frame, text="파일 선택",
                  command=self.select_file).grid(row=0, column=2, padx=5, pady=5)

        # 분석 유형 선택 프레임
        analysis_frame = ttk.LabelFrame(main_frame, text="분석 유형 선택", padding="10")
        analysis_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N), padx=(0, 10))

        self.analysis_type = tk.StringVar(value="avg_age")
        ttk.Radiobutton(analysis_frame, text="동별 평균연령", variable=self.analysis_type,
                       value="avg_age", command=self.on_analysis_type_changed).grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Radiobutton(analysis_frame, text="동별 연령별 인구 분포", variable=self.analysis_type,
                       value="age_distribution", command=self.on_analysis_type_changed).grid(row=1, column=0, sticky=tk.W, pady=5)

        # 지역 선택 프레임
        region_frame = ttk.LabelFrame(main_frame, text="지역 선택", padding="10")
        region_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10), pady=(10, 0))

        # 시/도 선택
        ttk.Label(region_frame, text="시/도:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.sido_combo = ttk.Combobox(region_frame, state="readonly", width=30)
        self.sido_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.sido_combo.bind('<<ComboboxSelected>>', self.on_sido_selected)

        # 시/군/구 선택
        ttk.Label(region_frame, text="시/군/구:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.sigungu_combo = ttk.Combobox(region_frame, state="readonly", width=30)
        self.sigungu_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.sigungu_combo.bind('<<ComboboxSelected>>', self.on_sigungu_selected)

        # 동/읍/면 선택 (c열)
        ttk.Label(region_frame, text="동/읍/면:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.dong_combo = ttk.Combobox(region_frame, state="readonly", width=30)
        self.dong_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.dong_combo.bind('<<ComboboxSelected>>', self.on_dong_selected)

        # 상세 동 선택 (d열)
        ttk.Label(region_frame, text="상세 동:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.detail_combo = ttk.Combobox(region_frame, state="readonly", width=30)
        self.detail_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.detail_combo.bind('<<ComboboxSelected>>', self.on_detail_selected)

        # 설명 라벨
        self.info_label = ttk.Label(region_frame, text="※ 시/도 선택: 해당 시/도의 모든 동 평균연령\n구 선택: 해당 구의 모든 동 평균연령",
                                    foreground="gray", font=("", 9))
        self.info_label.grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=5, pady=(10, 5))

        # 비교 지역 관리 프레임
        comparison_frame = ttk.LabelFrame(region_frame, text="지역 비교 (연령별 인구 분포용)", padding="5")
        comparison_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        ttk.Button(comparison_frame, text="현재 지역 추가", command=self.add_comparison_region).pack(side=tk.LEFT, padx=5)
        ttk.Button(comparison_frame, text="선택 지역 제거", command=self.remove_comparison_region).pack(side=tk.LEFT, padx=5)
        ttk.Button(comparison_frame, text="전체 지역 초기화", command=self.clear_comparison_regions).pack(side=tk.LEFT, padx=5)

        # 비교 지역 목록
        self.comparison_listbox = tk.Listbox(comparison_frame, height=3, font=("", 9))
        self.comparison_listbox.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # 버튼 프레임
        button_frame = ttk.Frame(region_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=(20, 0))

        ttk.Button(button_frame, text="그래프 그리기", command=self.draw_graph,
                  style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="그래프 저장", command=self.save_graph).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="초기화", command=self.reset_selection).pack(side=tk.LEFT, padx=5)

        # 그래프 표시 프레임
        graph_frame = ttk.LabelFrame(main_frame, text="그래프", padding="10")
        graph_frame.grid(row=1, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 그래프 캔버스
        self.figure = plt.Figure(figsize=(10, 7), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, graph_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        region_frame.columnconfigure(1, weight=1)

    def select_file(self):
        """파일 선택 다이얼로그"""
        filename = filedialog.askopenfilename(
            title="CSV 파일 선택",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if filename:
            self.data_file = filename
            self.file_label.config(text=os.path.basename(filename), foreground="blue")
            self.save_config()
            self.load_data()

    def load_data(self):
        """CSV 데이터 로드"""
        try:
            # 여러 인코딩 시도
            encodings = ['cp949', 'euc-kr', 'utf-8']
            for encoding in encodings:
                try:
                    # 헤더가 2번째 행에 있으므로 skiprows=1
                    self.df = pd.read_csv(self.data_file, encoding=encoding, skiprows=1)
                    print(f"데이터 로드 완료 (인코딩: {encoding}): {self.df.shape}")
                    print(f"컬럼: {self.df.columns.tolist()}")
                    break
                except UnicodeDecodeError:
                    continue

            if self.df is None:
                raise Exception("파일을 읽을 수 없습니다.")

            # 컬럼명 정리
            self.df.columns = self.df.columns.str.strip()

            # 필요한 컬럼 확인 및 이름 정규화
            # 예상 컬럼: 시도, 시군구, 동읍면, 인구수, 연령대별 인구
            col_names = self.df.columns.tolist()

            # 데이터 정제 - 빈 행 제거
            self.df = self.df.dropna(how='all')

            # 평균연령 계산 함수 추가
            self.calculate_average_age()

            # 지역 드롭다운 업데이트
            self.update_sido_combo()

            messagebox.showinfo("성공", f"데이터 로드 완료\n행 수: {len(self.df)}")

        except Exception as e:
            messagebox.showerror("오류", f"데이터 로드 중 오류 발생:\n{str(e)}")
            import traceback
            traceback.print_exc()

    def calculate_average_age(self):
        """연령대별 인구로부터 평균연령 계산"""
        try:
            # 컬럼명에서 연령대 컬럼 찾기
            age_cols = []
            for col in self.df.columns:
                if '0~9' in col or '10~19' in col or '20~29' in col or '30~39' in col or '40~49' in col or '50~59' in col or '60' in col:
                    age_cols.append(col)

            if len(age_cols) < 7:
                print(f"연령대 컬럼을 찾을 수 없습니다. 현재 컬럼: {self.df.columns.tolist()}")
                return

            # 각 연령대의 중간값
            age_midpoints = {
                age_cols[0]: 4.5,   # 0~9세
                age_cols[1]: 14.5,  # 10~19세
                age_cols[2]: 24.5,  # 20~29세
                age_cols[3]: 34.5,  # 30~39세
                age_cols[4]: 44.5,  # 40~49세
                age_cols[5]: 54.5,  # 50~59세
                age_cols[6]: 70     # 60세 이상
            }

            # 평균연령 계산
            def calc_avg_age(row):
                try:
                    total_pop = 0
                    weighted_sum = 0

                    for col, midpoint in age_midpoints.items():
                        pop = row[col]
                        if pd.notna(pop):
                            # 쉼표 제거 및 숫자 변환
                            if isinstance(pop, str):
                                pop = float(pop.replace(',', ''))
                            else:
                                pop = float(pop)

                            total_pop += pop
                            weighted_sum += pop * midpoint

                    if total_pop > 0:
                        return weighted_sum / total_pop
                    else:
                        return np.nan
                except:
                    return np.nan

            self.df['평균연령'] = self.df.apply(calc_avg_age, axis=1)
            print("평균연령 계산 완료")

        except Exception as e:
            print(f"평균연령 계산 오류: {e}")
            import traceback
            traceback.print_exc()

    def update_sido_combo(self):
        """시/도 콤보박스 업데이트"""
        if self.df is None or len(self.df) == 0:
            return

        # 첫 번째 컬럼이 시/도
        sido_col = self.df.columns[0]
        sido_list = self.df[sido_col].dropna().unique().tolist()

        # 코드값(-1100000000 같은) 제외
        sido_list = [s for s in sido_list if not (isinstance(s, str) and s.startswith('-'))]
        sido_list = sorted(sido_list)

        self.sido_combo['values'] = sido_list

        # 초기화
        self.sigungu_combo['values'] = []
        self.sigungu_combo.set('')
        self.dong_combo['values'] = []
        self.dong_combo.set('')
        self.detail_combo['values'] = []
        self.detail_combo.set('')

    def on_sido_selected(self, event=None):
        """시/도 선택 시"""
        self.selected_sido = self.sido_combo.get()

        if self.df is None:
            return

        # 시/군/구 목록 업데이트
        sido_col = self.df.columns[0]
        sigungu_col = self.df.columns[1]

        filtered = self.df[self.df[sido_col] == self.selected_sido]
        sigungu_list = filtered[sigungu_col].dropna().unique().tolist()

        # 코드값 제외
        sigungu_list = [s for s in sigungu_list if not (isinstance(s, str) and s.startswith('-'))]
        sigungu_list = sorted(sigungu_list)

        self.sigungu_combo['values'] = sigungu_list
        self.sigungu_combo.set('')
        self.dong_combo['values'] = []
        self.dong_combo.set('')
        self.detail_combo['values'] = []
        self.detail_combo.set('')

        self.selected_sigungu = None
        self.selected_dong = None
        self.selected_detail = None

    def on_sigungu_selected(self, event=None):
        """시/군/구 선택 시"""
        self.selected_sigungu = self.sigungu_combo.get()

        if self.df is None:
            return

        # 동/읍/면 목록 업데이트
        sido_col = self.df.columns[0]
        sigungu_col = self.df.columns[1]
        dong_col = self.df.columns[2]

        filtered = self.df[(self.df[sido_col] == self.selected_sido) &
                          (self.df[sigungu_col] == self.selected_sigungu)]
        dong_list = filtered[dong_col].dropna().unique().tolist()

        # 코드값 제외 및 괄호 안의 코드 제거
        dong_list_clean = []
        for d in dong_list:
            if isinstance(d, str) and not d.startswith('-'):
                # 괄호와 그 내용 제거
                import re
                clean_name = re.sub(r'\([^)]*\)', '', d).strip()
                dong_list_clean.append(clean_name)

        dong_list_clean = sorted(set(dong_list_clean))

        self.dong_combo['values'] = dong_list_clean
        self.dong_combo.set('')
        self.detail_combo['values'] = []
        self.detail_combo.set('')
        self.selected_dong = None
        self.selected_detail = None

    def on_dong_selected(self, event=None):
        """동/읍/면 선택 시"""
        self.selected_dong = self.dong_combo.get()

        if self.df is None:
            return

        # d열(상세 동) 목록 업데이트
        if len(self.df.columns) > 3:  # d열이 있는 경우
            sido_col = self.df.columns[0]
            sigungu_col = self.df.columns[1]
            dong_col = self.df.columns[2]
            detail_col = self.df.columns[3]

            filtered = self.df[(self.df[sido_col] == self.selected_sido) &
                              (self.df[sigungu_col] == self.selected_sigungu)]

            # 동 이름 매칭 (괄호 포함/불포함 모두 체크)
            import re
            for idx, row in filtered.iterrows():
                dong_name = str(row[dong_col])
                clean_name = re.sub(r'\([^)]*\)', '', dong_name).strip()
                if clean_name == self.selected_dong:
                    # 해당 동의 상세 동 목록 가져오기
                    detail_filtered = filtered[filtered[dong_col] == dong_name]
                    detail_list = detail_filtered[detail_col].dropna().unique().tolist()

                    # 코드값 제외 및 괄호 제거
                    detail_list_clean = []
                    for d in detail_list:
                        if isinstance(d, str) and not d.startswith('-'):
                            clean_detail = re.sub(r'\([^)]*\)', '', d).strip()
                            if clean_detail:  # 빈 문자열 제외
                                detail_list_clean.append(clean_detail)

                    if detail_list_clean:
                        detail_list_clean = sorted(set(detail_list_clean))
                        self.detail_combo['values'] = detail_list_clean
                    break

        self.detail_combo.set('')
        self.selected_detail = None

    def on_detail_selected(self, event=None):
        """상세 동 선택 시"""
        self.selected_detail = self.detail_combo.get()

    def add_comparison_region(self):
        """현재 선택된 지역을 비교 목록에 추가"""
        if not self.selected_dong:
            messagebox.showwarning("경고", "동/읍/면을 먼저 선택해주세요.")
            return

        if len(self.comparison_regions) >= 3:
            messagebox.showwarning("경고", "최대 3개 지역까지만 비교할 수 있습니다.")
            return

        # 지역 정보 저장
        region_info = {
            'sido': self.selected_sido,
            'sigungu': self.selected_sigungu,
            'dong': self.selected_dong,
            'detail': self.selected_detail
        }

        # 중복 체크
        for region in self.comparison_regions:
            if (region['sido'] == region_info['sido'] and
                region['sigungu'] == region_info['sigungu'] and
                region['dong'] == region_info['dong'] and
                region['detail'] == region_info['detail']):
                messagebox.showwarning("경고", "이미 추가된 지역입니다.")
                return

        self.comparison_regions.append(region_info)
        self.update_comparison_listbox()

    def remove_comparison_region(self):
        """선택된 지역을 비교 목록에서 제거"""
        selection = self.comparison_listbox.curselection()
        if not selection:
            messagebox.showwarning("경고", "제거할 지역을 선택해주세요.")
            return

        index = selection[0]
        del self.comparison_regions[index]
        self.update_comparison_listbox()

    def clear_comparison_regions(self):
        """모든 비교 지역 초기화"""
        self.comparison_regions = []
        self.update_comparison_listbox()

    def update_comparison_listbox(self):
        """비교 지역 목록 업데이트"""
        self.comparison_listbox.delete(0, tk.END)
        for region in self.comparison_regions:
            parts = [region['sido'], region['sigungu'], region['dong']]
            if region['detail']:
                parts.append(region['detail'])
            display_text = " > ".join(parts)
            self.comparison_listbox.insert(tk.END, display_text)

    def on_analysis_type_changed(self):
        """분석 유형 변경 시"""
        analysis = self.analysis_type.get()

        if analysis == "avg_age":
            self.info_label.config(text="※ 시/도 선택: 해당 시/도의 모든 동 평균연령\n구 선택: 해당 구의 모든 동 평균연령")
        else:
            self.info_label.config(text="※ 특정 동을 선택하여 연령별 인구 분포를 확인하세요\n※ 비교: 최대 3개 지역 추가 가능")

    def reset_selection(self):
        """선택 초기화"""
        self.sido_combo.set('')
        self.sigungu_combo.set('')
        self.sigungu_combo['values'] = []
        self.dong_combo.set('')
        self.dong_combo['values'] = []
        self.detail_combo.set('')
        self.detail_combo['values'] = []

        self.selected_sido = None
        self.selected_sigungu = None
        self.selected_dong = None
        self.selected_detail = None

        self.figure.clf()  # clf()로 모든 axes와 내용 완전히 제거
        self.canvas.draw_idle()

    def draw_graph(self):
        """그래프 그리기"""
        if self.df is None:
            messagebox.showwarning("경고", "데이터 파일을 먼저 선택해주세요.")
            return

        analysis = self.analysis_type.get()

        if analysis == "avg_age":
            self.draw_avg_age_graph()
        else:
            self.draw_age_distribution_graph()

    def draw_avg_age_graph(self):
        """동별 평균연령 그래프"""
        if not self.selected_sido:
            messagebox.showwarning("경고", "최소한 시/도를 선택해주세요.")
            return

        sido_col = self.df.columns[0]
        sigungu_col = self.df.columns[1]
        dong_col = self.df.columns[2]

        # d열 존재 여부 확인
        has_detail_col = len(self.df.columns) > 3
        if has_detail_col:
            detail_col = self.df.columns[3]

        import re

        # 선택 레벨에 따라 다르게 처리
        filtered = self.df.copy()
        title = ""
        xlabel = ""
        display_col = None

        if self.selected_dong:
            # 동/읍/면까지 선택: 상세 동의 평균연령
            if has_detail_col:
                filtered = filtered[(filtered[sido_col] == self.selected_sido) &
                                   (filtered[sigungu_col] == self.selected_sigungu)]

                # 해당 동에 속한 상세 동들 찾기
                dong_filtered = []
                for _, row in filtered.iterrows():
                    dong_name = str(row[dong_col])
                    clean_dong = re.sub(r'\([^)]*\)', '', dong_name).strip()
                    if clean_dong == self.selected_dong:
                        dong_filtered.append(row)

                if dong_filtered:
                    filtered = pd.DataFrame(dong_filtered)
                    # 상세 동만 선택 (d열에 값이 있는 것)
                    filtered = filtered[filtered[detail_col].notna()]
                    filtered = filtered[~filtered[detail_col].astype(str).str.startswith('-')]
                    filtered = filtered[filtered[detail_col].astype(str).str.strip() != '']

                    display_col = detail_col
                    xlabel = "상세 동"
                    title = f"{self.selected_sido} {self.selected_sigungu} {self.selected_dong} - 상세 동별 평균연령"
                else:
                    messagebox.showwarning("경고", "해당 동의 상세 데이터가 없습니다.")
                    return
            else:
                messagebox.showwarning("경고", "상세 동 데이터가 없습니다.")
                return

        elif self.selected_sigungu:
            # 시/군/구까지 선택: 동/읍/면의 평균연령
            filtered = filtered[(filtered[sido_col] == self.selected_sido) &
                               (filtered[sigungu_col] == self.selected_sigungu)]

            # 동 레벨만 선택 (d열이 비어있거나 없는 것)
            filtered = filtered[filtered[dong_col].notna()]
            filtered = filtered[~filtered[dong_col].astype(str).str.startswith('-')]

            # d열이 있으면 d열이 비어있는 행만 (동 전체 데이터)
            if has_detail_col:
                filtered = filtered[
                    (filtered[detail_col].isna()) |
                    (filtered[detail_col].astype(str).str.strip() == '') |
                    (filtered[detail_col].astype(str).str.startswith('-'))
                ]

            display_col = dong_col
            xlabel = "동/읍/면"
            title = f"{self.selected_sido} {self.selected_sigungu} - 동별 평균연령"

        else:
            # 시/도만 선택: 시/군/구의 평균연령
            filtered = filtered[filtered[sido_col] == self.selected_sido]

            # 시/군/구 레벨만 선택 (동이 비어있는 것)
            filtered = filtered[filtered[sigungu_col].notna()]
            filtered = filtered[~filtered[sigungu_col].astype(str).str.startswith('-')]
            filtered = filtered[
                (filtered[dong_col].isna()) |
                (filtered[dong_col].astype(str).str.strip() == '') |
                (filtered[dong_col].astype(str).str.startswith('-'))
            ]

            display_col = sigungu_col
            xlabel = "시/군/구"
            title = f"{self.selected_sido} - 시/군/구별 평균연령"

        if len(filtered) == 0:
            messagebox.showwarning("경고", "해당 지역의 데이터가 없습니다.")
            return

        # 평균연령이 있는 행만
        filtered = filtered[filtered['평균연령'].notna()]

        if len(filtered) == 0:
            messagebox.showwarning("경고", "평균연령 데이터가 없습니다.")
            return

        # 표시할 이름 정리 (괄호 제거)
        filtered['display_name'] = filtered[display_col].apply(lambda x: re.sub(r'\([^)]*\)', '', str(x)).strip())

        # 정렬 - 평균연령 낮은 순 (오름차순)
        filtered = filtered.sort_values('평균연령', ascending=True)

        # 그래프 그리기 - 이전 그래프 완전히 제거
        self.figure.clf()  # clf()로 모든 axes와 내용 완전히 제거

        # 가로 길이 조정 (데이터 개수에 따라)
        fig_width = max(12, len(filtered) * 0.4)
        self.figure.set_size_inches(fig_width, 7)

        ax = self.figure.add_subplot(111)

        # seaborn barplot 사용
        colors = sns.color_palette("coolwarm", len(filtered))
        bars = ax.bar(range(len(filtered)), filtered['평균연령'], color=colors,
                     edgecolor='black', linewidth=0.7, alpha=0.8)

        # x축 레이블 설정
        ax.set_xticks(range(len(filtered)))
        ax.set_xticklabels(filtered['display_name'], rotation=45, ha='right', fontsize=9)

        # 막대에 값 표시
        for i, (idx, row) in enumerate(filtered.iterrows()):
            ax.text(i, row['평균연령'] + 0.5, f"{row['평균연령']:.1f}",
                   ha='center', va='bottom', fontsize=8, fontweight='bold')

        ax.set_xlabel(xlabel, fontsize=11, fontweight='bold')
        ax.set_ylabel('평균 연령 (세)', fontsize=11, fontweight='bold')
        ax.set_title(title, fontsize=13, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3, axis='y', linestyle='--')

        # y축 범위 조정
        y_min = max(0, filtered['평균연령'].min() - 5)
        y_max = filtered['평균연령'].max() + 5
        ax.set_ylim(y_min, y_max)

        self.figure.tight_layout()
        self.canvas.draw_idle()  # draw_idle()로 효율적인 재그리기

    def draw_age_distribution_graph(self):
        """동별 연령별 인구 분포 그래프"""
        # 비교 지역이 있는지 확인
        use_comparison = len(self.comparison_regions) > 0

        if use_comparison:
            # 비교 모드: 여러 지역 비교
            self.draw_comparison_age_distribution()
        else:
            # 단일 지역 모드
            self.draw_single_age_distribution()

    def draw_single_age_distribution(self):
        """단일 지역의 연령별 인구 분포 그래프"""
        if not self.selected_dong:
            messagebox.showwarning("경고", "동/읍/면을 선택해주세요.")
            return

        sido_col = self.df.columns[0]
        sigungu_col = self.df.columns[1]
        dong_col = self.df.columns[2]

        # 해당 동 데이터 필터링
        filtered = self.df[(self.df[sido_col] == self.selected_sido) &
                          (self.df[sigungu_col] == self.selected_sigungu)]

        # 동 이름 매칭 (괄호 포함/불포함 모두 체크)
        import re
        dong_data = None

        # d열(상세 동)이 선택된 경우
        if self.selected_detail and len(self.df.columns) > 3:
            detail_col = self.df.columns[3]
            for _, row in filtered.iterrows():
                dong_name = str(row[dong_col])
                detail_name = str(row[detail_col])
                clean_dong = re.sub(r'\([^)]*\)', '', dong_name).strip()
                clean_detail = re.sub(r'\([^)]*\)', '', detail_name).strip()

                if clean_dong == self.selected_dong and clean_detail == self.selected_detail:
                    dong_data = row
                    break
        else:
            # 상세 동이 없거나 선택되지 않은 경우
            for _, row in filtered.iterrows():
                dong_name = str(row[dong_col])
                clean_name = re.sub(r'\([^)]*\)', '', dong_name).strip()
                if clean_name == self.selected_dong:
                    # d열이 있는 경우, d열이 비어있는 행만 선택 (전체 동 데이터)
                    if len(self.df.columns) > 3:
                        detail_col = self.df.columns[3]
                        if pd.isna(row[detail_col]) or str(row[detail_col]).strip() == '':
                            dong_data = row
                            break
                    else:
                        dong_data = row
                        break

        if dong_data is None:
            messagebox.showwarning("경고", "해당 동의 데이터를 찾을 수 없습니다.")
            return

        # 연령대별 인구 데이터 추출
        age_cols = []
        age_labels = []
        populations = []

        for col in self.df.columns:
            if '0~9' in col:
                age_cols.append(col)
                age_labels.append('0~9세')
            elif '10~19' in col:
                age_cols.append(col)
                age_labels.append('10~19세')
            elif '20~29' in col:
                age_cols.append(col)
                age_labels.append('20~29세')
            elif '30~39' in col:
                age_cols.append(col)
                age_labels.append('30~39세')
            elif '40~49' in col:
                age_cols.append(col)
                age_labels.append('40~49세')
            elif '50~59' in col:
                age_cols.append(col)
                age_labels.append('50~59세')
            elif '60' in col and '이상' in col:
                age_cols.append(col)
                age_labels.append('60세 이상')

        for col in age_cols:
            pop = dong_data[col]
            if isinstance(pop, str):
                pop = float(pop.replace(',', ''))
            populations.append(pop)

        # 총 인구 계산
        total_population = sum(populations)

        # 그래프 그리기 - seaborn 스타일 - 이전 그래프 완전히 제거
        self.figure.clf()  # clf()로 모든 axes와 내용 완전히 제거
        ax = self.figure.add_subplot(111)

        # seaborn 컬러 팔레트 사용
        colors = sns.color_palette("Set2", len(age_labels))
        bars = ax.bar(age_labels, populations, color=colors,
                     edgecolor='black', linewidth=0.8, alpha=0.85)

        # 막대 위에 인구수 표시 (가로 방향)
        for i, (label, pop) in enumerate(zip(age_labels, populations)):
            ax.text(i, pop, f'{int(pop):,}명', ha='center', va='bottom',
                   fontsize=10, fontweight='bold', rotation=0)

        # 막대 안에 비율(%) 표시
        for i, (bar, pop) in enumerate(zip(bars, populations)):
            if total_population > 0:
                percentage = (pop / total_population) * 100
                # 막대 중앙에 표시
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width() / 2., height / 2,
                       f'{percentage:.1f}%',
                       ha='center', va='center',
                       fontsize=12, fontweight='bold', color='white',
                       bbox=dict(boxstyle='round,pad=0.3', facecolor='black', alpha=0.7, edgecolor='none'))

        # 제목 생성 (상세 동 포함)
        title_parts = [self.selected_sido, self.selected_sigungu, self.selected_dong]
        if self.selected_detail:
            title_parts.append(self.selected_detail)
        title = " ".join(title_parts) + " - 연령별 인구 분포"

        ax.set_xlabel('연령대', fontsize=11, fontweight='bold')
        ax.set_ylabel('인구수 (명)', fontsize=11, fontweight='bold')
        ax.set_title(title, fontsize=13, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3, axis='y', linestyle='--')

        # y축 천단위 구분자
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

        # x축 레이블 회전
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha='center')

        self.figure.tight_layout()
        self.canvas.draw_idle()  # draw_idle()로 효율적인 재그리기

    def draw_comparison_age_distribution(self):
        """여러 지역의 연령별 인구 분포 비교 그래프"""
        import re

        # 모든 지역의 데이터 수집
        regions_data = []
        region_names = []

        sido_col = self.df.columns[0]
        sigungu_col = self.df.columns[1]
        dong_col = self.df.columns[2]

        for region_info in self.comparison_regions:
            # 해당 동 데이터 필터링
            filtered = self.df[(self.df[sido_col] == region_info['sido']) &
                              (self.df[sigungu_col] == region_info['sigungu'])]

            dong_data = None

            # d열(상세 동)이 선택된 경우
            if region_info['detail'] and len(self.df.columns) > 3:
                detail_col = self.df.columns[3]
                for _, row in filtered.iterrows():
                    dong_name = str(row[dong_col])
                    detail_name = str(row[detail_col])
                    clean_dong = re.sub(r'\([^)]*\)', '', dong_name).strip()
                    clean_detail = re.sub(r'\([^)]*\)', '', detail_name).strip()

                    if clean_dong == region_info['dong'] and clean_detail == region_info['detail']:
                        dong_data = row
                        break
            else:
                # 상세 동이 없거나 선택되지 않은 경우
                for _, row in filtered.iterrows():
                    dong_name = str(row[dong_col])
                    clean_name = re.sub(r'\([^)]*\)', '', dong_name).strip()
                    if clean_name == region_info['dong']:
                        if len(self.df.columns) > 3:
                            detail_col = self.df.columns[3]
                            if pd.isna(row[detail_col]) or str(row[detail_col]).strip() == '':
                                dong_data = row
                                break
                        else:
                            dong_data = row
                            break

            if dong_data is not None:
                regions_data.append(dong_data)
                # 지역 이름 생성
                parts = [region_info['dong']]
                if region_info['detail']:
                    parts.append(region_info['detail'])
                region_names.append(" ".join(parts))

        if len(regions_data) == 0:
            messagebox.showwarning("경고", "선택된 지역의 데이터를 찾을 수 없습니다.")
            return

        # 연령대별 인구 데이터 추출
        age_labels = []
        age_cols = []

        for col in self.df.columns:
            if '0~9' in col:
                age_cols.append(col)
                age_labels.append('0~9세')
            elif '10~19' in col:
                age_cols.append(col)
                age_labels.append('10~19세')
            elif '20~29' in col:
                age_cols.append(col)
                age_labels.append('20~29세')
            elif '30~39' in col:
                age_cols.append(col)
                age_labels.append('30~39세')
            elif '40~49' in col:
                age_cols.append(col)
                age_labels.append('40~49세')
            elif '50~59' in col:
                age_cols.append(col)
                age_labels.append('50~59세')
            elif '60' in col and '이상' in col:
                age_cols.append(col)
                age_labels.append('60세 이상')

        # 그래프 그리기 - 이전 그래프 완전히 제거
        self.figure.clf()  # clf()로 모든 axes와 내용 완전히 제거
        ax = self.figure.add_subplot(111)

        # 막대 너비 및 위치 계산
        x = np.arange(len(age_labels))
        width = 0.8 / len(regions_data)  # 막대 너비

        # 각 지역별로 막대 그리기 (지역별로 색상 구분)
        colors = sns.color_palette("Set2", len(regions_data))

        for i, (dong_data, region_name) in enumerate(zip(regions_data, region_names)):
            populations = []
            for col in age_cols:
                pop = dong_data[col]
                if isinstance(pop, str):
                    pop = float(pop.replace(',', ''))
                populations.append(pop)

            # 총 인구 계산
            total_population = sum(populations)

            # 막대 위치 조정
            offset = width * (i - len(regions_data) / 2 + 0.5)
            bars = ax.bar(x + offset, populations, width, label=region_name,
                         color=colors[i], edgecolor='black', linewidth=0.5, alpha=0.85)

            # 막대 위에 인구수 표시 (가로 방향)
            for j, (pos, pop) in enumerate(zip(x + offset, populations)):
                if pop > 0:  # 0이 아닐 때만 표시
                    ax.text(pos, pop, f'{int(pop):,}명', ha='center', va='bottom',
                           fontsize=8, rotation=0, fontweight='bold')

            # 막대 안에 비율(%) 표시
            for j, (bar, pop) in enumerate(zip(bars, populations)):
                if total_population > 0 and pop > 0:
                    percentage = (pop / total_population) * 100
                    height = bar.get_height()
                    # 막대가 충분히 클 때만 표시
                    if height > max(populations) * 0.05:  # 최대값의 5% 이상일 때만
                        ax.text(bar.get_x() + bar.get_width() / 2., height / 2,
                               f'{percentage:.1f}%',
                               ha='center', va='center',
                               fontsize=7, fontweight='bold', color='white',
                               bbox=dict(boxstyle='round,pad=0.2', facecolor='black', alpha=0.6, edgecolor='none'))

        ax.set_xlabel('연령대', fontsize=11, fontweight='bold')
        ax.set_ylabel('인구수 (명)', fontsize=11, fontweight='bold')
        ax.set_title('지역별 연령별 인구 분포 비교', fontsize=13, fontweight='bold', pad=20)
        ax.set_xticks(x)
        ax.set_xticklabels(age_labels)
        ax.legend(loc='upper right', fontsize=9)
        ax.grid(True, alpha=0.3, axis='y', linestyle='--')

        # y축 천단위 구분자
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

        self.figure.tight_layout()
        self.canvas.draw_idle()  # draw_idle()로 효율적인 재그리기

    def save_graph(self):
        """그래프를 이미지 파일로 저장"""
        if self.figure.get_axes():
            # 파일 저장 다이얼로그
            filename = filedialog.asksaveasfilename(
                title="그래프 저장",
                defaultextension=".png",
                filetypes=[
                    ("PNG 파일", "*.png"),
                    ("JPG 파일", "*.jpg"),
                    ("PDF 파일", "*.pdf"),
                    ("SVG 파일", "*.svg"),
                    ("모든 파일", "*.*")
                ]
            )

            if filename:
                try:
                    # 고해상도로 저장
                    self.figure.savefig(filename, dpi=300, bbox_inches='tight',
                                       facecolor='white', edgecolor='none')
                    messagebox.showinfo("성공", f"그래프가 저장되었습니다.\n{filename}")
                except Exception as e:
                    messagebox.showerror("오류", f"그래프 저장 중 오류 발생:\n{str(e)}")
        else:
            messagebox.showwarning("경고", "저장할 그래프가 없습니다.\n먼저 그래프를 그려주세요.")

    def save_config(self):
        """설정 저장"""
        config = {
            'data_file': self.data_file
        }

        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def load_config(self):
        """설정 로드"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.data_file = config.get('data_file')
            except Exception as e:
                print(f"설정 로드 실패: {e}")

def main():
    root = tk.Tk()
    app = PopulationAgeAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
