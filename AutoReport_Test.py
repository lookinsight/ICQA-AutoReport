import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import pandas as pd
import openpyxl 
from PIL import ImageGrab, Image, ImageDraw, ImageFont, ImageTk
import os
import ctypes
import json
import textwrap
from datetime import datetime
import random
import re
import io

# 윈도우 디스플레이 배율 무시
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
exceptException:
    pass

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

COORD_FILE = "capture_coords.json"
FONT_PATH = "font.ttf" 
ARROW_ICON_PATH = "arrow_icon.png" 

class ICQA_AutoReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AutoReport_test")
        self.center_window(self, 550, 800) 
        
        self.raw_filepath = None
        self.dive_filepath = None
        
        self.coords = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.load_coords()
        self.remote = None 
        self.guide_win = None
        self.barcode_candidates = {} 
        self.session_captures = [] # 💡 이번 세션에 찍은 파워 BI 캡처본들 저장용

        if not os.path.exists(FONT_PATH):
            messagebox.showerror("필수 파일 누락", f"프로그램 폴더 안에 '{FONT_PATH}' (한글 폰트) 파일이 반드시 있어야 합니다.\n\n프로그램을 종료합니다.")
            self.destroy()
            return
            
        self.auto_create_arrow_icon()

        # ==========================================
        # 📊 [1단계] 엑셀 데이터 파일 선택 UI
        # ==========================================
        frame_excel = ctk.CTkFrame(self)
        frame_excel.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(frame_excel, text="[1단계] 데이터 파일 입력 및 표 생성", font=("Arial", 16, "bold")).pack(pady=(10, 5))
        
        self.btn_raw = ctk.CTkButton(frame_excel, text="📁 1. Raw Data 엑셀 선택", fg_color="#2B547E", height=40, command=self.load_raw_data)
        self.btn_raw.pack(pady=5, padx=20, fill="x")
        
        self.btn_dive = ctk.CTkButton(frame_excel, text="📁 2. Dive-Deep(사유) 엑셀 선택", fg_color="#2B547E", height=40, command=self.load_dive_data)
        self.btn_dive.pack(pady=5, padx=20, fill="x")

        frame_date = ctk.CTkFrame(frame_excel, fg_color="transparent")
        frame_date.pack(pady=(10, 5), padx=20, fill="x")
        ctk.CTkLabel(frame_date, text="📅 보고 대상 날짜:", font=("Arial", 14, "bold")).pack(side="left", padx=(0, 10))
        self.date_combo = ctk.CTkComboBox(frame_date, values=["Raw Data를 먼저 넣어주세요"], width=180)
        self.date_combo.pack(side="left")

        # 💡 [룩희 피드백 반영] 교체 기능 UI 삭제 (Defect 창으로 이동)

        self.report_range = ctk.StringVar(value="top5")
        frame_range_opt = ctk.CTkFrame(frame_excel, fg_color="transparent")
        frame_range_opt.pack(padx=20, pady=(15, 0), fill="x")
        ctk.CTkLabel(frame_range_opt, text="📋 보고 표 범위:", font=("Arial", 12, "bold"), text_color="#00FFCC").pack(side="left")
        ctk.CTkRadioButton(frame_range_opt, text="Top 5 (기본)", variable=self.report_range, value="top5").pack(side="left", padx=(10, 5))
        ctk.CTkRadioButton(frame_range_opt, text="전체 데이터", variable=self.report_range, value="all").pack(side="left", padx=5)

        self.barcode_mode = ctk.StringVar(value="top1")
        frame_barcode_opt = ctk.CTkFrame(frame_excel, fg_color="transparent")
        frame_barcode_opt.pack(padx=20, pady=(5, 0), fill="x")
        ctk.CTkLabel(frame_barcode_opt, text="👇 바코드 추출 방식:", font=("Arial", 12, "bold"), text_color="yellow").pack(side="left")
        ctk.CTkRadioButton(frame_barcode_opt, text="1위 바코드", variable=self.barcode_mode, value="top1", command=self.update_barcode_text).pack(side="left", padx=(10, 5))
        ctk.CTkRadioButton(frame_barcode_opt, text="🎲랜덤 바코드", variable=self.barcode_mode, value="random", command=self.update_barcode_text).pack(side="left", padx=5)

        self.btn_run = ctk.CTkButton(frame_excel, text="🚀 Data 병합 및 Defect Type 선택", fg_color="green", hover_color="darkgreen", height=45, command=self.process_data)
        self.btn_run.pack(pady=15, padx=20, fill="x")

        self.result_box = ctk.CTkTextbox(frame_excel, height=80, font=("Arial", 14))
        self.result_box.pack(padx=20, pady=(5, 10), fill="x")

        # ==========================================
        # 📸 [2단계] 파워 BI 대시보드 캡처
        # ==========================================
        frame_capture = ctk.CTkFrame(self)
        frame_capture.pack(pady=10, padx=20, fill="both", expand=True)

        ctk.CTkLabel(frame_capture, text="[2단계] 파워 BI 대시보드 캡처", font=("Arial", 16, "bold")).pack(pady=(10, 5))
        
        self.coord_labels = {}
        for i in range(1, 6):
            row_frame = ctk.CTkFrame(frame_capture, fg_color="transparent")
            row_frame.pack(pady=3, fill="x", padx=10)
            
            btn_snip = ctk.CTkButton(row_frame, text=f"📍 {i}번 지정", width=100, command=lambda num=str(i): self.start_snip(num))
            btn_snip.pack(side="left", padx=5)
            
            status_text = "✅ 지정됨" if self.coords[str(i)] else "❌ 미지정"
            lbl = ctk.CTkLabel(row_frame, text=status_text, width=70)
            lbl.pack(side="left", padx=5)
            self.coord_labels[str(i)] = lbl
            
            btn_del = ctk.CTkButton(row_frame, text="❌ 삭제", width=50, fg_color="darkred", hover_color="maroon", command=lambda num=str(i): self.delete_coord(num))
            btn_del.pack(side="left", padx=5)

        remote_btn = ctk.CTkButton(frame_capture, text="🎛️ 항상 위 리모컨 띄우기", fg_color="#E56717", hover_color="#C35613", height=40, command=self.open_remote)
        remote_btn.pack(pady=15, padx=20, fill="x")

        footer_label = ctk.CTkLabel(self, text="💡 Developed by 룩희 & 재민", font=("Arial", 12, "bold", "italic"), text_color="gray")
        footer_label.pack(side="bottom", pady=10)

    def center_window(self, target_window, width, height):
        screen_width = target_window.winfo_screenwidth()
        screen_height = target_window.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        target_window.geometry(f"{width}x{height}+{x}+{y}")

    def auto_create_arrow_icon(self):
        # 💡 [룩희 피드백 반영] Navy #2B547E / Orange #E56717 도면 색상
        if not os.path.exists(ARROW_ICON_PATH):
            print(f"[{ARROW_ICON_PATH}] 파일이 없어 자동 생성합니다...")
            img = Image.new("RGBA", (100, 40), (255, 255, 255, 0)) # 투명 배경
            draw = ImageDraw.Draw(img)
            draw.rectangle([10, 10, 70, 30], fill="#2B547E") # Navy 몸통
            draw.polygon([(70, 0), (100, 20), (70, 40)], fill="#E56717") # Orange 머리
            img.save(ARROW_ICON_PATH)

    def clean_text(self, text):
        if pd.isna(text): return ""
        cleaned = str(text).strip()
        if cleaned.endswith('.0'): cleaned = cleaned[:-2] # 엑셀 소수점 지우기
        return cleaned

    def clean_barcode(self, val):
        if pd.isna(val): return ""
        b = str(val).strip().upper()
        if b.endswith('.0'): 
            b = b[:-2]  
        return b

    def get_text_width(self, font, text):
        try:
            return font.getlength(text)
        except:
            try:
                return font.getsize(text)[0]
            except:
                bbox = font.getbbox(text)
                return bbox[2] if bbox else len(text)*7

    def force_pixel_wrap(self, text, font, max_width):
        if not text: return ""
        text = str(text).replace('\r', '')
        text = re.sub(r'\n+', '\n', text).strip()
        
        final_lines = []
        for paragraph in text.split('\n'):
            words = paragraph.split(' ') 
            current_line = ""
            
            for word in words:
                spacer = " " if current_line else ""
                test_line = current_line + spacer + word
                
                if self.get_text_width(font, test_line) > max_width:
                    if current_line:
                        final_lines.append(current_line)
                        current_line = word
                    else:
                        char_line = ""
                        for char in word:
                            if self.get_text_width(font, char_line + char) < max_width:
                                char_line += char
                            else:
                                final_lines.append(char_line)
                                char_line = char
                        current_line = char_line
                else:
                    current_line = test_line
            
            if current_line:
                final_lines.append(current_line)
                
        return "\n".join(final_lines)

    def load_raw_data(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.raw_filepath = filepath
            filename = os.path.basename(filepath)
            self.btn_raw.configure(text=f"✅ {filename} (클릭하여 변경)", fg_color="#454545")
            
            try:
                df = pd.read_excel(filepath, engine='openpyxl')
                df.columns = df.columns.str.strip().str.replace('\n', '')
                if 'REPORT_DATE' in df.columns:
                    dates = pd.to_datetime(df['REPORT_DATE'], errors='coerce').dt.strftime('%Y-%m-%d').dropna().unique().tolist()
                    dates.sort(reverse=True)
                    if dates:
                        self.date_combo.configure(values=dates)
                        self.date_combo.set(dates[0])
            exceptException as e:
                print(f"날짜 불러오기 실패: {e}")

    def load_dive_data(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.dive_filepath = filepath
            filename = os.path.basename(filepath)
            self.btn_dive.configure(text=f"✅ {filename} (클릭하여 변경)", fg_color="#454545")

    def process_data(self):
        if not self.raw_filepath or not self.dive_filepath:
            messagebox.showwarning("경고", "Raw Data와 Dive-Deep 엑셀 파일을 모두 선택해주세요!")
            return
            
        target_date = self.date_combo.get()
        if not target_date or target_date == "Raw Data를 먼저 넣어주세요":
            messagebox.showwarning("경고", "보고 대상 날짜를 선택해주세요!")
            return

        try:
            with open(self.raw_filepath, 'rb') as f:
                df_raw = pd.read_excel(f, engine='openpyxl')
            
            df_raw.columns = df_raw.columns.str.strip().str.replace('\n', '')
            
            barcode_col = 'BARCODE' if 'BARCODE' in df_raw.columns else ('EXTERNALID' if 'EXTERNALID' in df_raw.columns else None)
            if not barcode_col:
                messagebox.showerror("오류", "Raw Data에 'BARCODE' 또는 'EXTERNALID' 열이 없습니다!")
                return
                
            req_raw = ['RESOLVETYPE', barcode_col, 'PROBLEM_QTY', 'MOVED_QTY', 'DESCRIPTION', 'REPORT_DATE']
            self.has_external_id = 'EXTERNALID' in df_raw.columns
            if self.has_external_id and 'EXTERNALID' not in req_raw:
                req_raw.append('EXTERNALID')
                
            for col in req_raw:
                if col not in df_raw.columns:
                    messagebox.showerror("오류", f"Raw Data에 '{col}' 열이 없습니다!")
                    return

            df_raw[barcode_col] = df_raw[barcode_col].apply(self.clean_barcode)
            if self.has_external_id:
                df_raw['EXTERNALID'] = df_raw['EXTERNALID'].astype(str)
                
            df_raw['REPORT_DATE'] = pd.to_datetime(df_raw['REPORT_DATE'], errors='coerce').dt.strftime('%Y-%m-%d')
            self.df_raw_sess = df_raw[df_raw['REPORT_DATE'] == target_date] # 💡 세션용 Raw 데이터 저장
            
            if self.df_raw_sess.empty:
                messagebox.showinfo("알림", f"Raw Data 파일에 {target_date} 날짜의 데이터가 없습니다.")
                return

            agg_dict = {
                'PROBLEM_QTY': 'sum',
                'MOVED_QTY': 'sum',
                barcode_col: 'count', 
                'DESCRIPTION': 'first' 
            }
            if self.has_external_id:
                agg_dict['EXTERNALID'] = 'first'

            self.df_grouped_sess = self.df_raw_sess.groupby(['RESOLVETYPE', barcode_col, 'REPORT_DATE']).agg(agg_dict).rename(columns={barcode_col: 'COUNT'}).reset_index()

            with open(self.dive_filepath, 'rb') as f:
                xls = pd.ExcelFile(f, engine='openpyxl')
                valid_dfs = []
                
                dive_date_col = 'Date' 
                req_dive = ['상품바코드', '문제유형', '사유', dive_date_col]
                
                for sheet_name in xls.sheet_names:
                    temp_df = pd.read_excel(xls, sheet_name=sheet_name)
                    clean_cols = temp_df.columns.astype(str).str.strip().str.replace('\n', '').str.replace(' ', '')
                    temp_df.columns = clean_cols 
                    
                    if all(col in clean_cols for col in req_dive):
                        valid_dfs.append(temp_df)
                
                if not valid_dfs:
                    messagebox.showerror("오류", f"어떤 시트에서도 필수 열({req_dive})을 모두 찾을 수 없습니다!")
                    return
                
                self.df_dive_sess = pd.concat(valid_dfs, ignore_index=True) # 💡 세션용 Dive 데이터 저장

            self.df_dive_sess['상품바코드'] = self.df_dive_sess['상품바코드'].apply(self.clean_barcode)
            self.df_dive_sess[dive_date_col] = pd.to_datetime(self.df_dive_sess[dive_date_col], errors='coerce').dt.strftime('%Y-%m-%d')
            self.df_dive_sess = self.df_dive_sess[self.df_dive_sess[dive_date_col] == target_date]

            self.recompute_final_report_list() # 💡 공통 병합 로직 호출
            self.open_defect_selector()

        except PermissionError:
            messagebox.showerror("권한 에러", "엑셀 파일이 열려있거나 동기화 중입니다!\n열려있는 엑셀을 닫고 다시 시도해주세요.")
        exceptException as e:
            messagebox.showerror("에러 발생", f"실행 중 문제가 발생했습니다:\n{str(e)}")

    # ==========================================
    # 💡 [핵심] 룩희 피드백: Top 5 목록 리프레시 및 교체 로직 공통화
    # ==========================================
    def recompute_final_report_list(self):
        self.final_report_data = []
        self.barcode_candidates = {} 
        self.barcode_col_name = 'BARCODE' if 'BARCODE' in self.df_raw_sess.columns else ('EXTERNALID' if 'EXTERNALID' in self.df_raw_sess.columns else None)
        dive_date_col = 'Date' 

        resolve_types = self.df_grouped_sess['RESOLVETYPE'].unique()
        
        for r_type in resolve_types:
            type_df = self.df_grouped_sess[self.df_grouped_sess['RESOLVETYPE'] == r_type].copy()
            
            # Top 5 소팅 (수량 기준 -> 이동 수량 기준)
            if self.report_range.get() == "top5":
                target_df = type_df.sort_values(by=['PROBLEM_QTY', 'MOVED_QTY'], ascending=[False, False]).head(5)
            else:
                target_df = type_df.sort_values(by=['PROBLEM_QTY', 'MOVED_QTY'], ascending=[False, False])
            
            self.barcode_candidates[r_type] = target_df[self.barcode_col_name].tolist()
            
            # 사유 데이터 병합
            merged = pd.merge(
                target_df, 
                self.df_dive_sess, 
                left_on=[self.barcode_col_name, 'REPORT_DATE'], 
                right_on=['상품바코드', dive_date_col], 
                how='left'
            )
            
            # 딕셔너리로 변환하여 사진 슬롯 초기화
            for index, row in merged.iterrows():
                row_dict = row.to_dict()
                row_dict['ATTACHED_IMAGES'] = {"1": None, "2": None, "3": None, "4": None} 
                row_dict['RANK'] = index + 1 # 표 순서 (Top 1~5)
                self.final_report_data.append(row_dict) 
        
        self.update_barcode_text()

    def update_barcode_text(self):
        self.result_box.delete("1.0", tk.END)
        mode = self.barcode_mode.get()
        
        for r_type, barcodes in self.barcode_candidates.items():
            if not barcodes: continue
            
            # 카톡 복사용 바코드 추출 (1위 또는 랜덤)
            if mode == "top1":
                selected_barcode = barcodes[0] 
            else:
                selected_barcode = random.choice(barcodes) 
                
            clean_b = self.clean_barcode(selected_barcode)
            self.result_box.insert(tk.END, f"[{r_type}] 검색 바코드: {clean_b}\n")

    # ==========================================
    # 💡 [룩희 피드백 반영] 결재 창 UI 개선 (바코드 즉시 교체 기능 추가)
    # ==========================================
    def open_defect_selector(self):
        self.sel_win = ctk.CTkToplevel(self)
        self.sel_win.title("Defect Type 및 사유 입력/사진 관리 (결재)")
        self.center_window(self.sel_win, 950, 750) 
        self.sel_win.attributes("-topmost", True)
        self.sel_win.focus_force()
        self.sel_win.grab_set() 

        main_label = ctk.CTkLabel(self.sel_win, text="📌 Defect Type을 선택하고 누락된 사유 입력 및 현장 사진을 첨부해주세", font=("Arial", 16, "bold"))
        main_label.pack(pady=(15, 5))

        # --- 🔄 [룩희 피드백] 데이터 교체(선수 교체) UI 결재창으로 이동 ---
        frame_swap = ctk.CTkFrame(self.sel_win, fg_color="transparent")
        frame_swap.pack(pady=5, padx=20, fill="x")
        ctk.CTkLabel(frame_swap, text="🔄 데이터 교체:", font=("Arial", 12, "bold"), text_color="#FFCC00").pack(side="left", padx=(0, 10))
        
        entry_old_b = ctk.CTkEntry(frame_swap, placeholder_text="뺄 바코드 (예: 111)", width=130)
        entry_old_b.pack(side="left", padx=5)
        
        ctk.CTkLabel(frame_swap, text="➡️", font=("Arial", 14, "bold")).pack(side="left")
        
        entry_new_b = ctk.CTkEntry(frame_swap, placeholder_text="넣을 바코드 (예: 222)", width=130)
        entry_new_b.pack(side="left", padx=5)

        btn_swap_exec = ctk.CTkButton(frame_swap, text="🔄 교체 실행", width=100, fg_color="dimgray", hover_color="black", command=lambda: self.execute_barcode_swap(entry_old_b, entry_new_b))
        btn_swap_exec.pack(side="left", padx=10)

        ctk.CTkLabel(self.sel_win, text="━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", text_color="gray").pack(pady=2)

        # 스크롤 영역
        scroll_frame = ctk.CTkScrollableFrame(self.sel_win, width=900, height=530)
        scroll_frame.pack(pady=5, padx=20, fill="both", expand=True)

        self.entries_data = [] # UI 요소들 저장용 (최종 병합 시 데이터 추출)
        
        for i, row_dict in enumerate(self.final_report_data):
            # 행(row) 프레임
            row_frame = ctk.CTkFrame(scroll_frame)
            row_frame.pack(fill="x", pady=8, padx=5)
            
            # 데이터 정보 라인 (1줄)
            top_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
            top_frame.pack(fill="x", padx=10, pady=(10, 2))
            
            b_code = self.clean_text(row_dict[self.barcode_col_name])
            qty = self.clean_text(row_dict['PROBLEM_QTY'])
            r_type = self.clean_text(row_dict['RESOLVETYPE'])
            dive_val = self.clean_text(row_dict['사유'])
            
            info_text = f"No.{i+1} [{r_type}] 바코드: {b_code} | 문제수량: {qty}"
            ctk.CTkLabel(top_frame, text=info_text, font=("Arial", 14, "bold")).pack(side="left")
            
            # 사진 관리 버튼
            btn_manage_photo = ctk.CTkButton(top_frame, text="📷 사진 관리", width=100, fg_color="#E56717", hover_color="#C35613", command=lambda r=row_dict: self.open_image_manager(r))
            btn_manage_photo.pack(side="right", padx=5)

            # Defect Type 선택 콤보박스
            combo = ctk.CTkComboBox(top_frame, values=["Found", "Loss", "DAMAGED_SKU"], width=130)
            combo.set("Found") # 기본값
            combo.pack(side="right", padx=(15, 5))
            
            # 사유 입력 라인 (1줄)
            bot_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
            bot_frame.pack(fill="x", padx=10, pady=(0, 10))
            
            ctk.CTkLabel(bot_frame, text="✏️ 사유 확인/수정:", font=("Arial", 12, "bold"), text_color="#F3E5AB").pack(side="left")
            entry = ctk.CTkEntry(bot_frame, font=("Arial", 13))
            entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
            
            # 사유 엑셀 값 넣어주기 (없으면 미기재)
            dive_text = dive_val if dive_val else "사유 미기재"
            entry.insert(0, dive_text)
            
            self.entries_data.append((row_dict, combo, entry, r_type)) # 저장

        # 최종 완료 버튼
        btn_finish = ctk.CTkButton(self.sel_win, text="✨ 입력 완료 및 최종 보고서 이미지 생성 ✨", height=60, font=("Arial", 16, "bold"), command=self.generate_final_tables)
        btn_finish.pack(pady=15)

    # ==========================================
    # 💡 [핵심 교체 로직] 룩희 피드백 반영: 데이터를 직접 치환하고 목록 새로고침
    # ==========================================
    def execute_barcode_swap(self, entry_old, entry_new):
        old_b_input = entry_old.get().strip()
        new_b_input = entry_new.get().strip()
        
        if not old_b_input or not new_b_input:
            messagebox.showwarning("입력 누락", "교체할 바코드와 새로 넣을 바코드를 모두 입력해주세요.")
            return

        old_b = self.clean_barcode(old_b_input)
        new_b = self.clean_barcode(new_b_input)

        if old_b == new_b:
            messagebox.showwarning("동일 바코드", "교체 대상과 목표 바코드가 같습니다.")
            return

        print(f"🔄 교체 시도: {old_b} ➡️ {new_b}")

        # 1. grouped 세션 데이터에서 뺄 바코드 찾기
        mask_old = (self.df_grouped_sess[self.barcode_col_name] == old_b)
        if not self.df_grouped_sess[mask_old].any().any():
            messagebox.showerror("교체 실패", f"데이터 목록에 바코드 '{old_b}' 가 존재하지 않습니다.\n(혹은 RESOLVETYPE이 다를 수 있습니다)")
            return

        # 2. Raw 데이터에서 넣을 바코드 데이터 가져오기
        mask_new_raw = (self.df_raw_sess[self.barcode_col_name] == new_b)
        if not self.df_raw_sess[mask_new_raw].any().any():
            messagebox.showerror("교체 실패", f"Raw Data 파일 안에 새로운 바코드 '{new_b}' 가 존재하지 않습니다.\n(날짜와 조건을 확인해주세요)")
            return

        # RESOLVETYPE 매칭을 위해 원본 데이터의 타입을 가져옴
        resolve_type_of_old = self.df_grouped_sess[mask_old]['RESOLVETYPE'].iloc[0]

        # 새 바코드 데이터 수집 및 그룹핑 (기존과 동일한 agg 방식)
        new_raw_df = self.df_raw_sess[mask_new_raw]
        # RESOLVETYPE은 기존 데이터의 것을 강제로 따름 (선수 교체니까!)
        new_raw_df['RESOLVETYPE'] = resolve_type_of_old 

        agg_dict = {'PROBLEM_QTY': 'sum', 'MOVED_QTY': 'sum', self.barcode_col_name: 'count', 'DESCRIPTION': 'first' }
        if self.has_external_id: agg_dict['EXTERNALID'] = 'first'

        new_grouped_row = new_raw_df.groupby(['RESOLVETYPE', self.barcode_col_name, 'REPORT_DATE']).agg(agg_dict).rename(columns={self.barcode_col_name: 'COUNT'}).reset_index()

        # 3. [치환 실행]
        # 기존 데이터 삭제
        self.df_grouped_sess = self.df_grouped_sess[~mask_old]
        # 새 데이터 추가
        self.df_grouped_sess = pd.concat([self.df_grouped_sess, new_grouped_row], ignore_index=True)

        print("✨ 데이터 치환 완료. Top 5 목록을 재계산합니다.")
        
        # 4. 목록 재계산 및 UI 새로고침
        self.recompute_final_report_list()
        
        self.sel_win.destroy() # 기존 결재창 닫고
        self.open_defect_selector() # 새로 고침된 목록으로 다시 열기

        messagebox.showinfo("교체 완료", f"바코드 '{old_b}' 데이터가 '{new_b}' 데이터로 성공적으로 교체되었습니다.")

    def open_image_manager(self, record_dict):
        manager_win = ctk.CTkToplevel(self.sel_win)
        manager_win.title(f"No.{record_dict['RANK']} 바코드: {record_dict[self.barcode_col_name]} 현장 사진 관리")
        self.center_window(manager_win, 650, 500)
        manager_win.attributes("-topmost", True)
        manager_win.focus_force()
        manager_win.grab_set()

        # 슬롯 정의
        slots_frame = ctk.CTkFrame(manager_win)
        slots_frame.pack(pady=20, padx=20, fill="both", expand=True)

        slot_names = {
            "1": "1번: 문제 로케이션 사진 (필수)",
            "2": "2번: 확인1 (조치 중/완료)",
            "3": "3번: 확인2 (선택)",
            "4": "4번: 확인3 (선택)"
        }

        # 사진 데이터 바인딩
        for slot_num, name in slot_names.items():
            slot_frame = ctk.CTkFrame(slots_frame, fg_color="transparent")
            slot_frame.pack(fill="x", pady=10, padx=10)

            ctk.CTkLabel(slot_frame, text=name, font=("Arial", 13, "bold"), width=200, anchor="w").pack(side="left", padx=10)
            
            # 현재 등록된 사진 경로 레이블
            current_path = record_dict['ATTACHED_IMAGES'][slot_num]
            lbl_path = ctk.CTkLabel(slot_frame, text=os.path.basename(current_path) if current_path else "사진 없음", text_color="gray", width=250, anchor="w")
            lbl_path.pack(side="left", padx=10)

            # 파일 찾기 버튼
            btn_find = ctk.CTkButton(slot_frame, text="파일 찾기", width=80, command=lambda s=slot_num, l=lbl_path: self.find_file_for_slot(s, l, record_dict))
            btn_find.pack(side="right", padx=10)

        ctk.CTkButton(manager_win, text="사진 저장 및 닫기", height=40, font=("Arial", 14, "bold"), fg_color="green", command=manager_win.destroy).pack(pady=20)

    def find_file_for_slot(self, slot_num, lbl_path, record_dict):
        filepath = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
        if filepath:
            # 딕셔너리에 데이터 저장
            record_dict['ATTACHED_IMAGES'][slot_num] = filepath
            filename = os.path.basename(filepath)
            lbl_path.configure(text=filename, text_color="white") # UI 업데이트

    # ==========================================
    # 🖨️ 최종 One-Page 보고서 병합 로직 (대규모 수정)
    # ==========================================
    def generate_final_tables(self):
        # UI 입력값들 최종 딕셔너리에 업데이트
        updated_report_list = []
        for index, (row_dict, combo, entry, r_type) in enumerate(self.entries_data):
            row_dict['DEFECT_TYPE'] = combo.get()
            row_dict['FINAL_DIVE_DEEP'] = entry.get() 
            updated_report_list.append(row_dict)
            
        self.sel_win.destroy() # 결재 창 닫기
        
        # --- [최종 리포트 병합 프로세스] ---
        # 1. 공통 리소스 로드 (폰트, 도면 등)
        try:
            font_title = ImageFont.truetype(FONT_PATH, 22)
            font_header = ImageFont.truetype(FONT_PATH, 14)
            font_row = ImageFont.truetype(FONT_PATH, 13)
        except:
            messagebox.showerror("폰트 로드 실패", f"'{FONT_PATH}' 파일을 로드할 수 없습니다.")
            return

        cols = [
            ("No.", 40), ("External ID", 110), ("SKU Name", 280), 
            ("Problem QTY", 100), ("Problem 건수", 100), 
            ("Problem Type", 150), ("Solve Type", 180), 
            ("Defect Type", 100), ("Dive-Deep 사유", 400)
        ]
        table_width = sum([w for _, w in cols])
        header_height = 40 

        # 2. 통합 표 리포트 생성 (가변 높이)
        df_final = pd.DataFrame(updated_report_list)
        
        # 각 Solve Type 별로 리포트를 그립니다. (표 아래 사진)
        for r_type in df_final['RESOLVETYPE'].unique():
            type_data = df_final[df_final['RESOLVETYPE'] == r_type]
            
            # --- 표 높이 계산 ---
            raw_title = f"[{r_type}] Problem Analysis & 현장 조치 확인 리포트"
            title_wrap = self.force_pixel_wrap(raw_title, font_title, table_width - 40)
            title_lines = title_wrap.count('\n') + 1
            title_height = (title_lines * 40) + 30 

            row_heights = []
            wrapped_rows = []
            
            for _, row in type_data.iterrows():
                ext_id = row['EXTERNALID'] if self.has_external_id else row[self.barcode_col_name]
                
                # 데이터 튜플 생성
                raw_vals = [
                    str(row['RANK']), 
                    self.clean_text(ext_id), 
                    self.clean_text(row['DESCRIPTION']),
                    self.clean_text(row['PROBLEM_QTY']),
                    self.clean_text(row['COUNT']),
                    self.clean_text(row['문제유형']),
                    self.clean_text(row['RESOLVETYPE']),
                    self.clean_text(row['DEFECT_TYPE']),
                    self.clean_text(row['FINAL_DIVE_DEEP']) 
                ]
                
                wrapped_vals = []
                max_lines = 1
                for j, val in enumerate(raw_vals):
                    col_w = cols[j][1]
                    wrap_text = self.force_pixel_wrap(val, font_row, col_w - 20)
                    wrapped_vals.append(wrap_text)
                    max_lines = max(max_lines, wrap_text.count('\n') + 1)
                
                row_heights.append(max(55, (max_lines * 18) + 20)) # 최소 55, 줄 당 18px
                wrapped_rows.append(wrapped_vals)
                
            total_table_height = title_height + header_height + sum(row_heights) + 20
            
            # --- [1단계] 표 이미지 생성 ---
            img_table = Image.new('RGB', (table_width, total_table_height), 'white')
            draw = ImageDraw.Draw(img_table)

            # 색상 테마 (Navy)
            color_navy = '#1A365D'; color_white = '#FFFFFF'; color_iceblue = '#F0F4F8'; color_border = '#808080'

            # 테두리
            draw.rectangle([0, 0, table_width-1, total_table_height-1], outline=color_border, width=2)

            # 메인 타이틀
            draw.text((15, 20), title_wrap, font=font_title, fill='black', spacing=10)

            y_off = title_height
            
            # 헤더
            draw.rectangle([0, y_off, table_width, y_off + header_height], fill=color_navy, outline=color_border)
            x_off = 0
            for name, w in cols:
                draw.rectangle([x_off, y_off, x_off+w, y_off + header_height], outline=color_border)
                
                try: text_w = font_header.getlength(name); text_h = 14
                except: text_w = len(name) * 8; text_h = 14
                
                draw.text((x_off + (w - text_w) / 2, y_off + (header_height - text_h) / 2 - 2), name, font=font_header, fill=color_white)
                x_off += w

            y_off += header_height
            
            # 데이터 행(row) 그리기
            for i, wrapped_vals in enumerate(wrapped_rows):
                bg_color = color_white if i % 2 == 0 else color_iceblue
                current_rh = row_heights[i]
                
                x_off = 0
                for j, (val, w) in enumerate(zip(wrapped_vals, [cw for _, cw in cols])):
                    draw.rectangle([x_off, y_off, x_off+w, y_off + current_rh], fill=bg_color, outline=color_border)
                    
                    try: text_h = len(val.split('\n')) * 16 
                    except: text_h = 16

                    center_y = y_off + (current_rh - text_h) / 2 - 2

                    if j == 2 or j == 8: # SKU Name, 사유는 왼쪽 정렬
                        draw.text((x_off + 10, center_y), val, font=font_row, fill='black', spacing=6, align='left') 
                    else: # 나머지는 가운데 정렬
                        try: text_w = max([font_row.getlength(line) for line in val.split('\n')])
                        except: text_w = len(val) * 7
                        draw.text((x_off + (w - text_w) / 2, center_y), val, font=font_row, fill='black', spacing=6, align='center') 
                    x_off += w
                y_off += current_rh

            # --- [2단계] 현장 사진 병합 (표 아래) ---
            # 룩희 도면 리소스 로드 (화살표 도면)
            try: img_arrow_raw = Image.open(ARROW_ICON_PATH)
            exceptException as e:
                messagebox.showerror("화살표 도면 실패", f"'{ARROW_ICON_PATH}' 파일을 로드할 수 없습니다.\n PNG 파일인지 확인해주세요.")
                return

            # 전체 사진 프레임 가변 높이 계산
            total_images_frame_height = 0
            image_blocks = [] # 병합 대상 데이터 블록들

            for _, row in type_data.iterrows():
                b_code = self.clean_text(row[self.barcode_col_name])
                deep_text = self.clean_text(row['FINAL_DIVE_DEEP']) 
                attached = row['ATTACHED_IMAGES']
                
                # 사진이나 사유가 하나라도 있을 때만 블록 생성
                valid_images = {k: v for k, v in attached.items() if v and os.path.exists(v)}
                if valid_images or deep_text: 
                    
                    block_title_raw = f"No.{row['RANK']} 바코드: {b_code} [{r_type}] 현장 조치 확인"
                    block_title_wrap = self.force_pixel_wrap(block_title_raw, font_header, table_width - 100)
                    deep_text_wrap = self.force_pixel_wrap(deep_text, font_row, table_width - 100)
                    
                    title_h = (block_title_wrap.count('\n') + 1) * 20 + 20
                    사유_h = (deep_text_wrap.count('\n') + 1) * 20 + 30
                    img_area_height = 350 # 사진 고정 높이 영역
                    
                    # 블록당 총 높이 = 타이틀 + 사진 + 사유
                    block_total_height = title_h + img_area_height + 사유_h + 30 
                    total_images_frame_height += block_total_height
                    
                    image_blocks.append({
                        "title": block_title_wrap,
                        "title_h": title_h,
                        "images": valid_images,
                        "img_area_h": img_area_height,
                        "사유_f": deep_text_wrap,
                        "사유_h": 사유_h,
                        "total_h": block_total_height
                    })
            
            # --- [3단계] 최종 마스터 이미지 생성 (웅장한 합체) ---
            margin_bottom = 50
            final_total_height = total_table_height + total_images_frame_height + margin_bottom
            img_final = Image.new('RGB', (table_width, final_total_height), 'white')
            
            # 표 붙이기
            img_final.paste(img_table, (0, 0))
            
            # 사진 블록 주르륵 붙이기
            current_y = total_table_height + 20
            draw_f = ImageDraw.Draw(img_final)
            color_border = '#808080'

            for block in image_blocks:
                # 테두리 (블록 전체)
                draw_f.rectangle([15, current_y, table_width - 15, current_y + block['total_h']], outline=color_border, width=1)
                # 타이틀
                draw_f.text((30, current_y + 10), block['title'], font=font_header, fill='black', spacing=4)
                
                current_y += block['title_h']
                
                # --- 사진 레이아웃 엔진 (도면 레이아웃 구현) ---
                img_area_y_start = current_y + 10
                
                # 1번 (로케이션) 사진
                loc_img_path = block['images'].get("1")
                conf_images_paths = [v for k, v in block['images'].items() if k in ["2", "3", "4"]]
                
                target_img_height = 300 
                
                if loc_img_path:
                    with Image.open(loc_img_path) as loc_img_raw:
                        # 높이 300px 고정 리사이즈 (가로비율 유지)
                        w_ratio = target_img_height / loc_img_raw.height 
                        final_w = int(loc_img_raw.width * w_ratio)
                        img_loc_final = loc_img_raw.resize((final_w, target_img_height), Image.Resampling.LANCZOS)
                        
                        # 💡 [버그 수정] 강조 둥근 네모 그리기 기능 탑재!
                        # 룩희 피드백 반영: 여기에 둥근 네모 그리는 미니 포토샵 기능 넣어야 함. 
                        # 일단은 사진만 붙임. (후속 개발 포인트)
                        
                        img_final.paste(img_loc_final, (30 + 10, img_area_y_start)) # 여백 10px
                        current_x = 30 + 10 + final_w + 30 # 화살표 시작 x좌표
                else:
                    # 1번 사진 없음
                    final_w = 400
                    draw_f.text((30 + 100, img_area_y_start + 100), "[1번: 로케이션 사진 없음]", font=font_header, fill='gray')
                    current_x = 30 + 10 + final_w + 30

                # 화살표 도면 그리기
                h_ratio = 40 / img_arrow_raw.height # 높이 40px 고정
                arrow_final = img_arrow_raw.resize((int(img_arrow_raw.width * h_ratio), 40), Image.Resampling.LANCZOS)
                # 💡 사진 가려지는 버그 수정: 사진 붙인 후 화살표를 그 x좌표 계산해서 붙임
                img_final.paste(arrow_final, (current_x, img_area_y_start + int(target_img_height/2) - 20)) # 가운데 정렬
                
                current_x += arrow_final.width + 30 
                
                # 확인(조치 완료) 사진들 주르륵 붙이기 (등간격 배분)
                if conf_images_paths:
                    conf_count = len(conf_images_paths)
                    avail_width = table_width - 30 - current_x - 10 # 남은 여유 가로 폭
                    
                    if conf_count > 0:
                        per_img_w = int((avail_width - (conf_count-1)*10) / conf_count) # 등간격 폭
                        
                        for idx, c_path in enumerate(conf_images_paths):
                            with Image.open(c_path) as c_img_raw:
                                # 가로 폭 per_img_w에 맞춘 리사이즈 (높이 300px 초과 시 300 고정)
                                w_ratio = per_img_w / c_img_raw.width
                                final_h = int(c_img_raw.height * w_ratio)
                                
                                if final_h > target_img_height:
                                    h_ratio = target_img_height / c_img_raw.height
                                    c_img_final = c_img_raw.resize((int(c_img_raw.width * h_ratio), target_img_height), Image.Resampling.LANCZOS)
                                else:
                                    c_img_final = c_img_raw.resize((per_img_w, final_h), Image.Resampling.LANCZOS)
                                
                                # 라벨링 (확인1, 확인2)
                                draw_f.text((current_x + idx*(per_img_w + 10) + 5, img_area_y_start - 18), f"확인{idx+1}", font=font_row, fill='gray')

                                img_final.paste(c_img_final, (current_x + idx*(per_img_w + 10), img_area_y_start))
                else:
                    draw_f.text((current_x + 10, img_area_y_start + 100), "[조치 확인 사진 없음]", font=font_row, fill='gray')
                
                current_y += block['img_area_h'] + 20
                
                # --- [도면] 사유 내용 둥근 박스 그리기 ---
                try: text_w = font_row.getlength(block['사유_f'].split('\n')[0]); center_x = (table_width - text_w) / 2
                except: text_w = len(deep_text_wrap) * 7; center_x = 100

                # 사유 박스 그리기
                draw_f.rectangle([30, current_y, table_width - 30, current_y + block['사유_h'] - 10], fill='#F7F9FC', outline='#D0D7DE')
                # 룩희 피드백: 둥근 네모로 펜 색상 변경 (후속 개발 포인트)

                draw_f.multiline_text((center_x, current_y + 10), block['사유_f'], font=font_row, fill='#1A365D', spacing=6, align='center') # 사유는 가운데 정렬
                
                current_y += block['사유_h']
            
            # --- [최종 저장] ---
            # Solve Type 이름에서 안전한 파일명 추출
            safe_name = "".join([c for c in r_type if c.isalpha() or c.isdigit() or c in " _-"]).rstrip()
            final_filename = f"Complete_ICQA_Report_{target_date}_{safe_name}.png"
            img_final.save(final_filename)

        messagebox.showinfo("완료", "결재 및 하단 사진 병합 리포트 이미지가 완벽하게 생성되었습니다!\n프로그램 폴더를 확인하세요.")

    def load_coords(self):
        if os.path.exists(COORD_FILE):
            with open(COORD_FILE, "r") as f:
                self.coords = json.load(f)

    def save_coords(self):
        with open(COORD_FILE, "w") as f:
            json.dump(self.coords, f)

    def delete_coord(self, num):
        self.coords[num] = None
        self.save_coords()
        self.coord_labels[num].configure(text="❌ 미지정")
        self.hide_guide()

    def start_snip(self, num):
        self.withdraw() 
        self.snip_window = tk.Toplevel(self)
        self.snip_window.attributes('-alpha', 0.3)
        self.snip_window.overrideredirect(True) 
        self.snip_window.config(cursor="cross")

        try:
            user32 = ctypes.windll.user32
            v_x = user32.GetSystemMetrics(76) 
            v_y = user32.GetSystemMetrics(77) 
            v_w = user32.GetSystemMetrics(78) 
            v_h = user32.GetSystemMetrics(79) 
            
            self.snip_window.geometry(f"{v_w}x{v_h}+{v_x}+{v_y}")
        except:
            self.snip_window.attributes('-fullscreen', True)

        self.snip_window.bind("<ButtonPress-1>", self.on_press)
        self.snip_window.bind("<B1-Motion>", self.on_drag)
        self.snip_window.bind("<ButtonRelease-1>", lambda event: self.on_release(event, num))
        
        self.canvas = tk.Canvas(self.snip_window, cursor="cross", bg="gray")
        self.canvas.pack(fill="both", expand=True)
        self.rect = None

    def on_press(self, event):
        self.start_x = self.snip_window.winfo_pointerx()
        self.start_y = self.snip_window.winfo_pointery()
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=3, fill="black")

    def on_drag(self, event):
        cur_x = self.snip_window.winfo_pointerx()
        cur_y = self.snip_window.winfo_pointery()
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_release(self, event, num):
        end_x = self.snip_window.winfo_pointerx()
        end_y = self.snip_window.winfo_pointery()
        self.snip_window.destroy()
        self.deiconify() 

        x1, y1 = min(self.start_x, end_x), min(self.start_y, end_y)
        x2, y2 = max(self.start_x, end_x), max(self.start_y, end_y)

        if (x2 - x1) > 10 and (y2 - y1) > 10:
            self.coords[num] = (x1, y1, x2, y2)
            self.save_coords()
            self.coord_labels[num].configure(text="✅ 지정됨")

    def open_remote(self):
        if self.remote is not None and self.remote.winfo_exists():
            self.remote.focus()
            return

        self.remote = ctk.CTkToplevel(self)
        self.remote.title("리모컨")
        self.center_window(self.remote, 280, 350)
        self.remote.attributes("-topmost", True)

        for i in range(1, 6):
            frame = ctk.CTkFrame(self.remote, fg_color="transparent")
            frame.pack(pady=5, padx=10, fill="x")
            
            btn_aim = ctk.CTkButton(frame, text=f"🔍 {i}번 조준", width=100, fg_color="gray", hover_color="dimgray", command=lambda num=str(i): self.show_guide(num))
            btn_aim.pack(side="left", padx=5)
            
            btn_shot = ctk.CTkButton(frame, text=f"📸 찰칵!", width=100, command=lambda num=str(i): self.take_screenshot(num))
            btn_shot.pack(side="right", padx=5)
            
        btn_clear = ctk.CTkButton(self.remote, text="❌ 조준선 끄기", fg_color="darkred", hover_color="maroon", command=self.hide_guide)
        btn_clear.pack(pady=15, fill="x", padx=15)

    def show_guide(self, num):
        self.hide_guide() 
        coord = self.coords[num]
        if not coord:
            return
            
        x1, y1, x2, y2 = coord
        w = x2 - x1
        h = y2 - y1
        
        self.guide_win = tk.Toplevel(self)
        self.guide_win.overrideredirect(True) 
        self.guide_win.attributes("-topmost", True) 
        
        transparent_color = "magenta"
        self.guide_win.config(bg=transparent_color)
        self.guide_win.attributes("-transparentcolor", transparent_color)
        
        self.guide_win.geometry(f"{w}x{h}+{x1}+{y1}")
        
        canvas = tk.Canvas(self.guide_win, bg=transparent_color, highlightthickness=3, highlightbackground="red")
        canvas.pack(fill="both", expand=True)

    def hide_guide(self):
        if self.guide_win:
            self.guide_win.destroy()
            self.guide_win = None

    def take_screenshot(self, num):
        coord = self.coords[num]
        if not coord:
            return
            
        self.hide_guide()
        if self.remote is not None and self.remote.winfo_exists():
            self.remote.withdraw()
        self.withdraw()
        
        self.after(300, lambda: self._do_capture(num, coord))
        
    def _do_capture(self, num, coord):
        bbox = (coord[0], coord[1], coord[2], coord[3])
        img = ImageGrab.grab(bbox=bbox, all_screens=True)
        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        filename = f"Capture_{num}_{time_str}.png"
        img.save(filename)
        
        self.session_captures.append(filename) # 💡 이번 세션 캡처 목록에 추가
        print(f"[{filename}] 캡처 완료 및 병합 목록에 저장됨!")

        if self.remote is not None and self.remote.winfo_exists():
            self.remote.deiconify()
        self.deiconify()

if __name__ == "__main__":
    app = ICQA_AutoReportApp()
    app.mainloop()
