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
except Exception:
    pass

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

COORD_FILE = "capture_coords.json"
FONT_PATH = "font.ttf" 

class ICQA_AutoReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AutoReport_Master")
        self.center_window(self, 650, 850) 
        
        self.raw_filepath = None
        self.dive_filepath = None
        
        self.coords = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.load_coords()
        self.remote = None 
        self.guide_win = None
        self.barcode_candidates = {} 
        self.selected_barcodes_dict = {} 
        
        self.latest_captures = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.bg_captures = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.bi_edit_coords = {"1": [], "2": [], "3": [], "4": [], "5": []}

        if not os.path.exists(FONT_PATH):
            messagebox.showerror("필수 파일 누락", f"프로그램 폴더 안에 '{FONT_PATH}' (한글 폰트) 파일이 반드시 있어야 합니다.\n\n프로그램을 종료합니다.")
            self.destroy()
            return

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

        self.report_range = ctk.StringVar(value="top5")
        frame_range_opt = ctk.CTkFrame(frame_excel, fg_color="transparent")
        frame_range_opt.pack(padx=20, pady=(10, 0), fill="x")
        ctk.CTkLabel(frame_range_opt, text="📋 보고 표 범위:", font=("Arial", 12, "bold"), text_color="#00FFCC").pack(side="left")
        ctk.CTkRadioButton(frame_range_opt, text="Top 5 (기본)", variable=self.report_range, value="top5").pack(side="left", padx=(10, 5))
        ctk.CTkRadioButton(frame_range_opt, text="전체 데이터", variable=self.report_range, value="all").pack(side="left", padx=5)

        self.barcode_mode = ctk.StringVar(value="top1")
        frame_barcode_opt = ctk.CTkFrame(frame_excel, fg_color="transparent")
        frame_barcode_opt.pack(padx=20, pady=(5, 0), fill="x")
        ctk.CTkLabel(frame_barcode_opt, text="👇 바코드 추출 방식:", font=("Arial", 12, "bold"), text_color="yellow").pack(side="left")
        ctk.CTkRadioButton(frame_barcode_opt, text="1위 바코드", variable=self.barcode_mode, value="top1", command=self.update_barcode_text).pack(side="left", padx=(10, 5))
        ctk.CTkRadioButton(frame_barcode_opt, text="🎲랜덤 바코드", variable=self.barcode_mode, value="random", command=self.update_barcode_text).pack(side="left", padx=5)

        self.btn_run = ctk.CTkButton(frame_excel, text="🚀 Data 병합 및 사유/현장사진 입력", fg_color="green", hover_color="darkgreen", height=45, command=self.process_data)
        self.btn_run.pack(pady=10, padx=20, fill="x")

        self.result_box = ctk.CTkTextbox(frame_excel, height=70, font=("Arial", 14))
        self.result_box.pack(padx=20, pady=(5, 10), fill="x")

        # ==========================================
        # 📸 [2단계] 파워 BI 대시보드 캡처
        # ==========================================
        frame_capture = ctk.CTkFrame(self)
        frame_capture.pack(pady=10, padx=20, fill="both", expand=True)

        ctk.CTkLabel(frame_capture, text="[2단계] 파워 BI 캡처 (1번:전체 / 2~5번:분할)", font=("Arial", 16, "bold")).pack(pady=(10, 5))
        
        self.coord_labels = {}
        self.btn_edits_bi = {} 
        
        for i in range(1, 6):
            row_frame = ctk.CTkFrame(frame_capture, fg_color="transparent")
            row_frame.pack(pady=3, fill="x", padx=10)
            
            btn_snip = ctk.CTkButton(row_frame, text=f"📍 {i}번 지정", width=80, command=lambda num=str(i): self.start_snip(num))
            btn_snip.pack(side="left", padx=5)
            
            status_text = "✅ 캡처완료" if self.latest_captures[str(i)] else ("✅ 지정됨" if self.coords[str(i)] else "❌ 미지정")
            text_color = "white" if self.coords[str(i)] else "gray"
            lbl = ctk.CTkLabel(row_frame, text=status_text, width=80, text_color=text_color)
            lbl.pack(side="left", padx=5)
            self.coord_labels[str(i)] = lbl
            
            btn_edit_bi = ctk.CTkButton(row_frame, text="🖍️ 에디터", width=70, fg_color="#2B547E", hover_color="#224263", state="disabled", command=lambda num=str(i): self.open_bi_editor(num))
            btn_edit_bi.pack(side="left", padx=5)
            self.btn_edits_bi[str(i)] = btn_edit_bi
            
            btn_del = ctk.CTkButton(row_frame, text="❌ 삭제", width=50, fg_color="darkred", hover_color="maroon", command=lambda num=str(i): self.delete_coord(num))
            btn_del.pack(side="left", padx=5)

        remote_btn = ctk.CTkButton(frame_capture, text="🎛️ 항상 위 리모컨 띄우기", fg_color="#E56717", hover_color="#C35613", height=40, command=self.open_remote)
        remote_btn.pack(pady=10, padx=20, fill="x")

        # ==========================================
        # 📝 [3단계] 최종 보고서 생성
        # ==========================================
        frame_final = ctk.CTkFrame(self, fg_color="transparent")
        frame_final.pack(pady=5, padx=20, fill="x")
        
        self.photo_layout_mode = ctk.StringVar(value="2col") 
        frame_layout_opt = ctk.CTkFrame(frame_final, fg_color="transparent")
        frame_layout_opt.pack(pady=(0, 10))
        ctk.CTkLabel(frame_layout_opt, text="🖼️ 현장사진 레이아웃:", font=("Arial", 13, "bold"), text_color="#00FFCC").pack(side="left")
        ctk.CTkRadioButton(frame_layout_opt, text="1단 보기(크게)", variable=self.photo_layout_mode, value="1col").pack(side="left", padx=(10, 5))
        ctk.CTkRadioButton(frame_layout_opt, text="2단 보기(압축)", variable=self.photo_layout_mode, value="2col").pack(side="left", padx=5)
        ctk.CTkRadioButton(frame_layout_opt, text="3단 보기(초압축)", variable=self.photo_layout_mode, value="3col").pack(side="left", padx=5)

        self.btn_final_report = ctk.CTkButton(frame_final, text="✨ [3단계] 최종 통합 One-Page 보고서 생성 ✨", height=50, font=("Arial", 16, "bold"), fg_color="#B8860B", hover_color="#8B6508", command=self.generate_final_tables)
        self.btn_final_report.pack(fill="x")

        footer_label = ctk.CTkLabel(self, text="💡 Developed by 룩희 & 재민", font=("Arial", 12, "bold", "italic"), text_color="gray")
        footer_label.pack(side="bottom", pady=5)

    def center_window(self, target_window, width, height):
        screen_width = target_window.winfo_screenwidth()
        screen_height = target_window.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        target_window.geometry(f"{width}x{height}+{x}+{y}")

    def clean_text(self, text):
        if pd.isna(text): return ""
        cleaned = str(text).strip()
        if cleaned.endswith('.0'): cleaned = cleaned[:-2]
        return cleaned

    def clean_barcode(self, val):
        if pd.isna(val): return ""
        b = str(val).strip().upper()
        if b.endswith('.0'): b = b[:-2]
        return b

    def get_text_width(self, font, text):
        try: return font.getlength(text)
        except Exception:
            try: return font.getsize(text)[0]
            except Exception: 
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
                            if self.get_text_width(font, char_line + char) < max_width: char_line += char
                            else: 
                                final_lines.append(char_line)
                                char_line = char
                        current_line = char_line
                else: current_line = test_line
            if current_line: final_lines.append(current_line)
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
            except Exception as e: print(f"날짜 불러오기 실패: {e}")

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
            with open(self.raw_filepath, 'rb') as f: df_raw = pd.read_excel(f, engine='openpyxl')
            df_raw.columns = df_raw.columns.str.strip().str.replace('\n', '')
            barcode_col = 'BARCODE' if 'BARCODE' in df_raw.columns else ('EXTERNALID' if 'EXTERNALID' in df_raw.columns else None)
            
            if not barcode_col: 
                messagebox.showerror("오류", "Raw Data에 'BARCODE' 또는 'EXTERNALID' 열이 없습니다!")
                return
                
            req_raw = ['RESOLVETYPE', barcode_col, 'PROBLEM_QTY', 'MOVED_QTY', 'DESCRIPTION', 'REPORT_DATE']
            self.has_external_id = 'EXTERNALID' in df_raw.columns
            if self.has_external_id and 'EXTERNALID' not in req_raw: req_raw.append('EXTERNALID')
                
            for col in req_raw:
                if col not in df_raw.columns: 
                    messagebox.showerror("오류", f"Raw Data에 '{col}' 열이 없습니다!")
                    return
                    
            df_raw[barcode_col] = df_raw[barcode_col].apply(self.clean_barcode)
            if self.has_external_id: df_raw['EXTERNALID'] = df_raw['EXTERNALID'].astype(str)
                
            df_raw['REPORT_DATE'] = pd.to_datetime(df_raw['REPORT_DATE'], errors='coerce').dt.strftime('%Y-%m-%d')
            self.df_raw_sess = df_raw[df_raw['REPORT_DATE'] == target_date]
            
            if self.df_raw_sess.empty: 
                messagebox.showinfo("알림", f"해당 날짜에 맞는 데이터가 없습니다.")
                return
                
            agg_dict = {'PROBLEM_QTY': 'sum', 'MOVED_QTY': 'sum', barcode_col: 'count', 'DESCRIPTION': 'first'}
            if self.has_external_id: agg_dict['EXTERNALID'] = 'first'
                
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
                    if all(col in clean_cols for col in req_dive): valid_dfs.append(temp_df)
                        
                if not valid_dfs: 
                    messagebox.showerror("오류", f"필수 열({req_dive})을 찾을 수 없습니다!")
                    return
                self.df_dive_sess = pd.concat(valid_dfs, ignore_index=True)

            self.df_dive_sess['상품바코드'] = self.df_dive_sess['상품바코드'].apply(self.clean_barcode)
            self.df_dive_sess[dive_date_col] = pd.to_datetime(self.df_dive_sess[dive_date_col], errors='coerce').dt.strftime('%Y-%m-%d')
            self.df_dive_sess = self.df_dive_sess[self.df_dive_sess[dive_date_col] == target_date]

            self.recompute_final_report_list(initial_load=True)
            self.open_defect_selector()
            
        except Exception as e: 
            messagebox.showerror("에러 발생", f"실행 중 문제가 발생했습니다:\n{str(e)}")

    def recompute_final_report_list(self, initial_load=False):
        self.final_report_data = []
        self.barcode_candidates = {}
        self.barcode_col_name = 'BARCODE' if 'BARCODE' in self.df_raw_sess.columns else ('EXTERNALID' if 'EXTERNALID' in self.df_raw_sess.columns else None)
        dive_date_col = 'Date' 
        
        resolve_types = self.df_grouped_sess['RESOLVETYPE'].unique()
        global_rank = 1
        
        for r_type in resolve_types:
            type_df = self.df_grouped_sess[self.df_grouped_sess['RESOLVETYPE'] == r_type].copy()
            
            if self.report_range.get() == "top5": target_df = type_df.sort_values(by=['PROBLEM_QTY', 'MOVED_QTY'], ascending=[False, False]).head(5)
            else: target_df = type_df.sort_values(by=['PROBLEM_QTY', 'MOVED_QTY'], ascending=[False, False])
                
            self.barcode_candidates[r_type] = target_df[self.barcode_col_name].tolist()
            merged = pd.merge(target_df, self.df_dive_sess, left_on=[self.barcode_col_name, 'REPORT_DATE'], right_on=['상품바코드', dive_date_col], how='left')
            
            for index, row in merged.iterrows(): 
                row_dict = row.to_dict()
                if 'ATTACHED_IMAGES' not in row_dict:
                    row_dict['ATTACHED_IMAGES'] = {"1": None, "2": None, "3": None, "4": None}
                    row_dict['BG_IMAGES'] = {"1": None, "2": None, "3": None, "4": None}
                    row_dict['EDIT_COORDS'] = {"1": [], "2": [], "3": [], "4": []}
                row_dict['RANK'] = index + 1 
                row_dict['GLOBAL_RANK'] = global_rank
                
                # 💡 [버그 수정] 화살표 색상을 위한 콤보박스 기본값은 무조건 "Found"로 고정!
                row_dict['DEFECT_TYPE'] = "Found" 
                
                reason_from_excel = self.clean_text(row_dict.get('사유', ''))
                row_dict['FINAL_DIVE_DEEP'] = reason_from_excel 
                
                global_rank += 1
                self.final_report_data.append(row_dict) 
                
        self.update_barcode_text(initial_load)

    def update_barcode_text(self, initial_load=False):
        self.result_box.delete("1.0", tk.END)
        mode = self.barcode_mode.get()
        
        if initial_load: self.selected_barcodes_dict.clear()
            
        for r_type, barcodes in self.barcode_candidates.items():
            if not barcodes: continue
            if initial_load:
                selected_barcode = barcodes[0] if mode == "top1" else random.choice(barcodes) 
                self.selected_barcodes_dict[r_type] = self.clean_barcode(selected_barcode)
            else:
                if self.selected_barcodes_dict.get(r_type) not in [self.clean_barcode(b) for b in barcodes]:
                    selected_barcode = barcodes[0] if mode == "top1" else random.choice(barcodes) 
                    self.selected_barcodes_dict[r_type] = self.clean_barcode(selected_barcode)
                    
            self.result_box.insert(tk.END, f"[{r_type}] 대표 바코드: {self.selected_barcodes_dict.get(r_type)}\n")

    def open_defect_selector(self):
        self.sel_win = ctk.CTkToplevel(self)
        self.sel_win.title("Defect Type 및 사유 입력/사진 관리")
        self.center_window(self.sel_win, 950, 750)
        self.sel_win.focus_force()
        self.sel_win.grab_set() 
        
        main_label = ctk.CTkLabel(self.sel_win, text="📌 딱 [대표 SKU]에만 사진을 등록하고 나머지는 사유 확인만 해주세요.", font=("Arial", 16, "bold"))
        main_label.pack(pady=(15, 5))
        
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
        
        scroll_frame = ctk.CTkScrollableFrame(self.sel_win, width=900, height=530)
        scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.entries_data = [] 
        for i, row_dict in enumerate(self.final_report_data):
            b_code = self.clean_text(row_dict[self.barcode_col_name])
            qty = self.clean_text(row_dict['PROBLEM_QTY'])
            r_type = self.clean_text(row_dict['RESOLVETYPE'])
            global_rank = row_dict['GLOBAL_RANK']
            
            is_representative = (b_code == self.selected_barcodes_dict.get(r_type))
            bg_color = "#3A3B3C" if is_representative else "transparent"
            
            row_frame = ctk.CTkFrame(scroll_frame, fg_color=bg_color)
            row_frame.pack(fill="x", pady=8, padx=5)
            
            top_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
            top_frame.pack(fill="x", padx=10, pady=(10, 2))
            
            ctk.CTkLabel(top_frame, text=f"No.{global_rank} [{r_type}] 바코드: {b_code} | 문제수량: {qty}", font=("Arial", 14, "bold"), text_color="white" if is_representative else "gray").pack(side="left")
            
            if is_representative:
                btn_manage_photo = ctk.CTkButton(top_frame, text="⭐ 대표 현장사진 등록", width=140, fg_color="#E56717", hover_color="#C35613", command=lambda r=row_dict: self.open_image_manager(r))
                btn_manage_photo.pack(side="right", padx=5)
            else:
                ctk.CTkLabel(top_frame, text="[사진 생략]", text_color="gray", font=("Arial", 12, "italic")).pack(side="right", padx=15)
                
            combo = ctk.CTkComboBox(top_frame, values=["Found", "Loss", "DAMAGED_SKU"], width=130)
            combo.set(row_dict.get('DEFECT_TYPE', 'Found'))
            combo.pack(side="right", padx=(15, 5))
            
            bot_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
            bot_frame.pack(fill="x", padx=10, pady=(0, 10))
            
            ctk.CTkLabel(bot_frame, text="✏️ 사유 확인/수정:", font=("Arial", 12, "bold"), text_color="#F3E5AB").pack(side="left")
            entry = ctk.CTkEntry(bot_frame, font=("Arial", 13))
            entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
            
            dive_val = row_dict.get('FINAL_DIVE_DEEP', '')
            entry.insert(0, dive_val if dive_val else "사유 미기재")
            
            self.entries_data.append((row_dict, combo, entry)) 
            
        ctk.CTkButton(self.sel_win, text="💾 사유 및 현장사진 임시저장", height=50, font=("Arial", 16, "bold"), fg_color="green", command=self.save_temporary_data).pack(pady=15)

    def save_temporary_data(self):
        for row_dict, combo, entry in self.entries_data:
            row_dict['DEFECT_TYPE'] = combo.get()
            row_dict['FINAL_DIVE_DEEP'] = entry.get() 
        self.sel_win.destroy()
        messagebox.showinfo("임시 저장 완료", "현장 데이터가 저장되었습니다.\n\n[2단계] 파워 BI 캡처를 진행한 후,\n[3단계]에서 최종 보고서를 생성해주세요!")

    def execute_barcode_swap(self, entry_old, entry_new):
        old_b_input = entry_old.get().strip()
        new_b_input = entry_new.get().strip()
        
        if not old_b_input or not new_b_input:
            messagebox.showwarning("입력 누락", "교체할 바코드와 새로 넣을 바코드를 모두 입력해주세요.")
            return

        old_b = self.clean_barcode(old_b_input)
        new_b = self.clean_barcode(new_b_input)

        mask_old = (self.df_grouped_sess[self.barcode_col_name] == old_b)
        if not self.df_grouped_sess[mask_old].any().any():
            messagebox.showerror("교체 실패", f"데이터 목록에 뺄 바코드 '{old_b}' 가 존재하지 않습니다.")
            return

        mask_new_raw = (self.df_raw_sess[self.barcode_col_name] == new_b)
        if not self.df_raw_sess[mask_new_raw].any().any():
            messagebox.showerror("교체 실패", f"Raw Data 엑셀에 새로운 바코드 '{new_b}' 가 존재하지 않습니다.")
            return

        resolve_type_of_old = self.df_grouped_sess[mask_old]['RESOLVETYPE'].iloc[0]
        new_raw_df = self.df_raw_sess[mask_new_raw]
        new_raw_df['RESOLVETYPE'] = resolve_type_of_old 

        agg_dict = {'PROBLEM_QTY': 'sum', 'MOVED_QTY': 'sum', self.barcode_col_name: 'count', 'DESCRIPTION': 'first'}
        if self.has_external_id: agg_dict['EXTERNALID'] = 'first'

        new_grouped_row = new_raw_df.groupby(['RESOLVETYPE', self.barcode_col_name, 'REPORT_DATE']).agg(agg_dict).rename(columns={self.barcode_col_name: 'COUNT'}).reset_index()

        self.df_grouped_sess = self.df_grouped_sess[~mask_old]
        self.df_grouped_sess = pd.concat([self.df_grouped_sess, new_grouped_row], ignore_index=True)

        self.recompute_final_report_list(initial_load=False)
        self.sel_win.destroy() 
        self.open_defect_selector() 
        messagebox.showinfo("교체 완료", f"선수 교체가 성공적으로 완료되었습니다!")

    def open_image_manager(self, record_dict):
        manager_win = ctk.CTkToplevel(self)
        manager_win.title(f"No.{record_dict['GLOBAL_RANK']} 바코드: {record_dict[self.barcode_col_name]} 현장 사진 관리")
        self.center_window(manager_win, 850, 550)
        manager_win.focus_force()
        manager_win.grab_set()
        
        slots_frame = ctk.CTkFrame(manager_win)
        slots_frame.pack(pady=20, padx=20, fill="both", expand=True)
        slot_names = {"1": "1번: 문제 로케이션 사진 (강조 추천)", "2": "2번: 확인1 (조치 중/완료)", "3": "3번: 확인2 (선택)", "4": "4번: 확인3 (선택)"}
        
        thumb_labels = {}

        def update_thumbnail(path, lbl_thumb):
            if path and os.path.exists(path):
                try:
                    img = Image.open(path).convert('RGB')
                    img.thumbnail((70, 70), Image.Resampling.LANCZOS)
                    ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
                    lbl_thumb.configure(image=ctk_img, text="")
                except Exception:
                    lbl_thumb.configure(image="", text="미리보기\n오류")
            else:
                lbl_thumb.configure(image="", text="사진\n없음")

        def find_file(slot_num, lbl_path, btn_edit, lbl_thumb):
            file_path = filedialog.askopenfilename(parent=manager_win, filetypes=[("Image files", "*.jpg *.jpeg *.png")])
            if file_path: 
                record_dict['BG_IMAGES'][slot_num] = file_path
                record_dict['ATTACHED_IMAGES'][slot_num] = file_path
                record_dict['EDIT_COORDS'][slot_num] = []
                
                lbl_path.configure(text=os.path.basename(file_path), text_color="white")
                btn_edit.configure(state="normal")
                update_thumbnail(file_path, lbl_thumb) 
                
        def open_editor(slot_num, lbl_path, lbl_thumb):
            bg_path = record_dict.get('BG_IMAGES', {}).get(slot_num)
            if not bg_path: 
                bg_path = record_dict['ATTACHED_IMAGES'][slot_num]
                record_dict['BG_IMAGES'][slot_num] = bg_path
                record_dict['EDIT_COORDS'][slot_num] = []
                
            if bg_path and os.path.exists(bg_path): 
                existing_coords = record_dict['EDIT_COORDS'][slot_num]
                def on_save_callback(new_bg_path, new_final_path, new_coords):
                    record_dict['BG_IMAGES'][slot_num] = new_bg_path
                    record_dict['ATTACHED_IMAGES'][slot_num] = new_final_path
                    record_dict['EDIT_COORDS'][slot_num] = new_coords
                    
                    lbl_path.configure(text=f"🖍️ [편집됨] {os.path.basename(new_final_path)}", text_color="#00FFCC")
                    update_thumbnail(new_final_path, lbl_thumb) 
                ImageEditorWindow(manager_win, bg_path, existing_coords, on_save_callback) 
                
        for slot_num, name in slot_names.items():
            slot_frame = ctk.CTkFrame(slots_frame, fg_color="transparent")
            slot_frame.pack(fill="x", pady=8, padx=10)
            
            lbl_thumb = ctk.CTkLabel(slot_frame, text="사진\n없음", width=70, height=70, fg_color="#3A3B3C", corner_radius=5)
            lbl_thumb.pack(side="left", padx=(0, 15))
            thumb_labels[slot_num] = lbl_thumb
            
            ctk.CTkLabel(slot_frame, text=name, font=("Arial", 13, "bold"), width=220, anchor="w").pack(side="left")
            
            btn_f = ctk.CTkFrame(slot_frame, fg_color="transparent")
            btn_f.pack(side="right", padx=10)
            
            current_path = record_dict['ATTACHED_IMAGES'].get(slot_num)
            lbl_path = ctk.CTkLabel(slot_frame, text=os.path.basename(current_path) if current_path else "선택된 파일 없음", text_color="gray", anchor="w")
            
            btn_edit = ctk.CTkButton(btn_f, text="🖍️ 편집/회전/강조", width=120, fg_color="#2B547E", hover_color="#224263", state="normal" if current_path else "disabled")
            btn_find = ctk.CTkButton(btn_f, text="📁 파일 찾기", width=90)
            
            btn_edit.configure(command=lambda s=slot_num, l=lbl_path, t=lbl_thumb: open_editor(s, l, t))
            btn_find.configure(command=lambda s=slot_num, l=lbl_path, b=btn_edit, t=lbl_thumb: find_file(s, l, b, t))
            
            btn_edit.pack(side="right", padx=3)
            btn_find.pack(side="right", padx=3)
            lbl_path.pack(side="left", fill="x", expand=True, padx=10)
            
            update_thumbnail(current_path, lbl_thumb)
            
        ctk.CTkButton(manager_win, text="사진 적용 완료", height=40, font=("Arial", 14, "bold"), fg_color="green", command=manager_win.destroy).pack(pady=15)

    def open_bi_editor(self, num):
        bg_path = self.bg_captures.get(num)
        if not bg_path:
            bg_path = self.latest_captures.get(num)
            self.bg_captures[num] = bg_path
            self.bi_edit_coords[num] = []
            
        if bg_path and os.path.exists(bg_path):
            existing_coords = self.bi_edit_coords[num]
            def on_save_callback(new_bg_path, new_final_path, new_coords):
                self.bg_captures[num] = new_bg_path
                self.latest_captures[num] = new_final_path
                self.bi_edit_coords[num] = new_coords
                
                self.coord_labels[num].configure(text="🖍️ 편집완료", text_color="#00FFCC")
                self.btn_edits_bi[num].configure(fg_color="green", hover_color="darkgreen") 
            ImageEditorWindow(self, bg_path, existing_coords, on_save_callback)

    def generate_final_tables(self):
        if not hasattr(self, 'final_report_data') or not self.final_report_data:
            messagebox.showwarning("데이터 없음", "[1단계]를 먼저 완료해주세요!")
            return

        try: 
            font_title = ImageFont.truetype(FONT_PATH, 22)
            font_header = ImageFont.truetype(FONT_PATH, 14)
            font_row = ImageFont.truetype(FONT_PATH, 13)
        except Exception as e: 
            messagebox.showerror("폰트 로드 실패", f"'{FONT_PATH}' 로드 실패: {e}")
            return
            
        layout_mode = self.photo_layout_mode.get()
            
        cols = [("No.", 40), ("External ID", 110), ("SKU Name", 280), ("Problem QTY", 100), ("Problem 건수", 100), ("Problem Type", 150), ("Solve Type", 180), ("Defect Type", 100), ("Dive-Deep 사유", 400)]
        table_width = sum([w for _, w in cols])
        report_segments = []
        color_navy = '#1A365D'
        color_white = '#FFFFFF'
        color_iceblue = '#F0F4F8'
        color_border = '#808080'

        main_title_img = Image.new('RGB', (table_width, 80), color_white)
        ImageDraw.Draw(main_title_img).text((20, 25), f"📊 ICQA Daily Auto-Report ({self.date_combo.get()})", font=ImageFont.truetype(FONT_PATH, 28), fill=color_navy)
        report_segments.append(main_title_img)

        TARGET_H = 350
        margin = 20
        full_w = table_width - 40
        half_w = (full_w - margin) // 2
        third_w = (full_w - margin * 2) // 3 
        
        row1_caps = [self.latest_captures.get("1")]
        row2_caps = [self.latest_captures.get("2"), self.latest_captures.get("3")]
        row3_caps = [self.latest_captures.get("4"), self.latest_captures.get("5")]

        def create_capture_row(img_paths, is_half=False):
            valid_paths = [p for p in img_paths if p and os.path.exists(p)]
            if not valid_paths: return None

            row_img = Image.new('RGB', (table_width, TARGET_H + 20), color_white)
            draw = ImageDraw.Draw(row_img)

            if is_half:
                for idx, path in enumerate(img_paths):
                    if path and os.path.exists(path):
                        box_x = 20 if idx == 0 else 20 + half_w + margin
                        try:
                            with Image.open(path) as cap:
                                cap_resized = cap.resize((half_w, TARGET_H), Image.Resampling.LANCZOS)
                                row_img.paste(cap_resized, (box_x, 10))
                        except Exception as e: print(e)
                        draw.rectangle([box_x, 10, box_x + half_w, 10 + TARGET_H], outline=color_border, width=2)
            else:
                path = img_paths[0]
                if path and os.path.exists(path):
                    box_x = 20
                    try:
                        with Image.open(path) as cap:
                            cap_resized = cap.resize((full_w, TARGET_H), Image.Resampling.LANCZOS)
                            row_img.paste(cap_resized, (box_x, 10))
                    except Exception as e: print(e)
                    draw.rectangle([box_x, 10, box_x + full_w, 10 + TARGET_H], outline=color_border, width=2)
            return row_img

        seg1 = create_capture_row(row1_caps, is_half=False)
        seg2 = create_capture_row(row2_caps, is_half=True)
        seg3 = create_capture_row(row3_caps, is_half=True)

        if seg1: report_segments.append(seg1)
        if seg2: report_segments.append(seg2)
        if seg3: report_segments.append(seg3)
        if seg1 or seg2 or seg3:
            report_segments.append(Image.new('RGB', (table_width, 20), color_white))

        df_final = pd.DataFrame(self.final_report_data)
        
        for r_type in df_final['RESOLVETYPE'].unique():
            type_data = df_final[df_final['RESOLVETYPE'] == r_type]
            title_wrap = self.force_pixel_wrap(f"[{r_type}] Problem Analysis", font_title, table_width - 40)
            title_height = (title_wrap.count('\n') + 1) * 40 + 20 
            row_heights = []
            wrapped_rows = []
            
            for i, row in type_data.iterrows():
                ext_id = row['EXTERNALID'] if self.has_external_id else row[self.barcode_col_name]
                raw_vals = [str(row['GLOBAL_RANK']), self.clean_text(ext_id), self.clean_text(row['DESCRIPTION']), self.clean_text(row['PROBLEM_QTY']), self.clean_text(row['COUNT']), self.clean_text(row['문제유형']), self.clean_text(row['RESOLVETYPE']), self.clean_text(row.get('DEFECT_TYPE','')), self.clean_text(row.get('FINAL_DIVE_DEEP',''))]
                wrapped_vals = []
                max_lines = 1
                for j, val in enumerate(raw_vals): 
                    wrap_text = self.force_pixel_wrap(val, font_row, cols[j][1] - 20)
                    wrapped_vals.append(wrap_text)
                    max_lines = max(max_lines, wrap_text.count('\n') + 1)
                row_heights.append(max(50, (max_lines * 18) + 20))
                wrapped_rows.append(wrapped_vals)
                
            total_table_height = title_height + 40 + sum(row_heights) + 10
            img_table = Image.new('RGB', (table_width, total_table_height), color_white)
            draw = ImageDraw.Draw(img_table)
            draw.rectangle([0, 0, table_width-1, total_table_height-1], outline=color_border, width=2)
            draw.text((15, 10), title_wrap, font=font_title, fill='black', spacing=10)
            
            y_off = title_height
            draw.rectangle([0, y_off, table_width, y_off + 40], fill=color_navy, outline=color_border)
            x_off = 0
            
            for name, w in cols:
                draw.rectangle([x_off, y_off, x_off+w, y_off + 40], outline=color_border)
                try: text_w = self.get_text_width(font_header, name)
                except Exception: text_w = len(name) * 8
                draw.text((x_off + (w - text_w) / 2, y_off + 11), name, font=font_header, fill=color_white)
                x_off += w
                
            y_off += 40
            for i, wrapped_vals in enumerate(wrapped_rows):
                bg_color = color_white if i % 2 == 0 else color_iceblue
                current_rh = row_heights[i]
                x_off = 0
                for j, (val, w) in enumerate(zip(wrapped_vals, [cw for _, cw in cols])):
                    draw.rectangle([x_off, y_off, x_off+w, y_off + current_rh], fill=bg_color, outline=color_border)
                    try: text_h = len(val.split('\n')) * 16
                    except Exception: text_h = 16
                    center_y = y_off + (current_rh - text_h) / 2 - 2
                    if j in [2, 8]: 
                        draw.text((x_off + 10, center_y), val, font=font_row, fill='black', spacing=6, align='left') 
                    else:
                        try: text_w = max([self.get_text_width(font_row, line) for line in val.split('\n')])
                        except Exception: text_w = len(val) * 7
                        draw.text((x_off + (w - text_w) / 2, center_y), val, font=font_row, fill='black', spacing=6, align='center') 
                    x_off += w
                y_off += current_rh
                
            report_segments.append(img_table)
            report_segments.append(Image.new('RGB', (table_width, 20), color_white))

        photo_title_img = Image.new('RGB', (table_width, 60), color_white)
        ImageDraw.Draw(photo_title_img).text((15, 20), "📸 현장 조치 확인 (각 항목별 대표 SKU)", font=font_title, fill=color_navy)
        report_segments.append(photo_title_img)
        
        def create_dynamic_arrow(color_hex):
            arr_img = Image.new("RGBA", (100, 40), (255, 255, 255, 0)) 
            draw_arr = ImageDraw.Draw(arr_img)
            draw_arr.rectangle([0, 10, 70, 30], fill=color_hex) 
            draw_arr.polygon([(70, 0), (100, 20), (70, 40)], fill=color_hex) 
            return arr_img

        color_map = {
            "Found": "#FFB300",       
            "Loss": "#212121",        
            "DAMAGED_SKU": "#9E9E9E"  
        }

        photo_blocks = [] 
        
        for i, row in df_final.iterrows():
            b_code = self.clean_text(row[self.barcode_col_name])
            r_type = self.clean_text(row['RESOLVETYPE'])
            
            if b_code != self.selected_barcodes_dict.get(r_type): 
                continue
                
            valid_images = {k: v for k, v in row.get('ATTACHED_IMAGES', {}).items() if v and os.path.exists(v)}
            if valid_images or row.get('FINAL_DIVE_DEEP'): 
                
                if layout_mode == "1col":
                    block_w = table_width
                    img_area_height = 350
                    pad_x = 30
                    inner_start_x = 40
                elif layout_mode == "2col":
                    block_w = half_w
                    img_area_height = 250
                    pad_x = 15
                    inner_start_x = 15
                else: 
                    block_w = third_w
                    img_area_height = 180 
                    pad_x = 10
                    inner_start_x = 10
                
                block_title_wrap = self.force_pixel_wrap(f"No.{row['GLOBAL_RANK']} 바코드: {b_code} [{r_type}]", font_header, block_w - (pad_x*2))
                deep_text_wrap = self.force_pixel_wrap(self.clean_text(row.get('FINAL_DIVE_DEEP','')), font_row, block_w - (pad_x*2) - 30)
                
                title_h = (block_title_wrap.count('\n') + 1) * 20 + 20
                사유_h = (deep_text_wrap.count('\n') + 1) * 20 + 30
                
                block_total_height = title_h + img_area_height + 사유_h + 30 
                
                block_img = Image.new('RGB', (block_w, block_total_height), color_white)
                draw_b = ImageDraw.Draw(block_img)
                
                if layout_mode == "1col":
                    draw_b.rectangle([15, 0, block_w - 16, block_total_height - 10], outline=color_border, width=2)
                    draw_b.text((30, 10), block_title_wrap, font=font_header, fill='black', spacing=4)
                else:
                    draw_b.rectangle([0, 0, block_w - 1, block_total_height - 10], outline=color_border, width=2)
                    draw_b.text((pad_x, 10), block_title_wrap, font=font_header, fill='black', spacing=4)
                
                img_area_y_start = title_h + 10
                current_x = inner_start_x
                
                if valid_images.get("1"):
                    with Image.open(valid_images.get("1")) as loc_img_raw:
                        max_first_photo_width = int(block_w * 0.35)
                        loc_img_raw.thumbnail((max_first_photo_width, img_area_height - 40), Image.Resampling.LANCZOS)
                        img_loc_final = loc_img_raw
                        
                        center_y_main = img_area_y_start + (img_area_height - img_loc_final.height) // 2
                        block_img.paste(img_loc_final, (current_x, center_y_main))
                        current_x += img_loc_final.width + pad_x 
                else: 
                    center_y_no_img = img_area_y_start + (img_area_height // 2) - 10
                    draw_b.text((current_x, center_y_no_img), "[사진 없음]", font=font_header, fill='gray')
                    current_x += int(block_w * 0.35) + pad_x
                    
                defect_val = self.clean_text(row.get('DEFECT_TYPE', 'Found'))
                arrow_color = color_map.get(defect_val, "#FFB300")
                img_arrow_raw = create_dynamic_arrow(arrow_color)
                
                arr_h = 40 if layout_mode == "1col" else (25 if layout_mode == "2col" else 15)
                arrow_resized = img_arrow_raw.resize((int(img_arrow_raw.width * (arr_h / img_arrow_raw.height)), arr_h), Image.Resampling.LANCZOS)
                
                center_y_arrow = img_area_y_start + (img_area_height - arrow_resized.height) // 2
                block_img.paste(arrow_resized, (current_x, center_y_arrow), mask=arrow_resized)
                current_x += arrow_resized.width + pad_x 
                    
                conf_images_paths = [v for k, v in valid_images.items() if k in ["2", "3", "4"]]
                if conf_images_paths:
                    avail_width = block_w - current_x - pad_x
                    per_img_w = int((avail_width - (len(conf_images_paths)-1)*10) / len(conf_images_paths))
                    if per_img_w > 0:
                        for idx, c_path in enumerate(conf_images_paths):
                            with Image.open(c_path) as c_img_raw:
                                c_img_raw.thumbnail((per_img_w, img_area_height - 40), Image.Resampling.LANCZOS)
                                paste_x = current_x + idx*(per_img_w + 10)
                                offset_x = paste_x + max(0, (per_img_w - c_img_raw.width) // 2)
                                font_size_conf = font_row if layout_mode != "3col" else ImageFont.truetype(FONT_PATH, 11)
                                
                                center_y_sub = img_area_y_start + (img_area_height - c_img_raw.height) // 2
                                draw_b.text((offset_x, center_y_sub - 18), f"확인{idx+1}", font=font_size_conf, fill='gray')
                                block_img.paste(c_img_raw, (offset_x, center_y_sub))
                else: 
                    center_y_no_sub = img_area_y_start + (img_area_height // 2) - 10
                    draw_b.text((current_x + 5, center_y_no_sub), "[조치 사진 없음]", font=font_row, fill='gray')
                    
                사유_y = title_h + img_area_height + 10
                
                box_x_start = pad_x
                box_x_end = block_w - pad_x
                draw_b.rectangle([box_x_start, 사유_y, box_x_end, 사유_y + 사유_h - 10], fill='#F7F9FC', outline='#D0D7DE')
                
                text_x = pad_x + 15  
                draw_b.multiline_text((text_x, 사유_y + 10), deep_text_wrap, font=font_row, fill='#1A365D', spacing=6, align='left')
                
                photo_blocks.append(block_img)

        # 블록 병합
        if layout_mode == "1col":
            for blk in photo_blocks:
                report_segments.append(blk)
        elif layout_mode == "2col":
            for i in range(0, len(photo_blocks), 2):
                h1 = photo_blocks[i].height
                h2 = photo_blocks[i+1].height if i+1 < len(photo_blocks) else 0
                row_h = max(h1, h2)
                row_img = Image.new('RGB', (table_width, row_h), color_white)
                row_img.paste(photo_blocks[i], (20, 0))
                if i+1 < len(photo_blocks):
                    row_img.paste(photo_blocks[i+1], (20 + half_w + margin, 0))
                report_segments.append(row_img)
        elif layout_mode == "3col":
            for i in range(0, len(photo_blocks), 3):
                h1 = photo_blocks[i].height
                h2 = photo_blocks[i+1].height if i+1 < len(photo_blocks) else 0
                h3 = photo_blocks[i+2].height if i+2 < len(photo_blocks) else 0
                row_h = max(h1, h2, h3)
                row_img = Image.new('RGB', (table_width, row_h), color_white)
                row_img.paste(photo_blocks[i], (20, 0))
                if i+1 < len(photo_blocks):
                    row_img.paste(photo_blocks[i+1], (20 + third_w + margin, 0))
                if i+2 < len(photo_blocks):
                    row_img.paste(photo_blocks[i+2], (20 + 2*(third_w + margin), 0))
                report_segments.append(row_img)

        img_final_master = Image.new('RGB', (table_width, sum(seg.height for seg in report_segments) + 30), color_white)
        current_y_paste = 0
        for seg in report_segments: 
            img_final_master.paste(seg, (0, current_y_paste))
            current_y_paste += seg.height
            
        final_filename = f"Complete_ICQA_OnePage_Report_{self.date_combo.get()}.png"
        img_final_master.save(final_filename)
        messagebox.showinfo("완료", f"🔥 완벽한 통합 One-Page 리포트가 생성되었습니다!\n폴더에서 '{final_filename}'을 확인하세요.")

    def load_coords(self):
        if os.path.exists(COORD_FILE):
            with open(COORD_FILE, "r") as f: self.coords = json.load(f)

    def save_coords(self):
        with open(COORD_FILE, "w") as f: json.dump(self.coords, f)

    def delete_coord(self, num): 
        self.coords[num] = None
        self.save_coords()
        self.latest_captures[num] = None
        self.bg_captures[num] = None
        self.bi_edit_coords[num] = []
        self.coord_labels[num].configure(text="❌ 미지정", text_color="gray")
        self.btn_edits_bi[num].configure(state="disabled", fg_color="#2B547E")
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
        except Exception: 
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
        self.canvas.coords(self.rect, self.start_x, self.start_y, self.snip_window.winfo_pointerx(), self.snip_window.winfo_pointery())

    def on_release(self, event, num): 
        end_x = self.snip_window.winfo_pointerx()
        end_y = self.snip_window.winfo_pointery()
        self.snip_window.destroy()
        self.deiconify() 
        
        x1, y1 = min(self.start_x, end_x), min(self.start_y, end_y)
        x2, y2 = max(self.start_x, end_x), max(self.start_y, end_y)
        if (x2 - x1) > 10 and (y2 - y1) > 10: 
            self.coords[num] = (int(x1), int(y1), int(x2), int(y2))
            self.save_coords()
            self.coord_labels[num].configure(text="✅ 지정됨", text_color="white")

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
            ctk.CTkButton(frame, text=f"🔍 {i}번 조준", width=100, fg_color="gray", hover_color="dimgray", command=lambda num=str(i): self.show_guide(num)).pack(side="left", padx=5)
            ctk.CTkButton(frame, text=f"📸 찰칵!", width=100, command=lambda num=str(i): self.take_screenshot(num)).pack(side="right", padx=5)
            
        ctk.CTkButton(self.remote, text="❌ 조준선 끄기", fg_color="darkred", hover_color="maroon", command=self.hide_guide).pack(pady=15, fill="x", padx=15)

    def show_guide(self, num): 
        self.hide_guide()
        coord = self.coords[num]
        if coord: 
            x1, y1, x2, y2 = coord
            self.guide_win = tk.Toplevel(self)
            self.guide_win.overrideredirect(True) 
            self.guide_win.attributes("-topmost", True)
            self.guide_win.config(bg="magenta")
            self.guide_win.attributes("-transparentcolor", "magenta")
            self.guide_win.geometry(f"{x2-x1}x{y2-y1}+{x1}+{y1}")
            tk.Canvas(self.guide_win, bg="magenta", highlightthickness=3, highlightbackground="red").pack(fill="both", expand=True)

    def hide_guide(self):
        if self.guide_win: 
            self.guide_win.destroy()
            self.guide_win = None

    def take_screenshot(self, num): 
        coord = self.coords[num]
        if coord: 
            self.hide_guide()
            if self.remote and self.remote.winfo_exists():
                self.remote.withdraw()
            self.withdraw()
            self.after(300, lambda: self._do_capture(num, coord))

    def _do_capture(self, num, coord): 
        try:
            bbox = (int(coord[0]), int(coord[1]), int(coord[2]), int(coord[3]))
            try: img = ImageGrab.grab(bbox=bbox, all_screens=True)
            except Exception: img = ImageGrab.grab(bbox=bbox)
                
            filename = f"Capture_{num}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            img.save(filename)
            
            self.latest_captures[num] = filename
            self.bg_captures[num] = filename
            self.bi_edit_coords[num] = []
            
            self.coord_labels[num].configure(text="✅ 캡처완료", text_color="#00FFCC")
            self.btn_edits_bi[num].configure(state="normal", fg_color="green", hover_color="darkgreen")
            
        except Exception as e:
            messagebox.showerror("캡처 오류", f"캡처 중 문제가 발생했습니다:\n{str(e)}")
        finally:
            if self.remote and self.remote.winfo_exists(): self.remote.deiconify()
            self.deiconify()

# ==========================================
# 🖌️ 둥근 강조 박스 및 사진 회전 에디터 (수정/되돌리기 지원)
# ==========================================
class ImageEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent_win, bg_img_path, existing_coords, on_save_callback):
        super().__init__(parent_win)
        self.title("ICQA 사진 강조 편집기")
        self.protocol("WM_DELETE_WINDOW", self.close_window)
        
        self.bg_img_path = bg_img_path
        self.on_save_callback = on_save_callback 
        
        self.current_pen_color = "#FF0000"
        self.current_line_width = 4  
        self.coords = existing_coords.copy() if existing_coords else []
        
        try: 
            self.original_pil_img = Image.open(bg_img_path).convert('RGB')
        except Exception: 
            messagebox.showerror("이미지 로드 실패", f"{bg_img_path} 로드 실패")
            self.close_window() 
            return
            
        self.setup_ui()
        self.refresh_canvas()

    def setup_ui(self):
        ctk.CTkLabel(self, text="💡 [90도 회전]으로 방향을 맞춘 후 마우스로 드래그하여 강조 네모를 그리세요.", font=("Arial", 14, "bold")).pack(pady=10)
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(pady=5, padx=10, fill="both", expand=True)
        
        self.canvas_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.canvas_frame.pack(side="left", fill="both", expand=True, padx=5)
        
        palette_frame = ctk.CTkFrame(self.main_frame, width=140)
        palette_frame.pack(side="right", fill="y", padx=5)
        
        ctk.CTkLabel(palette_frame, text="🎨 강조 색상", font=("Arial", 13, "bold")).pack(pady=(10, 5))
        colors = [("#FF0000", "빨강"), ("#0000FF", "파랑"), ("#FFFF00", "노랑"), ("#00FF00", "초록")]
        self.color_btns = {}
        for code, name in colors:
            btn = ctk.CTkButton(palette_frame, text=name, font=("Arial", 12), width=100, height=35, fg_color=code, text_color="black" if name in ["노랑", "초록"] else "white", hover_color=code, command=lambda c=code: self.change_pen_color(c))
            btn.pack(pady=3)
            self.color_btns[code] = btn
        self.color_btns[self.current_pen_color].configure(border_width=3, border_color="white")

        ctk.CTkLabel(palette_frame, text="━━━━━━━━━━", text_color="gray").pack(pady=5)
        ctk.CTkLabel(palette_frame, text="📏 선 굵기", font=("Arial", 13, "bold")).pack(pady=(5, 5))
        self.width_label = ctk.CTkLabel(palette_frame, text=f"현재 굵기: {self.current_line_width}", font=("Arial", 12))
        self.width_label.pack(pady=(0, 5))
        
        width_btn_frame = ctk.CTkFrame(palette_frame, fg_color="transparent")
        width_btn_frame.pack(pady=5)
        ctk.CTkButton(width_btn_frame, text="➖", width=40, font=("Arial", 14, "bold"), command=self.decrease_width).pack(side="left", padx=2)
        ctk.CTkButton(width_btn_frame, text="➕", width=40, font=("Arial", 14, "bold"), command=self.increase_width).pack(side="left", padx=2)
        
        ctk.CTkLabel(palette_frame, text="━━━━━━━━━━", text_color="gray").pack(pady=5)
        ctk.CTkButton(palette_frame, text="🔄 90도 회전", font=("Arial", 13, "bold"), width=100, height=40, fg_color="#E56717", hover_color="#C35613", command=self.rotate_image).pack(pady=5)

        bot_frame = ctk.CTkFrame(self, fg_color="transparent")
        bot_frame.pack(side="bottom", fill="x", pady=15, padx=20)
        
        ctk.CTkButton(bot_frame, text="🖍️ 전체 지우기", width=100, fg_color="darkred", hover_color="maroon", command=self.clear_canvas_lines).pack(side="left")
        ctk.CTkButton(bot_frame, text="↩️ 되돌리기", width=100, fg_color="#E56717", hover_color="#C35613", command=self.undo_last_line).pack(side="left", padx=5)
        
        ctk.CTkButton(bot_frame, text="❌ 취소 및 닫기", width=100, fg_color="#454545", command=self.close_window).pack(side="right", padx=5)
        ctk.CTkButton(bot_frame, text="✨ 편집 완료 및 적용 ✨", width=180, height=40, font=("Arial", 14, "bold"), fg_color="green", command=self.save_edits).pack(side="right", padx=15)

    def refresh_canvas(self):
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()
            
        max_w, max_h = 1000, 650
        self.scale_factor = min(max_w / self.original_pil_img.width, max_h / self.original_pil_img.height)
        final_w = int(self.original_pil_img.width * self.scale_factor)
        final_h = int(self.original_pil_img.height * self.scale_factor)
        
        self.display_pil_img = self.original_pil_img.resize((final_w, final_h), Image.Resampling.LANCZOS)
        self.photo_img = ImageTk.PhotoImage(self.display_pil_img)
        
        self.center_window(final_w + 180, final_h + 120)
        
        self.canvas = tk.Canvas(self.canvas_frame, width=final_w, height=final_h, bg="gray", cursor="cross")
        self.canvas.pack()
        
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        
        self.redraw_canvas()

    def redraw_canvas(self):
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.photo_img, anchor="nw")
        
        for coord in self.coords:
            orig_x1, orig_y1, orig_x2, orig_y2 = coord['bbox']
            cx1 = orig_x1 * self.scale_factor
            cy1 = orig_y1 * self.scale_factor
            cx2 = orig_x2 * self.scale_factor
            cy2 = orig_y2 * self.scale_factor
            self.canvas.create_rectangle(cx1, cy1, cx2, cy2, outline=coord['color'], width=coord['width'])

    def rotate_image(self):
        if self.coords:
            ans = messagebox.askyesno("회전 확인", "사진을 회전하면 기존에 그렸던 네모가 모두 초기화됩니다. 계속하시겠습니까?", parent=self)
            if not ans: return
            
        self.original_pil_img = self.original_pil_img.rotate(-90, expand=True)
        self.coords = [] 
        self.refresh_canvas()

    def center_window(self, width, height): 
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.geometry(f"{width}x{height}+{x}+{y}")

    def close_window(self):
        try: self.grab_release() 
        except Exception: pass
        self.withdraw() 
        self.after(10, self.destroy) 

    def change_pen_color(self, color_code): 
        self.color_btns[self.current_pen_color].configure(border_width=0)
        self.current_pen_color = color_code
        self.color_btns[self.current_pen_color].configure(border_width=3, border_color="white")
        
    def decrease_width(self):
        if self.current_line_width > 1:
            self.current_line_width -= 1
            self.width_label.configure(text=f"현재 굵기: {self.current_line_width}")

    def increase_width(self):
        if self.current_line_width < 15:
            self.current_line_width += 1
            self.width_label.configure(text=f"현재 굵기: {self.current_line_width}")

    def on_press(self, event): 
        self.start_x = event.x
        self.start_y = event.y
        self.current_rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline=self.current_pen_color, width=self.current_line_width)

    def on_drag(self, event): 
        self.canvas.coords(self.current_rect_id, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        if (abs(event.x - self.start_x) > 10 and abs(event.y - self.start_y) > 10): 
            orig_x1 = int(self.start_x / self.scale_factor)
            orig_y1 = int(self.start_y / self.scale_factor)
            orig_x2 = int(event.x / self.scale_factor)
            orig_y2 = int(event.y / self.scale_factor)
            
            self.coords.append({
                'bbox': (min(orig_x1, orig_x2), min(orig_y1, orig_y2), max(orig_x1, orig_x2), max(orig_y1, orig_y2)),
                'color': self.current_pen_color,
                'width': self.current_line_width 
            })
        self.redraw_canvas()

    def undo_last_line(self):
        if self.coords:
            self.coords.pop() 
            self.redraw_canvas()

    def clear_canvas_lines(self): 
        self.coords = []
        self.redraw_canvas()

    def save_edits(self):
        try:
            timestamp = datetime.now().strftime('%H%M%S')
            
            new_bg_filename = f"bg_{timestamp}.png"
            new_bg_path = os.path.abspath(new_bg_filename)
            self.original_pil_img.save(new_bg_path)
            
            final_img = self.original_pil_img.copy()
            draw = ImageDraw.Draw(final_img)
            for coord in self.coords: 
                scaled_width = max(1, int(coord['width'] / self.scale_factor))
                draw.rounded_rectangle(coord['bbox'], radius=int(scaled_width * 2), outline=coord['color'], width=scaled_width)
                
            final_filename = f"edited_{timestamp}.png"
            final_path = os.path.abspath(final_filename)
            final_img.save(final_path)
            
            self.on_save_callback(new_bg_path, final_path, self.coords)
            
            messagebox.showinfo("편집 완료", "사진 편집이 적용되었습니다!")
            self.close_window() 
        except Exception as e: 
            messagebox.showerror("저장 실패", f"사진을 저장할 수 없습니다.\n{e}")

if __name__ == "__main__": 
    ICQA_AutoReportApp().mainloop()
