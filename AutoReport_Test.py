import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl 
from PIL import ImageGrab, Image, ImageDraw, ImageFont
import os
import ctypes
import json
import textwrap
from datetime import datetime
import random

# 윈도우 디스플레이 배율 무시
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

COORD_FILE = "capture_coords.json"

class ICQA_AutoReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AutoReport_test")
        self.center_window(self, 550, 850) # UI가 줄었으니 창 높이도 살짝 줄여줌
        
        self.raw_filepath = None
        self.dive_filepath = None
        
        self.coords = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.load_coords()
        self.remote = None 
        self.guide_win = None
        self.barcode_candidates = {} 

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
        frame_range_opt.pack(padx=20, pady=(5, 0), fill="x")
        ctk.CTkLabel(frame_range_opt, text="📋 보고 표 범위:", font=("Arial", 12, "bold"), text_color="#00FFCC").pack(side="left")
        ctk.CTkRadioButton(frame_range_opt, text="Top 5 (기본)", variable=self.report_range, value="top5").pack(side="left", padx=(10, 5))
        ctk.CTkRadioButton(frame_range_opt, text="전체 데이터", variable=self.report_range, value="all").pack(side="left", padx=5)

        self.barcode_mode = ctk.StringVar(value="top1")
        frame_barcode_opt = ctk.CTkFrame(frame_excel, fg_color="transparent")
        frame_barcode_opt.pack(padx=20, pady=(5, 0), fill="x")
        ctk.CTkLabel(frame_barcode_opt, text="👇 바코드 추출 방식:", font=("Arial", 12, "bold"), text_color="yellow").pack(side="left")
        ctk.CTkRadioButton(frame_barcode_opt, text="1위 바코드", variable=self.barcode_mode, value="top1", command=self.update_barcode_text).pack(side="left", padx=(10, 5))
        ctk.CTkRadioButton(frame_barcode_opt, text="🎲랜덤 바코드", variable=self.barcode_mode, value="random", command=self.update_barcode_text).pack(side="left", padx=5)

        self.btn_run = ctk.CTkButton(frame_excel, text="🚀 VLOOKUP 병합 및 Defect Type 선택", fg_color="green", hover_color="darkgreen", height=45, command=self.process_data)
        self.btn_run.pack(pady=15, padx=20, fill="x")

        self.result_box = ctk.CTkTextbox(frame_excel, height=80, font=("Arial", 14))
        self.result_box.pack(padx=20, pady=(5, 10), fill="x")

        # ==========================================
        # 📸 [2단계] 파워 BI 캡처 리모컨 UI
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

    def clean_text(self, text):
        if pd.isna(text): return ""
        cleaned = str(text).strip()
        if cleaned.endswith('.0'): cleaned = cleaned[:-2]
        return cleaned

    def clean_barcode(self, val):
        if pd.isna(val): return ""
        b = str(val).strip().upper()
        if b.endswith('.0'): 
            b = b[:-2]  
        return b

    def force_pixel_wrap(self, text, font, max_width):
        if not text: return ""
        lines = []
        current_line = ""
        for char in str(text):
            test_line = current_line + char
            try:
                w = font.getlength(test_line)
            except:
                try:
                    w = font.getsize(test_line)[0]
                except:
                    bbox = font.getbbox(test_line)
                    w = bbox[2] if bbox else 0
            
            if w > max_width:
                lines.append(current_line)
                current_line = char
            else:
                current_line = test_line
        if current_line:
            lines.append(current_line)
        return "\n".join(lines)

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
            except Exception as e:
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
            has_external_id = 'EXTERNALID' in df_raw.columns
            if has_external_id and 'EXTERNALID' not in req_raw:
                req_raw.append('EXTERNALID')
                
            for col in req_raw:
                if col not in df_raw.columns:
                    messagebox.showerror("오류", f"Raw Data에 '{col}' 열이 없습니다!")
                    return

            df_raw[barcode_col] = df_raw[barcode_col].apply(self.clean_barcode)
            if has_external_id:
                df_raw['EXTERNALID'] = df_raw['EXTERNALID'].astype(str)
                
            df_raw['REPORT_DATE'] = pd.to_datetime(df_raw['REPORT_DATE'], errors='coerce').dt.strftime('%Y-%m-%d')
            df_raw = df_raw[df_raw['REPORT_DATE'] == target_date]
            
            if df_raw.empty:
                messagebox.showinfo("알림", f"Raw Data 파일에 {target_date} 날짜의 데이터가 없습니다.")
                return

            agg_dict = {
                'PROBLEM_QTY': 'sum',
                'MOVED_QTY': 'sum',
                barcode_col: 'count', 
                'DESCRIPTION': 'first' 
            }
            if has_external_id:
                agg_dict['EXTERNALID'] = 'first'

            grouped = df_raw.groupby(['RESOLVETYPE', barcode_col, 'REPORT_DATE']).agg(agg_dict).rename(columns={barcode_col: 'COUNT'}).reset_index()

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
                
                df_dive = pd.concat(valid_dfs, ignore_index=True)

            df_dive['상품바코드'] = df_dive['상품바코드'].apply(self.clean_barcode)
            df_dive[dive_date_col] = pd.to_datetime(df_dive[dive_date_col], errors='coerce').dt.strftime('%Y-%m-%d')
            df_dive = df_dive[df_dive[dive_date_col] == target_date]

            self.final_report_data = []
            self.barcode_candidates = {} 
            
            resolve_types = grouped['RESOLVETYPE'].unique()
            
            for r_type in resolve_types:
                type_df = grouped[grouped['RESOLVETYPE'] == r_type].copy()
                
                if self.report_range.get() == "top5":
                    target_df = type_df.sort_values(by=['PROBLEM_QTY', 'MOVED_QTY'], ascending=[False, False]).head(5)
                else:
                    target_df = type_df.sort_values(by=['PROBLEM_QTY', 'MOVED_QTY'], ascending=[False, False])
                
                self.barcode_candidates[r_type] = target_df[barcode_col].tolist()
                
                merged = pd.merge(
                    target_df, 
                    df_dive, 
                    left_on=[barcode_col, 'REPORT_DATE'], 
                    right_on=['상품바코드', dive_date_col], 
                    how='left'
                )
                
                for index, row in merged.iterrows():
                    ext_id = row['EXTERNALID'] if has_external_id else row[barcode_col]
                    
                    self.final_report_data.append({
                        'RESOLVETYPE': r_type,
                        'RANK': index + 1,
                        'EXTERNAL_ID': self.clean_text(ext_id), 
                        'BARCODE': self.clean_text(row[barcode_col]),
                        'SKU_NAME': self.clean_text(row['DESCRIPTION']),
                        'QTY': self.clean_text(row['PROBLEM_QTY']),
                        'COUNT': self.clean_text(row['COUNT']),
                        'PROB_TYPE': self.clean_text(row['문제유형']),
                        'DIVE_DEEP': self.clean_text(row['사유']),
                        'DEFECT_TYPE': 'Found' 
                    })
            
            self.update_barcode_text()
            self.open_defect_selector()

        except PermissionError:
            messagebox.showerror("권한 에러", "엑셀 파일이 열려있거나 동기화 중입니다!\n열려있는 엑셀을 닫고 다시 시도해주세요.")
        except Exception as e:
            messagebox.showerror("에러 발생", f"실행 중 문제가 발생했습니다:\n{str(e)}")

    def update_barcode_text(self):
        self.result_box.delete("1.0", tk.END)
        mode = self.barcode_mode.get()
        
        for r_type, barcodes in self.barcode_candidates.items():
            if not barcodes: continue
            
            if mode == "top1":
                selected_barcode = barcodes[0] 
            else:
                selected_barcode = random.choice(barcodes) 
                
            clean_b = self.clean_barcode(selected_barcode)
            self.result_box.insert(tk.END, f"[{r_type}] 검색 바코드: {clean_b}\n")

    def open_defect_selector(self):
        self.sel_win = ctk.CTkToplevel(self)
        self.sel_win.title("Defect Type 및 사유 입력 (결재)")
        self.center_window(self.sel_win, 850, 650)
        self.sel_win.attributes("-topmost", True)
        self.sel_win.grab_set() 

        ctk.CTkLabel(self.sel_win, text="📌 Defect Type을 선택하고 누락된 사유를 입력해주세요.", font=("Arial", 16, "bold")).pack(pady=15)

        scroll_frame = ctk.CTkScrollableFrame(self.sel_win, width=800, height=480)
        scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.entries_data = [] 
        
        for i, item in enumerate(self.final_report_data):
            row_frame = ctk.CTkFrame(scroll_frame)
            row_frame.pack(fill="x", pady=5, padx=5)
            
            top_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
            top_frame.pack(fill="x", padx=10, pady=(10, 2))
            
            info_text = f"[{item['RESOLVETYPE']}] 바코드: {item['BARCODE']} | 문제수량: {item['QTY']}"
            ctk.CTkLabel(top_frame, text=info_text, font=("Arial", 14, "bold")).pack(side="left")
            
            combo = ctk.CTkComboBox(top_frame, values=["Found", "Loss", "DAMAGED_SKU"], width=130)
            combo.set("Found")
            combo.pack(side="right")
            
            bot_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
            bot_frame.pack(fill="x", padx=10, pady=(0, 10))
            
            ctk.CTkLabel(bot_frame, text="✏️ 사유 입력/수정:", font=("Arial", 12, "bold"), text_color="#F3E5AB").pack(side="left")
            entry = ctk.CTkEntry(bot_frame, font=("Arial", 13))
            entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
            
            dive_text = item['DIVE_DEEP'] if item['DIVE_DEEP'] else "사유 미기재"
            entry.insert(0, dive_text)
            
            self.entries_data.append((item, combo, entry))

        btn_finish = ctk.CTkButton(self.sel_win, text="✨ 입력 완료 및 표 이미지 생성 ✨", height=50, command=self.generate_final_tables)
        btn_finish.pack(pady=15)

    def generate_final_tables(self):
        for item, combo, entry in self.entries_data:
            item['DEFECT_TYPE'] = combo.get()
            item['DIVE_DEEP'] = entry.get() 
        
        self.sel_win.destroy() 
        
        try:
            font_title = ImageFont.truetype("malgunbd.ttf", 22)
            font_header = ImageFont.truetype("malgunbd.ttf", 14)
            font_row = ImageFont.truetype("malgun.ttf", 13)
        except:
            font_title = font_header = font_row = ImageFont.load_default()

        cols = [
            ("NO", 50), ("External ID", 110), ("SKU Name", 300), 
            ("Problem QTY", 120), ("Problem 건수", 120), 
            ("Problem Type", 120), ("Solve Type", 120), 
            ("Defect Type", 100), ("Dive-Deep", 400)
        ]
        table_width = sum([w for _, w in cols])
        header_height = 40 

        df_final = pd.DataFrame(self.final_report_data)
        for r_type in df_final['RESOLVETYPE'].unique():
            type_data = df_final[df_final['RESOLVETYPE'] == r_type]
            
            raw_title = f"[{r_type}] Problem Analysis"
            
            safe_width = table_width - 40 
            title_wrap = self.force_pixel_wrap(raw_title, font_title, safe_width)
            title_lines = title_wrap.count('\n') + 1
            
            title_height = (title_lines * 40) + 30 

            row_heights = []
            wrapped_rows = []
            
            for _, row in type_data.iterrows():
                sku_wrap = self.force_pixel_wrap(str(row['SKU_NAME']), font_row, cols[2][1] - 30) 
                dive_wrap = self.force_pixel_wrap(str(row['DIVE_DEEP']), font_row, cols[8][1] - 30)
                
                max_lines = max(sku_wrap.count('\n'), dive_wrap.count('\n')) + 1
                calc_height = max(60, (max_lines * 20) + 20) 
                
                row_heights.append(calc_height)
                wrapped_rows.append((row, sku_wrap, dive_wrap))
                
            total_height = title_height + header_height + sum(row_heights) + 20
            
            img = Image.new('RGB', (table_width, total_height), 'white')
            draw = ImageDraw.Draw(img)

            color_navy = '#1A365D'; color_white = '#FFFFFF'; color_iceblue = '#F0F4F8'; color_border = '#808080'

            draw.rectangle([0, 0, table_width-1, total_height-1], outline=color_border, width=2)
            draw.text((15, 20), title_wrap, font=font_title, fill='black', spacing=10)

            y_off = title_height
            draw.rectangle([0, y_off, table_width, y_off + header_height], fill=color_navy, outline=color_border)
            
            x_off = 0
            for name, w in cols:
                draw.rectangle([x_off, y_off, x_off+w, y_off + header_height], outline=color_border)
                
                try:
                    text_bbox = font_header.getbbox(name)
                    text_w = text_bbox[2] - text_bbox[0]
                    text_h = text_bbox[3] - text_bbox[1]
                except:
                    text_w = len(name) * 8 
                    text_h = 14
                    
                center_x = x_off + (w - text_w) / 2
                center_y = y_off + (header_height - text_h) / 2 - 2
                
                draw.text((center_x, center_y), name, font=font_header, fill=color_white)
                x_off += w

            y_off += header_height
            
            for i, (row_data, sku_wrap, dive_wrap) in enumerate(wrapped_rows):
                bg_color = color_white if i % 2 == 0 else color_iceblue
                current_rh = row_heights[i]
                
                row_vals = [
                    str(row_data['RANK']), str(row_data['EXTERNAL_ID']), sku_wrap, str(row_data['QTY']),
                    str(row_data['COUNT']), str(row_data['PROB_TYPE']), str(row_data['RESOLVETYPE']), 
                    str(row_data['DEFECT_TYPE']), dive_wrap
                ]
                
                x_off = 0
                for j, (val, w) in enumerate(zip(row_vals, [cw for _, cw in cols])):
                    draw.rectangle([x_off, y_off, x_off+w, y_off + current_rh], fill=bg_color, outline=color_border)
                    
                    try:
                        lines = val.split('\n')
                        text_w = max([font_row.getlength(line) for line in lines]) if hasattr(font_row, 'getlength') else max([len(line)*7 for line in lines])
                        text_h = len(lines) * 16 
                    except:
                        text_w = len(val) * 7
                        text_h = 16

                    center_y = y_off + (current_rh - text_h) / 2 - 2

                    if j == 2 or j == 8:
                        draw.text((x_off + 15, center_y), val, font=font_row, fill='black', spacing=6) 
                    else:
                        center_x = x_off + (w - text_w) / 2
                        draw.text((center_x, center_y), val, font=font_row, fill='black', spacing=6) 

                    x_off += w
                
                y_off += current_rh

            safe_name = "".join([c for c in r_type if c.isalpha() or c.isdigit() or c in " _-"]).rstrip()
            img.save(f"Report_{safe_name}_Complete.png")

        messagebox.showinfo("완료", "결재 및 대형 표 생성이 완벽하게 끝났습니다!\n프로그램이 있는 폴더에서 'Report_...png'를 확인하세요.")

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
        print(f"[{filename}] 캡처 완료!")

        if self.remote is not None and self.remote.winfo_exists():
            self.remote.deiconify()
        self.deiconify()

if __name__ == "__main__":
    app = ICQA_AutoReportApp()
    app.mainloop()
