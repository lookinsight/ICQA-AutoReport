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
        self.center_window(self, 550, 800)
        
        self.raw_filepath = None
        self.dive_filepath = None
        
        self.coords = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.load_coords()
        self.remote = None 
        self.guide_win = None

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

        # 💡 [새 기능!] 보고 날짜 선택기 추가
        frame_date = ctk.CTkFrame(frame_excel, fg_color="transparent")
        frame_date.pack(pady=(10, 5), padx=20, fill="x")
        ctk.CTkLabel(frame_date, text="📅 보고 대상 날짜:", font=("Arial", 14, "bold")).pack(side="left", padx=(0, 10))
        self.date_combo = ctk.CTkComboBox(frame_date, values=["Raw Data를 먼저 넣어주세요"], width=180)
        self.date_combo.pack(side="left")

        self.btn_run = ctk.CTkButton(frame_excel, text="🚀 VLOOKUP 병합 및 Defect Type 선택", fg_color="green", hover_color="darkgreen", height=45, command=self.process_data)
        self.btn_run.pack(pady=15, padx=20, fill="x")

        ctk.CTkLabel(frame_excel, text="👇 [카톡 검색용] 각 유형별 1위 바코드", font=("Arial", 12, "bold"), text_color="yellow").pack(pady=(0, 2))
        self.result_box = ctk.CTkTextbox(frame_excel, height=80, font=("Arial", 14))
        self.result_box.pack(padx=20, pady=(0, 10), fill="x")

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

        # ==========================================
        # 💡 우리만의 시그니처 워터마크
        # ==========================================
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

    def load_raw_data(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.raw_filepath = filepath
            filename = os.path.basename(filepath)
            self.btn_raw.configure(text=f"✅ {filename} (클릭하여 변경)", fg_color="#454545")
            
            # 💡 [핵심] Raw Data를 쓱 읽어서 존재하는 날짜들을 콤보박스에 넣어줍니다!
            try:
                df = pd.read_excel(filepath, engine='openpyxl')
                df.columns = df.columns.str.strip().str.replace('\n', '')
                if 'REPORT_DATE' in df.columns:
                    # 날짜 형식으로 변환 후 YYYY-MM-DD 모양만 추출
                    dates = pd.to_datetime(df['REPORT_DATE'], errors='coerce').dt.strftime('%Y-%m-%d').dropna().unique().tolist()
                    dates.sort(reverse=True) # 최신 날짜가 맨 위로 오게 정렬
                    if dates:
                        self.date_combo.configure(values=dates)
                        self.date_combo.set(dates[0]) # 제일 최신 날짜로 자동 세팅
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
            self.result_box.delete("1.0", tk.END)
            
            with open(self.raw_filepath, 'rb') as f:
                df_raw = pd.read_excel(f, engine='openpyxl')
            
            df_raw.columns = df_raw.columns.str.strip().str.replace('\n', '')
            
            # 💡 BARCODE를 우선적으로 찾습니다! (Dive-Deep의 상품바코드와 매칭 확률 높이기)
            barcode_col = 'BARCODE' if 'BARCODE' in df_raw.columns else ('EXTERNALID' if 'EXTERNALID' in df_raw.columns else None)
            
            if not barcode_col:
                messagebox.showerror("오류", "Raw Data에 'BARCODE' 또는 'EXTERNALID' 열이 없습니다!")
                return
                
            df_raw[barcode_col] = df_raw[barcode_col].astype(str)

            req_raw = ['RESOLVETYPE', barcode_col, 'PROBLEM_QTY', 'MOVED_QTY', 'DESCRIPTION', 'REPORT_DATE']
            for col in req_raw:
                if col not in df_raw.columns:
                    messagebox.showerror("오류", f"Raw Data에 '{col}' 열이 없습니다!\n(열 이름 띄어쓰기를 확인해주세요)")
                    return

            # 날짜 변환 및 💡타겟 날짜만 쏙 필터링!
            df_raw['REPORT_DATE'] = pd.to_datetime(df_raw['REPORT_DATE'], errors='coerce').dt.strftime('%Y-%m-%d')
            df_raw = df_raw[df_raw['REPORT_DATE'] == target_date]
            
            if df_raw.empty:
                messagebox.showinfo("알림", f"선택하신 날짜({target_date})의 데이터가 Raw 파일에 존재하지 않습니다.")
                return

            grouped = df_raw.groupby(['RESOLVETYPE', barcode_col, 'REPORT_DATE']).agg({
                'PROBLEM_QTY': 'sum',
                'MOVED_QTY': 'sum',
                barcode_col: 'count', 
                'DESCRIPTION': 'first' 
            }).rename(columns={barcode_col: 'COUNT'}).reset_index()

            with open(self.dive_filepath, 'rb') as f:
                df_dive = pd.read_excel(f, engine='openpyxl')
            
            df_dive.columns = df_dive.columns.str.strip().str.replace('\n', '')
            
            if '상품바코드' in df_dive.columns:
                df_dive['상품바코드'] = df_dive['상품바코드'].astype(str)
            
            dive_date_col = 'Date' 
            req_dive = ['상품바코드', '문제유형', '사유', dive_date_col]
            for col in req_dive:
                if col not in df_dive.columns:
                    messagebox.showerror("오류", f"Dive-Deep 파일에 '{col}' 열이 없습니다!\n(실제 열 이름: {list(df_dive.columns)})")
                    return

            # 날짜 변환 및 💡타겟 날짜만 쏙 필터링!
            df_dive[dive_date_col] = pd.to_datetime(df_dive[dive_date_col], errors='coerce').dt.strftime('%Y-%m-%d')
            df_dive = df_dive[df_dive[dive_date_col] == target_date]

            self.final_report_data = []
            resolve_types = grouped['RESOLVETYPE'].unique()
            
            for r_type in resolve_types:
                type_df = grouped[grouped['RESOLVETYPE'] == r_type].copy()
                top5_df = type_df.sort_values(by=['PROBLEM_QTY', 'MOVED_QTY'], ascending=[False, False]).head(5)
                
                merged = pd.merge(
                    top5_df, 
                    df_dive, 
                    left_on=[barcode_col, 'REPORT_DATE'], 
                    right_on=['상품바코드', dive_date_col], 
                    how='left'
                )
                
                for index, row in merged.iterrows():
                    self.final_report_data.append({
                        'RESOLVETYPE': r_type,
                        'RANK': index + 1,
                        'BARCODE': self.clean_text(row[barcode_col]),
                        'SKU_NAME': self.clean_text(row['DESCRIPTION']),
                        'QTY': self.clean_text(row['PROBLEM_QTY']),
                        'COUNT': self.clean_text(row['COUNT']),
                        'PROB_TYPE': self.clean_text(row['문제유형']),
                        'DIVE_DEEP': self.clean_text(row['사유']),
                        'DEFECT_TYPE': 'Found' 
                    })
                
                if not top5_df.empty:
                    clean_barcode = self.clean_text(top5_df.iloc[0][barcode_col])
                    self.result_box.insert(tk.END, f"[{r_type}] 검색 바코드: {clean_barcode}\n")

            self.open_defect_selector()

        except PermissionError:
            messagebox.showerror("권한 에러", "엑셀 파일이 열려있거나 동기화 중입니다!\n열려있는 엑셀을 닫고 다시 시도해주세요.")
        except Exception as e:
            messagebox.showerror("에러 발생", f"실행 중 문제가 발생했습니다:\n{str(e)}")

    def open_defect_selector(self):
        self.sel_win = ctk.CTkToplevel(self)
        self.sel_win.title("Defect Type 선택 (결재)")
        self.center_window(self.sel_win, 800, 600)
        self.sel_win.attributes("-topmost", True)
        self.sel_win.grab_set() 

        ctk.CTkLabel(self.sel_win, text="📌 추출된 바코드의 Defect Type을 선택해주세요.", font=("Arial", 16, "bold")).pack(pady=15)

        scroll_frame = ctk.CTkScrollableFrame(self.sel_win, width=750, height=450)
        scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.comboboxes = []
        
        for i, item in enumerate(self.final_report_data):
            row_frame = ctk.CTkFrame(scroll_frame)
            row_frame.pack(fill="x", pady=5, padx=5)
            
            dive_text = item['DIVE_DEEP'][:15] + "..." if item['DIVE_DEEP'] else "사유 미기재(또는 날짜 불일치)"
            info_text = f"[{item['RESOLVETYPE']}] 바코드: {item['BARCODE']} | 수량: {item['QTY']} | 사유: {dive_text}"
            
            ctk.CTkLabel(row_frame, text=info_text, width=500, anchor="w").pack(side="left", padx=10, pady=10)
            
            combo = ctk.CTkComboBox(row_frame, values=["Found", "Loss", "DAMAGED_SKU"], width=150)
            combo.set("Found")
            combo.pack(side="right", padx=10, pady=10)
            self.comboboxes.append((item, combo))

        btn_finish = ctk.CTkButton(self.sel_win, text="✨ 선택 완료 및 표 이미지 생성 ✨", height=50, command=self.generate_final_tables)
        btn_finish.pack(pady=15)

    def generate_final_tables(self):
        for item, combo in self.comboboxes:
            item['DEFECT_TYPE'] = combo.get()
        
        self.sel_win.destroy() 
        
        try:
            font_title = ImageFont.truetype("malgunbd.ttf", 22)
            font_header = ImageFont.truetype("malgunbd.ttf", 14)
            font_row = ImageFont.truetype("malgun.ttf", 13)
        except:
            font_title = font_header = font_row = ImageFont.load_default()

        cols = [
            ("NO", 50), ("External ID", 110), ("SKU Name", 300), ("Problem QTY", 90), 
            ("Problem 건수", 90), ("Problem Type", 120), ("Solve Type", 120), 
            ("Defect Type", 100), ("Dive-Deep", 400)
        ]
        table_width = sum([w for _, w in cols])
        row_height = 60 
        title_height = 50

        df_final = pd.DataFrame(self.final_report_data)
        for r_type in df_final['RESOLVETYPE'].unique():
            type_data = df_final[df_final['RESOLVETYPE'] == r_type]
            
            total_height = title_height + row_height + (len(type_data) * row_height)
            img = Image.new('RGB', (table_width, total_height), 'white')
            draw = ImageDraw.Draw(img)

            color_navy = '#1A365D'; color_white = '#FFFFFF'; color_iceblue = '#F0F4F8'; color_border = '#808080'

            draw.text((10, 10), f"[{r_type}] Problem Analysis", font=font_title, fill='black')

            y_off = title_height
            draw.rectangle([0, y_off, table_width, y_off + row_height], fill=color_navy, outline=color_border)
            
            x_off = 0
            for name, w in cols:
                draw.rectangle([x_off, y_off, x_off+w, y_off + row_height], outline=color_border)
                draw.text((x_off + 10, y_off + 20), name, font=font_header, fill=color_white)
                x_off += w

            y_off += row_height
            for i, (_, row) in enumerate(type_data.iterrows()):
                bg_color = color_white if i % 2 == 0 else color_iceblue
                
                sku_wrap = textwrap.fill(row['SKU_NAME'], width=22)
                dive_wrap = textwrap.fill(row['DIVE_DEEP'], width=35)
                
                row_data = [
                    str(row['RANK']), str(row['BARCODE']), sku_wrap, str(row['QTY']),
                    str(row['COUNT']), row['PROB_TYPE'], row['RESOLVETYPE'], 
                    row['DEFECT_TYPE'], dive_wrap
                ]
                
                x_off = 0
                for j, (val, w) in enumerate(zip(row_data, [cw for _, cw in cols])):
                    draw.rectangle([x_off, y_off, x_off+w, y_off + row_height], fill=bg_color, outline=color_border)
                    draw.text((x_off + 10, y_off + 10), val, font=font_row, fill='black')
                    x_off += w
                
                y_off += row_height

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
        self.snip_window.attributes('-fullscreen', True)
        self.snip_window.config(cursor="cross")
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
