import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl 
from PIL import Image, ImageDraw, ImageFont
import os
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class TableGeneratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ICQA 표 그리기 엔진 (v1.4)")
        
        window_width = 450
        window_height = 450
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        ctk.CTkLabel(self, text="📊 엑셀 데이터 -> 표 이미지 변환기", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        
        btn = ctk.CTkButton(self, text="📁 Raw Data 엑셀 선택 및 실행", height=50, command=self.process_excel)
        btn.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(self, text="👇 [카톡 검색용] 각 유형별 1위 바코드 (복사하세요)", font=("Arial", 12, "bold"), text_color="yellow").pack(pady=(15, 5))
        self.result_box = ctk.CTkTextbox(self, width=400, height=150, font=("Arial", 14))
        self.result_box.pack(padx=20, pady=5)

    def clean_text(self, text):
        cleaned = str(text)
        if cleaned.endswith('.0'):
            cleaned = cleaned[:-2]
        return cleaned

    def process_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return

        try:
            self.result_box.delete("1.0", tk.END)
            target_barcode_col = 'BARCODE' 
            
            # 💡 [핵심 해결책] 엑셀이 열려있거나 동기화 중이어도, 강제로 복사본(rb) 모드로 안전하게 읽어옵니다!
            with open(filepath, 'rb') as f:
                df = pd.read_excel(f, engine='openpyxl', dtype={target_barcode_col: str})
            
            required_columns = ['RESOLVETYPE', target_barcode_col, 'PROBLEM_QTY']
            for col in required_columns:
                if col not in df.columns:
                    messagebox.showerror("오류", f"엑셀 파일에 '{col}' 열이 없습니다!\n실제 엑셀의 열 이름을 확인해주세요.")
                    return

            grouped = df.groupby(['RESOLVETYPE', target_barcode_col])['PROBLEM_QTY'].sum().reset_index()
            resolve_types = grouped['RESOLVETYPE'].unique()
            
            for r_type in resolve_types:
                type_df = grouped[grouped['RESOLVETYPE'] == r_type].copy()
                top5_df = type_df.sort_values(by='PROBLEM_QTY', ascending=False).head(5)
                
                self.draw_table_image(r_type, top5_df, target_barcode_col)
                
                if not top5_df.empty:
                    clean_barcode = self.clean_text(top5_df.iloc[0][target_barcode_col])
                    self.result_box.insert(tk.END, f"[{r_type}] 검색 바코드: {clean_barcode}\n")
                
            messagebox.showinfo("완료", "표 이미지 생성이 완료되었습니다!\n화면 아래의 바코드를 복사해서 카톡에 검색하세요.")

        except PermissionError:
            # 혹시라도 이 마법마저 뚫고 권한 에러가 나면, 친절한 한글 안내창을 띄웁니다.
            messagebox.showerror("권한 에러", "엑셀 파일이 열려있거나 동기화 중입니다!\n\n1. 열려있는 엑셀 창을 꺼주세요.\n2. 5초 정도 기다렸다가 다시 시도해주세요.")
        except Exception as e:
            messagebox.showerror("에러 발생", f"실행 중 문제가 발생했습니다:\n{str(e)}")

    def draw_table_image(self, resolve_type, df_top5, barcode_col):
        try:
            font_title = ImageFont.truetype("malgunbd.ttf", 20)
            font_header = ImageFont.truetype("malgunbd.ttf", 16)
            font_row = ImageFont.truetype("malgun.ttf", 16)
        except:
            font_title = font_header = font_row = ImageFont.load_default()

        col_widths = [250, 100] 
        table_width = sum(col_widths)
        row_height = 40
        title_height = 50
        
        total_height = title_height + row_height + (len(df_top5) * row_height)
        
        img = Image.new('RGB', (table_width, total_height), 'white')
        draw = ImageDraw.Draw(img)

        draw.text((10, 10), f"[{resolve_type}] Top 5", font=font_title, fill='black')

        y_offset = title_height
        draw.rectangle([0, y_offset, table_width, y_offset + row_height], fill='lightgray', outline='gray')
        draw.text((10, y_offset + 10), "바코드 (BARCODE)", font=font_header, fill='black')
        draw.text((col_widths[0] + 10, y_offset + 10), "수량", font=font_header, fill='black')

        y_offset += row_height
        for index, row in df_top5.iterrows():
            draw.rectangle([0, y_offset, table_width, y_offset + row_height], fill='#FFF2CC', outline='gray')
            
            sku_text = self.clean_text(row[barcode_col])
            qty_text = self.clean_text(row['PROBLEM_QTY'])
            
            draw.text((10, y_offset + 10), sku_text, font=font_row, fill='black')
            draw.text((col_widths[0] + 10, y_offset + 10), qty_text, font=font_row, fill='black')
            
            y_offset += row_height

        draw.line([(col_widths[0], title_height), (col_widths[0], total_height)], fill='gray', width=1)

        safe_filename = "".join([c for c in str(resolve_type) if c.isalpha() or c.isdigit() or c in " _-"]).rstrip()
        filename = f"Table_{safe_filename}_Top5.png"
        
        img.save(filename)

if __name__ == "__main__":
    app = TableGeneratorApp()
    app.mainloop()
