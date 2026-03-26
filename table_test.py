import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl  # 💡 조립 기계가 엑셀 도구를 빼먹지 않고 .exe에 포함하도록 강제 선언!
from PIL import Image, ImageDraw, ImageFont
import os
import ctypes

# 윈도우 디스플레이 배율 무시 (화면 크기 오류 방지)
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class TableGeneratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("표 그리기 엔진 테스트")
        
        # 💡 창을 모니터 정중앙에 띄우기 (가로 400, 세로 200)
        window_width = 400
        window_height = 200
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        ctk.CTkLabel(self, text="📊 엑셀 데이터 -> 표 이미지 변환기", font=("Arial", 18, "bold")).pack(pady=20)
        
        btn = ctk.CTkButton(self, text="📁 Raw Data 엑셀 선택 및 실행", height=50, command=self.process_excel)
        btn.pack(pady=10, padx=20, fill="x")

    def process_excel(self):
        # 1. 엑셀 파일 선택창 띄우기
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return

        try:
            # 2. Pandas로 엑셀 읽기 (엔진을 openpyxl로 확실하게 지정)
            df = pd.read_excel(filepath, engine='openpyxl')
            
            # 필수 열이 있는지 확인
            required_columns = ['RESOLVETYPE', 'EXTERNALID', 'PROBLEM_QTY']
            for col in required_columns:
                if col not in df.columns:
                    messagebox.showerror("오류", f"엑셀 파일에 '{col}' 열이 없습니다!\n컬럼명을 확인해주세요.")
                    return

            # 3. 데이터 그룹화 및 수량 합산
            grouped = df.groupby(['RESOLVETYPE', 'EXTERNALID'])['PROBLEM_QTY'].sum().reset_index()

            # 4. RESOLVETYPE 별로 쪼개서 Top 5 추출 및 이미지 그리기
            resolve_types = grouped['RESOLVETYPE'].unique()
            
            for r_type in resolve_types:
                # 해당 유형의 데이터만 필터링
                type_df = grouped[grouped['RESOLVETYPE'] == r_type].copy()
                
                # 수량 기준으로 내림차순 정렬 후 5개 자르기
                top5_df = type_df.sort_values(by='PROBLEM_QTY', ascending=False).head(5)
                
                # 5. 표 이미지 그리기 함수 호출
                self.draw_table_image(r_type, top5_df)
                
            messagebox.showinfo("완료", "데이터 분석 및 표 이미지 생성이 완료되었습니다!\n프로그램이 있는 폴더를 확인해주세요.")

        except Exception as e:
            messagebox.showerror("에러 발생", f"실행 중 문제가 발생했습니다:\n{str(e)}")

    def draw_table_image(self, resolve_type, df_top5):
        # 한글 폰트 설정 (윈도우 기본 폰트인 맑은 고딕 사용)
        try:
            font_title = ImageFont.truetype("malgunbd.ttf", 20)
            font_header = ImageFont.truetype("malgunbd.ttf", 16)
            font_row = ImageFont.truetype("malgun.ttf", 16)
        except:
            font_title = font_header = font_row = ImageFont.load_default()

        # 표 사이즈 설정
        col_widths = [250, 100] # SKU 열 너비, 수량 열 너비
        table_width = sum(col_widths)
        row_height = 40
        title_height = 50
        
        # 전체 도화지 크기 계산
        total_height = title_height + row_height + (len(df_top5) * row_height)
        
        # 하얀색 빈 도화지 생성
        img = Image.new('RGB', (table_width, total_height), 'white')
        draw = ImageDraw.Draw(img)

        # 🎨 제목 쓰기
        draw.text((10, 10), f"[{resolve_type}] Top 5 SKU", font=font_title, fill='black')

        # 🎨 회색 헤더 그리기
        y_offset = title_height
        draw.rectangle([0, y_offset, table_width, y_offset + row_height], fill='lightgray', outline='gray')
        draw.text((10, y_offset + 10), "EXTERNALID (SKU)", font=font_header, fill='black')
        draw.text((col_widths[0] + 10, y_offset + 10), "수량", font=font_header, fill='black')

        # 🎨 노란색 데이터 줄 그리기
        y_offset += row_height
        for index, row in df_top5.iterrows():
            draw.rectangle([0, y_offset, table_width, y_offset + row_height], fill='#FFF2CC', outline='gray')
            
            sku_text = str(row['EXTERNALID'])
            qty_text = str(row['PROBLEM_QTY'])
            
            draw.text((10, y_offset + 10), sku_text, font=font_row, fill='black')
            draw.text((col_widths[0] + 10, y_offset + 10), qty_text, font=font_row, fill='black')
            
            y_offset += row_height

        # 특수문자 제거하여 파일명 안전하게 만들기
        safe_filename = "".join([c for c in str(resolve_type) if c.isalpha() or c.isdigit() or c in " _-"]).rstrip()
        filename = f"Table_{safe_filename}_Top5.png"
        
        img.save(filename)

if __name__ == "__main__":
    app = TableGeneratorApp()
    app.mainloop()
