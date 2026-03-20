import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# 💡 영어 문제 유형을 이메일용 한글로 번역해 주는 사전
PROBLEM_TRANSLATION = {
    'NO_STOCK': '재고없음',
    'ETC': '기타',
    'MISTAKE': '피커실수',
    'EXISTED_SKU': '상품발견',
    'DAMAGED_SKU': '파손',
    'UNSCANABLE_SKU_BARCODE': '상품 바코드 인식 불가',
    'UNSCANABLE_LOCATION_BARCODE': '위치 바코드 인식 불가'
}

class AutoReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ICQA Auto Report Project_1 - v1.0")
        self.geometry("800x700")
        self.raw_data_path = ""

        # --- 상단: 엑셀 파일 선택 ---
        file_frame = ctk.CTkFrame(self, fg_color="transparent")
        file_frame.pack(pady=(30, 10), padx=20, fill="x")
        
        self.file_label = ctk.CTkLabel(file_frame, text="선택된 파일: 없음", font=ctk.CTkFont(size=14))
        self.file_label.pack(side="left", padx=10)
        
        file_btn = ctk.CTkButton(file_frame, text="📁 Raw Data 엑셀 선택", command=self.select_file)
        file_btn.pack(side="right", padx=10)

        # --- 중단: 수동 입력 값 (비율) ---
        input_frame = ctk.CTkFrame(self, corner_radius=10)
        input_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(input_frame, text="오늘의 문제보고 비율 (% 제외):", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=10, pady=15, sticky="e")
        self.today_ratio_entry = ctk.CTkEntry(input_frame, width=100)
        self.today_ratio_entry.grid(row=0, column=1, padx=10, pady=15, sticky="w")
        self.today_ratio_entry.insert(0, "0.105") # 예시 기본값

        ctk.CTkLabel(input_frame, text="전주 평균 비율 (% 제외):", font=ctk.CTkFont(size=14)).grid(row=0, column=2, padx=10, pady=15, sticky="e")
        self.last_week_ratio_entry = ctk.CTkEntry(input_frame, width=100)
        self.last_week_ratio_entry.grid(row=0, column=3, padx=10, pady=15, sticky="w")
        self.last_week_ratio_entry.insert(0, "0.133") # 예시 기본값

        # --- 하단: 생성 버튼 및 결과 텍스트박스 ---
        gen_btn = ctk.CTkButton(self, text="🚀 이메일 본문 자동 생성", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=self.generate_email_text)
        gen_btn.pack(pady=20, padx=20, fill="x")

        self.result_textbox = ctk.CTkTextbox(self, font=ctk.CTkFont(size=14), height=300)
        self.result_textbox.pack(pady=(0, 20), padx=20, fill="both", expand=True)

    def select_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.raw_data_path = filepath
            self.file_label.configure(text=f"선택된 파일: {filepath.split('/')[-1]}")

    def generate_email_text(self):
        if not self.raw_data_path:
            messagebox.showwarning("경고", "먼저 Raw Data 엑셀 파일을 선택해 주세요.")
            return

        try:
            today_ratio = float(self.today_ratio_entry.get())
            last_week_ratio = float(self.last_week_ratio_entry.get())
            gap = today_ratio - last_week_ratio

            df = pd.read_excel(self.raw_data_path)
            
            total_qty = df['PROBLEM_QTY'].sum()
            
            qty_by_type = df.groupby('PROBLEMTYPE')['PROBLEM_QTY'].sum().reset_index()
            qty_by_type['RATIO'] = (qty_by_type['PROBLEM_QTY'] / total_qty) * 100
            
            qty_by_type = qty_by_type.sort_values(by='RATIO', ascending=False)

            kr_parts = []
            en_parts = []
            
            for index, row in qty_by_type.iterrows():
                eng_type = row['PROBLEMTYPE']
                kr_type = PROBLEM_TRANSLATION.get(eng_type, eng_type) 
                ratio_str = f"{row['RATIO']:.2f}%"
                
                kr_parts.append(f"{kr_type} {ratio_str}")
                en_parts.append(f"{eng_type} {ratio_str}")

            kr_breakdown = " / ".join(kr_parts)
            en_breakdown = " / ".join(en_parts)

            today_str_kr = datetime.now().strftime("%Y년 %m월 %d일")
            today_str_en = datetime.now().strftime("%Y-%m-%d")

            email_template = f"""안녕하세요.
INC26 IC/QA Theo 입니다.

{today_str_kr} 발생한 Picking 문제보고 건에 대해 공유 드립니다.
(Picking 문제보고 기준시간은 00시 00분부터 23시 59분까지입니다.)

1. Picking 문제보고 비율 : {today_ratio:.3f}% ( 전주 평균 : {last_week_ratio:.3f}% / GAP {gap:+.3f}% )
2. 문제 해결 유형 : {kr_breakdown}  

--------------------------------------------------

Hello, this is Theo from INC26 IC/QA.
Please find below for the {today_str_en} Picking Problem Report.
[1day time is Picking Problem Reported From 00:00 to 23:59 and Trends.]

1. Failed Picking ratio : {today_ratio:.3f}% ( Avg. last week : {last_week_ratio:.3f}% / GAP {gap:+.3f}% ) 
2. Problem Resolution type : {en_breakdown}
"""
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", email_template)

        except Exception as e:
            messagebox.showerror("오류", f"데이터 처리 중 문제가 발생했습니다:\n{e}")

if __name__ == "__main__":
    app = AutoReportApp()
    app.mainloop()
