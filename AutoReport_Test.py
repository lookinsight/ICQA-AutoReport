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
            self.btn_raw.configure(text="✅ 1. Raw Data 선택 완료", fg_color="gray")

    def load_dive_data(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.dive_filepath = filepath
            self.btn_dive.configure(text="✅ 2. Dive-Deep 선택 완료", fg_color="gray")

    def process_data(self):
        if not self.raw_filepath or not self.dive_filepath:
            messagebox.showwarning("경고", "Raw Data와 Dive-Deep 엑셀 파일을 모두 선택해주세요!")
            return

        try:
            self.result_box.delete("1.0", tk.END)
            
            with open(self.raw_filepath, 'rb') as f:
