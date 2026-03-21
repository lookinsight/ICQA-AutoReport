import customtkinter as ctk
import tkinter as tk
from PIL import ImageGrab
import json
import os
from datetime import datetime

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# 좌표를 저장할 파일 이름
COORD_FILE = "capture_coords.json"

class CaptureApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ICQA Auto Report - 캡처 리모컨 테스트")
        self.geometry("400x500")
        
        # 5개의 좌표를 저장할 딕셔너리 (초기값은 빈 칸)
        self.coords = {
            "1": None, "2": None, "3": None, "4": None, "5": None
        }
        self.load_coords() # 저장된 좌표가 있으면 불러오기

        ctk.CTkLabel(self, text="[1단계] 5개 캡처 영역 고정하기", font=("Arial", 16, "bold")).pack(pady=(20, 10))
        
        self.coord_labels = {}
        # 1번부터 5번까지 영역 지정 버튼 만들기
        for i in range(1, 6):
            frame = ctk.CTkFrame(self, fg_color="transparent")
            frame.pack(pady=5, fill="x", padx=20)
            
            btn = ctk.CTkButton(frame, text=f"📍 {i}번 영역 지정", width=120, command=lambda num=str(i): self.start_snip(num))
            btn.pack(side="left", padx=10)
            
            status_text = "✅ 지정됨" if self.coords[str(i)] else "❌ 미지정"
            lbl = ctk.CTkLabel(frame, text=status_text)
            lbl.pack(side="left")
            self.coord_labels[str(i)] = lbl

        ctk.CTkLabel(self, text="[2단계] 실전 캡처 리모컨", font=("Arial", 16, "bold")).pack(pady=(30, 10))
        
        remote_btn = ctk.CTkButton(self, text="🎛️ 항상 위 리모컨 띄우기", fg_color="green", hover_color="darkgreen", height=40, command=self.open_remote)
        remote_btn.pack(pady=10, padx=20, fill="x")

    def load_coords(self):
        if os.path.exists(COORD_FILE):
            with open(COORD_FILE, "r") as f:
                self.coords = json.load(f)

    def save_coords(self):
        with open(COORD_FILE, "w") as f:
            json.dump(self.coords, f)

    def start_snip(self, num):
        # 화면을 어둡게 하고 마우스로 영역을 지정하는 투명 창 띄우기
        self.snip_window = tk.Toplevel(self)
        self.snip_window.attributes('-alpha', 0.3) # 반투명도 설정
        self.snip_window.attributes('-fullscreen', True)
        self.snip_window.config(cursor="cross") # 십자 커서
        
        self.snip_window.bind("<ButtonPress-1>", self.on_press)
        self.snip_window.bind("<B1-Motion>", self.on_drag)
        self.snip_window.bind("<ButtonRelease-1>", lambda event: self.on_release(event, num))
        
        self.canvas = tk.Canvas(self.snip_window, cursor="cross", bg="gray")
        self.canvas.pack(fill="both", expand=True)
        self.rect = None
        self.start_x = None
        self.start_y = None

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

        # 좌상단, 우하단 좌표 정리
        x1, y1 = min(self.start_x, end_x), min(self.start_y, end_y)
        x2, y2 = max(self.start_x, end_x), max(self.start_y, end_y)

        # 박스를 너무 작게 치면 무시 (클릭 실수 방지)
        if (x2 - x1) > 10 and (y2 - y1) > 10:
            self.coords[num] = (x1, y1, x2, y2)
            self.save_coords()
            self.coord_labels[num].configure(text="✅ 지정됨 (저장완료)")
            print(f"{num}번 영역 저장됨: {self.coords[num]}")

    def open_remote(self):
        # 🎛️ 항상 위에 떠 있는 미니 리모컨 창
        remote = ctk.CTkToplevel(self)
        remote.title("리모컨")
        remote.geometry("150x300")
        remote.attributes("-topmost", True) # ⭐️ 핵심: 다른 창에 가려지지 않음!

        for i in range(1, 6):
            btn = ctk.CTkButton(remote, text=f"📸 {i}번 찰칵!", command=lambda num=str(i): self.take_screenshot(num))
            btn.pack(pady=10, padx=10, fill="x")

    def take_screenshot(self, num):
        coord = self.coords[num]
        if not coord:
            print(f"{num}번 영역이 아직 지정되지 않았습니다!")
            return
        
        # 지정된 좌표로 화면 캡처
        bbox = (coord[0], coord[1], coord[2], coord[3])
        img = ImageGrab.grab(bbox=bbox, all_screens=True)
        
        # 파일로 저장 (예: 1_20260320_143000.png)
        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{num}_{time_str}.png"
        img.save(filename)
        print(f"[{filename}] 캡처 완료!")

if __name__ == "__main__":
    app = CaptureApp()
    app.mainloop()
