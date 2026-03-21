import customtkinter as ctk
import tkinter as tk
from PIL import ImageGrab
import json
import os
from datetime import datetime
import ctypes # 💡 새로 추가할 도구!

# 💡 윈도우 디스플레이 배율(확대) 무시하고 모니터 전체 크기를 100% 덮도록 강제하는 마법 주문
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass
    
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

COORD_FILE = "capture_coords.json"

class CaptureApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ICQA Auto Report - 캡처 리모컨 (v2.2)")
        self.geometry("400x500")
        
        self.coords = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.load_coords()
        
        self.guide_win = None # 💡 빨간색 조준선 창을 기억해둘 변수 추가

        # --- 메인 UI 부분 (이전과 동일) ---
        ctk.CTkLabel(self, text="[1단계] 5개 캡처 영역 고정하기", font=("Arial", 16, "bold")).pack(pady=(20, 10))
        self.coord_labels = {}
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
        self.withdraw() # 본체 창 숨기기
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
        self.deiconify() # 본체 창 다시 띄우기

        x1, y1 = min(self.start_x, end_x), min(self.start_y, end_y)
        x2, y2 = max(self.start_x, end_x), max(self.start_y, end_y)

        if (x2 - x1) > 10 and (y2 - y1) > 10:
            self.coords[num] = (x1, y1, x2, y2)
            self.save_coords()
            self.coord_labels[num].configure(text="✅ 지정됨 (저장완료)")

    # --- 🎛️ v2.2 업그레이드 리모컨 & 조준선 엔진 ---
    def open_remote(self):
        remote = ctk.CTkToplevel(self)
        remote.title("리모컨")
        remote.geometry("280x350") # 가로 폭을 살짝 넓혔습니다.
        remote.attributes("-topmost", True)

        for i in range(1, 6):
            frame = ctk.CTkFrame(remote, fg_color="transparent")
            frame.pack(pady=5, padx=10, fill="x")
            
            # 💡 [버튼 1] 조준선 켜기 버튼 (회색)
            btn_aim = ctk.CTkButton(frame, text=f"🔍 {i}번 조준", width=100, fg_color="gray", hover_color="dimgray", command=lambda num=str(i): self.show_guide(num))
            btn_aim.pack(side="left", padx=5)
            
            # 💡 [버튼 2] 진짜 사진 찍기 버튼 (파란색)
            btn_shot = ctk.CTkButton(frame, text=f"📸 찰칵!", width=100, command=lambda num=str(i): self.take_screenshot(num))
            btn_shot.pack(side="right", padx=5)
            
        # 조준선을 수동으로 지우는 버튼 추가
        btn_clear = ctk.CTkButton(remote, text="❌ 조준선 끄기", fg_color="darkred", hover_color="maroon", command=self.hide_guide)
        btn_clear.pack(pady=15, fill="x", padx=15)

    def show_guide(self, num):
        self.hide_guide() # 다른 조준선이 켜져 있으면 먼저 지웁니다.
        
        coord = self.coords[num]
        if not coord:
            print(f"{num}번 좌표가 없습니다!")
            return
            
        x1, y1, x2, y2 = coord
        w = x2 - x1
        h = y2 - y1
        
        # 조준선 창 만들기
        self.guide_win = tk.Toplevel(self)
        self.guide_win.overrideredirect(True) # 윈도우 창의 상단 X(닫기) 바를 없앰
        self.guide_win.attributes("-topmost", True) # 항상 위에 고정
        
        # 💡 [핵심 마법] 특정 색상(마젠타)을 완전 투명하게 만들고, 마우스 클릭을 통과시킵니다!
        transparent_color = "magenta"
        self.guide_win.config(bg=transparent_color)
        self.guide_win.attributes("-transparentcolor", transparent_color)
        
        # 지정된 좌표와 크기로 조준선 이동
        self.guide_win.geometry(f"{w}x{h}+{x1}+{y1}")
        
        # 빨간색 테두리 그리기
        canvas = tk.Canvas(self.guide_win, bg=transparent_color, highlightthickness=3, highlightbackground="red")
        canvas.pack(fill="both", expand=True)

    def hide_guide(self):
        # 조준선 창을 파괴(삭제)합니다.
        if self.guide_win:
            self.guide_win.destroy()
            self.guide_win = None

    def take_screenshot(self, num):
        coord = self.coords[num]
        if not coord:
            return
            
        # 💡 [가장 중요한 부분] 사진에 빨간 테두리가 찍히면 안 되니까, 찰칵! 하기 직전에 조준선을 지워버립니다.
        self.hide_guide()
        
        # 컴퓨터가 창을 지우는 시간을 아주 잠깐(0.2초) 벌어준 뒤에 진짜 카메라 셔터를 누릅니다.
        self.after(200, lambda: self._do_capture(num, coord))
        
    def _do_capture(self, num, coord):
        bbox = (coord[0], coord[1], coord[2], coord[3])
        img = ImageGrab.grab(bbox=bbox, all_screens=True)
        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{num}_{time_str}.png"
        img.save(filename)
        print(f"[{filename}] 캡처 완료!")

if __name__ == "__main__":
    app = CaptureApp()
    app.mainloop()
