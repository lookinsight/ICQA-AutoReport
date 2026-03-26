import customtkinter as ctk
import tkinter as tk
from PIL import ImageGrab
import json
import os
from datetime import datetime
import ctypes 

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
        self.title("ICQA Auto Report - 캡처 리모컨 (v2.5)")
        
        # 💡 [핵심 1] 메인 프로그램 창을 정중앙에 띄웁니다 (가로 450, 세로 550)
        self.center_window(self, 450, 550)
        
        self.coords = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.load_coords()
        
        self.guide_win = None 
        self.remote = None 

        # --- 메인 UI 부분 ---
        ctk.CTkLabel(self, text="[1단계] 5개 캡처 영역 고정하기", font=("Arial", 16, "bold")).pack(pady=(20, 10))
        self.coord_labels = {}
        
        for i in range(1, 6):
            frame = ctk.CTkFrame(self, fg_color="transparent")
            frame.pack(pady=5, fill="x", padx=20)
            
            btn = ctk.CTkButton(frame, text=f"📍 {i}번 영역 지정", width=120, command=lambda num=str(i): self.start_snip(num))
            btn.pack(side="left", padx=5)
            
            status_text = "✅ 지정됨" if self.coords[str(i)] else "❌ 미지정"
            lbl = ctk.CTkLabel(frame, text=status_text, width=80)
            lbl.pack(side="left", padx=5)
            self.coord_labels[str(i)] = lbl
            
            del_btn = ctk.CTkButton(frame, text="❌ 삭제", width=60, fg_color="darkred", hover_color="maroon", command=lambda num=str(i): self.delete_coord(num))
            del_btn.pack(side="left", padx=5)

        ctk.CTkLabel(self, text="[2단계] 실전 캡처 리모컨", font=("Arial", 16, "bold")).pack(pady=(30, 10))
        remote_btn = ctk.CTkButton(self, text="🎛️ 항상 위 리모컨 띄우기", fg_color="green", hover_color="darkgreen", height=40, command=self.open_remote)
        remote_btn.pack(pady=10, padx=20, fill="x")

    # 💡 [새로운 마법] 창을 화면 한가운데로 옮겨주는 전용 함수입니다.
    def center_window(self, target_window, width, height):
        # 1. 사용자의 모니터 해상도(가로, 세로 픽셀)를 알아냅니다.
        screen_width = target_window.winfo_screenwidth()
        screen_height = target_window.winfo_screenheight()
        
        # 2. 정중앙에 오기 위한 x, y 시작 좌표를 수학적으로 계산합니다.
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        
        # 3. 계산된 크기와 위치로 창을 띄웁니다.
        target_window.geometry(f"{width}x{height}+{x}+{y}")

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

    # --- 🎛️ 리모컨 & 조준선 엔진 ---
    def open_remote(self):
        if self.remote is not None and self.remote.winfo_exists():
            self.remote.focus()
            return

        self.remote = ctk.CTkToplevel(self)
        self.remote.title("리모컨")
        
        # 💡 [핵심 2] 리모컨 창도 띄울 때 정중앙에 띄웁니다 (가로 280, 세로 350)
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
            print(f"{num}번 좌표가 없습니다!")
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
        filename = f"{num}_{time_str}.png"
        img.save(filename)
        print(f"[{filename}] 캡처 완료!")

        if self.remote is not None and self.remote.winfo_exists():
            self.remote.deiconify()
        self.deiconify()

if __name__ == "__main__":
    app = CaptureApp()
    app.mainloop()
