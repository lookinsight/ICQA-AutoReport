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
        self.title("ICQA Auto Report - 캡처 리모컨 (v2.3)")
        self.geometry("400x500")
        
        self.coords = {"1": None, "2": None, "3": None, "4": None, "5": None}
        self.load_coords()
        
        self.guide_win = None 
        self.remote = None # 💡 리모컨 창을 기억할 변수를 추가했습니다.

        # --- 메인 UI 부분 ---
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
            self.coord_labels[num].configure(text="✅ 지정됨 (저장완료)")

    # --- 🎛️ 리모컨 & 조준선 엔진 ---
    def open_remote(self):
        # 이미 리모컨이 열려있으면 중복해서 열지 않도록 방어합니다.
        if self.remote is not None and self.remote.winfo_exists():
            self.remote.focus()
            return

        self.remote = ctk.CTkToplevel(self) # 💡 self.remote 로 변경했습니다.
        self.remote.title("리모컨")
        self.remote.geometry("280x350") 
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
            
        self.hide_guide() # 1. 조준선 숨기기
        
        # 💡 2. 리모컨 창도 캡처 영역을 가리지 않도록 숨깁니다! (투명 인간 마법)
        if self.remote is not None and self.remote.winfo_exists():
            self.remote.withdraw()
        
        # 창이 완벽하게 사라질 시간을 살짝(0.3초) 주고 캡처를 실행합니다.
        self.after(300, lambda: self._do_capture(num, coord))
        
    def _do_capture(self, num, coord):
        bbox = (coord[0], coord[1], coord[2], coord[3])
        img = ImageGrab.grab(bbox=bbox, all_screens=True)
        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{num}_{time_str}.png"
        img.save(filename)
        print(f"[{filename}] 캡처 완료!")

        # 💡 3. 사진이 다 찍혔으니 숨어있던 리모컨 창을 다시 부릅니다!
        if self.remote is not None and self.remote.winfo_exists():
            self.remote.deiconify()

if __name__ == "__main__":
    app = CaptureApp()
    app.mainloop()
