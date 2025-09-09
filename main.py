import sys
# --- PyInstaller resource path helper ---
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.abspath(relative_path)
import tkinter as tk
from PIL import Image, ImageTk, ImageSequence
import os
import time
from ui import ExcelMergerApp
from styles import DARK_BG, DARK_FG, DARK_ACCENT, MENU_WIDTH, MENU_HEIGHT, MENU_LABEL_FONT, MENU_BTN_FONT, MENU_BTN_WIDTH, MENU_BTN_HEIGHT, MENU_BTN_PADX, MENU_LABEL_PADY, MENU_BTN_PADY

# --- SplashScreen класс (адаптирован из auto_check_MOP) ---
class SplashScreen(tk.Toplevel):
    def __init__(self, master=None, gif_path=None, icon_path=None, duration=2.0, **kwargs):
        super().__init__(master, **kwargs)
        self.overrideredirect(True)
        self.configure(bg='#23262b')
        self.duration = duration
        self.frames = []
        self.frame_index = 0
        self.label = tk.Label(self, bg='#23262b')
        self.label.pack()
        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass
        if gif_path:
            self.load_gif(gif_path)
        self.attributes('-alpha', 0.0)
        self.fade_in()
        self.after(0, self.play_gif)
        self.after(int(self.duration * 1000), self.fade_out)

    def load_gif(self, gif_path):
        im = Image.open(gif_path)
        self.frames = []
        width, height = 340, 300
        for frame in ImageSequence.Iterator(im):
            orig = frame.copy().convert('RGBA')
            stretched = orig.resize((width, height), Image.LANCZOS)
            self.frames.append(ImageTk.PhotoImage(stretched))
        self.geometry(f"{width}x{height}+{self.winfo_screenwidth()//2-width//2}+{self.winfo_screenheight()//2-height//2}")

    def fade_in(self):
        for alpha in range(0, 101, 5):
            self.attributes('-alpha', alpha/100)
            self.update()
            time.sleep(0.01)

    def play_gif(self):
        if not hasattr(self, '_playing'):
            self._playing = True
        if self.frames and self._playing:
            self.label.config(image=self.frames[self.frame_index])
            self.frame_index = (self.frame_index + 1) % len(self.frames)
            self._after_id = self.after(120, self.play_gif)

    def fade_out(self):
        self._playing = False
        if hasattr(self, '_after_id'):
            self.after_cancel(self._after_id)
        for alpha in range(100, -1, -5):
            try:
                self.attributes('-alpha', alpha/100)
                self.update()
                time.sleep(0.01)
            except Exception:
                break
        self.destroy()

# --- Меню выбора режима ---
def show_mode_menu():
    menu = tk.Tk()
    menu.title('Выбор режима')
    menu_width = MENU_WIDTH
    menu_height = MENU_HEIGHT
    screen_width = menu.winfo_screenwidth()
    screen_height = menu.winfo_screenheight()
    x = (screen_width // 2) - (menu_width // 2)
    y = (screen_height // 2) - (menu_height // 2)
    menu.geometry(f'{menu_width}x{menu_height}+{x}+{y}')
    menu.configure(bg=DARK_BG)
    # Поднять окно поверх всех окон и сделать активным (только один раз, без дубликатов)
    menu.lift()
    menu.attributes('-topmost', True)
    menu.focus_force()
    menu.after(100, lambda: menu.attributes('-topmost', False))
    label = tk.Label(menu, text='Выберите режим работы:', font=MENU_LABEL_FONT, bg=DARK_BG, fg=DARK_FG)
    label.pack(pady=MENU_LABEL_PADY)
    def start_mop():
        menu.withdraw()
        from mop_app import MopApp
        mop_win = MopApp(menu)
        mop_win.wait_window()
        menu.deiconify()
        menu.lift()
        menu.focus_force()
    def start_flat():
        menu.withdraw()
        from ui import ExcelMergerApp
        flat_win = tk.Toplevel(menu)
        app = ExcelMergerApp(flat_win)
        flat_win.wait_window()
        menu.deiconify()
        menu.lift()
        menu.focus_force()
    btn_mop = tk.Button(menu, text='МОП (места общего пользования)', font=MENU_BTN_FONT, width=MENU_BTN_WIDTH, height=MENU_BTN_HEIGHT, bg=DARK_ACCENT, fg=DARK_FG, command=start_mop)
    btn_mop.pack(pady=MENU_BTN_PADY, fill='x', padx=MENU_BTN_PADX)
    btn_flat = tk.Button(menu, text='Квартира (жилые помещения)', font=MENU_BTN_FONT, width=MENU_BTN_WIDTH, height=MENU_BTN_HEIGHT, bg=DARK_ACCENT, fg=DARK_FG, command=start_flat)
    btn_flat.pack(pady=MENU_BTN_PADY, fill='x', padx=MENU_BTN_PADX)
    menu.mainloop()

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    splash = SplashScreen(master=root, gif_path=resource_path('splash.gif'), icon_path=resource_path('iconn.ico'), duration=2.0)
    def after_splash():
        root.destroy()
        show_mode_menu()
    splash.after(int(splash.duration * 1000), after_splash)
    splash.mainloop()

