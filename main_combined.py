import tkinter as tk
from auto_check_MOP.splash_screen import SplashScreen

def start_mop():
    from auto_check_MOP.auto_check import App as MopApp
    mop_app = MopApp()
    mop_app.mainloop()

def start_flat():
    from ui import ExcelMergerApp
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()

def show_mode_menu():
    menu = tk.Tk()
    menu.title('Выбор режима')
    menu.geometry('400x220')
    menu.configure(bg='#23262b')
    label = tk.Label(menu, text='Выберите режим работы:', font=('Segoe UI', 14), bg='#23262b', fg='#fff')
    label.pack(pady=30)
    btn_mop = tk.Button(menu, text='МОП', font=('Segoe UI', 13, 'bold'), width=16, height=2, bg='#3a3f4b', fg='#fff', command=lambda: (menu.destroy(), start_mop()))
    btn_mop.pack(pady=10)
    btn_flat = tk.Button(menu, text='Квартира', font=('Segoe UI', 13, 'bold'), width=16, height=2, bg='#3a3f4b', fg='#fff', command=lambda: (menu.destroy(), start_flat()))
    btn_flat.pack(pady=10)
    menu.mainloop()

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    def on_splash_close():
        root.destroy()
        show_mode_menu()
    splash = SplashScreen(master=root, gif_path='splash.gif', icon_path='iconn.ico', duration=2.0)
    splash.after(int(splash.duration * 1000), on_splash_close)
    splash.mainloop()
