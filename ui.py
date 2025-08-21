import tkinter as tk
from tkinter import filedialog, messagebox
from styles import DARK_BG, DARK_FG, DARK_ACCENT, DARK_BTN_BG, DARK_BTN_FG, DARK_ENTRY_BG, DARK_ENTRY_FG, DARK_HIGHLIGHT
from logic import (
    get_sheet_names_from_file,
    get_sheet_names_from_folder,
    merge_sheets_in_file,
    merge_all_files_in_folder,
    preview_merge_file,
    preview_merge_folder
)

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel Мерджер')
        # Центрируем окно
        window_width = 1000
        window_height = 600
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        self.root.configure(bg=DARK_BG)

        self.preview_df = None
        self.sheet_list = []
        self.last_result_path = None
        self.last_merge_file_path = None
        self.last_merge_folder_path = None

        # Splash Frame
        self.splash_frame = tk.Frame(root, width=340, height=300, bg=DARK_BG)
        self.splash_frame.place(relx=0.5, rely=0.5, anchor='center')
        self.splash_label = tk.Label(self.splash_frame, bg=DARK_BG)
        self.splash_label.pack()
        self.splash_frames = []
        self.splash_index = 0
        self.splash_running = True
        self.load_splash_gif()
        self.animate_splash()

        # Основной интерфейс (скрыт при старте)
        self.main_frame = tk.Frame(root, bg=DARK_BG)
        self.main_frame.pack(fill='both', expand=True)
        self.main_frame.pack_forget()

        # Верхняя панель с кнопками
        top_frame = tk.Frame(self.main_frame, bg=DARK_BG)
        top_frame.pack(side='top', fill='x')
        self.btn_file = tk.Button(top_frame, text='Объединить листы в файле', command=self.choose_file_and_merge, bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_file.pack(side='left', padx=5, pady=5)
        self.btn_folder = tk.Button(top_frame, text='Объединить файлы в папке', command=self.choose_folder_and_merge, bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_folder.pack(side='left', padx=5, pady=5)
        self.btn_export = tk.Button(top_frame, text='Выгрузить в Excel', command=self.export_to_excel, state='disabled', bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_export.pack(side='left', padx=5, pady=5)

        # Список листов слева с чекбоксами
        left_frame = tk.Frame(self.main_frame, width=220, bg=DARK_BG)
        left_frame.pack(side='left', fill='y')
        tk.Label(left_frame, text='Листы', bg=DARK_BG, fg=DARK_FG).pack(anchor='nw')
        self.sheet_vars = []  # список (var, name)
        self.sheet_labels = []  # для выделения активного листа
        self.active_sheet_name = None
        self.btn_delete_sheets = tk.Button(left_frame, text='Удалить выбранные листы', command=self.delete_selected_sheets, bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_delete_sheets.pack(pady=(5, 0), anchor='nw', fill='x')
        # Фрейм с чекбоксами и скроллбаром
        sheet_list_frame = tk.Frame(left_frame, bg=DARK_BG)
        sheet_list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.sheet_canvas = tk.Canvas(sheet_list_frame, bg=DARK_BG, highlightthickness=0)
        self.sheet_scrollbar = tk.Scrollbar(sheet_list_frame, orient='vertical', command=self.sheet_canvas.yview)
        self.sheet_checks_frame = tk.Frame(self.sheet_canvas, bg=DARK_BG)
        self.sheet_checks_frame.bind(
            '<Configure>',
            lambda e: self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox('all'))
        )
        self.sheet_canvas.create_window((0, 0), window=self.sheet_checks_frame, anchor='nw')
        self.sheet_canvas.configure(yscrollcommand=self.sheet_scrollbar.set)
        self.sheet_canvas.pack(side='left', fill='both', expand=True)
        self.sheet_scrollbar.pack(side='right', fill='y')

        # Центр — предпросмотр
        center_frame = tk.Frame(self.main_frame, bg=DARK_BG)
        center_frame.pack(side='left', fill='both', expand=True)
        tk.Label(center_frame, text='Предпросмотр объединения', bg=DARK_BG, fg=DARK_FG).pack(anchor='nw')
        self.text = tk.Text(center_frame, wrap='none', width=100, height=40, bg=DARK_ACCENT, fg=DARK_FG, insertbackground=DARK_FG, selectbackground=DARK_HIGHLIGHT)
        self.text.pack(fill='both', expand=True, padx=5, pady=5)
        self.text.config(state='disabled')

        # Скрыть splash через 2.5 секунды и показать основной интерфейс
        self.root.after(2500, self.hide_splash_and_show_main)

    def load_splash_gif(self):
        try:
            from PIL import Image, ImageTk
            gif = Image.open('splash.gif')
            self.splash_frames = []
            try:
                while True:
                    frame = gif.copy().resize((340, 300))
                    self.splash_frames.append(ImageTk.PhotoImage(frame))
                    gif.seek(len(self.splash_frames))
            except EOFError:
                pass
        except Exception:
            self.splash_frames = []

    def animate_splash(self):
        if self.splash_frames and self.splash_running:
            try:
                self.splash_label.config(image=self.splash_frames[self.splash_index])
                self.splash_index = (self.splash_index + 1) % len(self.splash_frames)
                self.root.after(60, self.animate_splash)
            except (tk.TclError, AttributeError):
                self.splash_running = False

    def hide_splash_and_show_main(self):
        self.splash_running = False
        self.splash_frame.destroy()
        self.main_frame.pack(fill='both', expand=True)

    def update_sheet_list(self):
        for widget in self.sheet_checks_frame.winfo_children():
            widget.destroy()
        self.sheet_vars = []
        self.sheet_labels = []
        for name in self.sheet_list:
            row = tk.Frame(self.sheet_checks_frame, bg=DARK_BG)
            row.pack(fill='x', anchor='w')
            var = tk.BooleanVar(value=False)
            chk = tk.Checkbutton(row, variable=var, bg=DARK_BG, fg=DARK_FG, selectcolor=DARK_ACCENT, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
            chk.pack(side='left')
            lbl = tk.Label(row, text=name, anchor='w', width=22, justify='left', bg=DARK_BG, fg=DARK_FG, cursor='hand2')
            lbl.pack(side='left', fill='x', expand=True)
            lbl.bind('<Button-1>', lambda e, n=name: self.show_sheet_content(n))
            self.sheet_vars.append((var, name))
            self.sheet_labels.append((lbl, name))
        self.update_active_sheet_highlight()
        self.sheet_canvas.update_idletasks()
        self.sheet_canvas.yview_moveto(0)

    def update_active_sheet_highlight(self):
        for lbl, name in self.sheet_labels:
            if name == self.active_sheet_name:
                lbl.config(bg=DARK_HIGHLIGHT, fg=DARK_BG)
            else:
                lbl.config(bg=DARK_BG, fg=DARK_FG)

    def show_sheet_content(self, sheet_name):
        self.active_sheet_name = sheet_name
        self.update_active_sheet_highlight()
        df = None
        if self.last_merge_file_path:
            import pandas as pd
            try:
                df = pd.read_excel(self.last_merge_file_path, sheet_name=sheet_name)
            except Exception as e:
                self.text.config(state='normal')
                self.text.delete('1.0', tk.END)
                self.text.insert('end', f'Ошибка чтения листа: {e}')
                self.text.config(state='disabled')
                return
        elif self.last_merge_folder_path:
            import os
            import pandas as pd
            folder = self.last_merge_folder_path
            for fname in os.listdir(folder):
                if fname.endswith('.xlsx'):
                    fpath = os.path.join(folder, fname)
                    try:
                        xl = pd.ExcelFile(fpath)
                        if sheet_name in xl.sheet_names:
                            df = xl.parse(sheet_name)
                            break
                    except Exception:
                        continue
        if df is not None:
            self.text.config(state='normal')
            self.text.delete('1.0', tk.END)
            self.text.insert('end', df.head(100).to_string(index=False))
            self.text.config(state='disabled')
        else:
            self.text.config(state='normal')
            self.text.delete('1.0', tk.END)
            self.text.insert('end', 'Не удалось загрузить лист или он пустой.')
            self.text.config(state='disabled')

    def show_merge_preview(self):
        self.active_sheet_name = None
        self.update_active_sheet_highlight()
        self.show_preview(self.preview_df)

    def delete_selected_sheets(self):
        to_delete = set(name for var, name in self.sheet_vars if var.get())
        self.sheet_list = [name for name in self.sheet_list if name not in to_delete]
        self.update_sheet_list()
        if self.last_merge_file_path:
            preview_df = preview_merge_file(self.last_merge_file_path, self.sheet_list)
            self.preview_df = preview_df
            self.show_preview(preview_df)
        elif self.last_merge_folder_path:
            preview_df = preview_merge_folder(self.last_merge_folder_path, self.sheet_list)
            self.preview_df = preview_df
            self.show_preview(preview_df)

    def choose_file_and_merge(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
        if file_path:
            self.sheet_list = get_sheet_names_from_file(file_path)
            self.update_sheet_list()
            self.last_merge_file_path = file_path
            self.last_merge_folder_path = None
            result_path, preview_df = merge_sheets_in_file(file_path, self.sheet_list)
            self.last_result_path = result_path
            self.preview_df = preview_df
            self.show_preview(preview_df)
            self.btn_export.config(state='normal')

    def choose_folder_and_merge(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.sheet_list = get_sheet_names_from_folder(folder_path)
            self.update_sheet_list()
            self.last_merge_folder_path = folder_path
            self.last_merge_file_path = None
            result_path, preview_df = merge_all_files_in_folder(folder_path, self.sheet_list)
            self.last_result_path = result_path
            self.preview_df = preview_df
            self.show_preview(preview_df)
            self.btn_export.config(state='normal')

    def show_preview(self, preview_df):
        self.text.config(state='normal')
        self.text.delete('1.0', tk.END)
        if preview_df is not None:
            self.text.insert('end', preview_df.to_string(index=False))
        self.text.config(state='disabled')
        self.active_sheet_name = None
        self.update_active_sheet_highlight()

    def export_to_excel(self):
        import os
        if self.last_result_path and os.path.exists(self.last_result_path):
            messagebox.showinfo('Выгрузка', f'Файл уже сохранён: {self.last_result_path}')
        elif self.preview_df is not None:
            save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
            if save_path:
                self.preview_df.to_excel(save_path, index=False)
                messagebox.showinfo('Выгрузка', f'Результат сохранён: {save_path}')
