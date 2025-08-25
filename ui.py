import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from styles import DARK_BG, DARK_FG, DARK_ACCENT, DARK_BTN_BG, DARK_BTN_FG, DARK_ENTRY_BG, DARK_ENTRY_FG, DARK_HIGHLIGHT
from styles import DARK_BTN_PREVIEW_ACTIVE, DARK_BTN_PREVIEW_INACTIVE
from styles import DARK_BTN_PREVIEW_ACTIVE, DARK_BTN_PREVIEW_INACTIVE
from styles import DARK_BTN_PREVIEW_ACTIVE, DARK_BTN_PREVIEW_INACTIVE
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
        # Стилизация Treeview (тёмная тема)
        style = ttk.Style()
        style.theme_use('default')
        style.configure('Treeview',
            background=DARK_ACCENT,
            foreground=DARK_FG,
            fieldbackground=DARK_ACCENT,
            rowheight=28)
        style.configure('Treeview.Heading',
            background=DARK_HIGHLIGHT,
            foreground=DARK_BG,
            font=('Segoe UI', 10, 'bold'))
        style.map('Treeview', background=[('selected', DARK_BTN_BG)])
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

    # Splash временно отключён

    # Основной интерфейс
        self.main_frame = tk.Frame(self.root, bg=DARK_BG)
        self.main_frame.pack(fill='both', expand=True)

    # Верхняя панель с кнопками
        top_frame = tk.Frame(self.main_frame, bg=DARK_BG)
        top_frame.pack(side='top', fill='x')
        self.btn_file = tk.Button(top_frame, text='Объединить листы в файле', command=self.choose_file_and_merge, bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_file.pack(side='left', padx=5, pady=5)
        self.btn_folder = tk.Button(top_frame, text='Объединить файлы в папке', command=self.choose_folder_and_merge, bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_folder.pack(side='left', padx=5, pady=5)
        self.btn_export = tk.Button(top_frame, text='Выгрузить в Excel', command=self.export_to_excel, state='disabled', bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_export.pack(side='left', padx=5, pady=5)
    # Кнопка возврата к предпросмотру
        self.btn_show_preview = tk.Button(top_frame, text='К предпросмотру', command=self.show_merge_preview, state='disabled', bg=DARK_BTN_BG, fg=DARK_BTN_FG, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
        self.btn_show_preview.pack(side='left', padx=5, pady=5)

    # Список листов слева с чекбоксами
        left_frame = tk.Frame(self.main_frame, width=220, bg=DARK_BG)
        left_frame.pack(side='left', fill='y')
    # --- Label с количеством листов ---
        self.sheet_count_var = tk.StringVar()
        self.sheet_count_label = tk.Label(left_frame, textvariable=self.sheet_count_var, bg=DARK_BG, fg=DARK_FG)
        self.sheet_count_label.pack(anchor='nw')
    # --- Поле поиска по листам ---
        search_frame = tk.Frame(left_frame, bg=DARK_BG)
        search_frame.pack(anchor='nw', fill='x', pady=(2, 4))
        tk.Label(search_frame, text='Поиск:', bg=DARK_BG, fg=DARK_FG).pack(side='left')
        self.sheet_search_var = tk.StringVar()
        self.sheet_search_var.trace_add('write', lambda *a: self.update_sheet_list())
        self.sheet_search_entry = tk.Entry(search_frame, textvariable=self.sheet_search_var, bg=DARK_ENTRY_BG, fg=DARK_ENTRY_FG, insertbackground=DARK_ENTRY_FG, relief='flat')
        self.sheet_search_entry.pack(side='left', fill='x', expand=True, padx=(4, 0))
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
        self.preview_label = tk.Label(center_frame, text='Предпросмотр: Результат', bg=DARK_BG, fg=DARK_FG)
        self.preview_label.pack(anchor='nw')
    # --- Treeview и скроллбары ---
        tree_frame = tk.Frame(center_frame, bg=DARK_BG)
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.tree = ttk.Treeview(tree_frame, show='headings')
        self.tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree_scroll_x = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        self.tree_scroll.grid(row=0, column=1, sticky='ns')
        self.tree_scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    # --- поиск по наименованиям внутри листа и фильтрация Treeview ---
    # (откат: этой функции нет)

        # Скрыть splash через 2.5 секунды и показать основной интерфейс
    # self.root.after(2500, self.hide_splash_and_show_main)  # splash отключён

    # --- splash временно отключён ---

    def update_sheet_list(self):
        # Фильтрация по поиску
        search = self.sheet_search_var.get().strip().lower() if hasattr(self, 'sheet_search_var') else ''
        if search:
            filtered = [name for name in self.sheet_list if search in name.lower()]
        else:
            filtered = list(self.sheet_list)
        # Обновить label с количеством листов
        self.sheet_count_var.set(f'Листы ({len(filtered)})')
        for widget in self.sheet_checks_frame.winfo_children():
            widget.destroy()
        self.sheet_vars = []
        self.sheet_labels = []
        for name in filtered:
            row = tk.Frame(self.sheet_checks_frame, bg=DARK_BG)
            row.pack(fill='x', anchor='w')
            var = tk.BooleanVar(value=False)
            chk = tk.Checkbutton(row, variable=var, bg=DARK_BG, fg=DARK_FG, selectcolor=DARK_ACCENT, activebackground=DARK_ACCENT, activeforeground=DARK_FG)
            chk.pack(side='left')
            lbl = tk.Label(row, text=name, anchor='w', justify='left', bg=DARK_BG, fg=DARK_FG, cursor='hand2')
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
        # Кнопка предпросмотра активна только если выбран отдельный лист
        from styles import DARK_BTN_PREVIEW_ACTIVE, DARK_BTN_PREVIEW_INACTIVE
        if self.active_sheet_name:
            self.btn_show_preview.config(state='normal', bg=DARK_BTN_BG, fg=DARK_BTN_FG)
        else:
            self.btn_show_preview.config(state='disabled', bg=DARK_BTN_BG, fg=DARK_BTN_FG)

    def show_sheet_content(self, sheet_name):
        self.active_sheet_name = sheet_name
        self.preview_label.config(text=f'Предпросмотр: "{sheet_name}"')
        self.update_active_sheet_highlight()
        df = None
        if self.last_merge_file_path:
            import pandas as pd
            try:
                df = pd.read_excel(self.last_merge_file_path, sheet_name=sheet_name)
            except Exception as e:
                self.show_tree_error(f'Ошибка чтения листа: {e}')
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
            # Только отображаем, не меняем self.preview_df!
            self.show_preview(df, is_sheet_preview=True)
        else:
            self.show_tree_error('Не удалось загрузить лист или он пустой.')

    def show_merge_preview(self):
        self.active_sheet_name = None
        self.preview_label.config(text='Предпросмотр: Результат')
        # Сбросить выделение чекбоксов
        for var, _ in self.sheet_vars:
            var.set(False)
        # Сбросить фильтр по результату
        if hasattr(self, 'result_search_var'):
            self.result_search_var.set('')
        self.show_preview(self.preview_df)
        self.update_active_sheet_highlight()
        self.btn_show_preview.config(state='disabled', bg=DARK_BTN_BG, fg=DARK_BTN_FG)

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
            preview_df = merge_sheets_in_file(file_path, self.sheet_list)
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
            preview_df = merge_all_files_in_folder(folder_path, self.sheet_list)
            self.preview_df = preview_df
            self.show_preview(preview_df)
            self.btn_export.config(state='normal')

    def show_preview(self, preview_df, is_sheet_preview=False):
        if not is_sheet_preview:
            self.preview_df = preview_df
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = ()
        self.tree['show'] = 'tree headings'
        # Сбросить поле поиска по наименованиям при новом предпросмотре
        if hasattr(self, 'name_search_var'):
            self.name_search_var.set('')
        if preview_df is not None and not preview_df.empty:
            columns = list(preview_df.columns)
            self.tree['columns'] = columns
            self.tree.heading('#0', text='№')
            self.tree.column('#0', anchor='center', width=60, minwidth=50, stretch=False)
            for col in preview_df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, anchor='w', width=200, minwidth=120)
            for idx, (_, row) in enumerate(preview_df.iterrows(), 1):
                alt = 'alt' if idx % 2 == 0 else ''
                tags = (alt,) if alt else ()
                num_text = f'{idx} ▸'
                self.tree.insert('', 'end', text=num_text, values=list(row), tags=tags)
            self.tree.tag_configure('alt', background='#23262b')
            style = ttk.Style()
            style.configure('Treeview', rowheight=28)
            style.configure('Treeview.Item', font=('Segoe UI', 10))
            style.configure('Treeview', foreground=DARK_FG)
            style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
            style.map('Treeview', foreground=[('selected', DARK_FG)])
        self.update_active_sheet_highlight()

    def show_tree_error(self, message):
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = ('Ошибка',)
        self.tree.heading('Ошибка', text='Ошибка')
        self.tree.column('Ошибка', anchor='center', width=400)
        self.tree.insert('', 'end', values=(message,))

    def export_to_excel(self):
        if self.preview_df is not None:
            save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
            if save_path:
                self.preview_df.to_excel(save_path, index=False)
                messagebox.showinfo('Выгрузка', f'Результат сохранён: {save_path}')
