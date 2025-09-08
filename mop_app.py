import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.filedialog as fd
from mop_logic import clean_and_aggregate, build_final_table_multi

class MopApp(tk.Toplevel):
    def show_floor_table(self, floor):
        # Найти DataFrame по этажу и показать предпросмотр
        if hasattr(self, 'floors') and floor in self.floors:
            df = self.floors[floor]
            # Передаём спец. флаг для синей темы
            self.show_preview(df, is_sheet_preview='blue')
        else:
            messagebox.showerror('Ошибка', f'Нет данных для этажа: {floor}')
    def apply_multiplier(self):
        # Применить множитель к выбранному этажу и обновить предпросмотр
        floor = self.typical_floor.get()
        try:
            mult = int(self.typical_mult.get())
        except Exception:
            mult = 1
        for f, v in self.multipliers.items():
            if f == floor:
                v.set(mult)
            else:
                v.set(1)
        self.recalc()
    def delete_selected_sheets(self):
        to_delete = set(name for var, name in self.sheet_vars if var.get())
        self.sheet_list = [name for name in self.sheet_list if name not in to_delete]
        self.update_sheet_list()
        if hasattr(self, 'last_merge_file_path') and self.last_merge_file_path:
            # Для МОП: просто обновить предпросмотр (можно доработать под бизнес-логику)
            pass

    def show_preview(self, preview_df, is_sheet_preview=False):
        # Определяем стиль: если is_sheet_preview == 'blue', то делаем синий фон и белый текст
        if is_sheet_preview == 'blue':
            self.style.configure('Treeview', background='#1a2340', foreground='#fff', fieldbackground='#1a2340', rowheight=28)
            self.style.configure('Treeview.Heading', background='#223', foreground='#fff', font=('Segoe UI', 10, 'bold'))
        else:
            self.style.configure('Treeview', background='#2a2d36', foreground='#000', fieldbackground='#2a2d36', rowheight=28)
            self.style.configure('Treeview.Heading', background='#444', foreground='#000', font=('Segoe UI', 10, 'bold'))
        if not is_sheet_preview:
            self.preview_df = preview_df
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = ()
        self.tree['show'] = 'tree headings'
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
    # self.update_active_sheet_highlight()  # убрано, чтобы не было AttributeError

    def show_tree_error(self, message):
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = ('Ошибка',)
        self.tree.heading('Ошибка', text='Ошибка')
        self.tree.column('Ошибка', anchor='center', width=400)
        self.tree.insert('', 'end', values=(message,))
    def __init__(self, master=None):
        super().__init__(master)
        self.title('Проверка Excel по этажам')
                # Центрирование окна
        window_width = 1200
        window_height = 700
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f'{window_width}x{window_height}+{x}+{y}')
        self.configure(bg='#222')
        self.lift()
        self.attributes('-topmost', True)
        self.focus_force()
        self.after(100, lambda: self.attributes('-topmost', False))
                # --- Стилизация Treeview (тёмная тема) ---
        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure('Treeview', background='#2a2d36', foreground='#000', fieldbackground='#2a2d36', rowheight=28)
        self.style.configure('Treeview.Heading', background='#444', foreground='#000', font=('Segoe UI', 10, 'bold'))
        self.style.map('Treeview', background=[('selected', '#444')])

                # --- Основной интерфейс ---
        self.main_frame = tk.Frame(self, bg='#222')
        self.main_frame.pack(fill='both', expand=True)

    # --- Верхняя панель с кнопками ---
        top_frame = tk.Frame(self.main_frame, bg='#222')
        top_frame.pack(side='top', fill='x')
        tk.Button(top_frame, text='Открыть Excel', command=self.open_file, bg='#333', fg='#fff').pack(side='left', padx=5, pady=5)
        tk.Button(top_frame, text='Обработать папку', command=self.open_folder, bg='#333', fg='#fff').pack(side='left', padx=5, pady=5)
        self.btn_export = tk.Button(top_frame, text='Экспортировать в Excel', command=self.export_to_excel, state='disabled', bg='#e8ffe8', fg='#222', relief='groove')
        self.btn_export.pack(side='left', padx=5, pady=5)
        mult_frame = tk.Frame(top_frame, bg='#222')
        mult_frame.pack(side='left', padx=10)
        tk.Label(mult_frame, text='Типовой этаж:', fg='#fff', bg='#222').grid(row=0, column=0)
        self.typical_floor = tk.StringVar()
        self.typical_mult = tk.StringVar(value='1')
        self.typical_floor_cb = ttk.Combobox(mult_frame, values=[], textvariable=self.typical_floor, state='readonly', width=8)
        self.typical_floor_cb.grid(row=0, column=1, padx=5)
        tk.Label(mult_frame, text='Множитель:', fg='#fff', bg='#222').grid(row=0, column=2)
        entry = tk.Entry(mult_frame, textvariable=self.typical_mult, width=5, bg='#333', fg='#fff', insertbackground='#fff')
        entry.grid(row=0, column=3, padx=5)
        tk.Button(mult_frame, text='Умножить', command=self.apply_multiplier, bg='#333', fg='#fff').grid(row=0, column=4, padx=5)
    def open_folder(self):
        folder_path = fd.askdirectory(title='Выберите папку с Excel-файлами')
        if not folder_path:
            return
        all_floors = {}
        for fname in os.listdir(folder_path):
            if fname.endswith('.xlsx') or fname.endswith('.xls'):
                fpath = os.path.join(folder_path, fname)
                try:
                    xl = pd.ExcelFile(fpath)
                    floors = clean_and_aggregate(xl)
                    for floor, df in floors.items():
                        if floor not in all_floors:
                            all_floors[floor] = []
                        all_floors[floor].append(df)
                except Exception:
                    continue
        # Объединить по этажам
        self.floors = {}
        for floor, dfs in all_floors.items():
            self.floors[floor] = pd.concat(dfs, ignore_index=True).groupby(['Марка_norm', 'Наименование_norm'], as_index=False).agg({'Марка':'first', 'Наименование':'first', 'Количество':'sum'})
        self.multipliers = {f: tk.IntVar(value=1) for f in self.floors}
        self.sheet_list = []
        for floor, df in self.floors.items():
            for name in df['Наименование'].unique():
                self.sheet_list.append(f'{floor}: {name}')
        self.update_sheet_list()
        self.recalc()
    # ...existing code...

                # --- Боковая панель с листами, поиском и чекбоксами ---
        left_frame = tk.Frame(self.main_frame, width=220, bg='#222')
        left_frame.pack(side='left', fill='y')
        self.sheet_count_var = tk.StringVar()
        self.sheet_count_label = tk.Label(left_frame, textvariable=self.sheet_count_var, bg='#222', fg='#fff')
        self.sheet_count_label.pack(anchor='nw')
        search_frame = tk.Frame(left_frame, bg='#222')
        search_frame.pack(anchor='nw', fill='x', pady=(2, 4))
        tk.Label(search_frame, text='Поиск:', bg='#222', fg='#fff').pack(side='left')
        self.sheet_search_var = tk.StringVar()
        self.sheet_search_var.trace_add('write', lambda *a: self.update_sheet_list())
        self.sheet_search_entry = tk.Entry(search_frame, textvariable=self.sheet_search_var, bg='#333', fg='#fff', insertbackground='#fff', relief='flat')
        self.sheet_search_entry.pack(side='left', fill='x', expand=True, padx=(4, 0))
        self.sheet_vars = []
        self.sheet_labels = []
        self.active_sheet_name = None
        self.btn_delete_sheets = tk.Button(left_frame, text='Удалить выбранные листы', command=self.delete_selected_sheets, bg='#333', fg='#fff')
        self.btn_delete_sheets.pack(pady=(5, 0), anchor='nw', fill='x')
        sheet_list_frame = tk.Frame(left_frame, bg='#222')
        sheet_list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.sheet_canvas = tk.Canvas(sheet_list_frame, bg='#222', highlightthickness=0)
        self.sheet_scrollbar = tk.Scrollbar(sheet_list_frame, orient='vertical', command=self.sheet_canvas.yview)
        self.sheet_checks_frame = tk.Frame(self.sheet_canvas, bg='#222')
        self.sheet_checks_frame.bind('<Configure>', lambda e: self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox('all')))
        self.sheet_canvas.create_window((0, 0), window=self.sheet_checks_frame, anchor='nw')
        self.sheet_canvas.configure(yscrollcommand=self.sheet_scrollbar.set)
        self.sheet_canvas.pack(side='left', fill='both', expand=True)
        self.sheet_scrollbar.pack(side='right', fill='y')

    # --- Центр — предпросмотр ---
        center_frame = tk.Frame(self.main_frame, bg='#222')
        center_frame.pack(side='left', fill='both', expand=True)
        self.preview_label = tk.Label(center_frame, text='Предпросмотр: Результат', bg='#222', fg='#fff')
        self.preview_label.pack(anchor='nw')
        tree_frame = tk.Frame(center_frame, bg='#222')
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.tree_frame = tree_frame
        self.tree = ttk.Treeview(tree_frame, show='headings')
        self.tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree_scroll_x = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        self.tree_scroll.grid(row=0, column=1, sticky='ns')
        self.tree_scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
    def open_file(self):
        path = fd.askopenfilename(title='Выберите Excel-файл', filetypes=[('Excel files', '*.xlsx *.xls')])
        if not path:
            return
        import traceback
        try:
            xl = pd.ExcelFile(path)
            self.floors = clean_and_aggregate(xl)
            print('DEBUG self.floors:', {k: v.shape for k, v in self.floors.items()})
            self.multipliers = {f: tk.IntVar(value=1) for f in self.floors}
            # Определяем этажи
            floor_nums = sorted([str(f) for f in self.floors if f not in ('0', '00', '-1')], key=lambda x: (len(x), x))
            self.typical_floor_cb['values'] = floor_nums
            if floor_nums:
                self.typical_floor.set(floor_nums[0])

            # --- КОМПАКТНЫЕ КНОПКИ В ВЕРХНЕЙ ПАНЕЛИ ---
            if hasattr(self, 'floor_btns_frame') and self.floor_btns_frame:
                self.floor_btns_frame.destroy()
            # Найти top_frame (верхняя панель)
            top_frame = None
            for widget in self.main_frame.winfo_children():
                if isinstance(widget, tk.Frame) and widget.winfo_manager() == 'pack' and widget.pack_info().get('side') == 'top':
                    top_frame = widget
                    break
            if not top_frame:
                top_frame = tk.Frame(self.main_frame, bg='#222')
                top_frame.pack(side='top', fill='x')
            self.floor_btns_frame = tk.Frame(top_frame, bg='#222')
            self.floor_btns_frame.pack(side='left', padx=10)

            # Кнопки управления (row=0)
            col0 = 0
            btn_result = tk.Button(self.floor_btns_frame, text='Результат', width=7, height=1, font=('Segoe UI', 9), command=lambda: self.show_table(), bg='#e8ffe8', fg='#222')
            btn_result.grid(row=0, column=col0, padx=2, pady=2)
            col0 += 1
            if '0' in self.floors:
                btn_svod = tk.Button(self.floor_btns_frame, text='Свод', width=7, height=1, font=('Segoe UI', 9), command=lambda: self.show_preview(self.floors['0'], is_sheet_preview='blue'), bg='#1a2340', fg='#fff')
                btn_svod.grid(row=0, column=col0, padx=2, pady=2)
                col0 += 1
            basement_floors = [f for f in self.floors if str(f) in ('00', '-1')]
            if basement_floors:
                btn_basement = tk.Button(self.floor_btns_frame, text='Подвал', width=7, height=1, font=('Segoe UI', 9), command=lambda: self.show_preview(self.floors[basement_floors[0]], is_sheet_preview='blue'), bg='#1a2340', fg='#fff')
                btn_basement.grid(row=0, column=col0, padx=2, pady=2)

            # Кнопки этажей (row=1+)
            max_cols = 8
            for idx, f in enumerate(floor_nums):
                row = idx // max_cols + 1
                col = idx % max_cols
                btn = tk.Button(self.floor_btns_frame, text=f, width=3, height=1, font=('Segoe UI', 9), command=lambda fl=f: self.show_preview(self.floors[fl], is_sheet_preview='blue'), bg='#1a2340', fg='#fff')
                btn.grid(row=row, column=col, padx=1, pady=1)

            # Гарантируем, что self.tree_frame существует
            if not hasattr(self, 'tree_frame') or self.tree_frame is None:
                center_frame = tk.Frame(self.main_frame, bg='#222')
                center_frame.pack(side='left', fill='both', expand=True)
                self.preview_label = tk.Label(center_frame, text='Предпросмотр: Результат', bg='#222', fg='#fff')
                self.preview_label.pack(anchor='nw')
                tree_frame = tk.Frame(center_frame, bg='#222')
                tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
                self.tree_frame = tree_frame
                self.tree = ttk.Treeview(tree_frame, show='headings')
                self.tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
                self.tree_scroll_x = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
                self.tree.configure(yscrollcommand=self.tree_scroll.set, xscrollcommand=self.tree_scroll_x.set)
                self.tree.grid(row=0, column=0, sticky='nsew')
                self.tree_scroll.grid(row=0, column=1, sticky='ns')
                self.tree_scroll_x.grid(row=1, column=0, sticky='ew')
                tree_frame.grid_rowconfigure(0, weight=1)
                tree_frame.grid_columnconfigure(0, weight=1)
            self.recalc()
            print('DEBUG final_df:', getattr(self, 'final_df', None))
        except Exception as e:
            tb = traceback.format_exc()
            messagebox.showerror('Ошибка', f'Не удалось загрузить файл:\n{e}\n\nTraceback:\n{tb}')
            print('ERROR:', tb)
    def recalc(self):
        mults = {f: v.get() if hasattr(v, 'get') else v for f, v in self.multipliers.items()}
        self.final_df = build_final_table_multi(self.floors, mults)
        self.show_table()
    def show_table(self):
        for widget in self.tree_frame.winfo_children():
            widget.destroy()
        if self.final_df is None or self.final_df.empty:
            tk.Label(self.tree_frame, text='Нет данных для отображения.', fg='#222', bg='#fff').pack()
            return
        columns = list(self.final_df.columns)
        columns_with_idx = ['№'] + columns
        self.tree = ttk.Treeview(self.tree_frame, columns=columns_with_idx, show='headings', height=25)
        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#444", foreground="#000")
        self.style.configure("Treeview", font=("Segoe UI", 10), rowheight=24, borderwidth=0, background="#23262b", fieldbackground="#23262b", foreground="#000")
        self.style.map("Treeview", background=[('selected', '#444')])
        for col in columns_with_idx:
            if col == '№':
                self.tree.heading(col, text=col, anchor='center')
                self.tree.column(col, width=40, anchor='center', minwidth=30, stretch=False)
            else:
                self.tree.heading(col, text=col, anchor='center')
                self.tree.column(col, width=110, anchor='center', minwidth=60, stretch=True)
        yscroll = ttk.Scrollbar(self.tree_frame, orient='vertical', command=self.tree.yview)
        xscroll = ttk.Scrollbar(self.tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(side='left', fill='both', expand=True)
        yscroll.pack(side='right', fill='y')
        xscroll.pack(side='bottom', fill='x')
        for i, row in self.final_df.iterrows():
            vals = [i+1]
            for v in row:
                if isinstance(v, float):
                    vals.append(f'{v:.1f}')
                else:
                    vals.append(v)
            check = abs(row.get('Проверка', 0))
            if check > 1e-2:
                tag = 'err'
            elif check > 1e-6:
                tag = 'warn'
            else:
                tag = 'ok'
            self.tree.insert('', 'end', values=vals, tags=(tag,))
        self.tree.tag_configure('err', background='#ffe5ec')
        self.tree.tag_configure('warn', background='#fffbe5')
        self.tree.tag_configure('ok', background='#e8ffe8')
        for col in columns_with_idx:
            if col not in ('Марка', 'Наименование', '№'):
                self.tree.column(col, anchor='e')
    def export_to_excel(self):
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill
        self.output_path = fd.asksaveasfilename(title='Сохранить результат как...', defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
        if not self.output_path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = 'Сводная'
        ws.append(list(self.final_df.columns))
        for i, row in self.final_df.iterrows():
            ws.append([v if not isinstance(v, float) else round(v, 1) for v in row])
        red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        for r in range(2, ws.max_row+1):
            val = ws.cell(row=r, column=ws.max_column).value
            try:
                if abs(float(val)) > 1e-6:
                    ws.cell(row=r, column=ws.max_column).fill = red_fill
            except:
                pass
        for f in sorted(self.floors, key=lambda x: (len(str(x)), str(x))):
            df = self.floors[f]
            ws_floor = wb.create_sheet(title=f'{f}')
            ws_floor.append(list(df.columns))
            for _, row in df.iterrows():
                ws_floor.append([v if not isinstance(v, float) else round(v, 1) for v in row])
        wb.save(self.output_path)
        messagebox.showinfo('Готово', f'Результат сохранён: {self.output_path}')

if __name__ == '__main__':
    app = MopApp()
    app.mainloop()
