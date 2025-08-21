
import re
import numpy as np
from datetime import datetime

def extract_name_and_qty_columns(df):
    # Ключевые слова для поиска
    name_keys = ['наимен', 'опис', 'обозн']
    qty_keys = ['кол', 'кол-во', 'количество', 'qty', 'count', 'площадь']
    exclude = ['марка', 'изображ', 'условн', 'фото', 'примеч']
    norm_cols = [str(c).replace(' ', '').lower() for c in df.columns]
    # Найти первый подходящий столбец для наименования
    name_col = None
    for i, c in enumerate(norm_cols):
        if any(k in c for k in name_keys) and not any(e in c for e in exclude):
            name_col = df.columns[i]
            break
    # Найти первый подходящий столбец для количества
    qty_col = None
    for i, c in enumerate(norm_cols):
        if any(k in c for k in qty_keys) and not any(e in c for e in exclude):
            qty_col = df.columns[i]
            break
    # Fallback: если не нашли — взять первые два столбца
    if name_col is None:
        name_col = df.columns[0] if len(df.columns) > 0 else None
    if qty_col is None or qty_col == name_col:
        # Ищем столбец с максимальным количеством числовых значений (но не совпадающий с name_col)
        max_numeric = 0
        best_col = None
        for col in df.columns:
            if col == name_col:
                continue
            nums = df[col].apply(lambda v: pd.notnull(smart_number(v))).sum()
            if nums > max_numeric:
                max_numeric = nums
                best_col = col
        qty_col = best_col
    # Обработка значений
    name_series = df[name_col].astype(str).str.strip() if name_col else pd.Series(dtype=str)
    if qty_col:
        def debug_smart_number(val):
            parsed = smart_number(val)
            print(f"DEBUG qty_col: raw='{val}' -> parsed={parsed}")
            return parsed
        qty_series = df[qty_col].apply(debug_smart_number)
    else:
        qty_series = pd.Series(dtype=float)
    # Собрать итоговый DataFrame
    result = pd.DataFrame({'Наименование': name_series, 'Кол-во': qty_series})
    # Удалить полностью пустые строки
    result = result.dropna(how='all')
    # Оставить только строки, где есть наименование и количество
    result = result[(result['Наименование'].notna()) & (result['Кол-во'].notna())]
    # Удалить дубли по наименованию, если есть хотя бы одна строка с Кол-во != 0, оставить её
    def keep_nonzero_dupes(df):
        nonzero = df[df['Кол-во'] != 0]
        if not nonzero.empty:
            return nonzero.iloc[0]
        return df.iloc[0]
    result = result.groupby('Наименование', as_index=False).apply(keep_nonzero_dupes).reset_index(drop=True)
    return result

def normalize_key(s):
    s = str(s).strip().lower()
    s = re.sub(r'[\u2013\u2014\u2212]', '-', s)
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', s)
    return s

def smart_number(val):
    if pd.isnull(val):
        return np.nan
    if isinstance(val, (int, float, np.integer, np.floating)):
        return float(val)
    s = str(val).strip().replace(',', '.').replace(' ', '')
    try:
        if s.isdigit() and 35000 < int(s) < 50000:
            dt = datetime(1899, 12, 30) + pd.to_timedelta(int(s), unit='D')
            return float(dt.day)
    except Exception:
        pass
    try:
        return float(s)
    except Exception:
        pass
    # Если не число — вернуть np.nan
    return np.nan
def clean_dataframe(df):
    # Ключевые слова для поиска нужных столбцов
    # Привести все имена столбцов к нижнему регистру без пробелов для поиска
    # Если первый столбец — марка, всё равно ищем только наименование и кол-во
    return extract_name_and_qty_columns(df)
import os
import pandas as pd
from openpyxl import load_workbook

def get_sheet_names_from_file(file_path):
    wb = load_workbook(file_path, read_only=True)
    return [f'{sheet}' for sheet in wb.sheetnames]

def get_sheet_names_from_folder(folder_path):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    sheet_names = []
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        wb = load_workbook(file_path, read_only=True)
        for sheet in wb.sheetnames:
            sheet_names.append(f'{file} | {sheet}')
    return sheet_names

def merge_sheets_in_file(file_path, sheet_list):
    wb = load_workbook(file_path, read_only=True)
    preview_rows = []
    for sheet_name in wb.sheetnames:
        if sheet_name not in sheet_list:
            continue
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df_clean = clean_dataframe(df)
        if not df_clean.empty:
            preview_rows.append(df_clean)
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return preview_df

def merge_all_files_in_folder(folder_path, sheet_list):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    preview_rows = []
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        wb = load_workbook(file_path, read_only=True)
        for sheet_name in wb.sheetnames:
            sheet_full_name = f'{file} | {sheet_name}'
            if sheet_full_name not in sheet_list:
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df_clean = clean_dataframe(df)
            if not df_clean.empty:
                preview_rows.append(df_clean)
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return preview_df

def preview_merge_file(file_path, sheet_list):
    wb = load_workbook(file_path, read_only=True)
    preview_rows = []
    for sheet_name in wb.sheetnames:
        if sheet_name not in sheet_list:
            continue
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df_clean = clean_dataframe(df)
        if not df_clean.empty:
            preview_rows.append(df_clean)
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return preview_df

def preview_merge_folder(folder_path, sheet_list):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    preview_rows = []
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        wb = load_workbook(file_path, read_only=True)
        for sheet_name in wb.sheetnames:
            sheet_full_name = f'{file} | {sheet_name}'
            if sheet_full_name not in sheet_list:
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df_clean = clean_dataframe(df)
            if not df_clean.empty:
                preview_rows.append(df_clean)
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return preview_df
