import pandas as pd
import re
import numpy as np
from datetime import datetime
import hashlib

def extract_floor_from_sheet(sheet_name):
    name = str(sheet_name).strip()
    if re.match(r'^00(\D|$)', name):
        return '00'
    if re.match(r'^-?1(\D|$)', name):
        if name.startswith('-1'):
            return '-1'
        else:
            return '1'
    if re.match(r'^0(\D|$)', name):
        return '0'
    m = re.match(r'^(\d+)', name)
    if m:
        return m.group(1)
    return None

def normalize_key(s):
    s = str(s).strip().lower()
    s = re.sub(r'[\u2013\u2014\u2212]', '-', s)
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', s)
    return s

def smart_number(val):
    if pd.isnull(val):
        return 0.0
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
    if re.search(r'[a-zA-Zа-яА-ЯёЁxхXХ]', s):
        return 0.0
    return 0.0

def content_hash(df):
    if not {'Марка_norm', 'Наименование_norm', 'Количество'}.issubset(df.columns):
        return None
    arr = df[['Марка_norm', 'Наименование_norm', 'Количество']].sort_values(['Марка_norm', 'Наименование_norm']).to_string(index=False)
    return hashlib.md5(arr.encode('utf-8')).hexdigest()

def clean_and_aggregate(xl):
    floors = {}
    hashes = {}
    for sheet in xl.sheet_names:
        floor = extract_floor_from_sheet(sheet)
        if floor is None:
            continue
        df = xl.parse(sheet, header=None)
        df = df.dropna(how='all')
        if df.shape[0] == 0:
            continue
        header_row = None
        for i in range(min(2, df.shape[0])):
            cols = [str(x).lower() for x in df.iloc[i]]
            if any('марка' in c for c in cols) and (any('наимен' in c for c in cols) or any('опис' in c for c in cols)):
                header_row = i
                break
        if header_row is not None:
            df.columns = df.iloc[header_row].astype(str)
            df = df.iloc[header_row+1:]
            cols = list(df.columns)
            mark_col = next((c for c in cols if re.search(r'марка', c, re.I)), cols[0])
            name_col = next((c for c in cols if re.search(r'наимен|опис', c, re.I)), cols[1])
            name_idx = cols.index(name_col)
            qty_col = cols[name_idx+1] if name_idx+1 < len(cols) else cols[-1]
            df = df[[mark_col, name_col, qty_col]].copy()
            df.columns = ['Марка', 'Наименование', 'Количество']
        else:
            cols = list(df.columns)[:3]
            df = df[cols].copy()
            df.columns = ['Марка', 'Наименование', 'Количество']
        df['Количество'] = df['Количество'].apply(smart_number)
        df['Марка_norm'] = df['Марка'].apply(normalize_key)
        df['Наименование_norm'] = df['Наименование'].apply(normalize_key)
        df = df.groupby(['Марка_norm', 'Наименование_norm'], as_index=False).agg({'Марка':'first', 'Наименование':'first', 'Количество':'sum'})
        h = content_hash(df)
        if floor not in floors:
            floors[floor] = []
            hashes[floor] = set()
        if h and h not in hashes[floor]:
            floors[floor].append(df)
            hashes[floor].add(h)
    for floor in floors:
        floors[floor] = pd.concat(floors[floor], ignore_index=True).groupby(['Марка_norm', 'Наименование_norm'], as_index=False).agg({'Марка':'first', 'Наименование':'first', 'Количество':'sum'})
    return floors

def build_final_table_multi(floors, multipliers):
    all_keys = set()
    for df in floors.values():
        for _, row in df.iterrows():
            all_keys.add((row['Марка_norm'], row['Наименование_norm']))
    df0 = floors.get('0', None)
    if df0 is not None:
        for _, row in df0.iterrows():
            all_keys.add((row['Марка_norm'], row['Наименование_norm']))
    all_keys = sorted(all_keys)
    podval_keys = [f for f in floors if f in ('00', '-1')]
    floor_nums = sorted([f for f in floors if f not in ('0', '00', '-1')], key=lambda x: (len(str(x)), str(x)))
    data = []
    for key in all_keys:
        row = {}
        for df in floors.values():
            found = df[(df['Марка_norm'] == key[0]) & (df['Наименование_norm'] == key[1])]
            if not found.empty:
                row['Марка'] = found.iloc[0]['Марка']
                row['Наименование'] = found.iloc[0]['Наименование']
                break
        df0 = floors.get('0', None)
        if df0 is None:
            row['Сводная'] = 0
        else:
            found = df0[(df0['Марка_norm'] == key[0]) & (df0['Наименование_norm'] == key[1])]
            row['Сводная'] = float(found['Количество'].iloc[0]) if not found.empty else 0
        podval_val = 0
        if podval_keys:
            for f in podval_keys:
                dfp = floors[f]
                found = dfp[(dfp['Марка_norm'] == key[0]) & (dfp['Наименование_norm'] == key[1])]
                val = float(found['Количество'].iloc[0]) * multipliers.get(f, 1) if not found.empty else 0
                podval_val += val
                row['Подвал'] = podval_val
        sum_etazhi = 0
        for f in floor_nums:
            dff = floors[f]
            found = dff[(dff['Марка_norm'] == key[0]) & (dff['Наименование_norm'] == key[1])]
            val = float(found['Количество'].iloc[0]) * multipliers.get(f, 1) if not found.empty else 0
            row[f] = val
            sum_etazhi += val
        sum_total = sum_etazhi + podval_val if podval_keys else sum_etazhi
        row['Сумма этажей'] = sum_total
        row['Проверка'] = row.get('Сводная', 0) - sum_total
        data.append(row)
    columns = ['Марка', 'Наименование', 'Сводная']
    if podval_keys:
        columns.append('Подвал')
    columns += floor_nums
    columns += ['Сумма этажей', 'Проверка']
    df_final = pd.DataFrame(data, columns=columns)
    return df_final
