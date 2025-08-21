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
    writer = pd.ExcelWriter(file_path.replace('.xlsx', '_merged.xlsx'), engine='openpyxl')
    row_offset = 0
    preview_rows = []
    for sheet_name in wb.sheetnames:
        if sheet_name not in sheet_list:
            continue
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        header = pd.DataFrame({df.columns[0]: [f'Лист: {sheet_name}']})
        header.to_excel(writer, index=False, header=False, startrow=row_offset)
        row_offset += 1
        df.to_excel(writer, index=False, startrow=row_offset)
        row_offset += len(df) + 2
        preview_rows.append(header)
        preview_rows.append(df.head(5))
    writer.close()
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return file_path.replace('.xlsx', '_merged.xlsx'), preview_df.head(20)

def merge_all_files_in_folder(folder_path, sheet_list):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    output_file = os.path.join(folder_path, 'merged_folder_preview.xlsx')
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    row_offset = 0
    preview_rows = []
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        wb = load_workbook(file_path, read_only=True)
        for sheet_name in wb.sheetnames:
            sheet_full_name = f'{file} | {sheet_name}'
            if sheet_full_name not in sheet_list:
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            header = pd.DataFrame({df.columns[0]: [f'Файл: {file} | Лист: {sheet_name}']})
            header.to_excel(writer, index=False, header=False, startrow=row_offset)
            row_offset += 1
            df.to_excel(writer, index=False, startrow=row_offset)
            row_offset += len(df) + 2
            preview_rows.append(header)
            preview_rows.append(df.head(5))
    writer.close()
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return output_file, preview_df.head(20)

def preview_merge_file(file_path, sheet_list):
    wb = load_workbook(file_path, read_only=True)
    preview_rows = []
    for sheet_name in wb.sheetnames:
        if sheet_name not in sheet_list:
            continue
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        header = pd.DataFrame({df.columns[0]: [f'Лист: {sheet_name}']})
        preview_rows.append(header)
        preview_rows.append(df.head(5))
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return preview_df.head(20)

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
            header = pd.DataFrame({df.columns[0]: [f'Файл: {file} | Лист: {sheet_name}']})
            preview_rows.append(header)
            preview_rows.append(df.head(5))
    preview_df = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame()
    return preview_df.head(20)
