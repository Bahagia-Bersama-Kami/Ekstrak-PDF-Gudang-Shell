import pandas as pd
import openpyxl
import glob
import os
from datetime import datetime

def format_date_source(date_val):
    try:
        if isinstance(date_val, str):
            dt_obj = pd.to_datetime(date_val)
            return dt_obj.strftime('%d/%m/%Y')
        elif isinstance(date_val, datetime):
            return date_val.strftime('%d/%m/%Y')
        return None
    except:
        return None

def format_date_target(date_val):
    try:
        if isinstance(date_val, datetime):
            return date_val.strftime('%d/%m/%Y')
        elif isinstance(date_val, str):
            return date_val 
        return str(date_val)
    except:
        return str(date_val)

print("--> Memulai proses...")

source_files = glob.glob("data_export_*.xlsx")
if not source_files:
    print("--> Error: File data_export_.....xlsx tidak ditemukan.")
    exit()

source_file = source_files[0]
print(f"--> File sumber ditemukan: {source_file}")

try:
    df_source = pd.read_excel(source_file, sheet_name="data", dtype={'Nomor Faktur Pajak': str})
except Exception as e:
    print(f"--> Error saat membaca file sumber: {e}")
    exit()

lookup_data = {}

print("--> Membuat indeks data dari file sumber (Sistem Antrean)...")
for index, row in df_source.iterrows():
    raw_date = row.get('Tanggal Faktur Pajak')
    dpp_val = row.get('Harga Jual/Penggantian/DPP')
    no_fp = row.get('Nomor Faktur Pajak')
    
    clean_date = format_date_source(raw_date)
    
    if clean_date and pd.notna(dpp_val):
        key = (clean_date, float(dpp_val))
        
        if key not in lookup_data:
            lookup_data[key] = []
            
        lookup_data[key].append(no_fp)

target_filename = "Hasil_Ekstrak_temp.xlsx"
if not os.path.exists(target_filename):
    print(f"--> Error: File {target_filename} tidak ditemukan.")
    exit()

print(f"--> Membuka file target: {target_filename}")
wb = openpyxl.load_workbook(target_filename)

for sheet in wb.worksheets:
    print(f"--> Memproses Sheet: {sheet.title}")
    
    headers = {}
    for cell in sheet[1]:
        if cell.value:
            headers[cell.value] = cell.column
            
    if "Tanggal" not in headers or "DPP" not in headers or "No FP" not in headers:
        print(f"--> Lewati Sheet {sheet.title}: Kolom yang dibutuhkan tidak lengkap.")
        continue
        
    col_tanggal = headers["Tanggal"]
    col_dpp = headers["DPP"]
    col_no_fp = headers["No FP"]
    
    match_count = 0
    
    for row in sheet.iter_rows(min_row=2):
        cell_tanggal = row[col_tanggal - 1]
        cell_dpp = row[col_dpp - 1]
        cell_no_fp = row[col_no_fp - 1]
        
        val_tanggal = cell_tanggal.value
        val_dpp = cell_dpp.value
        
        if val_tanggal is not None and val_dpp is not None:
            target_date_str = format_date_target(val_tanggal)
            try:
                target_dpp_float = float(val_dpp)
            except:
                continue
                
            search_key = (target_date_str, target_dpp_float)
            
            if search_key in lookup_data and len(lookup_data[search_key]) > 0:
            	
                found_no_fp = lookup_data[search_key].pop(0)
                
                cell_no_fp.value = found_no_fp
                match_count += 1

    print(f"--> Selesai memproses {sheet.title}. Ditemukan {match_count} kecocokan.")

print("--> Menyimpan file...")
wb.save(target_filename)
print("--> Selesai. Data berhasil disimpan dengan metode ticketing.")