import openpyxl

print("--> Memulai proses penghapusan duplikat...")

filename = "Hasil_Ekstrak_temp.xlsx"

try:
    wb = openpyxl.load_workbook(filename)
except Exception as e:
    print(f"--> Gagal membuka file: {e}")
    exit()

for sheet in wb.worksheets:
    print(f"--> Memproses Sheet: {sheet.title}")
    
    seen_rows = set()
    rows_to_delete = []
    
    for i, row in enumerate(sheet.iter_rows(values_only=True), 1):
        if row in seen_rows:
            rows_to_delete.append(i)
        else:
            seen_rows.add(row)
    
    for row_idx in reversed(rows_to_delete):
        sheet.delete_rows(row_idx, 1)
        
    print(f"--> Berhasil menghapus {len(rows_to_delete)} baris duplikat pada Sheet {sheet.title}")

print("--> Menyimpan file...")
wb.save(filename)
print("--> Selesai.")