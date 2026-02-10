import openpyxl

def clean_excel_data(filename):
    print(f"--> Membaca file: {filename}")
    
    try:
        wb = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        print(f"--> File {filename} tidak ditemukan.")
        return

    check_columns = ['Tanggal', 'No Inv', 'No SJ', 'No PO', 'Tgl FP', 'DPP']
    priority_column = 'No FP'

    for sheet in wb.worksheets:
        print(f"--> Memproses sheet: {sheet.title}")
        
        headers = {}
        for cell in sheet[1]:
            if cell.value:
                headers[cell.value] = cell.column - 1 
        
        missing_cols = [col for col in check_columns + [priority_column] if col not in headers]
        if missing_cols:
            print(f"--> Kolom tidak lengkap di sheet {sheet.title}. Melewati sheet ini.")
            continue

        row_groups = {}
        rows_to_delete = []

        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            key_values = []
            for col_name in check_columns:
                val = row[headers[col_name]]
                key_values.append(val if val is not None else "")
            
            key = tuple(key_values)
            no_fp_val = row[headers[priority_column]]
            
            if key not in row_groups:
                row_groups[key] = []
            row_groups[key].append({'row_idx': row_idx, 'no_fp': no_fp_val})

        for key, group in row_groups.items():
            if len(group) > 1:
                group.sort(key=lambda x: (x['no_fp'] is None or str(x['no_fp']).strip() == '', x['row_idx']))
                
                for item in group[1:]:
                    rows_to_delete.append(item['row_idx'])

        rows_to_delete.sort(reverse=True)
        
        count = 0
        for r_idx in rows_to_delete:
            sheet.delete_rows(r_idx)
            count += 1
        
        print(f"--> Menghapus {count} baris duplikat di sheet {sheet.title}")

    print(f"--> Menyimpan file...")
    wb.save(filename)
    print(f"--> Selesai.")

if __name__ == "__main__":
    clean_excel_data('Hasil_Ekstrak_temp.xlsx')
