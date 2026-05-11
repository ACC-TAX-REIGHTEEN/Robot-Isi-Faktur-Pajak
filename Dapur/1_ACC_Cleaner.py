import pandas as pd
from openpyxl.utils import get_column_letter

input_file = 'Acc.xls'
output_file = 'Acc_Bersih.xlsx'

try:
    df = pd.read_excel(input_file, header=3)
    
    target_columns = [
        'No. Faktur', 'Tgl Faktur', 'No. Pelanggan', 
        'Nama Pelanggan', 'Nilai Faktur', 'Terutang', 'Keterangan'
    ]
    
    df_selected = df[target_columns].copy()
    
    df_clean = df_selected.dropna(subset=['No. Faktur']).copy()
    
    cols_to_clean = ['No. Faktur', 'Nilai Faktur', 'Terutang', 'No. Pelanggan']
    
    for col in cols_to_clean:
        df_clean[col] = df_clean[col].astype(str)
        df_clean[col] = df_clean[col].str.replace(',00', '', regex=False)
        df_clean[col] = df_clean[col].str.replace(r'\.0$', '', regex=True)
        df_clean[col] = df_clean[col].str.replace('nan', '', regex=False)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_clean.to_excel(writer, index=False, sheet_name='Data')
        
        worksheet = writer.sheets['Data']
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if cell.value:
                        cell_len = len(str(cell.value))
                        if cell_len > max_length:
                            max_length = cell_len
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"Berhasil memproses: {input_file} -> {output_file}")

except Exception as e:
    print(f"Terjadi kesalahan: {e}")

input("Tekan Enter untuk keluar")