import warnings
import pandas as pd
from openpyxl.utils import get_column_letter

def baca_excel_adaptif(file_path, sheet_name, dtype=str):
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, dtype=dtype, engine='calamine')
    except Exception:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return pd.read_excel(file_path, sheet_name=sheet_name, dtype=dtype, engine='openpyxl')

try:
    df_acc = baca_excel_adaptif('Acc_Bersih.xlsx', sheet_name='Data', dtype=str)
    df_ctx = baca_excel_adaptif('Ctx.xlsx', sheet_name='data', dtype=str)

    df_acc['No. Faktur'] = df_acc['No. Faktur'].str.strip()
    df_ctx['Referensi'] = df_ctx['Referensi'].str.strip()

    df_ctx = df_ctx.drop_duplicates(subset=['Referensi'], keep='first')

    df_merged = pd.merge(
        df_acc[['No. Faktur']],
        df_ctx[['Referensi', 'Nomor Faktur Pajak']],
        left_on='No. Faktur',
        right_on='Referensi',
        how='left'
    )

    df_merged['Nomor Faktur Pajak'] = (
        df_merged['Nomor Faktur Pajak']
        .replace(r'^\s*$', pd.NA, regex=True)
        .fillna(df_merged['No. Faktur'])
    )

    df_final = df_merged[['No. Faktur', 'Nomor Faktur Pajak']]

    with pd.ExcelWriter('Faktur.xlsx', engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Faktur')
        worksheet = writer.sheets['Faktur']
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if cell.value:
                        cell_len = len(str(cell.value))
                        if cell_len > max_length:
                            max_length = cell_len
                except Exception:
                    pass
            
            worksheet.column_dimensions[column_letter].width = max_length + 2

    print("--> Berhasil menyimpan Faktur.xlsx")

except Exception as e:
    print(f"--> Terjadi kesalahan: {e}")

input("Tekan Enter untuk keluar")