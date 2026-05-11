import pandas as pd
from openpyxl.utils import get_column_letter

try:
    df_acc = pd.read_excel('Acc_Bersih.xlsx', sheet_name='Data', dtype=str)
    df_ctx = pd.read_excel('Ctx.xlsx', sheet_name='data', dtype=str)

    df_acc['No. Faktur'] = df_acc['No. Faktur'].str.strip()
    df_ctx['Referensi'] = df_ctx['Referensi'].str.strip()

    df_merged = pd.merge(
        df_acc[['No. Faktur']],
        df_ctx[['Referensi', 'Nomor Faktur Pajak']],
        left_on='No. Faktur',
        right_on='Referensi',
        how='left'
    )

    df_merged['Nomor Faktur Pajak'] = df_merged['Nomor Faktur Pajak'].fillna(df_merged['No. Faktur'])
    
    mask_empty = df_merged['Nomor Faktur Pajak'].str.strip() == ''
    df_merged.loc[mask_empty, 'Nomor Faktur Pajak'] = df_merged.loc[mask_empty, 'No. Faktur']

    df_merged.loc[df_merged['Nomor Faktur Pajak'].isna(), 'Nomor Faktur Pajak'] = df_merged.loc[df_merged['Nomor Faktur Pajak'].isna(), 'No. Faktur']

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
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print("Berhasil menyimpan Faktur.xlsx")

except Exception as e:
    print(f"Terjadi kesalahan: {e}")

input("Tekan Enter untuk keluar")