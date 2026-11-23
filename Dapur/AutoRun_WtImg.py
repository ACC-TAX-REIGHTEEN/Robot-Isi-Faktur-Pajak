import os
import time
import ast
import warnings
import configparser
import pandas as pd
import pyautogui
import pyperclip

CONFIG_FILE = 'config.conf'
FILE_EXCEL = 'Faktur.xlsx'
NAMA_SHEET = 'Faktur'

POSISI_KOLOM_TEMPEL = (165, 555)
POSISI_TOMBOL_SIMPAN = (1200, 680)
GAMBAR_ERROR_OK = 'tombol_ok.png'

if os.path.exists(CONFIG_FILE):
    try:
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE, encoding='utf-8')
        if 'GENERAL' in config:
            sec = config['GENERAL']
            if 'POSISI_KOLOM_TEMPEL' in sec:
                POSISI_KOLOM_TEMPEL = ast.literal_eval(sec['POSISI_KOLOM_TEMPEL'])
            if 'POSISI_TOMBOL_SIMPAN' in sec:
                POSISI_TOMBOL_SIMPAN = ast.literal_eval(sec['POSISI_TOMBOL_SIMPAN'])
            if 'GAMBAR_ERROR_OK' in sec:
                val = sec['GAMBAR_ERROR_OK'].strip()
                if (val.startswith("'") and val.endswith("'")) or (val.startswith('"') and val.endswith('"')):
                    GAMBAR_ERROR_OK = ast.literal_eval(val)
                else:
                    GAMBAR_ERROR_OK = val
    except Exception as e:
        print(f"--> Peringatan: Gagal membaca {CONFIG_FILE}, memakai default: {e}")
else:
    print(f"--> Peringatan: {CONFIG_FILE} tidak ditemukan, memakai pengaturan default.")


def baca_excel_adaptif(file_path, sheet_name, usecols="B", dtype=str):
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, usecols=usecols, header=0, dtype=dtype, engine='calamine')
    except Exception:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return pd.read_excel(file_path, sheet_name=sheet_name, usecols=usecols, header=0, dtype=dtype, engine='openpyxl')


def jalankan_robot():
    print("--> Sedang membaca file Excel...")
    try:
        df = baca_excel_adaptif(FILE_EXCEL, sheet_name=NAMA_SHEET, usecols="B", dtype=str)
        data_faktur = df.iloc[:, 0].dropna().tolist()
        print(f"--> Ditemukan {len(data_faktur)} data untuk diproses.")
    except Exception as e:
        print(f"--> Gagal membaca Excel: {e}")
        return

    if not os.path.exists(GAMBAR_ERROR_OK):
        print(f"--> PERINGATAN: File gambar '{GAMBAR_ERROR_OK}' tidak ditemukan.")
        print("--> Fitur cek error otomatis mungkin tidak akan berjalan.")
       
    print("--> \n=== PERSIAPAN ===")
    print(f"--> Target Kolom Tempel : {POSISI_KOLOM_TEMPEL}")
    print(f"--> Target Tombol Simpan: {POSISI_TOMBOL_SIMPAN}")
    print(f"--> File Deteksi Popup  : {GAMBAR_ERROR_OK}")
    print("--> 1. Buka Accurate.")
    print("--> 2. Pastikan List Faktur terbuka.")
    print("--> 3. KLIK/SOROT baris data paling atas (Data Nomor 1) tapi JANGAN dibuka dulu.")
    print("--> 4. Robot akan mulai dalam 5 detik...")
    
    for i in range(5, 0, -1):
        print(f"--> {i}...", end=' ', flush=True)
        time.sleep(1)
    print("--> \nMULAI!")

    pyautogui.FAILSAFE = True 

    for index, nilai_faktur in enumerate(data_faktur):
        nilai_str = str(nilai_faktur).strip()
        
        print(f"--> Memproses data ke-{index+1}: {nilai_str}")

        pyautogui.press('enter') 
        time.sleep(2)

        pyautogui.doubleClick(POSISI_KOLOM_TEMPEL)
        time.sleep(1)

        pyautogui.press('backspace')
        time.sleep(0.1)

        pyperclip.copy(nilai_str)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        pyautogui.click(POSISI_TOMBOL_SIMPAN)
   
        time.sleep(3) 

        try:
            lokasi_error = pyautogui.locateCenterOnScreen(GAMBAR_ERROR_OK, confidence=0.9, grayscale=True)
            
            if lokasi_error is not None:
                print("-->    -> Terdeteksi Popup Error/Info! Mengklik OK...")
                pyautogui.click(lokasi_error)
                time.sleep(3)
        except Exception:
            pass
 
        pyautogui.press('down')
        time.sleep(0.5)

    print("--> \n=== TUGAS SELESAI ===")
    
    pyautogui.alert(
        text=f'Sukses! Robot telah memproses {len(data_faktur)} data.\nSilakan cek hasil di Accurate.', 
        title='Robot Python - Selesai', 
        button='OK Siap'
    )

if __name__ == "__main__":
    jalankan_robot()