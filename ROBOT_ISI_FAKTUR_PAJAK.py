import os
import sys
import shutil
import subprocess

def main():
    folder_dapur = 'Dapur'
    
    skrip_berurutan = [
        '1_ACC_Cleaner.py',
        '2_MergedxLook.py',
        'AutoRun_WtImg.py',
        'tombol_ok.png',
        'AutoRun_1.py',
        'cek_koordinatXY_layar.py'
    ]
    
    file_sumber = [
        'Acc.xls', 
        'Ctx.xlsx'
    ]

    if not os.path.exists(folder_dapur):
        print(f"Error: Folder '{folder_dapur}' tidak ditemukan.")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    list_skrip_hilang = []
    for nama_file in skrip_berurutan:
        path_lengkap = os.path.join(folder_dapur, nama_file)
        if not os.path.exists(path_lengkap):
            list_skrip_hilang.append(nama_file)

    if list_skrip_hilang:
        print(f"Error: File script berikut tidak ditemukan di dalam folder '{folder_dapur}':")
        for f in list_skrip_hilang:
            print(f"- {f}")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    list_sumber_hilang = []
    for nama_file in file_sumber:
        if not os.path.exists(nama_file):
            list_sumber_hilang.append(nama_file)

    if list_sumber_hilang:
        print("Error: File sumber berikut tidak ditemukan di folder utama:")
        for f in list_sumber_hilang:
            print(f"- {f}")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    try:
        for nama_file in os.listdir(folder_dapur):
            if nama_file.lower().endswith(('.xls', '.xlsx')):
                file_path = os.path.join(folder_dapur, nama_file)
                os.remove(file_path)
    except Exception as e:
        print(f"Gagal membersihkan file Excel lama di Dapur: {e}")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    try:
        print("Menyalin file data ke folder sistem...")
        for nama_file in file_sumber:
            tujuan = os.path.join(folder_dapur, nama_file)
            shutil.copy(nama_file, tujuan)
        print("Penyalinan berhasil.\n")
    except Exception as e:
        print(f"Gagal menyalin file: {e}")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    cwd_awal = os.getcwd()

    try:
        os.chdir(folder_dapur)
        
        for skrip in skrip_berurutan:
            print(f"=== Menjalankan: {skrip} ===")
            
            if skrip in ['1_ACC_Cleaner.py', '2_MergedxLook.py']:
                try:
                    subprocess.run([sys.executable, skrip], input=b'\n', check=True)
                except subprocess.CalledProcessError:
                    print(f"\nError: Terjadi kesalahan saat menjalankan {skrip}.")
                    input("Tekan Enter untuk melihat detail error dan keluar...")
                    sys.exit()
            else:
                subprocess.run([sys.executable, skrip], check=True)
            
            print(f"=== Selesai: {skrip} ===\n")
        
        print("SELURUH RANGKAIAN PROSES SELESAI.")

    except subprocess.CalledProcessError:
        print(f"\nProses berhenti karena terjadi error pada langkah terakhir.")
        input("Tekan Enter untuk keluar...")
    except Exception as e:
        print(f"\nTerjadi kesalahan tak terduga: {e}")
        input("Tekan Enter untuk keluar...")
    finally:
        os.chdir(cwd_awal)

if __name__ == "__main__":
    main()