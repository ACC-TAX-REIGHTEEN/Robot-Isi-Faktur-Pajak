# 🤖 Robot Isi Faktur Pajak

> **Otomasi pengisian Nomor Faktur Pajak ke dalam Accurate — robot klik dan tempel satu per satu, tanpa sentuh keyboard**

Pipeline Python dua tahap: tahap data (bersihkan ekspor Accurate + lookup nomor faktur pajak dari Coretax) menghasilkan daftar `Faktur.xlsx`, lalu robot GUI mengambil alih — membuka setiap rekaman di Accurate, menempel nomor faktur pajak ke kolom yang tepat, menekan Simpan, dan berpindah ke rekaman berikutnya secara otomatis hingga seluruh daftar selesai.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan (Lengkap)](#-cara-penggunaan-lengkap)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
- [Kalibrasi Koordinat Layar](#-kalibrasi-koordinat-layar)
- [Format File Input](#-format-file-input)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting & Keamanan](#-catatan-penting--keamanan)

---

## 🗂️ Gambaran Umum

Pengisian `Nomor Faktur Pajak` di Accurate harus dilakukan satu per satu secara manual: buka rekaman → klik kolom → ketik nomor → simpan → buka rekaman berikutnya. Untuk ratusan faktur, ini memakan berjam-jam.

Robot ini mengotomasi seluruh proses tersebut:

```
Acc.xls (ekspor Accurate) + Ctx.xlsx (ekspor Coretax)
                  ↓
    1. Bersihkan & ekstrak data Accurate
    2. Lookup Nomor Faktur Pajak dari Coretax → Faktur.xlsx
                  ↓
    3. Robot mengambil alih layar:
       Enter → klik kolom → paste → Simpan → pindah ke berikutnya
       (otomatis untuk setiap baris di Faktur.xlsx)
```

> ⚠️ **Windows only.** Robot menggunakan `pyautogui` untuk mengontrol mouse dan keyboard secara nyata. Jangan sentuh mouse/keyboard selama robot berjalan.

---

## 🔧 Prasyarat

### Python
Python **3.8+** pada **Windows**.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl pyautogui pyperclip pillow
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Skrip 1, 2, AutoRun | Baca dan transformasi data Excel |
| `openpyxl` | Skrip 1, 2 | Tulis `.xlsx` dan auto-fit kolom |
| `pyautogui` | AutoRun, cek_koordinat | Kontrol mouse & keyboard otomatis |
| `pyperclip` | AutoRun | Copy-paste via clipboard |
| `pillow` | AutoRun_WtImg | Pengenalan gambar layar (image recognition) |
| `xlrd` | Skrip 1 | Baca `Acc.xls` format lama |

> **`xlrd`:** Pastikan versi yang kompatibel dengan `.xls`:
> ```bash
> pip install "xlrd>=1.0.0,<2.0.0"
> ```

### Aplikasi yang harus terbuka
- **Accurate** — Harus sudah terbuka di komputer yang sama, di layar yang sama, dengan tampilan **List Faktur** aktif sebelum robot dimulai.

---

## 📁 Struktur Folder & File

```
📦 Robot-Isi-Faktur-Pajak/
│
├── 📄 ROBOT_ISI_FAKTUR_PAJAK.py         ← Titik masuk utama. Jalankan ini
│
├── 📄 Acc.xls                           ← [INPUT] Ekspor faktur dari Accurate
├── 📄 Ctx.xlsx                          ← [INPUT] Ekspor data dari Coretax DJP
│
└── 📁 Dapur/                            ← Folder pipeline (jangan diubah)
    ├── 📄 __init__.py
    ├── 📄 1_ACC_Cleaner.py              ← Bersihkan Acc.xls → Acc_Bersih.xlsx
    ├── 📄 2_MergedxLook.py             ← Lookup & merge → Faktur.xlsx
    ├── 📄 AutoRun_WtImg.py             ← Robot utama (dengan deteksi popup error)
    ├── 📄 AutoRun_1.py                 ← Robot sederhana (tanpa deteksi popup)
    ├── 📄 tombol_ok.png                ← Referensi gambar tombol OK untuk deteksi popup
    └── 📄 cek_koordinatXY_layar.py     ← Alat kalibrasi koordinat layar
```

---

## 🚀 Cara Penggunaan (Lengkap)

### Langkah 1 — Kalibrasi koordinat layar (wajib, lakukan sekali)

Koordinat klik robot harus disesuaikan dengan posisi elemen di layar **dan resolusi monitor Anda**. Ikuti panduan di [Kalibrasi Koordinat Layar](#-kalibrasi-koordinat-layar).

### Langkah 2 — Siapkan file input

Letakkan dua file berikut di folder utama (sejajar dengan `ROBOT_ISI_FAKTUR_PAJAK.py`):

| File | Nama wajib | Sumber |
|---|---|---|
| Ekspor Accurate | `Acc.xls` | List Faktur dari Accurate, diekspor ke `.xls` |
| Ekspor Coretax | `Ctx.xlsx` | Unduh data Faktur Pajak dari Coretax DJP |

### Langkah 3 — Buka Accurate dan siapkan tampilan

1. Buka **Accurate**.
2. Navigasi ke **List Faktur** (daftar faktur penjualan).
3. **Sorot/klik baris pertama** (data nomor 1) — jangan dibuka dulu.
4. Pastikan tidak ada dialog atau popup lain yang terbuka.

### Langkah 4 — Jalankan robot

```bash
python ROBOT_ISI_FAKTUR_PAJAK.py
```

Pipeline akan:
1. Memvalidasi semua file syarat
2. Membersihkan file lama di `Dapur/`
3. Menyalin `Acc.xls` dan `Ctx.xlsx` ke `Dapur/`
4. Menjalankan `1_ACC_Cleaner.py` (data prep)
5. Menjalankan `2_MergedxLook.py` (lookup & merge)
6. Menjalankan robot `AutoRun_WtImg.py` (atau `AutoRun_1.py`)

### Langkah 5 — Jangan sentuh komputer

Setelah robot mulai (countdown 5 detik selesai), **jangan gerakkan mouse atau tekan keyboard** hingga robot selesai. Jika harus dihentikan darurat: **gerakkan mouse ke pojok kiri atas layar** — robot akan berhenti otomatis (PyAutoGUI Failsafe).

### Langkah 6 — Selesai

Popup `"Sukses! Robot telah memproses N data."` muncul ketika semua baris selesai.

---

## 🔄 Alur Kerja Pipeline

```
[ROBOT_ISI_FAKTUR_PAJAK.py]
   │
   ├─── Validasi folder Dapur/ + 6 file syarat
   ├─── Hapus *.xls, *.xlsx lama dari Dapur/
   ├─── Salin Acc.xls & Ctx.xlsx → Dapur/
   ├─── Pindah ke direktori Dapur/
   │
   ├─── [1] 1_ACC_Cleaner.py
   │       Baca Acc.xls (header di baris ke-4, index=3)
   │       Pilih 7 kolom: No. Faktur, Tgl Faktur, No. Pelanggan,
   │                      Nama Pelanggan, Nilai Faktur, Terutang, Keterangan
   │       Hapus baris tanpa No. Faktur
   │       Bersihkan artefak numerik (.0, ,00) dari kolom angka
   │       Auto-fit kolom → simpan: Acc_Bersih.xlsx
   │
   ├─── [2] 2_MergedxLook.py
   │       Baca Acc_Bersih.xlsx (sheet: Data) → kolom No. Faktur
   │       Baca Ctx.xlsx (sheet: data) → kolom Referensi & Nomor Faktur Pajak
   │       Hapus duplikat Referensi (pertahankan yang pertama)
   │       LEFT JOIN: No. Faktur ← Referensi
   │       Fallback: jika tidak cocok → Nomor Faktur Pajak = No. Faktur asli
   │       Simpan: Faktur.xlsx (sheet: Faktur)
   │         Kolom A: No. Faktur
   │         Kolom B: Nomor Faktur Pajak ← kolom yang dibaca robot
   │
   ├─── [3] AutoRun_WtImg.py  (robot utama, dengan deteksi popup)
   │       Baca kolom B Faktur.xlsx → daftar Nomor Faktur Pajak
   │       Tampilkan instruksi + countdown 5 detik
   │       Per baris:
   │         1. Enter → buka rekaman faktur (tunggu 2 detik)
   │         2. Double-click POSISI_KOLOM_TEMPEL → fokus ke kolom No. FP
   │         3. Backspace → hapus nilai lama
   │         4. Paste via clipboard (Ctrl+V)
   │         5. Click POSISI_TOMBOL_SIMPAN → simpan (tunggu 3 detik)
   │         6. Cek popup error via image recognition (tombol_ok.png)
   │            → jika terdeteksi: klik OK, tunggu 3 detik lagi
   │         7. Down arrow → pindah ke baris berikutnya (tunggu 0.5 detik)
   │       Tampilkan popup selesai via pyautogui.alert()
   │
   └─── [4] AutoRun_1.py  (robot sederhana, tanpa deteksi popup)
           Alur sama dengan AutoRun_WtImg.py
           tapi TANPA langkah 6 (tidak cek popup error)
```

> **Catatan urutan eksekusi:** Orkestrator menjalankan `AutoRun_WtImg.py` terlebih dahulu, diikuti `AutoRun_1.py`. Artinya robot berjalan **dua kali** berurutan. Untuk menjalankan hanya satu robot, eksekusi `1_ACC_Cleaner.py` dan `2_MergedxLook.py` terlebih dahulu secara manual, lalu pilih satu robot (`AutoRun_WtImg.py` atau `AutoRun_1.py`) dan jalankan sendiri.

---

## 🔍 Detail Tiap Skrip

### `1_ACC_Cleaner.py`

Membaca `Acc.xls` dengan header di baris ke-4 (indeks 3 — tiga baris pertama biasanya berisi metadata laporan Accurate). Tujuh kolom dipilih berdasarkan nama:

| Kolom | Keterangan |
|---|---|
| `No. Faktur` | Nomor faktur penjualan (kunci JOIN) |
| `Tgl Faktur` | Tanggal faktur |
| `No. Pelanggan` | Kode pelanggan |
| `Nama Pelanggan` | Nama pelanggan |
| `Nilai Faktur` | Nilai total faktur |
| `Terutang` | Sisa piutang |
| `Keterangan` | Keterangan tambahan |

Kolom `No. Faktur`, `Nilai Faktur`, `Terutang`, dan `No. Pelanggan` dibersihkan dari artefak desimal: `,00` dan `.0` yang muncul saat Accurate mengekspor angka ke format teks.

---

### `2_MergedxLook.py`

Melakukan LEFT JOIN antara data Accurate dan Coretax:

```
Acc_Bersih.xlsx          Ctx.xlsx
   No. Faktur    ←JOIN→   Referensi
                           Nomor Faktur Pajak
```

**Logika fallback:** Jika `No. Faktur` dari Accurate tidak ditemukan di kolom `Referensi` Coretax (hasil join NULL atau kosong), kolom `Nomor Faktur Pajak` diisi dengan `No. Faktur` asli. Ini memastikan robot selalu punya nilai untuk ditempel, meski mapping tidak ditemukan.

**Output `Faktur.xlsx` (sheet: Faktur):**

| Kolom A | Kolom B |
|---|---|
| `No. Faktur` | `Nomor Faktur Pajak` ← **dibaca robot** |

---

### `AutoRun_WtImg.py` — Robot Utama (Direkomendasikan)

Versi robot dengan **deteksi popup error otomatis** menggunakan pengenalan gambar (`tombol_ok.png`). Jika setelah menekan Simpan muncul dialog error atau konfirmasi dari Accurate, robot secara otomatis mendeteksi dan mengklik tombol OK sebelum melanjutkan ke rekaman berikutnya.

**Dua konstanta koordinat yang harus dikalibrasi:**

```python
POSISI_KOLOM_TEMPEL = (165, 555)   # Koordinat kolom "No. Faktur Pajak" di Accurate
POSISI_TOMBOL_SIMPAN = (1200, 680) # Koordinat tombol Simpan di Accurate
```

---

### `AutoRun_1.py` — Robot Sederhana

Versi tanpa deteksi popup. Lebih sederhana dan cocok jika Accurate di lingkungan Anda tidak pernah memunculkan dialog error/konfirmasi saat menyimpan.

Dua konstanta koordinat yang sama harus dikalibrasi seperti `AutoRun_WtImg.py`.

---

### `cek_koordinatXY_layar.py` — Alat Kalibrasi

Menampilkan posisi koordinat mouse secara real-time di terminal. Digunakan untuk menentukan nilai `POSISI_KOLOM_TEMPEL` dan `POSISI_TOMBOL_SIMPAN`. Hentikan dengan **Ctrl+C**.

```
Posisi Mouse: X=165, Y=555
```

---

## 🎯 Kalibrasi Koordinat Layar

Nilai koordinat `POSISI_KOLOM_TEMPEL` dan `POSISI_TOMBOL_SIMPAN` **berbeda-beda di setiap komputer** tergantung resolusi layar, zoom Windows, dan posisi jendela Accurate. Kalibrasi **wajib** dilakukan sebelum pertama kali menggunakan robot.

### Langkah kalibrasi

**1. Buka Accurate dan buka satu rekaman faktur.**

Buka rekaman faktur mana saja di Accurate sehingga terlihat form edit faktur lengkap dengan kolom `No. Faktur Pajak` dan tombol Simpan.

**2. Jalankan `cek_koordinatXY_layar.py`:**

```bash
cd Dapur
python cek_koordinatXY_layar.py
```

**3. Cari koordinat kolom `No. Faktur Pajak`:**

Arahkan kursor ke **dalam kotak input kolom "No. Faktur Pajak"** di form Accurate. Catat nilai `X` dan `Y` yang muncul di terminal.

**4. Cari koordinat tombol Simpan:**

Arahkan kursor ke **tengah tombol Simpan** di form Accurate. Catat nilai `X` dan `Y`.

**5. Tutup `cek_koordinatXY_layar.py`** dengan Ctrl+C.

**6. Update konstanta di kedua file robot:**

Buka `Dapur/AutoRun_WtImg.py` dan `Dapur/AutoRun_1.py`, ubah baris:

```python
# Ganti dengan hasil kalibrasi Anda
POSISI_KOLOM_TEMPEL = (165, 555)   # ← masukkan koordinat kolom No. FP
POSISI_TOMBOL_SIMPAN = (1200, 680) # ← masukkan koordinat tombol Simpan
```

### Tips kalibrasi

- Jalankan kalibrasi dengan jendela Accurate **dalam posisi yang sama** seperti saat robot akan berjalan (ukuran jendela, posisi di layar).
- Jika layar memiliki scaling Windows (misalnya 125% atau 150%), gunakan koordinat yang ditampilkan `cek_koordinatXY_layar.py` — sudah memperhitungkan scaling.
- Jangan memindahkan atau mengubah ukuran jendela Accurate setelah kalibrasi dan sebelum robot selesai.

---

## 📋 Format File Input

### `Acc.xls` — Ekspor Accurate

Diekspor dari menu **List Faktur Penjualan** di Accurate ke format `.xls`. Skrip membaca baris ke-4 (index 3) sebagai header — artinya tiga baris pertama file dianggap metadata laporan dan dilewati.

**Kolom yang wajib ada (nama persis):**

```
No. Faktur | Tgl Faktur | No. Pelanggan | Nama Pelanggan | Nilai Faktur | Terutang | Keterangan
```

### `Ctx.xlsx` — Ekspor Coretax DJP

Diekspor dari portal Coretax DJP. **Wajib memiliki sheet bernama `data`** (huruf kecil) dengan minimal dua kolom:

| Kolom | Keterangan |
|---|---|
| `Referensi` | Nomor faktur di Accurate yang menjadi kunci pencarian |
| `Nomor Faktur Pajak` | Nomor faktur pajak dari Coretax yang akan diisikan ke Accurate |

Jika ada duplikat pada kolom `Referensi`, hanya baris pertama yang digunakan.

---

## 🛠️ Troubleshooting

### ❌ `File Acc.xls tidak ditemukan di folder utama`
Letakkan file ekspor Accurate di folder utama (sejajar dengan `ROBOT_ISI_FAKTUR_PAJAK.py`) dengan nama persis `Acc.xls`.

### ❌ `Terjadi kesalahan: 'No. Faktur' tidak ditemukan`
Header kolom tidak ditemukan. Pastikan ekspor Accurate menggunakan header `No. Faktur` dan header tersebut berada di baris ke-4 (index 3). Jika posisi header berbeda, ubah nilai `header=3` di `1_ACC_Cleaner.py`.

### ❌ `Terjadi kesalahan: Worksheet named 'data' not found`
Sheet di `Ctx.xlsx` tidak bernama `data` (huruf kecil). Buka file Coretax dan periksa nama tab sheet — ubah menjadi `data` atau ganti `sheet_name='data'` di `2_MergedxLook.py` sesuai nama aktual.

### ❌ Robot mengklik posisi yang salah (mengenai elemen lain)
Koordinat belum dikalibrasi atau jendela Accurate bergeser. Jalankan ulang `cek_koordinatXY_layar.py` dan kalibrasi ulang. Pastikan jendela Accurate tidak dipindahkan setelah kalibrasi.

### ❌ Robot berjalan terlalu cepat / Accurate belum siap
Nilai `time.sleep()` mungkin terlalu kecil untuk spesifikasi komputer Anda. Naikkan nilai delay di baris-baris berikut di `AutoRun_WtImg.py` / `AutoRun_1.py`:

```python
time.sleep(2)   # setelah Enter (buka rekaman) → naikkan ke 3-5
time.sleep(3)   # setelah Simpan → naikkan ke 4-6
```

### ❌ Robot gagal mendeteksi popup error (AutoRun_WtImg)
File `tombol_ok.png` tidak cocok dengan tampilan dialog di komputer Anda. Ambil screenshot baru dari tombol OK yang muncul di Accurate, crop rapat, simpan sebagai `Dapur/tombol_ok.png`. Pastikan ukuran dan warna persis sama.

### ❌ `ModuleNotFoundError: No module named 'pyautogui'`
```bash
pip install pyautogui pyperclip pillow
```

### ❌ Robot tiba-tiba berhenti (FailSafe Triggered)
Mouse tergerak ke pojok kiri atas layar. Ini adalah fitur keamanan `pyautogui`. Pindahkan mouse menjauh dari pojok kiri atas dan jalankan ulang robot dari awal (data yang sudah diproses tidak akan otomatis diulang — perlu disesuaikan manual).

---

## ⚠️ Catatan Penting & Keamanan

- **Jangan sentuh mouse atau keyboard selama robot berjalan.** Robot mengontrol input perangkat secara nyata — gangguan apa pun dapat menyebabkan klik di posisi yang salah.
- **FAILSAFE aktif.** Gerakkan mouse ke pojok kiri atas layar (koordinat 0,0) kapan saja untuk menghentikan robot secara darurat. Ini adalah fitur bawaan PyAutoGUI (`FAILSAFE = True`).
- **Robot tidak bisa di-undo.** Setelah data tersimpan di Accurate, tidak ada mekanisme rollback otomatis. Pastikan data di `Faktur.xlsx` sudah benar sebelum menjalankan robot.
- **Kalibrasi wajib per perangkat.** Koordinat `POSISI_KOLOM_TEMPEL` dan `POSISI_TOMBOL_SIMPAN` hardcoded dan berbeda di setiap komputer. Nilai default (`165, 555` dan `1200, 680`) kemungkinan tidak valid di komputer Anda.
- **Resolusi dan zoom layar berpengaruh.** Jika Windows Display Scaling berubah (misalnya dari 100% ke 125%), lakukan kalibrasi ulang.
- **Dua robot berjalan berurutan di orchestrator.** `ROBOT_ISI_FAKTUR_PAJAK.py` menjalankan `AutoRun_WtImg.py` kemudian `AutoRun_1.py` secara berurutan. Untuk menghindari pemrosesan ganda, jalankan satu robot saja secara manual setelah step persiapan data (`1_ACC_Cleaner.py` dan `2_MergedxLook.py`) selesai.
- **Layar tidak boleh terkunci.** Screen lock atau screen saver akan menghentikan robot karena tidak ada elemen Accurate yang bisa diklik.
- **Fallback di `2_MergedxLook.py`.** Jika nomor faktur tidak ditemukan di Coretax, robot akan menempelkan nomor faktur Accurate itu sendiri. Verifikasi `Faktur.xlsx` sebelum menjalankan robot untuk memastikan hasilnya sesuai.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

* Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)
