# Screenshot OPD ke Excel Google Drive

Program memantau screenshot Windows, membaca tanggal akhir grafik dengan OCR, memilih file Excel tanggal yang sama di Google Drive for Desktop, lalu menampilkan popup untuk memilih project serta Metro/Broadband sebelum gambar dimasukkan.

Program tidak menambahkan teks ke spreadsheet. Sebelum file Drive diubah, salinan cadangan otomatis disimpan di folder lokal `backups`.

## Alur

```text
Screenshot Ruijie
→ OCR membaca tanggal akhir
→ cari DD SEPTEMBER YYYY.xlsx
→ tampilkan project dari nama-nama sheet
→ pilih project dan Metro/Broadband pada popup
→ cari judul Link Broadband/Metro
→ potong bagian Speed Summary
→ masukkan gambar
→ Google Drive menyinkronkan file
→ kirim notifikasi
```

## Persyaratan

- Windows 10/11.
- Python 3.11 atau lebih baru.
- Google Drive for Desktop aktif.
- Tesseract OCR.
- File Excel tujuan tidak sedang dibuka saat gambar dimasukkan.

## Instalasi

Buka PowerShell di folder project:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\install.ps1
```

Jika Tesseract belum tersedia, jalankan:

```powershell
.\install_tesseract.ps1
```

Setelah instalasi Tesseract, tutup dan buka kembali PowerShell.

## Konfigurasi `.env`

Jika `.env` lama masih menggunakan `REPORTS_FOLDER`, ganti isinya dengan konfigurasi berikut:

```env
SCREENSHOT_FOLDER=C:\Users\acera\OneDrive\Pictures\Screenshots
DRIVE_REPORT_FOLDER=G:\My Drive\Latihan OPD Automation\SEPTEMBER
BACKUP_FOLDER=backups

TESSERACT_COMMAND=C:\Program Files\Tesseract-OCR\tesseract.exe
MAX_IMAGE_WIDTH_PX=1220
MAX_IMAGE_HEIGHT_PX=330
CROP_SPEED_SUMMARY=true
ALLOW_REPLACE_EXISTING=false
TIMEZONE_OFFSET_HOURS=7

NOTIFICATION_MODE=console
N8N_WEBHOOK_URL=
N8N_WEBHOOK_TOKEN=
```

`ALLOW_REPLACE_EXISTING=false` mencegah program menimpa screenshot yang sudah ada pada slot yang sama.
Jika slot sudah terisi, PowerShell menampilkan notifikasi `DUPLIKAT / SLOT SUDAH TERISI`
tanpa traceback error. Screenshot tidak dimasukkan dan file Excel tidak diubah.

## Nama screenshot

Nama screenshot otomatis dari Windows boleh langsung digunakan, misalnya:

```text
Screenshot 2026-09-15 101500.png
```

Program akan menampilkan popup. Daftar project di popup dibaca otomatis dari nama-nama sheet file Excel yang sesuai dengan tanggal pada screenshot.

Format lama berikut juga tetap didukung dan akan diproses tanpa popup:

```text
24-RG-BPBD-PROVSU_METRO.png
24-RG-BPBD-PROVSU_BROADBAND.png
```

Program mengabaikan token `RG` ketika mencocokkan project, sehingga `24-RG-BPBD-PROVSU` dapat masuk ke sheet `24-BPBD-PROVSU`.


## Aturan tanggal

Jika judul screenshot adalah:

```text
2026/09/14–2026/09/15 Speed Summary
```

program menggunakan tanggal akhir dan memilih `15 SEPTEMBER 2026.xlsx`.

## Menjalankan

```powershell
.\run.ps1
```

Jika PowerShell memblokir script:

```powershell
.\.venv\Scripts\python.exe .\main.py --watch
```

## Pengujian satu screenshot

Gunakan file tanggal yang masih kosong. Tutup Excel, kemudian jalankan screenshot dengan nama apa adanya:

```powershell
.\test_once.ps1 "C:\Users\acera\OneDrive\Pictures\Screenshots\Screenshot 2026-09-15 101500.png"
```

Setelah tanggal terbaca, popup muncul. Pilih nama project dan jenis jaringan, kemudian klik `Proses`.

Jika berhasil, PowerShell menampilkan file, sheet, dan posisi gambar. Periksa juga status sinkronisasi Google Drive.

## Perlindungan file

- File asli dicadangkan ke `backups\YYYY-MM-DD` sebelum perubahan.
- Slot yang sudah memiliki gambar tidak ditimpa secara default.
- File sementara diverifikasi bisa dibuka sebelum menggantikan file Drive.
- Screenshot yang sama tidak diproses dua kali.
- Jika file sedang dibuka di Excel, program menolak menyimpan dan menampilkan error.

## Notifikasi

Gunakan mode console selama pengujian:

```env
NOTIFICATION_MODE=console
```

Setelah penyisipan ke Drive stabil, ubah ke n8n:

```env
NOTIFICATION_MODE=n8n
N8N_WEBHOOK_URL=https://nnael.app.n8n.cloud/webhook/screenshot-excel
N8N_WEBHOOK_TOKEN=isi-token-sendiri
```

Token tidak boleh disimpan di GitHub atau dikirim melalui percakapan.
