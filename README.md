# OPD Automation v5

Screenshot tetap manual. Setelah OCR membaca tanggal akhir, pilih project dan
Metro/Broadband melalui popup. Nama screenshot tidak perlu diubah.

## Memperbarui v4
Hentikan program. Salin main.py, START.bat, dan README.md dari paket ini ke folder
project lama. Pertahankan .env dan .venv. Tidak perlu instal ulang Tesseract.

Jalankan dengan klik dua kali START.bat atau:
python .\main.py --watch

## Perubahan
- Daftar sheet dibaca langsung dari metadata XLSX, tanpa memuat workbook lengkap.
- Ketik nama project untuk menyaring dropdown, lalu pilih nama persis dari daftar.
- Project dan jaringan wajib dipilih; tidak memilih project pertama secara diam-diam.
- Detektor folder berjalan di thread terpisah dengan interval 0,4 detik.
- Antrean tetap menerima gambar saat popup atau penyimpanan berjalan.
- Workbook disimpan dan diperiksa di disk lokal, lalu disalin ke Drive.
- Pemeriksaan CRC arsip menggantikan pembukaan ulang semua gambar untuk verifikasi.
- Perubahan ukuran/waktu file tujuan diperiksa sebelum penggantian.
  Ini membantu mendeteksi perubahan bersamaan, bukan kunci kolaborasi penuh.
- Backup unik, notifikasi duplikat tanpa traceback, dan durasi per tahap.
- START.bat tidak menjalankan skrip PowerShell sehingga tidak terkena kebijakan tanda tangan ps1.
- Batas waktu OCR 45 detik.

## Konfigurasi
Gunakan .env lama, dengan NOTIFICATION_MODE=console selama pengujian.
ALLOW_REPLACE_EXISTING=false menjaga gambar lama.
Daftar project mengikuti jumlah sheet file tanggal, bukan angka tetap 36.

## Batasan
Popup masih menunggu OCR dan file tanggal tersedia.
Antrean berada di memori, tidak dipulihkan setelah program ditutup.
File yang sudah ada saat program mulai tidak otomatis dimasukkan.
Ulangi file gagal/batal menggunakan:
python .\main.py --once "C:\lokasi\Screenshot.png"

Penyimpanan berhasil berarti file di folder Drive tersimpan, bukan bukti selesai
sinkronisasi online. Periksa aplikasi Drive untuk status unggahan.
Kecepatan belum diukur pada laptop Windows pengguna.
OCR/crop dan ukuran tampilan gambar tetap seperti versi sebelumnya.
Pengaturan MAX_IMAGE mengubah tampilan di Excel, tidak mengompresi piksel sumber.
Tutup Excel tujuan saat memproses.

Integrasi deteksi NO_DATA dan workflow WhatsApp belum selesai. Versi ini mencegah
notifikasi eksternal untuk hasil normal. Jangan menganggap mode n8n sebagai
monitor gangguan aktif. Tidak ada perubahan atau pesan dikirim ke layanan luar
saat pengujian paket.

## Verifikasi pengembang
Diuji pada salinan template lokal: gambar tersisip, arsip valid, backup dibuat,
slot duplikat tidak mengubah file. OCR dimock pada uji penyimpanan.
Popup Windows dan sinkronisasi Drive harus diuji pada laptop pengguna.
