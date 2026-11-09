# Pasangkan Foto – Panduan Singkat

## Pasang di komputer lain
**Cara termudah: `Aplikasi\PasangkanFoto_Setup.exe` (34 MB).** Salin file itu ke komputer tujuan, klik dua kali, lalu **Pasang**.
- Tidak perlu akses administrator. Terpasang di `%LOCALAPPDATA%\Programs\Pasangkan Foto`, dengan shortcut di Start Menu (dan Desktop, bisa dimatikan).
- Memasang lagi = memperbarui versi lama (pengaturan tidak hilang). Kalau aplikasi sedang terbuka, installer menawarkan menutupnya.
- Menghapus: Windows **Settings › Apps › Pasangkan Foto › Uninstall**.
- Butuh **Windows 10 / 11 64-bit**. Pertama kali dibuka, Windows mungkin menampilkan "Windows protected your PC" (aplikasi belum bertanda tangan digital): **More info → Run anyway** (sekali per komputer).
- Alternatif tanpa installer: salin folder `PasangkanFoto_Cepat` utuh, atau file `PasangkanFoto.exe`.

## Pakai (3 langkah)
1. Buka salah satu:
   - `Aplikasi\PasangkanFoto_Cepat\PasangkanFoto_Cepat.exe` – **versi folder, cepat dibuka (±1,5 dtk)**. Salin/pindahkan **seluruh folder** `PasangkanFoto_Cepat` (jangan hanya file .exe-nya). Untuk shortcut: klik kanan .exe → Send to → Desktop.
   - `Aplikasi\PasangkanFoto.exe` – satu file, mudah dikirim, tapi butuh 3–8 dtk untuk terbuka.
2. **Pilih file data…** (`.xlsx` atau `.csv`; kalau Excel punya banyak sheet, pilih **Sheet**) dan **Pilih folder foto…**. Hasil muncul **otomatis**, tanpa tombol Proses.
3. Perbaiki baris berwarna (pilih barisnya):
   - **Ganti foto** – pilih foto yang benar
   - **Edit tanggal** – ketik tanggal yang benar
   - **Reset** – kembali ke hasil otomatis
   - klik dua kali baris – buka fotonya

Lalu **Buka hasil** → file `<nama> - <sheet>_merge.txt`. Itu yang dipakai di InDesign:
`Window › Utilities › Data Merge › Select Data Source`. Kolom foto bernama `@Foto`.

Tombol lain: **↻ Proses ulang** (setelah memperbaiki nama file foto), **Semua data** (melihat seluruh isi sheet), **⚙** (ganti kolom nama / tanggal kalau salah terdeteksi).

## Arti angka & warna
| | |
|---|---|
| **Cocok** | foto ketemu dan yakin |
| **Cek** (kuning) | foto mirip tapi ragu → cek manual. Dikosongkan di file hasil |
| **Tanpa foto** (merah) | foto tidak ketemu |
| **Tanggal cek** (oranye) | tanggal ragu / tidak terbaca / mungkin tertukar |
| abu / biru | baris diabaikan (mis. judul) / foto tidak terpakai |

## Tanggal
- **Format tanggal** menyeragamkan semua tanggal (mis. `21 Agustus 2000`). Data aslimu tidak diubah.
- Banner merah **"kemungkinan tertukar hari↔bulan"**: Excel kadang salah membaca `06-09-2009` (6 September) sebagai 9 Juni. Cek 1–2 nama ke database; kalau memang tertukar, centang **Tukar hari↔bulan** di banner itu.

## Catatan
- File hasil selalu **UTF-16 (.txt)** – hanya format ini yang terbaca InDesign-mu.
- File bayangan Mac (`._Nama.jpg`) dan file tersembunyi otomatis diabaikan.
- Pilih folder foto **per kelas** bila bisa, supaya tidak tercampur kelas lain.
- Pintasan: `Ctrl+Enter` = Proses ulang, `Ctrl+O` = pilih file.
- Windows mungkin menampilkan peringatan SmartScreen saat pertama dibuka: **More info → Run anyway**.

## Untuk pengembang
- Bangun ulang `.exe`: `python -m PyInstaller --noconfirm --onefile --windowed --name PasangkanFoto --icon ikon.ico --distpath Aplikasi aplikasi_foto.py`
- Installer: setelah folder `Aplikasi\PasangkanFoto_Cepat` dibangun ulang (perintah `--onedir`, nama `PasangkanFoto_Cepat`), jalankan `python buat_installer.py` → `Aplikasi\PasangkanFoto_Setup.exe`
- Tes: `python tes_inti.py` dan `python tes_tanggal.py`
- Baris perintah: `python cocokkan_foto.py data.xlsx "folder foto" --sheet "IPA 1" --kolom-tanggal auto` (`--utf8` untuk CSV biasa, `--tukar-hari-bulan`)
