# Sistem Penugasan Pegawai

Aplikasi desktop untuk membantu pengelolaan penugasan pegawai berdasarkan kegiatan, periode, anggaran, komposisi pegawai, dan kalender penugasan.

## 📥 Download

Versi aplikasi terbaru dapat diunduh melalui halaman berikut:

**[Download Sistem Penugasan Pegawai (Latest Release)](../../releases/latest)**

Installer terbaru:

`SistemPenugasan-Setup-v1.2.0.exe`

## ✨ Fitur

- 👥 Pengelolaan data pegawai
- 📊 Import data pegawai dari Excel
- ✏️ Tambah, edit, nonaktifkan, aktifkan kembali, dan hapus permanen pegawai
- 📋 Pengelolaan kegiatan
- 👤 Pengaturan Penanggung Jawab (PJ)
- 💰 Penentuan pegawai berdasarkan anggaran kegiatan
- ⚙️ Penentuan komposisi pegawai berdasarkan golongan dan tarif
- 📅 Kalender penugasan bulanan
- 🚫 Pencegahan penugasan pegawai pada kegiatan berbeda di tanggal yang sama
- 📈 Rekap penugasan bulanan
- 📤 Export data ke Excel

## 🛠️ Teknologi

- **Python**
- **Tkinter**
- **SQLite**
- **OpenPyXL**
- **PyInstaller**
- **Inno Setup**

## 💻 Instalasi

1. Buka halaman [Releases](../../releases).
2. Pilih versi terbaru.
3. Download file installer `.exe`.
4. Jalankan installer.
5. Ikuti proses instalasi sampai selesai.
6. Jalankan **Sistem Penugasan Pegawai** melalui shortcut yang dibuat oleh installer.

## 🗃️ Data Aplikasi

Data aplikasi disimpan secara lokal pada komputer pengguna.

Database disimpan di luar folder instalasi aplikasi, yaitu pada direktori data aplikasi pengguna. Dengan demikian, data tetap dapat dipertahankan ketika aplikasi diperbarui ke versi berikutnya.

## 🔄 Pembaruan Aplikasi

Aplikasi dilengkapi dengan mekanisme pemeriksaan pembaruan yang mengambil informasi versi terbaru melalui `update.json` pada repositori GitHub.

Jika pembaruan tersedia, pengguna dapat memperoleh informasi versi terbaru dan mengunduh installer melalui tautan rilis yang disediakan.

## 📌 Status

**Version 1.2.0 - Latest Release**

Versi 1.2.0 mencakup penyempurnaan pengelolaan penugasan, penentuan komposisi pegawai berdasarkan golongan dan anggaran, pengaturan PJ, validasi jadwal, serta peningkatan mekanisme pembaruan aplikasi.

Lihat [Release Notes](../../releases) untuk mengetahui riwayat versi dan perubahan pada setiap rilis.