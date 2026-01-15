# Excel VBA: Automated Attendance Splitter (Modular)

Script VBA ini dirancang untuk mengolah data absensi karyawan secara otomatis. Program akan memisahkan data dari satu lembar kerja utama (*Master*) menjadi beberapa lembar kerja (*Sheets*) berdasarkan Nama atau ID karyawan, sekaligus menghitung keterlambatan dan kepulangan awal secara otomatis.

## ✨ Fitur Utama
- **Modular Design**: Kode dipecah menjadi fungsi-fungsi kecil agar mudah dimodifikasi.
- **Auto-Split**: Membuat sheet baru untuk setiap karyawan secara otomatis.
- **Logika Jam Kerja Dinamis**:
  - **Jumat**: Masuk 07:00, Pulang 11:30.
  - **Sabtu**: Pulang 13:15.
  - **Senin-Kamis**: Masuk 07:15, Pulang 14:00.
- **Smart Validation**: 
  - Mendeteksi selisih waktu yang tidak wajar (di atas 3 jam) sebagai error.
  - Membersihkan karakter terlarang (`/`, `*`, `?`, dll) agar nama sheet tidak error.
  - Validasi format data (menghindari *crash* jika ada sel kosong atau teks).

## 📊 Persyaratan Format Data
Agar script berjalan lancar, pastikan Sheet Master Anda memiliki struktur kolom sebagai berikut:

| Kolom | Nama Data | Keterangan |
|-------|-----------|------------|
| A     | Tanggal   | Format: Date (DD/MM/YYYY) |
| B     | Nama/ID   | Kriteria pemisah sheet |
| C     | Jam Log   | Format: Time (HH:MM:SS) |
| D     | Status    | Harus mengandung kata "IN" atau "OUT" |

## 🚀 Cara Penggunaan
1. Download file `CreateSheet.bas` dari repositori ini.
2. Buka file Excel Anda (simpan sebagai `.xlsm`).
3. Tekan `Alt + F11` untuk membuka VBA Editor.
4. Klik kanan pada folder *Modules* > **Import File...** > Pilih `AbsensiModular.bas`.
5. Kembali ke Excel, tekan `Alt + F8`, pilih `PisahAbsensiModular`, lalu klik **Run**.

## 📁 Struktur Repositori
- `template/`: Contoh file Excel untuk pengujian.
- `README.md`: Dokumentasi proyek.
