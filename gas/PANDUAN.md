# Panduan Menerapkan Perubahan (Google Apps Script)

File final: `Code.gs` dan `Index.html` di folder ini. Data & fitur lama tetap utuh.

## Perbaikan REV 9 (terbaru)
1. Peringatan kelas 11: nama hari di sheet Jadwal kini dinormalisasi ("Jum'at",
   "ahad", huruf kecil, dsb tetap dikenali). Kelas yang nama hari jadwalnya tidak
   pernah cocok dengan tanggal nyata diperlakukan seperti tanpa jadwal sehingga
   TETAP diperingatkan di semua hari aktif.
2. Dashboard hanya menampilkan RINGKASAN jumlah peringatan (badge + tombol).
   Fungsi lama penampil nama/kelas di dashboard dinonaktifkan total. Daftar nama
   + aksi sembunyikan (dismiss) hanya di menu "Kelola Peringatan": Admin/Moderator
   bisa dismiss, Pengampu read-only (sesuai desain).
3. Grafik kehadiran di Dashboard Publik: bila minggu ini belum ada data, otomatis
   menampilkan minggu terakhir yang punya data (maks 12 minggu ke belakang).
4. Baca bahan ajar: tombol Perbesar kini memakai maximize milik aplikasi (modal
   hampir penuh layar) karena tombol fullscreen bawaan Drive viewer diblokir
   sandbox Apps Script. Ditambah tombol "Buka di tab baru" untuk kontrol penuh.

## Langkah
1. Buka project Apps Script Anda (dari spreadsheet: Extensions → Apps Script).
2. Salin seluruh isi `Code.gs` (timpa isi lama) dan `Index.html` (timpa isi lama).
3. Simpan, lalu jalankan sekali fungsi `setupDatabase` (Run di editor) — ini merapikan header,
   merename sheet `AbsensiPengampuh` → `AbsensiPengampu` (data ikut), dan mengubah Role
   `Pengampuh` → `Pengampu` di sheet Auth secara otomatis.
4. Deploy ulang: Deploy → Manage deployments → Edit → New version → Deploy.
   PENTING untuk akses Pengunjung: akses "Who has access" = **Anyone** (execute as: Me).
5. Bahan Ajar: saat upload pertama, folder "Bahan Ajar Katekisasi" dibuat otomatis di
   Google Drive akun pemilik script dan dibagikan "anyone with link can view" agar
   preview/unduh berfungsi untuk semua (termasuk pengunjung).

## Konfigurasi (opsional)
Di bagian atas `Code.gs` ada blok `KONFIG`:
- `GANJIL_MULAI_BULAN` (default 7 = Juli) & `GENAP_MULAI_BULAN` (default 1 = Januari)
- `TICKER_PX_PER_DETIK` (kecepatan running text, default 90)
- `BAHAN_AJAR_FOLDER` (nama folder Drive)
- `PERINGATAN_MAX_HARI` (berapa lama peringatan menengok ke belakang, default 120)

## Catatan
- Status absensi baru tertulis "Alpa"; data lama "Alfa" tetap terbaca normal di semua
  statistik/rekap/peringatan (dinormalisasi otomatis).
- Peringatan kini berbasis Jadwal Mengajar: kelas diperingatkan pada hari jadwalnya
  walau belum pernah mengisi absen (perbaikan bug kelas 11). Hari libur bertanda
  tetap dikecualikan.
