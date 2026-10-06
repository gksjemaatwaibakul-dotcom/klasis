# Panduan Menerapkan Perubahan (Google Apps Script)

File final: `Code.gs` dan `Index.html` di folder ini. Data & fitur lama tetap utuh.

## Perbaikan REV 13 (terbaru)
1. Libur kelompok kini PASTI juga menekan peringatan absen pengampu, jurnal, dan
   buku mingguan: nama kelas dinormalisasi saat mencocokkan ("11A" = "11 A" =
   "Kelas 11 A"), sehingga beda ejaan tidak lagi membuat pengecualian libur gagal.
2. Catatan perilaku: pengampu baru dikecualikan bila SEMUA kelas yang diampunya
   libur pada tanggal tsb. Pengampu yang mengampu kelas di dua kelompok berbeda
   tetap diperingatkan bila salah satu kelompoknya tidak libur. Pastikan juga kolom
   kelas diampu pada data pengampu (sheet Auth) terisi.

## Perbaikan REV 12
1. Hari Libur kini bisa diisi BANYAK TANGGAL sekaligus: mode Satu hari, Rentang
   tanggal, Satu minggu (Sen–Min), Satu bulan penuh, atau Beberapa tanggal terpilih.
   Opsi "Hanya tandai hari yang ada jadwalnya" menyaring hari tanpa jadwal.
2. Setiap libur disimpan per-tanggal untuk kelompok kelas tertentu — SPESIFIK pada
   tanggal yang dipilih, TIDAK berlaku permanen. Otomatis menekan peringatan absen
   siswa DAN absen pengampu untuk kelompok & tanggal terkait.
3. Inilah cara yang benar untuk "meniadakan peringatan" (menggantikan dismiss
   permanen). Untuk mengembalikan, cukup hapus baris libur pada tanggal tsb.

## Perbaikan REV 11
1. Akar masalah kelas 11 ditemukan lewat Diagnosa: seluruh peringatannya tertutup
   aturan "sembunyikan" (dismiss) yang pernah dibuat. Diagnosa kini juga MENAMPILKAN
   daftar aturan dismiss yang menutup absen siswa (kelompok, tanggal, siapa pembuatnya).
2. Tombol baru "Tampilkan Semua" (Admin) di menu Kelola Peringatan: menghapus semua
   aturan sembunyikan sehingga seluruh peringatan (termasuk kelas 11) tampil kembali.
   Setelah itu, sembunyikan ulang hanya tanggal tertentu yang memang perlu.

## Perbaikan REV 10
1. Bahan ajar: modal Bootstrap DIGANTI TOTAL dengan overlay buatan aplikasi
   (murni CSS+JS). Modal Bootstrap di dalam iframe sandbox Apps Script memang
   sering freeze — halaman tidak bisa ditekan sama sekali. Kini tombol Tutup,
   Zoom +/-, Layar Penuh, dan Buka-di-Tab-Baru dijamin berfungsi.
2. Dashboard kembali menampilkan DAFTAR LENGKAP nama kelas & pengampu yang
   belum mengisi (absen siswa, absen pengampu, jurnal, buku mingguan) — read-only
   tanpa checkbox. Aksi sembunyikan (dismiss) hanya ada di menu Kelola Peringatan
   untuk Admin/Moderator.
3. Tombol "Diagnosa (Admin)" baru di menu Kelola Peringatan: menelusuri per kelas
   apakah punya jadwal, berapa tanggal wajib isi, berapa yang belum diisi, dan
   berapa yang disembunyikan aturan dismiss — untuk melacak kasus kelas yang
   peringatannya tidak muncul (mis. kelas 11). Salin hasilnya ke pengembang.

## Perbaikan REV 9
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
5. Auto Changelog: nomor revisi + tanggal build kini tampil di footer halaman
   login dan footer aplikasi (diambil dari `APP_REV` & `APP_BUILD_DATE` di atas
   Code.gs — naikkan nilainya setiap kali Anda deploy perubahan).

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
