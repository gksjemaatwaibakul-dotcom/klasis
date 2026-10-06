# PRD — Perbaikan & Peningkatan Aplikasi Katekisasi SMA Kristen Waibakul

## Problem Statement (ringkas)
Aplikasi Web Google Apps Script (Code.gs + Index.html, database Google Sheets). Perbaikan tanpa rewrite:
1. Bug peringatan absensi tidak muncul untuk kelas 11 (akar: hari aktif dari absensi yang sudah ada)
2. Menu terpisah kelola peringatan (dashboard hanya ringkasan); dismiss hanya Admin/Moderator per wilayah
3. Peringatan lintas kelas satu kelompok untuk pengampu (read-only)
4. Running text terlalu cepat saat pengumuman banyak (durasi dinamis, px/detik konstan)
5. Statistik: per kelompok, filter minggu(default)/bulan/triwulan/semester, toggle pie(default)/bar
6. Tombol Pengunjung/Pengamat tanpa login (dashboard publik read-only)
7. Halaman Bahan Ajar: upload PDF ke Google Drive (Admin/Moderator), download semua, baca via Drive viewer embed
Tambahan user: ganti istilah Pengampuh→Pengampu, Alfa→Alpa (dengan kompatibilitas data lama).
Konfigurasi default: Ganjil Jul–Des, Genap Jan–Jun, triwulan per 3 bulan (diubah via KONFIG di Code.gs).

## Persona
- Admin: full akses, kelola peringatan semua wilayah, upload bahan ajar
- Moderator: kelola peringatan & bahan ajar per kelompok wilayahnya
- Pengampu: isi absen/jurnal, lihat peringatan satu kelompok (read-only), unduh bahan ajar
- Pengunjung/Pengamat: tanpa akun, dashboard publik read-only

## File Kerja
- /app/gas/Code.gs — backend GAS (final)
- /app/gas/Index.html — frontend GAS (final)
- /app/gas/test/simulate.js — harness mock SpreadsheetApp/DriveApp untuk uji logika (node test/simulate.js)
- /app/gas/test/preview.html — preview UI dengan mock google.script.run (serve: python3 -m http.server)
- /app/gas/PANDUAN.md — panduan menerapkan ke project Apps Script

## Log Import (2026-10-06)
- Repo GitHub https://github.com/gksjemaatwaibakul-dotcom/klasis (branch main, commit fd265a0) diimpor ke /app; git remote origin terhubung dan tracking origin/main
- Dependencies: pip install -r backend/requirements.txt OK; yarn install OK; semua service supervisor RUNNING; backend /api merespons; frontend 200 (splash template bawaan repo); gas/test/simulate.js 41/41 PASS

## Yang Sudah Diimplementasikan (2026-10-06)
- Fase 1: _hariAktifDariJadwal() menggantikan _hariAktifDariAbsensi; kelas tanpa jadwal tetap ikut loop; libur global/per-kelompok dikecualikan; running text durasi dinamis (fbTickerSetSpeed, 90 px/detik)
- Fase 2: halaman fbPeringatan (menu Kelola Peringatan), dashboard hanya ringkasan jumlah per topik + tombol; dismiss tetap server-side Admin/Moderator per wilayah
- Fase 3: _kelasSetPeringatan (pengampu melihat satu kelompok); statistik scope Semua/Kelompok:/kelas, getPeriodeRange(minggu/bulan/triwulan/semester) via KONFIG, toggle pie(default)/bar + %/jumlah
- Fase 4: tombol Pengunjung/Pengamat, getPublikDashboard (kehadiran minggu, ibadah semester dari BukuMingguan, siswa umum+per kelompok, kelas per kelompok, daftar pengampu+moderator+admin tanpa data sensitif), CSS role-pengunjung
- Fase 5: sheet BahanAjar + uploadBahanAjar/getBahanAjarList/hapusBahanAjar, folder Drive otomatis (PropertiesService), viewer iframe /preview, link unduh /uc?export=download
- Terminologi: Pengampuh→Pengampu & Alfa→Alpa global; migrasi Role di _ensureAuthSchema; rename/merge sheet AbsensiPengampuh→AbsensiPengampu; normalisasi /^al/i untuk data lama
- Uji: 41/41 tes simulasi Node lolos; screenshot UI (login, publik, dashboard admin, kelola peringatan, bahan ajar) OK

## Perbaikan REV 9 (2026-10-06)
- Bug kelas 11: `_normHari()` menormalisasi nama hari dari sheet Jadwal ("Jum'at"/"ahad"/huruf kecil); fallback — kelas dengan hari jadwal yang tak pernah cocok tanggal nyata diperlakukan tanpa jadwal (tetap diperingatkan)
- Dashboard: fungsi lama `loadDashPeringatan` (penampil nama/kelas) dinonaktifkan jadi stub; dashboard hanya ringkasan badge; detail + dismiss hanya di menu Kelola Peringatan (Admin/Moderator dismiss, Pengampu read-only — sesuai desain)
- Dashboard publik: `getPublikDashboard` mundur ke minggu terakhir yang punya data (maks 12 minggu) bila minggu ini kosong + flag `mingguIni` & label dinamis di UI
- Bahan ajar: maximize modal via CSS sendiri (`.ba-max`) karena fullscreen Drive viewer diblokir sandbox GAS; iframe `allowfullscreen`; tombol "buka di tab baru"
- Uji: 47/47 tes simulasi Node lolos (6 tes baru REV 9); file unduhan di frontend/public/download diperbarui

## Backlog / Next
- P1: Verifikasi di environment Google Apps Script nyata oleh user (deploy ulang Web App)
- P2: Opsi edit tanggal mulai/akhir semester via UI Pengaturan (saat ini via KONFIG)
- P2: Batas ukuran file PDF bahan ajar bila diperlukan
- P3: Notifikasi WA otomatis terjadwal untuk peringatan
