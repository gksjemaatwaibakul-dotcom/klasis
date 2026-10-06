/**
 * DASHBOARD GURU (REV 4 - CUSTOM NAMA SEKOLAH & PRINT BROWSER)
 * 
 * Perbaikan & fitur baru:
 * - ID auto-generate di semua sheet (hapus reliable, no XSS)
 * - Validasi server-side (trim, format tanggal, cek duplikat, jam mulai<selesai)
 * - Sheet Mapel baru (multi-mata pelajaran)
 * - Sheet Aktivitas (log audit ringan)
 * - Fungsi edit per modul (siswa, jadwal, nilai, jurnal)
 * - Import siswa massal
 * - Dashboard stats endpoint
 * - Export Excel (CSV) endpoint
 * - Format kolom tanggal di-set eksplisit (cegah bug rekap)
 * - Trim username/password saat verifikasi
 * - Nama aplikasi mengikuti Nama Sekolah dari Pengaturan
 * - Rekap cetak memakai print browser, tanpa penyimpanan ke Drive
 * - REV 7: Filter mata pelajaran di menu Rekap (Absensi, Nilai, Jurnal)
 * - REV 7 FAST: Optimasi login, simpan/input, dan hapus data agar lebih cepat
 * - FIX IMPORT: Import template lengkap sekaligus membuat kelas & siswa
 */

// ==========================================
// KONSTANTA & UTIL
// ==========================================
const SHEET_NAMES = {
  KELAS: 'Kelas',
  SISWA: 'Siswa',
  MAPEL: 'Mapel',
  JADWAL: 'Jadwal',
  ABSENSI: 'Absensi',
  NILAI: 'Nilai',
  JURNAL: 'Jurnal',
  SETTING: 'Setting',
  AUTH: 'Auth',
  AKTIVITAS: 'Aktivitas',
  PENGAMPU: 'Pengampu',
  ABSENSI_PENGAMPU: 'AbsensiPengampu'
};

const DEFAULT_APP_NAME = 'Sistem Manajemen Kelas Katekisasi SMA Kristen Waibakul';

// Header multi-role. Auth sheet menyimpan admin & pengampu, dibedakan dgn kolom Role.
// KelasDiampu: comma-separated (kosong = semua kelas / untuk admin).
const AUTH_HEADERS = ['ID', 'Username', 'Password', 'Role', 'Nama', 'NIP', 'NoHP', 'KelasDiampu', 'Active'];
const ABSENSI_PENGAMPU_HEADERS = ['ID', 'Tanggal', 'Username', 'Nama', 'JamMasuk', 'JamPulang', 'Status', 'Keterangan'];
const DEFAULT_ADMIN = { username: 'admin', password: 'admin123', nama: 'Administrator' };

// ==========================================
// KONFIGURASI AKADEMIK & FITUR (ubah di sini bila aturan sekolah berbeda)
// ==========================================
const KONFIG = {
  // Semester Ganjil mulai bulan Juli (7) s.d. Desember; Genap Januari (1) s.d. Juni.
  // Triwulan = blok 3 bulan di dalam semester berjalan (Ganjil: TW1 Jul-Sep, TW2 Okt-Des; Genap: TW3 Jan-Mar, TW4 Apr-Jun).
  GANJIL_MULAI_BULAN: 7,
  GENAP_MULAI_BULAN: 1,
  // Kecepatan running text pengumuman (piksel/detik). Durasi animasi dihitung otomatis
  // dari panjang konten sehingga kecepatan tetap stabil berapa pun jumlah pengumuman.
  TICKER_PX_PER_DETIK: 90,
  // Nama folder Google Drive tempat file Bahan Ajar disimpan (dibuat otomatis sekali).
  BAHAN_AJAR_FOLDER: 'Bahan Ajar Katekisasi',
  // Seberapa jauh (hari ke belakang) peringatan "belum mengisi" ditelusuri.
  PERINGATAN_MAX_HARI: 120
};

// Versi aplikasi — ditampilkan di footer (halaman login & aplikasi) agar mudah
// memastikan versi yang sedang berjalan. Naikkan setiap kali deploy perubahan.
var APP_REV = 'REV 10';
var APP_BUILD_DATE = '2026-10-06';


const SISWA_HEADERS = [
  'ID', 'No Urut', 'Nama Siswa', 'Kelas',
  'NIS', 'NISN', 'Jenis Kelamin', 'Tempat Lahir', 'Tanggal Lahir',
  'Agama', 'Alamat Siswa', 'Status', 'Keterangan Siswa',
  'Nama Ayah', 'Pekerjaan Ayah', 'No HP Ayah',
  'Nama Ibu', 'Pekerjaan Ibu', 'No HP Ibu',
  'Alamat Orang Tua', 'Keterangan Orang Tua'
];

const SISWA_EXTRA_KEYS = [
  'nis', 'nisn', 'jenisKelamin', 'tempatLahir', 'tanggalLahir',
  'agama', 'alamatSiswa', 'status', 'keteranganSiswa',
  'namaAyah', 'pekerjaanAyah', 'noHpAyah',
  'namaIbu', 'pekerjaanIbu', 'noHpIbu',
  'alamatOrangTua', 'keteranganOrangTua'
];


function _blankSetting() {
  return ['', '', '', '', '', ''];
}

function _getSettingValues() {
  const sheet = _ss().getSheetByName(SHEET_NAMES.SETTING);
  if (!sheet || sheet.getLastRow() < 2) return _blankSetting();
  const row = sheet.getRange(2, 1, 1, 7).getDisplayValues()[0] || [];
  while (row.length < 7) row.push('');
  return row.map(_t);
}

function _getAppName() {
  const setting = _getSettingValues();
  return setting[0] || DEFAULT_APP_NAME;
}

function getPublicAppInfo() {
  const s = _getSettingValues();
  return {
    appName: s[0] || DEFAULT_APP_NAME,
    namaSekolah: s[0] || '',
    alamat: s[1] || '',
    namaGuru: s[2] || '',
    nipGuru: s[3] || '',
    namaKepsek: s[4] || '',
    nipKepsek: s[5] || '',
    logoUrl: s[6] || '',
    rev: APP_REV,
    buildDate: APP_BUILD_DATE,
    versi: APP_REV + ' \u2022 Build ' + APP_BUILD_DATE
  };
}

function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function _sheet(name) {
  const s = _ss().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" tidak ditemukan. Jalankan Setup Database.');
  return s;
}


// Pastikan sheet Absensi punya kolom Mata Pelajaran tanpa merusak data lama.
// Struktur baru: [ID, Timestamp, Tanggal, Kelas, Mata Pelajaran, No Urut, Nama Siswa, Status]
function _ensureAbsensiMapelColumn() {
  const ss = _ss();
  let sheet = ss.getSheetByName(SHEET_NAMES.ABSENSI);
  if (!sheet) return;
  const expected = ['ID', 'Timestamp', 'Tanggal', 'Kelas', 'Mata Pelajaran', 'No Urut', 'Nama Siswa', 'Status'];
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  if (headers.indexOf('Mata Pelajaran') === -1) {
    sheet.insertColumnAfter(4); // sisipkan setelah Kelas; data lama otomatis bergeser aman
  }
  sheet.getRange(1, 1, 1, expected.length).setValues([expected]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 2).setNumberFormat('@'); // Timestamp + Tanggal
  }
}

function _absensiCols() {
  _ensureAbsensiMapelColumn();
  return { id: 0, ts: 1, tanggal: 2, kelas: 3, mapel: 4, no: 5, nama: 6, status: 7 };
}


function _ensureSiswaLengkapColumns() {
  const ss = _ss();
  const sheet = ss.getSheetByName(SHEET_NAMES.SISWA);
  if (!sheet) return;
  if (sheet.getMaxColumns() < SISWA_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), SISWA_HEADERS.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, SISWA_HEADERS.length).setValues([SISWA_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (sheet.getMaxRows() > 1) {
    const textCols = [2, 3, 4, 5, 6, 9, 16, 19]; // No, Nama, Kelas, NIS, NISN, Tgl Lahir, HP Ayah, HP Ibu
    textCols.forEach(function(col) {
      try { sheet.getRange(2, col, sheet.getMaxRows() - 1).setNumberFormat('@'); } catch (e) {}
    });
  }
}


// Generate ID unik: timestamp-based + random suffix (tahan duplikasi)
function _newId(prefix) {
  return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).substring(2, 6);
}

// Trim string aman (tanpa error kalau null/undefined/number)
function _t(v) { return (v == null) ? '' : String(v).trim(); }

// Timestamp sortable: format "YYYY-MM-DD HH:mm:ss" (timezone script)
// Penting untuk sort string-based di rekap "input terakhir"
function _nowStamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// Display stamp: format friendly Indonesia "DD MMM YYYY HH:mm" — untuk tampilan saja
function _displayStamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
}

// Lock helper: mencegah race condition saat 2 user/tab submit bersamaan.
// Penting untuk operasi yang baca-modify-tulis (replace mode).
function _withLock(fn) {
  const lock = LockService.getDocumentLock();
  // Tunggu maks 15 detik untuk dapat lock
  try {
    lock.waitLock(15000);
  } catch (e) {
    throw new Error('Sistem sedang sibuk memproses request lain. Coba lagi dalam beberapa detik.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// Validasi tanggal format YYYY-MM-DD (ketat: cek tanggal yang benar-benar valid)
function _isValidDate(s) {
  if (!s || typeof s !== 'string') return false;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return false;
  const yyyy = parseInt(m[1]), mm = parseInt(m[2]), dd = parseInt(m[3]);
  if (yyyy < 1900 || yyyy > 2100) return false;
  if (mm < 1 || mm > 12) return false;
  if (dd < 1 || dd > 31) return false;
  // Cek roundtrip: kalau Date object beda dengan input, berarti overflow (mis. 30 Feb → 2 Mar)
  const d = new Date(yyyy, mm - 1, dd);
  return d.getFullYear() === yyyy && d.getMonth() === mm - 1 && d.getDate() === dd;
}

// Cari baris berdasarkan ID di kolom A. Return: { rowIndex (1-based), values[] } atau null.
function _findRowById(sheetName, id) {
  const sheet = _sheet(sheetName);
  if (sheet.getLastRow() < 2) return null;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      const rowIdx = i + 2;
      const values = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
      return { rowIndex: rowIdx, values: values };
    }
  }
  return null;
}

// Log aktivitas (best-effort, non-blocking jika sheet belum ada). actor = user pelaku (opsional).
function _logActivity(action, detail, actor) {
  try {
    const sheet = _ss().getSheetByName(SHEET_NAMES.AKTIVITAS || 'Aktivitas');
    if (!sheet) return;
    sheet.appendRow([_displayStamp(), action, detail || '', _t(actor) || 'Sistem']);
  } catch (e) { /* ignore */ }
}

// Ambil log aktivitas terbaru (untuk admin). limit default 200. Terbaru di atas.
function getAktivitasLog(limit) {
  const sheet = _ss().getSheetByName(SHEET_NAMES.AKTIVITAS || 'Aktivitas');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lim = parseInt(limit) > 0 ? parseInt(limit) : 200;
  const last = sheet.getLastRow();
  const start = Math.max(2, last - lim + 1);
  const rows = sheet.getRange(start, 1, last - start + 1, 4).getDisplayValues();
  return rows.map(function(r){
    return { waktu:_t(r[0]), aksi:_t(r[1]), detail:_t(r[2]), oleh:_t(r[3]) || 'Sistem' };
  }).reverse();
}

function hapusAktivitasLog(role) {
  if (_t(role).toLowerCase() !== 'admin') throw new Error('Hanya Admin yang boleh mengosongkan log.');
  return _withLock(function(){
    const sheet = _ss().getSheetByName(SHEET_NAMES.AKTIVITAS || 'Aktivitas');
    if (sheet && sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(),4)).clearContent();
    return 'Log aktivitas dikosongkan.';
  });
}

// ==========================================
// 1. ENTRY POINT
// ==========================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(_getAppName() + ' - Dashboard Guru')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 2. LOGIN & MULTI-ROLE AUTH
// ==========================================
// Cari user berdasarkan username. Return row values + rowIndex, atau null.
function _findUserByUsername(username) {
  const u = _t(username).toLowerCase();
  if (!u) return null;
  const sheet = _ss().getSheetByName(SHEET_NAMES.AUTH);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, AUTH_HEADERS.length).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (_t(rows[i][1]).toLowerCase() === u) {
      return { rowIndex: i + 2, values: rows[i] };
    }
  }
  return null;
}

// Public info dari sebuah user row (tanpa password) untuk client.
function _userToPublic(vals) {
  return {
    id: _t(vals[0]),
    username: _t(vals[1]),
    role: _t(vals[3]) || 'Pengampu',
    nama: _t(vals[4]),
    nip: _t(vals[5]),
    noHp: _t(vals[6]),
    kelasDiampu: _t(vals[7]).split(',').map(_t).filter(Boolean),
    active: (_t(vals[8]).toLowerCase() !== 'tidak')
  };
}

// Legacy compat: kembalikan admin pertama, atau default.
function getCredentials() {
  const rows = _readSheet(SHEET_NAMES.AUTH);
  const admin = rows.find(r => _t(r[3]).toLowerCase() === 'admin');
  if (admin) return { username: _t(admin[1]), password: _t(admin[2]) };
  return { username: DEFAULT_ADMIN.username, password: DEFAULT_ADMIN.password };
}

function cekLogin(username, password) {
  const u = _t(username);
  const p = _t(password);
  if (!u || !p) return { success: false, message: 'Username dan password wajib diisi!' };

  // Auto-migrate & seed admin jika belum ada
  _ensureAuthSchema();
  _seedDefaultAdmin();

  const found = _findUserByUsername(u);
  if (!found) return { success: false, message: 'Username atau Password salah!' };
  const vals = found.values;
  if (_t(vals[2]) !== p) return { success: false, message: 'Username atau Password salah!' };
  if (_t(vals[8]).toLowerCase() === 'tidak') return { success: false, message: 'Akun Anda dinonaktifkan. Hubungi Admin.' };

  _logActivity('Login', _t(vals[1]) + ' (' + _t(vals[3]) + ')');
  return { success: true, message: 'Login berhasil!', user: _userToPublic(vals) };
}

// Login cepat: verifikasi + kirim data awal + dashboard dalam 1 panggilan server.
function cekLoginCepat(username, password) {
  const res = cekLogin(username, password);
  if (!res.success) return res;
  return {
    success: true,
    message: res.message,
    user: res.user,
    initialData: getAllInitialData(),
    dashboardStats: getDashboardStats()
  };
}

// Ganti username & password untuk user saat ini (identified by usernameLama+passwordLama).
function gantiCredentials(usernameLama, passwordLama, usernameBaru, passwordBaru) {
  const uLama = _t(usernameLama);
  const pLama = _t(passwordLama);
  const uBaru = _t(usernameBaru);
  const pBaru = _t(passwordBaru);
  if (!uBaru || uBaru.length < 3) throw new Error('Username baru minimal 3 karakter!');
  if (!pBaru || pBaru.length < 6) throw new Error('Password baru minimal 6 karakter!');

  _ensureAuthSchema();
  const found = _findUserByUsername(uLama);
  if (!found || _t(found.values[2]) !== pLama) {
    throw new Error('Username atau Password lama salah!');
  }
  // Cek username baru tidak dipakai user lain
  if (uBaru.toLowerCase() !== uLama.toLowerCase()) {
    const dup = _findUserByUsername(uBaru);
    if (dup) throw new Error('Username "' + uBaru + '" sudah dipakai user lain!');
  }
  const sheet = _sheet(SHEET_NAMES.AUTH);
  sheet.getRange(found.rowIndex, 2, 1, 2).setNumberFormat('@');
  sheet.getRange(found.rowIndex, 2, 1, 2).setValues([[uBaru, pBaru]]);
  _logActivity('Ganti Kredensial', uLama + ' → ' + uBaru);
  return 'Username dan Password berhasil diperbarui!';
}


// ==========================================
// 3. SETUP DATABASE (struktur baru pakai ID)
// ==========================================
function setupDatabase() {
  const ss = _ss();
  const config = {
    'Kelas':    ['ID', 'Nama Kelas'],
    'Siswa':    SISWA_HEADERS,
    'Mapel':    ['ID', 'Nama Mata Pelajaran'],
    'Jadwal':   ['ID', 'Hari', 'Jam Mulai', 'Jam Selesai', 'Mata Pelajaran', 'Kelas'],
    'Absensi':  ['ID', 'Timestamp', 'Tanggal', 'Kelas', 'Mata Pelajaran', 'No Urut', 'Nama Siswa', 'Status'],
    'Nilai':    ['ID', 'Timestamp', 'Tanggal', 'Kelas', 'Mata Pelajaran', 'Kategori', 'No Urut', 'Nama Siswa', 'Nilai', 'Bab', 'Tujuan Pembelajaran', 'Bentuk'],
    'Jurnal':   ['ID', 'Timestamp', 'Tanggal', 'Jam Ke', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', 'Kegiatan Pembelajaran', 'Keterangan'],
    'Setting':  ['Nama Aplikasi / Sekolah', 'Alamat Lengkap', 'Nama Guru', 'NIP Guru', 'Nama Kepala Sekolah', 'NIP Kepsek', 'Logo URL'],
    'Auth':     AUTH_HEADERS,
    'Aktivitas':['Waktu', 'Aksi', 'Detail', 'Oleh'],
    'AbsensiPengampu': ABSENSI_PENGAMPU_HEADERS,
    'HariLibur': ['ID', 'Tanggal', 'Keterangan']
  };
  
  for (const sheetName in config) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    if (sheetName === 'Absensi') {
      _ensureAbsensiMapelColumn();
      sheet = ss.getSheetByName(sheetName);
    }
    if (sheetName === 'Siswa') {
      _ensureSiswaLengkapColumns();
      sheet = ss.getSheetByName(sheetName);
    }
    const headers = config[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Paksa kolom Tanggal & Timestamp jadi text supaya format konsisten dan sort string-based reliable
    if (['Absensi', 'Nilai', 'Jurnal'].indexOf(sheetName) >= 0) {
      const tglColIdx = headers.indexOf('Tanggal') + 1;
      const tsColIdx = headers.indexOf('Timestamp') + 1;
      if (tglColIdx > 0) {
        sheet.getRange(2, tglColIdx, sheet.getMaxRows() - 1).setNumberFormat('@');
      }
      if (tsColIdx > 0) {
        sheet.getRange(2, tsColIdx, sheet.getMaxRows() - 1).setNumberFormat('@');
      }
    }
    // Sheet Aktivitas juga: kolom Waktu sebagai text supaya display stamp tetap konsisten
    if (sheetName === 'Aktivitas') {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1).setNumberFormat('@');
    }
  }
  
  _rapikanNomorSiswaPerKelas();
  _ensureAuthSchema();
  _seedDefaultAdmin();
  return 'Database berhasil disiapkan! Struktur multi-role & sheet Pengampu sudah tersedia.';
}

// ==========================================
// 3B. MIGRASI AUTH LAMA & SEED ADMIN DEFAULT
// ==========================================
function _ensureAuthSchema() {
  const ss = _ss();
  let sheet = ss.getSheetByName(SHEET_NAMES.AUTH);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.AUTH);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, AUTH_HEADERS.length)).getDisplayValues()[0];

  // Deteksi skema lama: cuma [Username, Password]
  if (headers[0] === 'Username' || (headers[0] !== 'ID' && sheet.getLastRow() >= 2)) {
    // Ambil kredensial lama
    let oldUser = '', oldPass = '';
    if (sheet.getLastRow() >= 2) {
      const d = sheet.getRange(2, 1, 1, 2).getDisplayValues()[0];
      oldUser = _t(d[0]); oldPass = _t(d[1]);
    }
    // Reset sheet dan tulis skema baru
    sheet.clear();
    sheet.getRange(1, 1, 1, AUTH_HEADERS.length).setValues([AUTH_HEADERS]).setFontWeight('bold');
    // Migrasi user lama -> Admin (kalau ada)
    if (oldUser && oldPass) {
      sheet.appendRow([_newId('U'), oldUser, oldPass, 'Admin', 'Administrator', '', '', '', 'Ya']);
    }
  } else {
    sheet.getRange(1, 1, 1, AUTH_HEADERS.length).setValues([AUTH_HEADERS]).setFontWeight('bold');
  }
  // Migrasi istilah peran lama -> 'Pengampu' (prefix 'pengampu' mencakup ejaan lama 'pengampu*h').
  if (sheet.getLastRow() >= 2) {
    const roleVals = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
    let roleDirty = false;
    for (let ri = 0; ri < roleVals.length; ri++) {
      const rv = String(roleVals[ri][0] || '').trim();
      if (rv && rv.toLowerCase().indexOf('pengampu') === 0 && rv !== 'Pengampu') { roleVals[ri][0] = 'Pengampu'; roleDirty = true; }
    }
    if (roleDirty) sheet.getRange(2, 4, roleVals.length, 1).setValues(roleVals);
  }
  sheet.setFrozenRows(1);
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, AUTH_HEADERS.length).setNumberFormat('@');
}

function _seedDefaultAdmin() {
  const sheet = _sheet(SHEET_NAMES.AUTH);
  const rows = _readSheet(SHEET_NAMES.AUTH);
  const hasAdmin = rows.some(r => _t(r[3]).toLowerCase() === 'admin');
  if (!hasAdmin) {
    sheet.appendRow([_newId('U'), DEFAULT_ADMIN.username, DEFAULT_ADMIN.password, 'Admin', DEFAULT_ADMIN.nama, '', '', '', 'Ya']);
  }
}

// ==========================================
// 4. LOAD DATA
// ==========================================
function _readSheet(name) {
  const sheet = _ss().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
}

function getAllInitialData() {
  return {
    kelas: _readSheet(SHEET_NAMES.KELAS),
    mapel: _readSheet(SHEET_NAMES.MAPEL),
    setting: _readSheet(SHEET_NAMES.SETTING),
    appInfo: getPublicAppInfo(),
    pengampu: getPengampuData()
  };
}

function getSiswaData() {
  // Nomor siswa sudah dirapikan saat import/edit/hapus/setup.
  // Tidak dirapikan ulang di setiap load supaya absensi/nilai/jurnal terasa lebih cepat.
  return _sortSiswaRows(_readSheet(SHEET_NAMES.SISWA));
}
function getJadwalData() { return _readSheet(SHEET_NAMES.JADWAL); }
function getMapelData() { return _readSheet(SHEET_NAMES.MAPEL); }

function getSiswaByKelas(namaKelas) {
  const all = _readSheet(SHEET_NAMES.SISWA);
  // Struktur baru: [ID, No, Nama, Kelas]
  return _sortSiswaRows(all.filter(row => row[3] === namaKelas));
}

function _sortSiswaRows(rows) {
  const kelasOrder = {};
  _readSheet(SHEET_NAMES.KELAS).forEach(function(r, i) {
    kelasOrder[r[1]] = i;
  });
  return (rows || []).slice().sort(function(a, b) {
    const ka = kelasOrder[a[3]] !== undefined ? kelasOrder[a[3]] : 999999;
    const kb = kelasOrder[b[3]] !== undefined ? kelasOrder[b[3]] : 999999;
    if (ka !== kb) return ka - kb;
    const na = parseInt(a[1], 10);
    const nb = parseInt(b[1], 10);
    const va = isNaN(na) ? 999999 : na;
    const vb = isNaN(nb) ? 999999 : nb;
    if (va !== vb) return va - vb;
    return String(a[2] || '').localeCompare(String(b[2] || ''));
  });
}

function _rapikanNomorSiswaPerKelas(targetKelas) {
  const sheet = _ss().getSheetByName(SHEET_NAMES.SISWA);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4);
  const data = range.getDisplayValues();
  const groups = {};

  const target = _t(targetKelas);
  data.forEach(function(row, idx) {
    const kelas = _t(row[3]);
    if (!kelas) return;
    if (target && kelas !== target) return;
    if (!groups[kelas]) groups[kelas] = [];
    groups[kelas].push({
      idx: idx,
      oldNo: _t(row[1]),
      nama: _t(row[2]),
      kelas: kelas
    });
  });

  const changes = [];
  Object.keys(groups).forEach(function(kelas) {
    groups[kelas].sort(function(a, b) {
      const na = parseInt(a.oldNo, 10);
      const nb = parseInt(b.oldNo, 10);
      const va = isNaN(na) ? 999999 : na;
      const vb = isNaN(nb) ? 999999 : nb;
      if (va !== vb) return va - vb;
      return a.idx - b.idx;
    });

    groups[kelas].forEach(function(item, pos) {
      const newNo = String(pos + 1);
      if (_t(data[item.idx][1]) !== newNo) {
        changes.push({
          kelas: kelas,
          nama: item.nama,
          oldNo: item.oldNo,
          newNo: newNo
        });
        data[item.idx][1] = newNo;
      }
    });
  });

  if (changes.length > 0) {
    range.setValues(data);
    _updateNomorSiswaInHistory(changes);
    _logActivity('Rapikan Nomor Siswa', changes.length + ' nomor disesuaikan per kelas');
  }
  return changes.length;
}

function _updateNomorSiswaInHistory(changes) {
  if (!Array.isArray(changes) || changes.length === 0) return;
  const map = {};
  changes.forEach(function(c) {
    map[c.kelas + '||' + c.nama + '||' + c.oldNo] = c.newNo;
  });

  // ABSENSI: [ID, TS, Tgl, Kelas, Mapel, No, Nama, Status]
  const absSheet = _ss().getSheetByName(SHEET_NAMES.ABSENSI);
  if (absSheet && absSheet.getLastRow() > 1) {
    const c = _absensiCols();
    const range = absSheet.getRange(2, 1, absSheet.getLastRow() - 1, Math.max(absSheet.getLastColumn(), c.status + 1));
    const rows = range.getDisplayValues();
    let changed = false;
    rows.forEach(function(row) {
      const key = row[c.kelas] + '||' + row[c.nama] + '||' + row[c.no];
      if (map[key]) {
        row[c.no] = map[key];
        changed = true;
      }
    });
    if (changed) range.setValues(rows);
  }

  // NILAI: [ID, TS, Tgl, Kelas, Mapel, Kategori, No, Nama, Nilai, Bab, Tujuan, Bentuk]
  const nilSheet = _ss().getSheetByName(SHEET_NAMES.NILAI);
  if (nilSheet && nilSheet.getLastRow() > 1) {
    const range = nilSheet.getRange(2, 1, nilSheet.getLastRow() - 1, nilSheet.getLastColumn());
    const rows = range.getDisplayValues();
    let changed = false;
    rows.forEach(function(row) {
      const key = row[3] + '||' + row[7] + '||' + row[6];
      if (map[key]) {
        row[6] = map[key];
        changed = true;
      }
    });
    if (changed) range.setValues(rows);
  }
}

// ==========================================
// 5. DASHBOARD STATS
// ==========================================
function getDashboardStats() {
  const ss = _ss();
  const stats = { 
    totalKelas: 0, totalSiswa: 0, totalJadwal: 0,
    jadwalHariIni: []
  };
  
  const kelas = _readSheet(SHEET_NAMES.KELAS);
  const siswa = _readSheet(SHEET_NAMES.SISWA);
  const jadwal = _readSheet(SHEET_NAMES.JADWAL);
  
  stats.totalKelas = kelas.length;
  stats.totalSiswa = siswa.length;
  stats.totalJadwal = jadwal.length;
  
  // Jadwal hari ini — pakai timezone script (set ke Asia/Jakarta di appsscript.json)
  const tz = Session.getScriptTimeZone();
  const hariIniStr = Utilities.formatDate(new Date(), tz, 'EEEE'); // 'Monday', 'Tuesday', dst
  const hariMap = {
    'Sunday': 'Minggu', 'Monday': 'Senin', 'Tuesday': 'Selasa',
    'Wednesday': 'Rabu', 'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu'
  };
  const hariIni = hariMap[hariIniStr] || hariIniStr;
  // Struktur: [ID, Hari, Jam Mulai, Jam Selesai, Mapel, Kelas]
  stats.jadwalHariIni = jadwal.filter(r => r[1] === hariIni)
    .map(r => ({ jamMulai: r[2], jamSelesai: r[3], mapel: r[4], kelas: r[5] }))
    .sort((a, b) => String(a.jamMulai || '').padStart(5, '0').localeCompare(String(b.jamMulai || '').padStart(5, '0')));
  
  return stats;
}

// Rentang minggu berjalan: Senin s.d. Minggu (timezone script)
function _currentWeekRange() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const dow = parseInt(Utilities.formatDate(now, tz, 'u'), 10); // 1=Senin .. 7=Minggu
  const monday = new Date(now.getTime() - (dow - 1) * 86400000);
  const sunday = new Date(monday.getTime() + 6 * 86400000);
  return {
    start: Utilities.formatDate(monday, tz, 'yyyy-MM-dd'),
    end: Utilities.formatDate(sunday, tz, 'yyyy-MM-dd')
  };
}

// Statistik kehadiran siswa (H/S/I/A/T) untuk dashboard chart.
// Default periode = minggu berjalan. kelas='Semua' -> agregat; per kelas -> difilter.
// Role pengampu otomatis difilter ke kelas yang diampu via _kelasSetFromCtx.
function getKehadiranStats(tglMulai, tglAkhir, kelas, userCtx) {
  const range = _currentWeekRange();
  const tm = _isValidDate(_t(tglMulai)) ? _t(tglMulai) : range.start;
  const ta = _isValidDate(_t(tglAkhir)) ? _t(tglAkhir) : range.end;
  const kls = _t(kelas);
  const set = _kelasSetFromCtx(userCtx);
  const c = _absensiCols();

  // Filter per kelompok kelas: kirim "Kelompok: <id>" sebagai parameter kelas.
  let kelompokSet = null;
  let kelompokNama = '';
  if (kls.indexOf('Kelompok:') === 0) {
    const g = _kelompokById(_t(kls.substring(9)));
    kelompokSet = {};
    (g ? g.kelas : []).forEach(function(k){ kelompokSet[k] = true; });
    kelompokNama = g ? g.nama : kls;
  }

  // Normalisasi status: terima kata penuh & kode singkat.
  // Data lama dieja "Alfa" tetap dihitung via normalisasi /^al/i di bawah.
  const NORM = {
    'Hadir':'H','H':'H','Sakit':'S','S':'S','Izin':'I','I':'I',
    'Alpa':'A','A':'A','Terlambat':'T','T':'T'
  };
  const counts = { H:0, S:0, I:0, A:0, T:0 };

  _readSheet(SHEET_NAMES.ABSENSI).forEach(function(r) {
    const tgl = _t(r[c.tanggal]);
    if (tgl < tm || tgl > ta) return;
    const rowKelas = _t(r[c.kelas]);
    if (set && !set.has(rowKelas)) return;                 // batasan role pengampu
    if (kelompokSet) { if (!kelompokSet[rowKelas]) return; } // filter kelompok kelas
    else if (kls && kls !== 'Semua' && rowKelas !== kls) return; // filter kelas spesifik
    let stRaw = _t(r[c.status]);
    if (/^al/i.test(stRaw)) stRaw = 'Alpa'; // kompatibel ejaan lama pada data tersimpan
    const code = NORM[stRaw];
    if (code) counts[code]++;
  });

  const total = counts.H + counts.S + counts.I + counts.A + counts.T;
  const pct = {};
  ['H','S','I','A','T'].forEach(function(k){
    pct[k] = total > 0 ? Math.round((counts[k] / total) * 1000) / 10 : 0;
  });

  return {
    tglMulai: tm,
    tglAkhir: ta,
    kelas: kelompokNama || kls || 'Semua',
    labels: ['Hadir','Sakit','Izin','Alpa','Terlambat'],
    keys: ['H','S','I','A','T'],
    counts: counts,
    pct: pct,
    total: total,
    totalPct: total > 0 ? 100 : 0
  };
}

// ==========================================
// 6. SIMPAN DATA - KELAS
// ==========================================
function simpanKelas(namaKelas) {
  return _withLock(function() {
  const nama = _t(namaKelas);
  if (!nama) throw new Error('Nama kelas tidak boleh kosong!');
  
  // Cek duplikat
  const existing = _readSheet(SHEET_NAMES.KELAS);
  if (existing.some(r => r[1].toLowerCase() === nama.toLowerCase())) {
    throw new Error('Nama kelas "' + nama + '" sudah ada!');
  }
  
  const id = _newId('K');
  _sheet(SHEET_NAMES.KELAS).appendRow([id, nama]);
  _logActivity('Tambah Kelas', nama);
  return 'Kelas "' + nama + '" berhasil ditambahkan!';
  });
}


function simpanKelasMassal(dataArray) {
  if (!Array.isArray(dataArray) || dataArray.length === 0) {
    throw new Error('Tidak ada data kelas untuk disimpan.');
  }
  return _withLock(function() {
    const existing = _readSheet(SHEET_NAMES.KELAS);
    const existingNames = new Set(existing.map(r => _t(r[1]).toLowerCase()).filter(Boolean));
    const seen = new Set();
    const errors = [];
    const toInsert = [];

    dataArray.forEach((item, idx) => {
      const nama = _t(typeof item === 'string' ? item : item.nama);
      if (!nama) return;
      const key = nama.toLowerCase();
      if (existingNames.has(key)) {
        errors.push('Baris ' + (idx + 1) + ': kelas "' + nama + '" sudah ada');
        return;
      }
      if (seen.has(key)) {
        errors.push('Baris ' + (idx + 1) + ': kelas "' + nama + '" duplikat di input');
        return;
      }
      seen.add(key);
      toInsert.push([_newId('K'), nama]);
    });

    if (toInsert.length === 0) {
      throw new Error(errors.length ? errors.join('\n') : 'Isi minimal 1 nama kelas.');
    }

    const sheet = _sheet(SHEET_NAMES.KELAS);
    sheet.getRange(sheet.getLastRow() + 1, 1, toInsert.length, 2).setValues(toInsert);
    _logActivity('Tambah Kelas Massal', toInsert.length + ' kelas');

    return 'Berhasil menyimpan ' + toInsert.length + ' kelas' +
      (errors.length ? ' (' + errors.length + ' baris dilewati).' : '.');
  });
}

function editKelas(id, namaBaru) {
  return _withLock(function() {
  const nama = _t(namaBaru);
  if (!nama) throw new Error('Nama kelas tidak boleh kosong!');
  
  const found = _findRowById(SHEET_NAMES.KELAS, id);
  if (!found) throw new Error('Kelas tidak ditemukan!');
  
  // Cek duplikat (selain dirinya sendiri)
  const existing = _readSheet(SHEET_NAMES.KELAS);
  if (existing.some(r => r[0] !== id && r[1].toLowerCase() === nama.toLowerCase())) {
    throw new Error('Nama kelas "' + nama + '" sudah dipakai!');
  }
  
  const namaLama = found.values[1];
  _sheet(SHEET_NAMES.KELAS).getRange(found.rowIndex, 2).setValue(nama);
  
  // Update juga semua sheet yang merefer nama kelas (Siswa, Jadwal, Absensi, Nilai, Jurnal)
  _renameKelasInRefs(namaLama, nama);
  
  _logActivity('Edit Kelas', namaLama + ' → ' + nama);
  return 'Kelas berhasil diperbarui!';
  });
}

function _renameKelasInRefs(namaLama, namaBaru) {
  // Siswa: kolom 4 (Kelas)
  _bulkUpdate(SHEET_NAMES.SISWA, 4, namaLama, namaBaru);
  // Jadwal: kolom 6 (Kelas)
  _bulkUpdate(SHEET_NAMES.JADWAL, 6, namaLama, namaBaru);
  // Absensi: kolom 4
  _bulkUpdate(SHEET_NAMES.ABSENSI, 4, namaLama, namaBaru);
  // Nilai: kolom 4
  _bulkUpdate(SHEET_NAMES.NILAI, 4, namaLama, namaBaru);
  // Jurnal: kolom 5
  _bulkUpdate(SHEET_NAMES.JURNAL, 5, namaLama, namaBaru);
}

function _bulkUpdate(sheetName, colIndex, oldVal, newVal) {
  const sheet = _ss().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const range = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1);
  const values = range.getValues();
  let changed = false;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === oldVal) { values[i][0] = newVal; changed = true; }
  }
  if (changed) range.setValues(values);
}

// ==========================================
// 7. SIMPAN DATA - SISWA
// ==========================================
function simpanSiswa(noUrut, nama, kelas, actor) {
  return _withLock(function() {
  const no = _t(noUrut);
  const nm = _t(nama);
  const kls = _t(kelas);
  if (!no || !nm || !kls) throw new Error('Semua kolom siswa wajib diisi!');
  
  // Cek duplikat No Urut di kelas yang sama (PERBAIKAN BUG #6)
  const existing = _readSheet(SHEET_NAMES.SISWA);
  if (existing.some(r => r[1] === no && r[3] === kls)) {
    throw new Error('No urut "' + no + '" sudah ada di kelas ' + kls + '!');
  }
  
  const id = _newId('S');
  _sheet(SHEET_NAMES.SISWA).appendRow([id, no, nm, kls]);
  _rapikanNomorSiswaPerKelas();
  _logActivity('Tambah Siswa', nm + ' (' + kls + ')', actor);
  return 'Siswa "' + nm + '" berhasil ditambahkan! Nomor urut sudah dirapikan per kelas.';
  });
}

function editSiswa(id, noUrut, nama, kelas, actor) {
  return _withLock(function() {
  const no = _t(noUrut);
  const nm = _t(nama);
  const kls = _t(kelas);
  if (!no || !nm || !kls) throw new Error('Semua kolom siswa wajib diisi!');
  
  const found = _findRowById(SHEET_NAMES.SISWA, id);
  if (!found) throw new Error('Siswa tidak ditemukan!');
  
  // Cek duplikat (selain dirinya)
  const existing = _readSheet(SHEET_NAMES.SISWA);
  if (existing.some(r => r[0] !== id && r[1] === no && r[3] === kls)) {
    throw new Error('No urut "' + no + '" sudah ada di kelas ' + kls + '!');
  }
  
  // Data lama untuk update history
  const oldNo = String(found.values[1]);
  const oldNama = String(found.values[2]);
  const oldKelas = String(found.values[3]);
  
  _sheet(SHEET_NAMES.SISWA).getRange(found.rowIndex, 2, 1, 3).setValues([[no, nm, kls]]);
  
  // Update history Absensi & Nilai kalau ada perubahan
  // Absensi: [ID, TS, Tgl, Kelas (3), No (4), Nama (5), Status]
  // Nilai:   [ID, TS, Tgl, Kelas (3), Mapel, Kategori, No (6), Nama (7), Nilai, ...]
  if (oldNo !== no || oldNama !== nm || oldKelas !== kls) {
    _updateSiswaInHistory(oldNo, oldNama, oldKelas, no, nm, kls);
  }
  
  _rapikanNomorSiswaPerKelas();
  _logActivity('Edit Siswa', oldNama + ' → ' + nm, actor);
  return 'Data siswa berhasil diperbarui! Nomor urut sudah dirapikan per kelas.';
  });
}

// Update record history Absensi & Nilai untuk reflect perubahan siswa
function _updateSiswaInHistory(oldNo, oldNama, oldKelas, newNo, newNama, newKelas) {
  // ABSENSI
  const absSheet = _ss().getSheetByName(SHEET_NAMES.ABSENSI);
  if (absSheet && absSheet.getLastRow() > 1) {
    const c = _absensiCols();
    const range = absSheet.getRange(2, 1, absSheet.getLastRow() - 1, Math.max(absSheet.getLastColumn(), c.status + 1));
    const data = range.getDisplayValues();
    let changed = false;
    data.forEach(row => {
      if (row[c.kelas] === oldKelas && row[c.no] === oldNo && row[c.nama] === oldNama) {
        row[c.kelas] = newKelas; row[c.no] = newNo; row[c.nama] = newNama;
        changed = true;
      }
    });
    if (changed) range.setValues(data);
  }
  
  // NILAI
  const nilSheet = _ss().getSheetByName(SHEET_NAMES.NILAI);
  if (nilSheet && nilSheet.getLastRow() > 1) {
    const range = nilSheet.getRange(2, 1, nilSheet.getLastRow() - 1, nilSheet.getLastColumn());
    const data = range.getDisplayValues();
    let changed = false;
    data.forEach(row => {
      // [ID, TS, Tgl, Kelas, Mapel, Kategori, No, Nama, Nilai, Bab, Tujuan, Bentuk]
      if (row[3] === oldKelas && row[6] === oldNo && row[7] === oldNama) {
        row[3] = newKelas; row[6] = newNo; row[7] = newNama;
        changed = true;
      }
    });
    if (changed) range.setValues(data);
  }
}

// Import siswa massal lengkap: array of {kelas, no, nama, nis, nisn, ...}
// Sekali import bisa membuat kelas baru sekaligus mengisi siswa.
function importSiswaMassal(dataArray, actor) {
  if (!Array.isArray(dataArray) || dataArray.length === 0) {
    throw new Error('Tidak ada data untuk diimpor!');
  }

  return _withLock(function() {
    _ensureSiswaLengkapColumns();

    const existingSiswa = _readSheet(SHEET_NAMES.SISWA);
    const existingSiswaKey = new Set(
      existingSiswa
        .map(function(r) { return _t(r[3]).toLowerCase() + '||' + _t(r[2]).toLowerCase(); })
        .filter(function(k) { return k !== '||'; })
    );

    const existingKelas = _readSheet(SHEET_NAMES.KELAS);
    const existingKelasNames = new Set(existingKelas.map(function(r) { return _t(r[1]).toLowerCase(); }).filter(Boolean));

    const errors = [];
    const toInsertKelas = [];
    const toInsertSiswa = [];
    const seenKelasBaru = new Set();
    const seenSiswaBatch = new Set();

    dataArray.forEach(function(row, idx) {
      const kls = _t(row.kelas);
      const no = _t(row.no);
      const nm = _t(row.nama);
      const lineNum = idx + 1;

      if (!kls || !no || !nm) {
        errors.push('Baris ' + lineNum + ': kelas, no urut, dan nama siswa wajib diisi');
        return;
      }

      const kelasKey = kls.toLowerCase();
      if (!existingKelasNames.has(kelasKey) && !seenKelasBaru.has(kelasKey)) {
        seenKelasBaru.add(kelasKey);
        existingKelasNames.add(kelasKey);
        toInsertKelas.push([_newId('K'), kls]);
      }

      const siswaKey = kelasKey + '||' + nm.toLowerCase();
      if (existingSiswaKey.has(siswaKey)) {
        errors.push('Baris ' + lineNum + ': siswa "' + nm + '" sudah ada di kelas ' + kls);
        return;
      }
      if (seenSiswaBatch.has(siswaKey)) {
        errors.push('Baris ' + lineNum + ': siswa "' + nm + '" duplikat di file impor untuk kelas ' + kls);
        return;
      }

      seenSiswaBatch.add(siswaKey);
      existingSiswaKey.add(siswaKey);

      toInsertSiswa.push([
        _newId('S'), no, nm, kls,
        _t(row.nis), _t(row.nisn), _t(row.jenisKelamin), _t(row.tempatLahir), _t(row.tanggalLahir),
        _t(row.agama), _t(row.alamatSiswa), _t(row.status), _t(row.keteranganSiswa),
        _t(row.namaAyah), _t(row.pekerjaanAyah), _t(row.noHpAyah),
        _t(row.namaIbu), _t(row.pekerjaanIbu), _t(row.noHpIbu),
        _t(row.alamatOrangTua), _t(row.keteranganOrangTua)
      ]);
    });

    if (toInsertSiswa.length === 0 && toInsertKelas.length === 0) {
      throw new Error(errors.length ? errors.join('\n') : 'Tidak ada data valid untuk diimpor.');
    }

    if (toInsertKelas.length > 0) {
      const kelasSheet = _sheet(SHEET_NAMES.KELAS);
      kelasSheet.getRange(kelasSheet.getLastRow() + 1, 1, toInsertKelas.length, 2).setValues(toInsertKelas);
    }

    let nomorDirapikan = 0;
    if (toInsertSiswa.length > 0) {
      const siswaSheet = _sheet(SHEET_NAMES.SISWA);
      siswaSheet.getRange(siswaSheet.getLastRow() + 1, 1, toInsertSiswa.length, SISWA_HEADERS.length).setValues(toInsertSiswa);
      nomorDirapikan = _rapikanNomorSiswaPerKelas();
    }

    _logActivity('Import Kelas & Siswa', toInsertKelas.length + ' kelas, ' + toInsertSiswa.length + ' siswa', actor);

    return {
      sukses: toInsertSiswa.length,
      kelasBaru: toInsertKelas.length,
      gagal: errors.length,
      errors: errors,
      pesan: 'Berhasil impor ' + toInsertSiswa.length + ' siswa' +
             (toInsertKelas.length > 0 ? ' dan membuat ' + toInsertKelas.length + ' kelas baru' : '') +
             '. Nomor urut sudah dirapikan otomatis per kelas' +
             (nomorDirapikan > 0 ? ' (' + nomorDirapikan + ' nomor disesuaikan)' : '') +
             (errors.length > 0 ? ', ' + errors.length + ' baris dilewati' : '')
    };
  });
}

// ==========================================
// 8. SIMPAN DATA - MAPEL
// ==========================================
function simpanMapel(nama) {
  return _withLock(function() {
  const nm = _t(nama);
  if (!nm) throw new Error('Nama mata pelajaran tidak boleh kosong!');
  
  const existing = _readSheet(SHEET_NAMES.MAPEL);
  if (existing.some(r => r[1].toLowerCase() === nm.toLowerCase())) {
    throw new Error('Mata pelajaran "' + nm + '" sudah ada!');
  }
  
  const id = _newId('M');
  _sheet(SHEET_NAMES.MAPEL).appendRow([id, nm]);
  _logActivity('Tambah Mapel', nm);
  return 'Mata pelajaran "' + nm + '" ditambahkan!';
  });
}


function simpanMapelMassal(dataArray) {
  if (!Array.isArray(dataArray) || dataArray.length === 0) {
    throw new Error('Tidak ada mata pelajaran untuk disimpan.');
  }
  return _withLock(function() {
    const existing = _readSheet(SHEET_NAMES.MAPEL);
    const existingNames = new Set(existing.map(r => _t(r[1]).toLowerCase()).filter(Boolean));
    const seen = new Set();
    const errors = [];
    const toInsert = [];

    dataArray.forEach((item, idx) => {
      const nama = _t(typeof item === 'string' ? item : item.nama);
      if (!nama) return;
      const key = nama.toLowerCase();
      if (existingNames.has(key)) {
        errors.push('Baris ' + (idx + 1) + ': mapel "' + nama + '" sudah ada');
        return;
      }
      if (seen.has(key)) {
        errors.push('Baris ' + (idx + 1) + ': mapel "' + nama + '" duplikat di input');
        return;
      }
      seen.add(key);
      toInsert.push([_newId('M'), nama]);
    });

    if (toInsert.length === 0) {
      throw new Error(errors.length ? errors.join('\n') : 'Isi minimal 1 mata pelajaran.');
    }

    const sheet = _sheet(SHEET_NAMES.MAPEL);
    sheet.getRange(sheet.getLastRow() + 1, 1, toInsert.length, 2).setValues(toInsert);
    _logActivity('Tambah Mapel Massal', toInsert.length + ' mapel');

    return 'Berhasil menyimpan ' + toInsert.length + ' mata pelajaran' +
      (errors.length ? ' (' + errors.length + ' baris dilewati).' : '.');
  });
}

function editMapel(id, namaBaru) {
  return _withLock(function() {
  const nm = _t(namaBaru);
  if (!nm) throw new Error('Nama tidak boleh kosong!');
  
  const found = _findRowById(SHEET_NAMES.MAPEL, id);
  if (!found) throw new Error('Mata pelajaran tidak ditemukan!');
  
  const existing = _readSheet(SHEET_NAMES.MAPEL);
  if (existing.some(r => r[0] !== id && r[1].toLowerCase() === nm.toLowerCase())) {
    throw new Error('Nama mapel sudah dipakai!');
  }
  
  _sheet(SHEET_NAMES.MAPEL).getRange(found.rowIndex, 2).setValue(nm);
  _logActivity('Edit Mapel', nm);
  return 'Mata pelajaran diperbarui!';
  });
}

// ==========================================
// 9. SIMPAN DATA - JADWAL
// ==========================================
function simpanJadwal(hari, jamMulai, jamSelesai, mapel, kelas) {
  return _withLock(function() {
  const h = _t(hari), jm = _t(jamMulai), js = _t(jamSelesai);
  const m = _t(mapel), k = _t(kelas);
  if (!h || !jm || !js || !m || !k) throw new Error('Semua kolom jadwal wajib diisi!');
  
  // PERBAIKAN BUG #5: validasi jam mulai < jam selesai
  if (jm >= js) throw new Error('Jam mulai harus lebih awal dari jam selesai!');
  
  const id = _newId('J');
  _sheet(SHEET_NAMES.JADWAL).appendRow([id, h, jm, js, m, k]);
  _logActivity('Tambah Jadwal', h + ' ' + jm + ' ' + m);
  return 'Jadwal disimpan!';
  });
}

function editJadwal(id, hari, jamMulai, jamSelesai, mapel, kelas) {
  return _withLock(function() {
  const h = _t(hari), jm = _t(jamMulai), js = _t(jamSelesai);
  const m = _t(mapel), k = _t(kelas);
  if (!h || !jm || !js || !m || !k) throw new Error('Semua kolom wajib diisi!');
  if (jm >= js) throw new Error('Jam mulai harus lebih awal dari jam selesai!');
  
  const found = _findRowById(SHEET_NAMES.JADWAL, id);
  if (!found) throw new Error('Jadwal tidak ditemukan!');
  
  _sheet(SHEET_NAMES.JADWAL).getRange(found.rowIndex, 2, 1, 5).setValues([[h, jm, js, m, k]]);
  _logActivity('Edit Jadwal', h + ' ' + m);
  return 'Jadwal diperbarui!';
  });
}


// ==========================================
// 9B. WRAPPER CEPAT UNTUK UI
// Mengurangi bolak-balik server: simpan/edit + data terbaru dalam 1 request.
// ==========================================
function simpanKelasMassalCepat(dataArray) {
  const message = simpanKelasMassal(dataArray);
  return { message: message, kelas: _readSheet(SHEET_NAMES.KELAS) };
}

function editKelasCepat(id, namaBaru) {
  const message = editKelas(id, namaBaru);
  return {
    message: message,
    kelas: _readSheet(SHEET_NAMES.KELAS),
    siswa: getSiswaData(),
    jadwal: getJadwalData()
  };
}

function simpanSiswaCepat(noUrut, nama, kelas, actor) {
  const message = simpanSiswa(noUrut, nama, kelas, actor);
  return { message: message, siswa: getSiswaData() };
}

function editSiswaCepat(id, noUrut, nama, kelas, actor) {
  const message = editSiswa(id, noUrut, nama, kelas, actor);
  return { message: message, siswa: getSiswaData() };
}

function importSiswaMassalCepat(dataArray, actor) {
  const res = importSiswaMassal(dataArray, actor);
  res.kelas = _readSheet(SHEET_NAMES.KELAS);
  res.siswa = getSiswaData();
  return res;
}

function simpanMapelMassalCepat(dataArray) {
  const message = simpanMapelMassal(dataArray);
  return { message: message, mapel: _readSheet(SHEET_NAMES.MAPEL) };
}

function editMapelCepat(id, namaBaru) {
  const message = editMapel(id, namaBaru);
  return { message: message, mapel: _readSheet(SHEET_NAMES.MAPEL) };
}

function simpanJadwalCepat(hari, jamMulai, jamSelesai, mapel, kelas) {
  const message = simpanJadwal(hari, jamMulai, jamSelesai, mapel, kelas);
  return { message: message, jadwal: getJadwalData() };
}

function editJadwalCepat(id, hari, jamMulai, jamSelesai, mapel, kelas) {
  const message = editJadwal(id, hari, jamMulai, jamSelesai, mapel, kelas);
  return { message: message, jadwal: getJadwalData() };
}

// ==========================================
// 10. SIMPAN DATA - JURNAL
// ==========================================
function simpanJurnal(tanggal, jamKe, kelas, mapel, materi, kegiatan, keterangan, actor) {
  const tgl = _t(tanggal);
  if (!_isValidDate(tgl)) throw new Error('Format tanggal tidak valid!');
  if (!_t(kelas) || !_t(mapel)) throw new Error('Kelas dan mata pelajaran wajib diisi!');
  
  const kls = _t(kelas), mpl = _t(mapel), jam = _t(jamKe);
  
  return _withLock(function() {
    // Auto-replace: kalau ada jurnal di slot yang sama, hapus dulu
    const deleted = _bulkDeleteRows(SHEET_NAMES.JURNAL, function(row) {
      return row[2] === tgl && row[4] === kls && row[5] === mpl && row[3] === jam;
    });
    
    const id = _newId('JR');
    const ts = _nowStamp();
    _sheet(SHEET_NAMES.JURNAL).appendRow([
      id, ts, tgl, jam, kls, mpl, 
      _t(materi), _t(kegiatan), _t(keterangan)
    ]);
    
    const action = deleted > 0 ? 'Update Jurnal' : 'Tambah Jurnal';
    _logActivity(action, kls + ' ' + mpl + (deleted > 0 ? ' (replace)' : ''), actor);
    
    return deleted > 0 
      ? 'Jurnal DIPERBARUI! (mengganti jurnal lama untuk slot tanggal & jam yang sama)'
      : 'Jurnal berhasil disimpan!';
  });
}

function editJurnal(id, tanggal, jamKe, kelas, mapel, materi, kegiatan, keterangan, actor) {
  return _withLock(function() {
  const tgl = _t(tanggal);
  if (!_isValidDate(tgl)) throw new Error('Format tanggal tidak valid!');
  
  const found = _findRowById(SHEET_NAMES.JURNAL, id);
  if (!found) throw new Error('Jurnal tidak ditemukan!');
  
  _sheet(SHEET_NAMES.JURNAL).getRange(found.rowIndex, 3, 1, 7).setValues([[
    tgl, _t(jamKe), _t(kelas), _t(mapel), _t(materi), _t(kegiatan), _t(keterangan)
  ]]);
  _logActivity('Edit Jurnal', _t(kelas), actor);
  return 'Jurnal diperbarui!';
  });
}

// ==========================================
// 11. SIMPAN SETTING
// ==========================================
function simpanSetting(data) {
  if (!Array.isArray(data) || data.length < 6) throw new Error('Data setting tidak lengkap!');
  const namaAplikasi = _t(data[0]);
  if (!namaAplikasi) throw new Error('Nama aplikasi / sekolah wajib diisi!');
  while (data.length < 7) data.push('');

  const sheet = _sheet(SHEET_NAMES.SETTING);
  // Pastikan header punya kolom Logo URL
  const headerLen = Math.max(sheet.getLastColumn(), 6);
  if (headerLen < 7) {
    sheet.getRange(1, 7).setValue('Logo URL').setFontWeight('bold');
  }
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  }
  sheet.getRange(2, 1, 1, 7).setValues([[
    namaAplikasi, _t(data[1]), _t(data[2]),
    _t(data[3]), _t(data[4]), _t(data[5]), _t(data[6])
  ]]);
  _logActivity('Update Pengaturan', namaAplikasi);
  return 'Pengaturan berhasil disimpan. Nama aplikasi sekarang: ' + namaAplikasi;
}

// ==========================================
// 12. HAPUS DATA (PAKAI ID)
// ==========================================
function hapusDataById(sheetName, id, actor) {
  return _withLock(function() {
  const found = _findRowById(sheetName, id);
  if (!found) throw new Error('Data tidak ditemukan!');
  const kelasSiswa = (sheetName === SHEET_NAMES.SISWA || sheetName === 'Siswa') ? _t(found.values[3]) : '';
  _sheet(sheetName).deleteRow(found.rowIndex);
  if (sheetName === SHEET_NAMES.SISWA || sheetName === 'Siswa') {
    _rapikanNomorSiswaPerKelas(kelasSiswa);
  }
  _logActivity('Hapus ' + sheetName, String(id) + (kelasSiswa ? ' (kelas ' + kelasSiswa + ')' : ''), actor);
  return (sheetName === SHEET_NAMES.SISWA || sheetName === 'Siswa')
    ? 'Data berhasil dihapus! Nomor urut siswa sudah dirapikan per kelas.'
    : 'Data berhasil dihapus!';
  });
}

// Versi cepat untuk UI: hapus + kirim data terbaru yang relevan dalam 1 panggilan.
function hapusDataByIdCepat(sheetName, id, actor) {
  const message = hapusDataById(sheetName, id, actor);
  const res = { message: message, sheetName: sheetName };
  if (sheetName === SHEET_NAMES.KELAS || sheetName === 'Kelas') {
    res.kelas = _readSheet(SHEET_NAMES.KELAS);
    res.siswa = getSiswaData();
    res.jadwal = getJadwalData();
  } else if (sheetName === SHEET_NAMES.SISWA || sheetName === 'Siswa') {
    res.siswa = getSiswaData();
  } else if (sheetName === SHEET_NAMES.MAPEL || sheetName === 'Mapel') {
    res.mapel = _readSheet(SHEET_NAMES.MAPEL);
  } else if (sheetName === SHEET_NAMES.JADWAL || sheetName === 'Jadwal') {
    res.jadwal = getJadwalData();
  }
  return res;
}

// ==========================================
// 13. SIMPAN MASAL - ABSENSI & NILAI
// Mode 'replace': hapus data existing untuk kunci yang sama, lalu insert baru.
// Mode 'append' (default): hanya insert. Dipakai kalau tidak ada existing.
// ==========================================

// === HELPER: hapus baris berdasarkan filter ===
// Strategi: rewrite sheet tanpa baris yang match predicate.
// JAUH lebih cepat daripada deleteRow per-1 untuk batch banyak baris,
// karena cuma 2 round-trip ke Sheets API (clearContent + setValues) tidak peduli jumlah baris.
function _bulkDeleteRows(sheetName, predicate) {
  const sheet = _ss().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
  
  // Pisahkan: kept rows (tidak match predicate) vs deleted count
  const kept = [];
  let deletedCount = 0;
  data.forEach(row => {
    if (predicate(row)) deletedCount++;
    else kept.push(row);
  });
  
  if (deletedCount === 0) return 0;
  
  // Strategi rewrite:
  // 1. Clear semua content di area data
  // 2. Tulis kembali kept rows mulai dari baris 2
  const totalRowsData = data.length;
  sheet.getRange(2, 1, totalRowsData, lastCol).clearContent();
  if (kept.length > 0) {
    sheet.getRange(2, 1, kept.length, lastCol).setValues(kept);
  }
  
  return deletedCount;
}

// === CEK ABSENSI EXISTING (untuk auto-load) ===
// Kunci baru: Tanggal + Kelas + Mata Pelajaran
// Return: { exists: boolean, data: { 'noUrut': 'Status', ... } }
function cekAbsensiTanggal(tanggal, kelas, mapel) {
  const tgl = _t(tanggal);
  const kls = _t(kelas);
  const mpl = _t(mapel);
  if (!tgl || !kls || !mpl) return { exists: false, data: {} };
  const c = _absensiCols();
  const rows = _readSheet(SHEET_NAMES.ABSENSI).filter(r =>
    r[c.tanggal] === tgl && r[c.kelas] === kls && r[c.mapel] === mpl
  );
  const data = {};
  rows.forEach(r => { data[r[c.no]] = r[c.status]; }); // No → Status
  return { exists: rows.length > 0, count: rows.length, data: data };
}

// === CEK NILAI EXISTING (untuk auto-load per siswa) ===
// Kunci: Tanggal + Kelas + Mapel + Kategori (Bab tidak masuk)
// Return: { exists, data: { 'noUrut': { nilai, bab, tujuan, bentuk }, ... } }
function cekNilaiTanggal(tanggal, kelas, mapel, kategori) {
  const tgl = _t(tanggal);
  const kls = _t(kelas);
  const mpl = _t(mapel);
  const ktg = _t(kategori);
  if (!tgl || !kls || !mpl || !ktg) return { exists: false, data: {} };
  
  // Struktur Nilai: [ID, TS, Tgl, Kelas, Mapel, Kategori, No, Nama, Nilai, Bab, Tujuan, Bentuk]
  const rows = _readSheet(SHEET_NAMES.NILAI).filter(r => 
    r[2] === tgl && r[3] === kls && r[4] === mpl && r[5] === ktg
  );
  const data = {};
  rows.forEach(r => { 
    data[r[6]] = { 
      nilai: r[8], 
      bab: r[9] === '-' ? '' : r[9], 
      tujuan: r[10] === '-' ? '' : r[10], 
      bentuk: r[11] === '-' ? '' : r[11] 
    }; 
  });
  // Untuk meta (bab/tujuan/bentuk), ambil dari record pertama (asumsi konsisten)
  const meta = rows.length > 0 ? {
    bab: data[rows[0][6]].bab,
    tujuan: data[rows[0][6]].tujuan,
    bentuk: data[rows[0][6]].bentuk
  } : { bab: '', tujuan: '', bentuk: '' };
  
  return { exists: rows.length > 0, count: rows.length, data: data, meta: meta };
}

// === CEK JURNAL EXISTING ===
// Kunci: Tanggal + Kelas + Mapel + Jam Ke
// Return: { exists, data: { id, materi, kegiatan, keterangan } }
function cekJurnalSlot(tanggal, kelas, mapel, jamKe) {
  const tgl = _t(tanggal);
  const kls = _t(kelas);
  const mpl = _t(mapel);
  const jam = _t(jamKe);
  if (!tgl || !kls || !mpl) return { exists: false };
  
  // Struktur Jurnal: [ID, TS, Tgl, JamKe, Kelas, Mapel, Materi, Kegiatan, Keterangan]
  const rows = _readSheet(SHEET_NAMES.JURNAL).filter(r => 
    r[2] === tgl && r[4] === kls && r[5] === mpl && r[3] === jam
  );
  if (rows.length === 0) return { exists: false };
  const r = rows[0];
  return {
    exists: true,
    data: { id: r[0], materi: r[6], kegiatan: r[7], keterangan: r[8] }
  };
}

function simpanAbsensiMasal(tanggal, kelas, mapel, dataSiswa, actor) {
  const tgl = _t(tanggal);
  const kls = _t(kelas);
  const mpl = _t(mapel);
  if (!_isValidDate(tgl)) throw new Error('Format tanggal tidak valid!');
  if (!kls) throw new Error('Kelas wajib diisi!');
  if (!mpl) throw new Error('Mata pelajaran wajib diisi!');
  if (!Array.isArray(dataSiswa) || dataSiswa.length === 0) {
    throw new Error('Tidak ada data absensi untuk disimpan.');
  }
  const statusValid = { 'Hadir': true, 'Sakit': true, 'Izin': true, 'Alpa': true, 'Terlambat': true };
  dataSiswa.forEach(function(s) {
    const st = _t(s.status) || 'Hadir';
    if (!statusValid[st]) throw new Error('Status absensi tidak valid: ' + st);
  });
  _ensureAbsensiMapelColumn();
  const c = _absensiCols();
  
  return _withLock(function() {
    // Mode replace: kunci tanggal + kelas + mata pelajaran
    const deleted = _bulkDeleteRows(SHEET_NAMES.ABSENSI, function(row) {
      return row[c.tanggal] === tgl && row[c.kelas] === kls && row[c.mapel] === mpl;
    });
    
    const ts = _nowStamp();
    const rows = dataSiswa.map(s => [
      _newId('A'), ts, tgl, kls, mpl, _t(s.no), _t(s.nama), _t(s.status)
    ]);
    
    const sheet = _sheet(SHEET_NAMES.ABSENSI);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    
    const action = deleted > 0 ? 'Update Absensi' : 'Absensi';
    _logActivity(action, kls + ' - ' + mpl + ' (' + rows.length + ' siswa)' + (deleted > 0 ? ', replace ' + deleted : ''), actor);
    
    const msg = deleted > 0 
      ? 'Absensi ' + kls + ' - ' + mpl + ' tanggal ' + tgl + ' berhasil DIPERBARUI! (mengganti ' + deleted + ' data lama)'
      : 'Absensi ' + kls + ' - ' + mpl + ' tanggal ' + tgl + ' berhasil disimpan!';
    return msg;
  });
}

function simpanNilaiMasal(tanggal, kelas, mapel, kategori, bab, tujuan, bentuk, dataSiswa) {
  const tgl = _t(tanggal);
  if (!_isValidDate(tgl)) throw new Error('Format tanggal tidak valid!');
  if (!_t(kelas) || !_t(mapel) || !_t(kategori)) {
    throw new Error('Kelas, mata pelajaran, dan kategori wajib diisi!');
  }
  if (!Array.isArray(dataSiswa) || dataSiswa.length === 0) {
    throw new Error('Tidak ada data nilai untuk disimpan.');
  }
  
  const kls = _t(kelas), mpl = _t(mapel), ktg = _t(kategori);
  
  // Validasi & normalisasi nilai per siswa (read-only, di luar lock)
  const filtered = [];
  const invalidErrors = [];
  dataSiswa.forEach((s, idx) => {
    const raw = _t(s.nilai);
    if (raw === '') return; // kosong = tidak ikut
    // Terima koma sebagai desimal Indonesia, ubah ke titik
    const normalized = raw.replace(',', '.');
    const num = parseFloat(normalized);
    if (isNaN(num)) {
      invalidErrors.push('Nilai "' + raw + '" untuk siswa no ' + s.no + ' bukan angka.');
      return;
    }
    if (num < 0 || num > 100) {
      invalidErrors.push('Nilai ' + num + ' untuk siswa no ' + s.no + ' di luar range 0-100.');
      return;
    }
    // Simpan dengan 2 desimal max kalau ada desimal, atau integer kalau bulat
    const finalVal = (num === Math.floor(num)) ? String(Math.floor(num)) : num.toFixed(2);
    filtered.push({ no: _t(s.no), nama: _t(s.nama), nilai: finalVal });
  });
  
  if (invalidErrors.length > 0) {
    throw new Error('Nilai tidak valid:\n' + invalidErrors.slice(0, 5).join('\n') + 
      (invalidErrors.length > 5 ? '\n... dan ' + (invalidErrors.length - 5) + ' error lain' : ''));
  }
  
  return _withLock(function() {
    // Mode replace: hapus dulu data existing dengan kunci tgl+kelas+mapel+kategori
    // (Bab tidak masuk kunci sesuai keputusan UX)
    const deleted = _bulkDeleteRows(SHEET_NAMES.NILAI, function(row) {
      return row[2] === tgl && row[3] === kls && row[4] === mpl && row[5] === ktg;
    });
    
    if (filtered.length === 0) {
      // Kalau user hapus semua nilai, anggap sebagai aksi "hapus semua nilai untuk slot ini"
      if (deleted > 0) {
        _logActivity('Hapus Semua Nilai', kls + ' ' + mpl + ' ' + ktg);
        return 'Semua nilai untuk ' + kls + ' (' + mpl + ' - ' + ktg + ') tanggal ' + tgl + ' telah dihapus.';
      }
      throw new Error('Tidak ada nilai yang diisi.');
    }
    
    const ts = _nowStamp();
    const rows = filtered.map(s => [
      _newId('N'), ts, tgl, kls, mpl, ktg,
      s.no, s.nama, s.nilai,
      _t(bab) || '-', _t(tujuan) || '-', _t(bentuk) || '-'
    ]);
    
    const sheet = _sheet(SHEET_NAMES.NILAI);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    
    const action = deleted > 0 ? 'Update Nilai' : 'Nilai';
    _logActivity(action, kls + ' ' + mpl + ' (' + rows.length + ' siswa)' + (deleted > 0 ? ', replace ' + deleted : ''));
    
    const msg = deleted > 0 
      ? 'Nilai ' + kls + ' (' + mpl + ' - ' + ktg + ') berhasil DIPERBARUI! (' + rows.length + ' siswa, mengganti ' + deleted + ' data lama)'
      : 'Nilai ' + kls + ' (' + mpl + ') berhasil disimpan! (' + rows.length + ' siswa)';
    return msg;
  });
}

// ==========================================
// 14. EDIT NILAI (per baris)
// ==========================================
function editNilai(id, nilaiBaru) {
  return _withLock(function() {
  const found = _findRowById(SHEET_NAMES.NILAI, id);
  if (!found) throw new Error('Data nilai tidak ditemukan!');
  // Kolom Nilai: index 9 (1-based)
  _sheet(SHEET_NAMES.NILAI).getRange(found.rowIndex, 9).setValue(_t(nilaiBaru));
  _logActivity('Edit Nilai', String(id));
  return 'Nilai diperbarui!';
  });
}

// ==========================================
// 15. REKAP - HELPER
// ==========================================
function getDataRekapanHelper(kelas, jenis, tglMulai, tglAkhir, mapel) {
  const sheet = _ss().getSheetByName(jenis);
  let result = { headers: [], data: [], isGrouped: false };
  if (!sheet || sheet.getLastRow() < 2) return result;
  
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  
  // Mapping kolom (struktur baru pakai ID di kolom 0):
  // Absensi: [ID, TS, Tgl, Kelas, Mapel, No, Nama, Status]
  // Nilai:   [ID, TS, Tgl, Kelas, Mapel, Kategori, No, Nama, Nilai, Bab, Tujuan, Bentuk]
  // Jurnal:  [ID, TS, Tgl, JamKe, Kelas, Mapel, Materi, Kegiatan, Keterangan]
  
  const tglIdx = 2; // sama untuk ketiganya
  const klsIdx = (jenis === 'Jurnal') ? 4 : 3;
  const mapelIdx = (jenis === 'Absensi') ? _absensiCols().mapel : ((jenis === 'Jurnal') ? 5 : 4);
  const selectedMapel = _t(mapel);
  
  const filtered = allData.filter(row => {
    const tgl = row[tglIdx];
    const kls = row[klsIdx];
    const mpl = row[mapelIdx];
    const matchKelas = (kelas === 'Semua') || (kls === kelas);
    const matchMapel = !selectedMapel || selectedMapel === 'Semua' || mpl === selectedMapel;
    const matchTgl = (tgl >= tglMulai && tgl <= tglAkhir);
    return matchKelas && matchMapel && matchTgl;
  });
  
  if (jenis === 'Absensi') {
    result.headers = ['No', 'Nama Lengkap', 'Hadir', 'Sakit', 'Izin', 'Alpa', 'Terlambat'];
    const c = _absensiCols();
    const rekap = {};
    filtered.forEach(row => {
      const no = row[c.no], nama = row[c.nama];
      let status = _t(row[c.status]);
      if (/^al/i.test(status)) status = 'Alpa'; // kompatibel ejaan lama pada data tersimpan
      const key = no + '||' + nama;
      if (!rekap[key]) rekap[key] = { no, nama, H: 0, S: 0, I: 0, A: 0, T: 0 };
      if (status === 'Hadir') rekap[key].H++;
      else if (status === 'Sakit') rekap[key].S++;
      else if (status === 'Izin') rekap[key].I++;
      else if (status === 'Alpa') rekap[key].A++;
      else if (status === 'Terlambat') rekap[key].T++;
    });
    result.data = Object.values(rekap)
      .map(s => [s.no, s.nama, s.H, s.S, s.I, s.A, s.T])
      .sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
      
  } else if (jenis === 'Nilai') {
    result.isGrouped = true;
    const grouped = {};
    filtered.forEach(row => {
      const tgl = row[2], mapel = row[4], kategori = row[5];
      const no = row[6], nama = row[7], nilai = row[8];
      const bab = row[9] || '-', tujuan = row[10] || '-', bentuk = row[11] || '-';
      const key = tgl + '|' + mapel + '|' + kategori + '|' + bab + '|' + bentuk;
      if (!grouped[key]) {
        grouped[key] = { tanggal: tgl, mapel, kategori, bab, tujuan, bentuk, siswa: [] };
      }
      grouped[key].siswa.push([no, nama, nilai]);
    });
    Object.values(grouped).forEach(g => {
      g.siswa.sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
    });
    result.data = Object.values(grouped);
    
  } else if (jenis === 'Jurnal') {
    result.headers = ['Tanggal', 'Jam Ke', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', 'Kegiatan', 'Keterangan'];
    result.data = filtered.map(row => [row[2], row[3], row[4], row[5], row[6], row[7], row[8]]);
  }
  
  return result;
}

function getPreviewRekap(kelas, jenis, tglMulai, tglAkhir, mode, bulan, tahun, mapel) {
  const range = _resolveRange(mode || 'custom', bulan, tahun, tglMulai, tglAkhir);
  const selectedMapel = _t(mapel);
  
  // Untuk Absensi: kembalikan format pivot sesuai kelas dan mata pelajaran
  if (jenis === 'Absensi' && kelas !== 'Semua' && kelas) {
    const pivot = getRekapAbsensiPivot(kelas, range.tglMulai, range.tglAkhir, selectedMapel);
    return {
      isPivot: true,
      kelas: kelas,
      mapel: selectedMapel,
      periodeLabel: range.label,
      headerTanggal: pivot.headerTanggal,
      rows: pivot.rows
    };
  }
  
  // Untuk Nilai/Jurnal: format standar/grouped sesuai mata pelajaran
  const result = getDataRekapanHelper(kelas, jenis, range.tglMulai, range.tglAkhir, selectedMapel);
  result.mapel = selectedMapel;
  result.periodeLabel = range.label;
  return result;
}

// ==========================================
// 15.5 HELPER REKAP BULANAN ABSENSI (PIVOT)
// ==========================================

// Konversi bulan & tahun ke rentang tanggal YYYY-MM-DD
function _bulanToRange(bulan, tahun) {
  const b = parseInt(bulan); // 1-12
  const t = parseInt(tahun);
  if (!b || b < 1 || b > 12 || !t || t < 2000 || t > 2100) {
    throw new Error('Bulan/tahun tidak valid!');
  }
  const lastDay = new Date(t, b, 0).getDate(); // hari terakhir bulan
  const mm = (b < 10 ? '0' : '') + b;
  return {
    tglMulai: t + '-' + mm + '-01',
    tglAkhir: t + '-' + mm + '-' + (lastDay < 10 ? '0' + lastDay : lastDay),
    namaBulan: ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][b-1],
    tahun: t,
    lastDay: lastDay
  };
}

// Resolve range tanggal berdasarkan mode (bulanan / custom)
function _resolveRange(mode, bulan, tahun, tglMulai, tglAkhir) {
  if (mode === 'bulanan') {
    const r = _bulanToRange(bulan, tahun);
    return { tglMulai: r.tglMulai, tglAkhir: r.tglAkhir, label: r.namaBulan + ' ' + r.tahun, meta: r };
  }
  if (!_isValidDate(tglMulai) || !_isValidDate(tglAkhir)) {
    throw new Error('Rentang tanggal tidak valid!');
  }
  return { tglMulai: tglMulai, tglAkhir: tglAkhir, label: tglMulai + ' s.d ' + tglAkhir, meta: null };
}

// Generate pivot absensi:
// - Baris: SEMUA siswa terdaftar di kelas (meski tidak ada record)
// - Kolom: hari aktif sekolah (terdeteksi dari data + skip Minggu) atau tanggal di range
// - Sel: status (H/S/I/A) atau kosong
// - Akhir: total H/S/I/A/T per siswa
//
// kelas: nama kelas spesifik (TIDAK 'Semua' untuk pivot — tidak masuk akal)
function getRekapAbsensiPivot(kelas, tglMulai, tglAkhir, mapel) {
  if (!kelas || kelas === 'Semua') {
    throw new Error('Pivot absensi memerlukan kelas spesifik (bukan Semua)');
  }
  
  const selectedMapel = _t(mapel);
  
  // 1. Ambil semua siswa di kelas
  const siswaAll = _readSheet(SHEET_NAMES.SISWA)
    .filter(r => r[3] === kelas)
    .sort((a, b) => (parseInt(a[1]) || 0) - (parseInt(b[1]) || 0));
  
  // 2. Ambil semua absensi di range dan mapel yang dipilih
  const c = _absensiCols();
  const absensi = _readSheet(SHEET_NAMES.ABSENSI).filter(r => {
    const matchMapel = !selectedMapel || selectedMapel === 'Semua' || r[c.mapel] === selectedMapel;
    return r[c.kelas] === kelas && matchMapel && r[c.tanggal] >= tglMulai && r[c.tanggal] <= tglAkhir;
  });
  
  // 3. Tentukan kolom tanggal yang ditampilkan
  // Strategi: ambil semua tanggal unik dari data absensi yang ada.
  // Kalau tidak ada data sama sekali, fallback ke semua hari Senin-Sabtu di range.
  const tanggalSet = new Set();
  absensi.forEach(r => tanggalSet.add(r[c.tanggal]));
  
  let tanggalKolom;
  if (tanggalSet.size > 0) {
    tanggalKolom = Array.from(tanggalSet).sort();
  } else {
    // Fallback: enumerate Senin-Sabtu di range
    tanggalKolom = [];
    const start = new Date(tglMulai + 'T00:00:00');
    const end = new Date(tglAkhir + 'T00:00:00');
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      if (d.getDay() !== 0) { // 0 = Minggu
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        tanggalKolom.push(yyyy + '-' + mm + '-' + dd);
      }
    }
  }
  
  // 4. Build map: key = noUrut + '|' + nama, value = { tanggal: status, ... }
  const dataMap = {};
  absensi.forEach(r => {
    const key = r[c.no] + '|' + r[c.nama]; // No|Nama
    if (!dataMap[key]) dataMap[key] = {};
    // Jika dalam satu tanggal ada beberapa mapel, rekap harian menampilkan status terakhir pada tanggal itu.
    let stNorm = _t(r[c.status]);
    if (/^al/i.test(stNorm)) stNorm = 'Alpa'; // kompatibel ejaan lama pada data tersimpan
    dataMap[key][r[c.tanggal]] = stNorm;
  });
  
  // 5. Build rows untuk SEMUA siswa terdaftar
  const STATUS_KODE = { 'Hadir':'H', 'Sakit':'S', 'Izin':'I', 'Alpa':'A', 'Terlambat':'T' };
  const rows = siswaAll.map(s => {
    const no = s[1], nama = s[2];
    const key = no + '|' + nama;
    const records = dataMap[key] || {};
    const cells = tanggalKolom.map(tgl => STATUS_KODE[records[tgl]] || '');
    let H = 0, S = 0, I = 0, A = 0, T = 0;
    Object.values(records).forEach(st => {
      if (st === 'Hadir') H++;
      else if (st === 'Sakit') S++;
      else if (st === 'Izin') I++;
      else if (st === 'Alpa') A++;
      else if (st === 'Terlambat') T++;
    });
    return { no: no, nama: nama, cells: cells, H: H, S: S, I: I, A: A, T: T };
  });
  
  // 6. Format tanggal kolom jadi tanggal saja (dd) untuk header singkat
  const headerTanggal = tanggalKolom.map(t => parseInt(t.substring(8, 10))); // ambil DD
  
  return {
    kelas: kelas,
    mapel: selectedMapel,
    tanggalKolom: tanggalKolom,
    headerTanggal: headerTanggal,
    rows: rows,
    totalSiswa: siswaAll.length
  };
}

// ==========================================
// 16. EXPORT CSV (LEGACY - TANPA DRIVE)
// CSV dikembalikan sebagai teks agar bisa diunduh langsung dari browser.
// Fungsi utama di UI sekarang memakai download CSV lokal di browser.
// ==========================================
function exportRekapCSV(kelas, jenis, tglMulai, tglAkhir, mode, bulan, tahun, mapel) {
  const range = _resolveRange(mode || 'custom', bulan, tahun, tglMulai, tglAkhir);
  const tglM = range.tglMulai;
  const tglA = range.tglAkhir;
  const selectedMapel = _t(mapel);
  let csv = '';
  
  if (jenis === 'Absensi') {
    if (kelas === 'Semua') {
      throw new Error('Rekap Absensi Pivot harus pilih kelas spesifik (tidak bisa Semua Kelas).');
    }
    const pivot = getRekapAbsensiPivot(kelas, tglM, tglA, selectedMapel);
    csv += 'REKAPITULASI ABSENSI\n';
    csv += 'Kelas:,' + _csvEscape(kelas) + '\n';
    csv += 'Mata Pelajaran:,' + _csvEscape(selectedMapel || '-') + '\n';
    csv += 'Periode:,' + _csvEscape(range.label) + '\n';
    csv += 'Tanggal Cetak:,' + _csvEscape(_displayStamp()) + '\n\n';
    const headerRow = ['No', 'Nama Siswa'];
    pivot.headerTanggal.forEach(d => headerRow.push(d));
    headerRow.push('H', 'S', 'I', 'A', 'T');
    csv += headerRow.map(_csvEscape).join(',') + '\n';
    pivot.rows.forEach(r => {
      const row = [r.no, r.nama].concat(r.cells).concat([r.H, r.S, r.I, r.A, r.T]);
      csv += row.map(_csvEscape).join(',') + '\n';
    });
    csv += '\nKeterangan:,H = Hadir, S = Sakit, I = Izin, A = Alpa, T = Terlambat\n';
  } else if (jenis === 'Nilai') {
    const rekap = getDataRekapanHelper(kelas, jenis, tglM, tglA, selectedMapel);
    csv = 'REKAPITULASI NILAI\n';
    csv += 'Kelas:,' + _csvEscape(kelas) + '\n';
    csv += 'Mata Pelajaran:,' + _csvEscape(selectedMapel || '-') + '\n';
    csv += 'Periode:,' + _csvEscape(range.label) + '\n\n';
    csv += 'Tanggal,Mata Pelajaran,Kategori,Bab,Bentuk,Tujuan Pembelajaran,No,Nama Siswa,Nilai\n';
    rekap.data.forEach(g => {
      g.siswa.forEach(s => {
        csv += [
          _csvEscape(g.tanggal), _csvEscape(g.mapel), _csvEscape(g.kategori),
          _csvEscape(g.bab), _csvEscape(g.bentuk), _csvEscape(g.tujuan),
          _csvEscape(s[0]), _csvEscape(s[1]), _csvEscape(s[2])
        ].join(',') + '\n';
      });
    });
  } else {
    const rekap = getDataRekapanHelper(kelas, jenis, tglM, tglA, selectedMapel);
    csv = 'REKAPITULASI JURNAL\n';
    csv += 'Kelas:,' + _csvEscape(kelas) + '\n';
    csv += 'Mata Pelajaran:,' + _csvEscape(selectedMapel || '-') + '\n';
    csv += 'Periode:,' + _csvEscape(range.label) + '\n\n';
    csv += rekap.headers.map(_csvEscape).join(',') + '\n';
    rekap.data.forEach(row => {
      csv += row.map(_csvEscape).join(',') + '\n';
    });
  }
  return {
    filename: 'Rekap_' + jenis + '_' + kelas + '_' + (selectedMapel || 'Semua_Mapel') + '_' + tglM + '_' + tglA + '.csv',
    content: '\uFEFF' + csv,
    mimeType: 'text/csv; charset=utf-8'
  };
}

function _exportRekapCSVLegacy(kelas, jenis, tglMulai, tglAkhir) {
  return exportRekapCSV(kelas, jenis, tglMulai, tglAkhir, 'custom', null, null, null);
}

function _csvEscape(v) {
  const s = (v == null) ? '' : String(v);
  if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

// ==========================================
// 17. CETAK REKAP (LEGACY - TANPA DRIVE)
// UI utama memakai print browser dari sisi client.
// Fungsi ini tetap disediakan bila ada pemanggilan lama, namun hanya mengembalikan HTML.
// ==========================================
function prosesRekapCetak(kelas, jenis, tglMulai, tglAkhir, mode, bulan, tahun, mapel) {
  const kop = _getSettingValues();
  const range = _resolveRange(mode || 'custom', bulan, tahun, tglMulai, tglAkhir);
  const selectedMapel = _t(mapel);
  let html;
  if (jenis === 'Absensi') {
    if (kelas === 'Semua') {
      throw new Error('Rekap Absensi memerlukan kelas spesifik (tidak bisa Semua Kelas).');
    }
    html = _buildAbsensiPivotPdf(kop, kelas, range, selectedMapel);
  } else {
    html = _buildStandardPdf(kop, kelas, jenis, range.tglMulai, range.tglAkhir, range.label, selectedMapel);
  }
  return {
    filename: 'Laporan_' + jenis + '_' + kelas + '_' + (selectedMapel || 'Semua_Mapel') + '_' + range.tglMulai + '.html',
    html: html
  };
}

// ----- PDF ABSENSI PIVOT (LANDSCAPE) -----
function _buildAbsensiPivotPdf(kop, kelas, range, mapel) {
  const selectedMapel = _t(mapel);
  const pivot = getRekapAbsensiPivot(kelas, range.tglMulai, range.tglAkhir, selectedMapel);
  const totalCols = 2 + pivot.headerTanggal.length + 5; // No, Nama, ...tanggal, H, S, I, A, T
  
  let tableHtml;
  if (pivot.rows.length === 0) {
    tableHtml = '<p style="text-align:center; font-weight:bold; padding:30px;">Belum ada siswa terdaftar di kelas ini.</p>';
  } else {
    // Header row
    let headerCells = '<th class="col-no">No</th><th class="col-nama">Nama Siswa</th>';
    pivot.headerTanggal.forEach(d => {
      headerCells += '<th class="col-tgl">' + d + '</th>';
    });
    headerCells += '<th class="col-rekap">H</th><th class="col-rekap">S</th><th class="col-rekap">I</th><th class="col-rekap">A</th><th class="col-rekap">T</th>';
    
    // Body rows
    let bodyRows = '';
    pivot.rows.forEach(r => {
      let cells = '<td class="col-no">' + _esc(r.no) + '</td><td class="col-nama">' + _esc(r.nama) + '</td>';
      r.cells.forEach(c => {
        const cls = c ? 'cell-' + c : '';
        cells += '<td class="col-tgl ' + cls + '">' + _esc(c) + '</td>';
      });
      cells += '<td class="col-rekap">' + r.H + '</td>';
      cells += '<td class="col-rekap">' + r.S + '</td>';
      cells += '<td class="col-rekap">' + r.I + '</td>';
      cells += '<td class="col-rekap">' + r.A + '</td>';
      cells += '<td class="col-rekap">' + r.T + '</td>';
      bodyRows += '<tr>' + cells + '</tr>';
    });
    
    tableHtml = 
      '<table class="pivot-table">' +
        '<thead><tr>' + headerCells + '</tr></thead>' +
        '<tbody>' + bodyRows + '</tbody>' +
      '</table>';
  }
  
  return (
    '<!DOCTYPE html><html><head>' +
    '<meta charset="UTF-8">' +
    '<style>' +
      '@page { size: A4 landscape; margin: 1cm; }' +
      'body { font-family: Helvetica, Arial, sans-serif; font-size: 10px; }' +
      '.header { text-align: center; border-bottom: 2px solid black; padding-bottom: 8px; margin-bottom: 12px; }' +
      '.header h1 { margin: 0; font-size: 16px; text-transform: uppercase; }' +
      '.header p { margin: 1px 0; font-size: 10px; }' +
      'h3 { text-align: center; text-transform: uppercase; margin: 5px 0; font-size: 13px; }' +
      '.periode-info { text-align: center; margin-bottom: 10px; font-size: 11px; }' +
      '.periode-info b { display: inline-block; margin: 0 8px; }' +
      '.pivot-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 9px; }' +
      '.pivot-table th, .pivot-table td { border: 1px solid #555; padding: 3px 2px; text-align: center; vertical-align: middle; }' +
      '.pivot-table th { background-color: #e0e0e0; font-weight: bold; }' +
      '.col-no { width: 25px; }' +
      '.col-nama { width: 140px; text-align: left !important; padding-left: 5px !important; }' +
      '.col-tgl { font-weight: bold; }' +
      '.col-rekap { width: 26px; background-color: #f5f5f5; font-weight: bold; }' +
      '.cell-H { color: #198754; }' +
      '.cell-S { color: #b8860b; background: #fff8e1; }' +
      '.cell-I { color: #0a6c8a; background: #e1f5fe; }' +
      '.cell-A { color: #fff; background: #dc3545; font-weight: bold; }' +
      '.cell-T { color: #fff; background: #6f42c1; font-weight: bold; }' +
      '.legenda { margin-top: 8px; font-size: 9px; }' +
      '.legenda span { display: inline-block; margin-right: 12px; }' +
      '.ttd-container { width: 100%; margin-top: 25px; page-break-inside: avoid; }' +
      '.ttd-kiri { float: left; width: 40%; text-align: center; }' +
      '.ttd-kanan { float: right; width: 40%; text-align: center; }' +
      '.ttd-container p { margin: 2px 0; font-size: 10px; }' +
      '.clear { clear: both; }' +
    '</style></head><body>' +
      '<div class="header"><h1>' + _esc(kop[0]) + '</h1><p>' + _esc(kop[1]) + '</p></div>' +
      '<h3>Rekapitulasi Absensi Siswa</h3>' +
      '<p class="periode-info"><b>Kelas: ' + _esc(kelas) + '</b> | <b>Mapel: ' + _esc(selectedMapel || '-') + '</b> | <b>Periode: ' + _esc(range.label) + '</b></p>' +
      tableHtml +
      '<p class="legenda"><b>Keterangan:</b> ' +
        '<span><b style="color:#198754;">H</b> = Hadir</span>' +
        '<span><b style="color:#b8860b;">S</b> = Sakit</span>' +
        '<span><b style="color:#0a6c8a;">I</b> = Izin</span>' +
        '<span><b style="color:#dc3545;">A</b> = Alpa</span>' +
        '<span><b style="color:#6f42c1;">T</b> = Terlambat</span>' +
        '<span>(sel kosong = tidak ada absensi tercatat)</span>' +
      '</p>' +
      '<div class="ttd-container">' +
        '<div class="ttd-kiri"><p>Mengetahui,</p><p>Kepala Sekolah</p><br><br><br>' +
          '<p><b><u>' + _esc(kop[4]) + '</u></b></p><p>NIP. ' + _esc(kop[5]) + '</p></div>' +
        '<div class="ttd-kanan"><p>................., .........................</p><p>Guru Mata Pelajaran</p><br><br><br>' +
          '<p><b><u>' + _esc(kop[2]) + '</u></b></p><p>NIP. ' + _esc(kop[3]) + '</p></div>' +
        '<div class="clear"></div>' +
      '</div>' +
    '</body></html>'
  );
}

// ----- PDF NILAI / JURNAL (PORTRAIT, format lama) -----
function _buildStandardPdf(kop, kelas, jenis, tglM, tglA, periodeLabel, mapel) {
  const selectedMapel = _t(mapel);
  const rekap = getDataRekapanHelper(kelas, jenis, tglM, tglA, selectedMapel);
  const textKelas = (kelas === 'Semua') ? 'Semua Kelas' : 'Kelas ' + kelas;
  let tabelHtml = '';
  
  if (rekap.isGrouped && jenis === 'Nilai') {
    if (rekap.data.length === 0) {
      tabelHtml = '<p style="text-align:center; font-weight:bold;">Tidak ada data pada periode ini.</p>';
    } else {
      const parts = [];
      rekap.data.forEach(g => {
        parts.push(
          '<div class="meta-nilai"><table>' +
            '<tr><td width="160px"><b>Mata Pelajaran</b></td><td>: ' + _esc(g.mapel) + '</td></tr>' +
            '<tr><td><b>Jenis</b></td><td>: ' + _esc(g.kategori) + '</td></tr>' +
            '<tr><td><b>Tanggal</b></td><td>: ' + _esc(g.tanggal) + '</td></tr>' +
            '<tr><td><b>Bab</b></td><td>: ' + _esc(g.bab) + '</td></tr>' +
            '<tr><td><b>Tujuan</b></td><td>: ' + _esc(g.tujuan) + '</td></tr>' +
            '<tr><td><b>Bentuk</b></td><td>: ' + _esc(g.bentuk) + '</td></tr>' +
          '</table></div>' +
          '<table class="data-table">' +
            '<thead><tr><th width="50px">No</th><th>Nama Siswa</th><th width="100px">Nilai</th></tr></thead>' +
            '<tbody>' +
              g.siswa.map(s => 
                '<tr><td style="text-align:center">' + _esc(s[0]) + '</td>' +
                '<td>' + _esc(s[1]) + '</td>' +
                '<td style="text-align:center">' + _esc(s[2]) + '</td></tr>'
              ).join('') +
            '</tbody>' +
          '</table><br>'
        );
      });
      tabelHtml = parts.join('');
    }
  } else {
    const headers = rekap.headers;
    tabelHtml = '<table class="data-table"><thead><tr>' +
      headers.map(h => '<th>' + _esc(h) + '</th>').join('') +
      '</tr></thead><tbody>' +
      (rekap.data.length > 0
        ? rekap.data.map(row => '<tr>' + row.map(c => '<td>' + _esc(c) + '</td>').join('') + '</tr>').join('')
        : '<tr><td colspan="' + headers.length + '" style="text-align:center;">Tidak ada data pada periode ini.</td></tr>'
      ) + '</tbody></table>';
  }
  
  return (
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      'body { font-family: Helvetica, Arial, sans-serif; font-size: 12px; }' +
      '.header { text-align: center; border-bottom: 3px solid black; padding-bottom: 10px; margin-bottom: 20px; }' +
      '.header h1 { margin: 0; font-size: 18px; text-transform: uppercase; }' +
      '.header p { margin: 2px 0; font-size: 12px; }' +
      'h3 { text-align: center; text-transform: uppercase; margin-bottom: 5px; }' +
      '.periode { text-align: center; margin-bottom: 20px; font-style: italic; }' +
      '.meta-nilai table { width: 100%; border: none; margin-bottom: 5px; }' +
      '.meta-nilai td { border: none; padding: 2px; text-align: left; }' +
      '.data-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }' +
      '.data-table th, .data-table td { border: 1px solid #000; padding: 6px 8px; text-align: left; }' +
      '.data-table th { background-color: #f2f2f2; text-align: center; }' +
      '.ttd-container { width: 100%; margin-top: 40px; page-break-inside: avoid; }' +
      '.ttd-kiri { float: left; width: 45%; text-align: center; }' +
      '.ttd-kanan { float: right; width: 45%; text-align: center; }' +
      '.clear { clear: both; }' +
    '</style></head><body>' +
      '<div class="header"><h1>' + _esc(kop[0]) + '</h1><p>' + _esc(kop[1]) + '</p></div>' +
      '<h3>Laporan Rekapitulasi ' + _esc(jenis) + '</h3>' +
      '<p class="periode">' + _esc(textKelas) + ' | Mapel: ' + _esc(selectedMapel || '-') + ' | Periode: ' + _esc(periodeLabel) + '</p>' +
      tabelHtml +
      '<div class="ttd-container">' +
        '<div class="ttd-kiri"><p>Mengetahui,</p><p>Kepala Sekolah</p><br><br><br><br>' +
          '<p><b><u>' + _esc(kop[4]) + '</u></b></p><p>NIP. ' + _esc(kop[5]) + '</p></div>' +
        '<div class="ttd-kanan"><p>................., .........................</p><p>Guru Mata Pelajaran</p><br><br><br><br>' +
          '<p><b><u>' + _esc(kop[2]) + '</u></b></p><p>NIP. ' + _esc(kop[3]) + '</p></div>' +
        '<div class="clear"></div>' +
      '</div>' +
    '</body></html>'
  );
}

// PERBAIKAN BUG #3: escape HTML supaya aman dari XSS & nama dengan karakter spesial
function _esc(v) {
  if (v == null) return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ==========================================
// 18. DIAGNOSTIK (untuk customer support)
// Jalankan dari Apps Script editor → pilih checkDatabaseHealth → Run → lihat Logger output
// Customer copy-paste output ini saat lapor bug ke developer
// ==========================================
function checkDatabaseHealth() {
  const ss = _ss();
  const report = [];
  report.push('=== DASHBOARD GURU - HEALTH CHECK ===');
  report.push('Spreadsheet: ' + ss.getName());
  report.push('Script Timezone: ' + Session.getScriptTimeZone());
  report.push('Tanggal Cek: ' + _displayStamp());
  report.push('');
  
  // Cek setiap sheet
  const expectedSheets = {
    'Kelas':    ['ID', 'Nama Kelas'],
    'Siswa':    SISWA_HEADERS,
    'Mapel':    ['ID', 'Nama Mata Pelajaran'],
    'Jadwal':   ['ID', 'Hari', 'Jam Mulai', 'Jam Selesai', 'Mata Pelajaran', 'Kelas'],
    'Absensi':  ['ID', 'Timestamp', 'Tanggal', 'Kelas', 'Mata Pelajaran', 'No Urut', 'Nama Siswa', 'Status'],
    'Nilai':    ['ID', 'Timestamp', 'Tanggal', 'Kelas', 'Mata Pelajaran', 'Kategori', 'No Urut', 'Nama Siswa', 'Nilai', 'Bab', 'Tujuan Pembelajaran', 'Bentuk'],
    'Jurnal':   ['ID', 'Timestamp', 'Tanggal', 'Jam Ke', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', 'Kegiatan Pembelajaran', 'Keterangan'],
    'Setting':  ['Nama Aplikasi / Sekolah', 'Alamat Lengkap', 'Nama Guru', 'NIP Guru', 'Nama Kepala Sekolah', 'NIP Kepsek', 'Logo URL'],
    'Auth':     ['Username', 'Password'],
    'Aktivitas':['Waktu', 'Aksi', 'Detail']
  };
  
  let issuesFound = 0;
  
  for (const name in expectedSheets) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      report.push('❌ Sheet "' + name + '" TIDAK ADA');
      issuesFound++;
      continue;
    }
    const expected = expectedSheets[name];
    const actual = sheet.getRange(1, 1, 1, Math.min(expected.length, sheet.getLastColumn())).getValues()[0];
    const headerMatch = expected.every((h, i) => actual[i] === h);
    const rowCount = Math.max(0, sheet.getLastRow() - 1);
    
    let line = (headerMatch ? '✓' : '⚠️') + ' Sheet "' + name + '": ' + rowCount + ' baris data';
    if (!headerMatch) {
      line += ' — HEADER TIDAK COCOK!';
      line += ' (Diharapkan: ' + expected.join(', ') + ')';
      line += ' (Aktual: ' + actual.join(', ') + ')';
      issuesFound++;
    }
    report.push(line);
    
    // Cek format kolom tanggal & timestamp untuk sheet relevan
    if (['Absensi', 'Nilai', 'Jurnal'].indexOf(name) >= 0 && sheet.getLastRow() > 1) {
      const tglColIdx = expected.indexOf('Tanggal') + 1;
      const tsColIdx = expected.indexOf('Timestamp') + 1;
      if (tglColIdx > 0) {
        const fmt = sheet.getRange(2, tglColIdx).getNumberFormat();
        if (fmt !== '@') {
          report.push('  ⚠️ Format kolom Tanggal bukan "Plain text" (saat ini: ' + fmt + ') — bisa menyebabkan bug rekap');
          issuesFound++;
        }
      }
      if (tsColIdx > 0) {
        const fmt = sheet.getRange(2, tsColIdx).getNumberFormat();
        if (fmt !== '@') {
          report.push('  ⚠️ Format kolom Timestamp bukan "Plain text" (saat ini: ' + fmt + ')');
          issuesFound++;
        }
      }
    }
  }
  
  // Cek orphan: siswa di sheet Siswa dengan kelas tidak terdaftar
  const validKelas = new Set(_readSheet(SHEET_NAMES.KELAS).map(r => r[1]));
  const orphanSiswa = _readSheet(SHEET_NAMES.SISWA).filter(r => !validKelas.has(r[3]));
  if (orphanSiswa.length > 0) {
    report.push('');
    report.push('⚠️ ' + orphanSiswa.length + ' siswa terdaftar di kelas yang tidak ada di sheet Kelas:');
    orphanSiswa.slice(0, 5).forEach(r => report.push('  - ' + r[2] + ' (kelas: ' + r[3] + ')'));
    issuesFound++;
  }
  
  report.push('');
  if (issuesFound === 0) {
    report.push('✓ Database sehat, tidak ada masalah ditemukan.');
  } else {
    report.push('Total isu ditemukan: ' + issuesFound);
    report.push('Saran: jalankan setupDatabase() untuk perbaiki header & format kolom.');
  }
  
  const output = report.join('\n');
  Logger.log(output);
  return output;
}
// ==========================================
// 19. DATA PENGAMPU (CRUD) - Admin only
// ==========================================
// Pengampu disimpan di sheet Auth dengan Role='Pengampu'.

function getPengampuData() {
  _ensureAuthSchema();
  const rows = _readSheet(SHEET_NAMES.AUTH);
  return rows
    .filter(r => _t(r[3]).toLowerCase() === 'pengampu')
    .map(r => ({
      id: _t(r[0]),
      username: _t(r[1]),
      // password sengaja tidak dibuka; hanya panjang untuk indikator
      nama: _t(r[4]),
      nip: _t(r[5]),
      noHp: _t(r[6]),
      kelasDiampu: _t(r[7]).split(',').map(_t).filter(Boolean),
      active: (_t(r[8]).toLowerCase() !== 'tidak')
    }));
}

function simpanPengampu(data) {
  const nama = _t(data && data.nama);
  const nip = _t(data && data.nip);
  const noHp = _t(data && data.noHp);
  const username = _t(data && data.username);
  const password = _t(data && data.password);
  const kelasDiampu = Array.isArray(data && data.kelasDiampu) ? data.kelasDiampu.map(_t).filter(Boolean) : [];
  if (!nama) throw new Error('Nama pengampu wajib diisi!');
  if (!username || username.length < 3) throw new Error('Username minimal 3 karakter!');
  if (!password || password.length < 4) throw new Error('Password minimal 4 karakter!');

  _ensureAuthSchema();
  if (_findUserByUsername(username)) throw new Error('Username "' + username + '" sudah dipakai!');

  const sheet = _sheet(SHEET_NAMES.AUTH);
  sheet.appendRow([_newId('U'), username, password, 'Pengampu', nama, nip, noHp, kelasDiampu.join(','), 'Ya']);
  _logActivity('Tambah Pengampu', nama + ' (' + username + ')');
  return 'Pengampu "' + nama + '" berhasil ditambahkan!';
}

function editPengampu(id, data) {
  const nama = _t(data && data.nama);
  const nip = _t(data && data.nip);
  const noHp = _t(data && data.noHp);
  const username = _t(data && data.username);
  const password = _t(data && data.password);  // opsional; kosong = tidak ganti
  const kelasDiampu = Array.isArray(data && data.kelasDiampu) ? data.kelasDiampu.map(_t).filter(Boolean) : [];
  const active = data && data.active === false ? 'Tidak' : 'Ya';

  if (!nama) throw new Error('Nama pengampu wajib diisi!');
  if (!username || username.length < 3) throw new Error('Username minimal 3 karakter!');

  const found = _findRowById(SHEET_NAMES.AUTH, id);
  if (!found) throw new Error('Pengampu tidak ditemukan!');
  if (_t(found.values[3]).toLowerCase() !== 'pengampu') throw new Error('User ini bukan pengampu.');

  // Cek duplikat username selain dirinya
  const dup = _findUserByUsername(username);
  if (dup && _t(dup.values[0]) !== _t(id)) throw new Error('Username "' + username + '" sudah dipakai user lain!');

  const newPass = password ? password : _t(found.values[2]);
  const sheet = _sheet(SHEET_NAMES.AUTH);
  sheet.getRange(found.rowIndex, 1, 1, AUTH_HEADERS.length).setNumberFormat('@');
  sheet.getRange(found.rowIndex, 1, 1, AUTH_HEADERS.length).setValues([[
    _t(found.values[0]), username, newPass, 'Pengampu', nama, nip, noHp, kelasDiampu.join(','), active
  ]]);
  _logActivity('Edit Pengampu', nama + ' (' + username + ')');
  return 'Data pengampu diperbarui!';
}

function hapusPengampu(id) {
  const found = _findRowById(SHEET_NAMES.AUTH, id);
  if (!found) throw new Error('Pengampu tidak ditemukan!');
  if (_t(found.values[3]).toLowerCase() === 'admin') throw new Error('Akun Admin tidak boleh dihapus!');
  _sheet(SHEET_NAMES.AUTH).deleteRow(found.rowIndex);
  _logActivity('Hapus Pengampu', _t(found.values[4]) + ' (' + _t(found.values[1]) + ')');
  return 'Pengampu dihapus!';
}

// Wrapper cepat untuk UI
function simpanPengampuCepat(data) {
  return { message: simpanPengampu(data), pengampu: getPengampuData() };
}
function editPengampuCepat(id, data) {
  return { message: editPengampu(id, data), pengampu: getPengampuData() };
}
function hapusPengampuCepat(id) {
  return { message: hapusPengampu(id), pengampu: getPengampuData() };
}

// ==========================================
// 20. ABSENSI PENGAMPU
// Sheet AbsensiPengampu: [ID, Tanggal, Username, Nama, JamMasuk, JamPulang, Status, Keterangan]
// - Pengampu: self check-in / check-out (auto Hadir)
// - Admin: set status harian (Hadir/Sakit/Izin/Alpa) + keterangan
// ==========================================
function _ensureAbsensiPengampuSheet() {
  const ss = _ss();
  let sheet = ss.getSheetByName(SHEET_NAMES.ABSENSI_PENGAMPU);
  // Migrasi ejaan lama (nama sheet + 'h'): rename, atau gabungkan bila sheet baru telanjur dibuat.
  const legacy = ss.getSheetByName(SHEET_NAMES.ABSENSI_PENGAMPU + 'h');
  if (legacy) {
    if (!sheet) { legacy.setName(SHEET_NAMES.ABSENSI_PENGAMPU); sheet = legacy; }
    else {
      if (legacy.getLastRow() >= 2) {
        const oldRows = legacy.getRange(2, 1, legacy.getLastRow() - 1, ABSENSI_PENGAMPU_HEADERS.length).getValues();
        const startRow = Math.max(sheet.getLastRow(), 1) + 1;
        sheet.getRange(startRow, 1, oldRows.length, ABSENSI_PENGAMPU_HEADERS.length).setValues(oldRows);
      }
      ss.deleteSheet(legacy);
    }
  }
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.ABSENSI_PENGAMPU);
  sheet.getRange(1, 1, 1, ABSENSI_PENGAMPU_HEADERS.length).setValues([ABSENSI_PENGAMPU_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, ABSENSI_PENGAMPU_HEADERS.length).setNumberFormat('@');
  return sheet;
}

function _findAbsensiPengampu(tanggal, username) {
  const sheet = _ensureAbsensiPengampuSheet();
  if (sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ABSENSI_PENGAMPU_HEADERS.length).getDisplayValues();
  const t = _t(tanggal), u = _t(username).toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    if (_t(rows[i][1]) === t && _t(rows[i][2]).toLowerCase() === u) {
      return { rowIndex: i + 2, values: rows[i] };
    }
  }
  return null;
}

function _todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Format 'yyyy-MM-dd' -> 'Rabu, 06/08/2026' (untuk tampilan cetak)
function _formatTanggalId(s) {
  s = _t(s);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return s;
  var d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  var hari = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  return hari[d.getDay()] + ', ' + m[3] + '/' + m[2] + '/' + m[1];
}

function _nowTime() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
}

// Self check-in: pengampu menekan tombol "Absen Masuk"
function checkInPengampu(username) {
  const u = _t(username);
  if (!u) throw new Error('Username wajib diisi!');
  const found = _findUserByUsername(u);
  if (!found) throw new Error('User tidak ditemukan!');
  const nama = _t(found.values[4]) || u;
  const tgl = _todayStr();
  const jam = _nowTime();

  return _withLock(function() {
    const sheet = _ensureAbsensiPengampuSheet();
    const rec = _findAbsensiPengampu(tgl, u);
    if (rec) {
      if (_t(rec.values[4])) throw new Error('Anda sudah check-in hari ini pukul ' + _t(rec.values[4]) + '.');
      sheet.getRange(rec.rowIndex, 5).setValue(jam);
      sheet.getRange(rec.rowIndex, 7).setValue('Hadir');
    } else {
      sheet.appendRow([_newId('AP'), tgl, u, nama, jam, '', 'Hadir', '']);
    }
    _logActivity('Check-in Pengampu', nama + ' pukul ' + jam);
    return 'Check-in berhasil pada pukul ' + jam;
  });
}

// Self check-out: pengampu menekan tombol "Absen Pulang"
function checkOutPengampu(username) {
  const u = _t(username);
  if (!u) throw new Error('Username wajib diisi!');
  const tgl = _todayStr();
  const jam = _nowTime();
  return _withLock(function() {
    const sheet = _ensureAbsensiPengampuSheet();
    const rec = _findAbsensiPengampu(tgl, u);
    if (!rec) throw new Error('Anda belum check-in hari ini.');
    if (_t(rec.values[5])) throw new Error('Anda sudah check-out hari ini pukul ' + _t(rec.values[5]) + '.');
    sheet.getRange(rec.rowIndex, 6).setValue(jam);
    _logActivity('Check-out Pengampu', _t(rec.values[3]) + ' pukul ' + jam);
    return 'Check-out berhasil pada pukul ' + jam;
  });
}

// Set status harian (Admin) atau update keterangan
function setStatusAbsensiPengampu(data) {
  const tgl = _t(data && data.tanggal);
  const username = _t(data && data.username);
  const status = _t(data && data.status);
  const keterangan = _t(data && data.keterangan);
  if (!_isValidDate(tgl)) throw new Error('Tanggal tidak valid!');
  if (!username) throw new Error('Username wajib diisi!');
  const validStatus = { 'Hadir': true, 'Sakit': true, 'Izin': true, 'Alpa': true, 'Terlambat': true };
  if (!validStatus[status]) throw new Error('Status tidak valid!');
  const found = _findUserByUsername(username);
  if (!found) throw new Error('User tidak ditemukan!');
  const nama = _t(found.values[4]) || username;

  return _withLock(function() {
    const sheet = _ensureAbsensiPengampuSheet();
    const rec = _findAbsensiPengampu(tgl, username);
    if (rec) {
      sheet.getRange(rec.rowIndex, 7).setValue(status);
      sheet.getRange(rec.rowIndex, 8).setValue(keterangan);
    } else {
      sheet.appendRow([_newId('AP'), tgl, username, nama, '', '', status, keterangan]);
    }
    _logActivity('Set Status Pengampu', nama + ' - ' + tgl + ' - ' + status);
    return 'Status absensi pengampu disimpan!';
  });
}

// Lapor ketidakhadiran oleh pengampu sendiri (langsung tercatat; admin tetap bisa edit/hapus).
function laporKetidakhadiranPengampu(username, status, keterangan) {
  const u = _t(username);
  const st = _t(status);
  const ket = _t(keterangan);
  if (!u) throw new Error('Username wajib diisi!');
  const validStatus = { 'Sakit': true, 'Izin': true, 'Alpa': true, 'Terlambat': true };
  if (!validStatus[st]) throw new Error('Status tidak valid! Pilih Sakit, Izin, Alpa, atau Terlambat.');
  const found = _findUserByUsername(u);
  if (!found) throw new Error('User tidak ditemukan!');
  const nama = _t(found.values[4]) || u;
  const tgl = _todayStr();

  return _withLock(function() {
    const sheet = _ensureAbsensiPengampuSheet();
    const rec = _findAbsensiPengampu(tgl, u);
    if (rec) {
      sheet.getRange(rec.rowIndex, 7).setValue(st);
      sheet.getRange(rec.rowIndex, 8).setValue(ket);
    } else {
      sheet.appendRow([_newId('AP'), tgl, u, nama, '', '', st, ket]);
    }
    _logActivity('Lapor Ketidakhadiran', nama + ' - ' + tgl + ' - ' + st);
    return 'Laporan ketidakhadiran (' + st + ') tercatat untuk tanggal ' + tgl + '.';
  });
}

// Get data absensi. Untuk pengampu: kirim usernameFilter untuk data sendiri.
function getAbsensiPengampuData(tglMulai, tglAkhir, usernameFilter) {
  _ensureAbsensiPengampuSheet();
  const rows = _readSheet(SHEET_NAMES.ABSENSI_PENGAMPU);
  const tm = _t(tglMulai), ta = _t(tglAkhir), uf = _t(usernameFilter).toLowerCase();
  return rows.filter(r => {
    const t = _t(r[1]);
    if (tm && t < tm) return false;
    if (ta && t > ta) return false;
    if (uf && _t(r[2]).toLowerCase() !== uf) return false;
    return true;
  }).map(r => ({
    id: _t(r[0]), tanggal: _t(r[1]), username: _t(r[2]), nama: _t(r[3]),
    jamMasuk: _t(r[4]), jamPulang: _t(r[5]), status: _t(r[6]) || '-', keterangan: _t(r[7])
  })).sort((a, b) => (b.tanggal + b.nama).localeCompare(a.tanggal + a.nama));
}

// Status hari ini untuk pengampu yang sedang login
function getAbsensiPengampuHariIni(username) {
  const rec = _findAbsensiPengampu(_todayStr(), username);
  if (!rec) return { tanggal: _todayStr(), jamMasuk: '', jamPulang: '', status: '', keterangan: '' };
  return {
    tanggal: _t(rec.values[1]),
    jamMasuk: _t(rec.values[4]),
    jamPulang: _t(rec.values[5]),
    status: _t(rec.values[6]),
    keterangan: _t(rec.values[7])
  };
}

function hapusAbsensiPengampu(id) {
  const found = _findRowById(SHEET_NAMES.ABSENSI_PENGAMPU, id);
  if (!found) throw new Error('Data tidak ditemukan!');
  _sheet(SHEET_NAMES.ABSENSI_PENGAMPU).deleteRow(found.rowIndex);
  _logActivity('Hapus Absensi Pengampu', _t(found.values[3]) + ' - ' + _t(found.values[1]));
  return 'Data absensi pengampu dihapus!';
}

// ==========================================
// 21. GUARD PRINT UNTUK ROLE PENGAMPU
// UI seharusnya menyembunyikan tombol print untuk Pengampu,
// tapi server juga memvalidasi untuk keamanan.
// ==========================================
function prosesRekapCetakGuarded(role, kelas, jenis, tglMulai, tglAkhir, mode, bulan, tahun, mapel) {
  if (_t(role).toLowerCase() === 'pengampu') {
    throw new Error('Akses ditolak: Pengampu tidak diizinkan mencetak laporan.');
  }
  return prosesRekapCetak(kelas, jenis, tglMulai, tglAkhir, mode, bulan, tahun, mapel);
}

function exportRekapCSVGuarded(role, kelas, jenis, tglMulai, tglAkhir, mode, bulan, tahun, mapel) {
  if (_t(role).toLowerCase() === 'pengampu') {
    throw new Error('Akses ditolak: Pengampu tidak diizinkan mengekspor laporan.');
  }
  return exportRekapCSV(kelas, jenis, tglMulai, tglAkhir, mode, bulan, tahun, mapel);
}

// Endpoint bantuan untuk dropdown UI
function getKelasList() {
  return _readSheet(SHEET_NAMES.KELAS).map(r => ({ id: _t(r[0]), nama: _t(r[1]) }));
}

// ==========================================
// 22. ROLE-BASED FILTERING (Kelas milik Pengampu)
// Endpoint baru yang menerima userCtx = { role, username, kelasDiampu:[] }
// Digunakan oleh client saat role=Pengampu; admin tetap pakai endpoint asli.
// ==========================================
function _kelasSetFromCtx(userCtx) {
  if (!userCtx || String(userCtx.role||'').toLowerCase() === 'admin') return null; // null = akses semua
  const arr = Array.isArray(userCtx.kelasDiampu) ? userCtx.kelasDiampu : [];
  return new Set(arr.map(function(k){ return String(k||'').trim(); }).filter(Boolean));
}

// Scope khusus PERINGATAN: Admin = semua kelas; Moderator = kelas wilayahnya;
// Pengampu = kelasnya sendiri + kelas lain dalam KELOMPOK yang sama (read-only,
// tujuan saling mengingatkan). Aksi dismiss tetap hanya Admin/Moderator.
function _kelasSetPeringatan(userCtx) {
  const set = _kelasSetFromCtx(userCtx);
  if (!set) return null;
  if (_fbRole(userCtx) !== 'pengampu') return set;
  const out = {};
  set.forEach(function(k){ out[k] = true; });
  getKelompokKelas().forEach(function (g) {
    const intersect = (g.kelas || []).some(function (k) { return set.has(k); });
    if (intersect) (g.kelas || []).forEach(function (k) { out[k] = true; });
  });
  return new Set(Object.keys(out));
}

function getSiswaDataForUser(userCtx) {
  const rows = _sortSiswaRows(_readSheet(SHEET_NAMES.SISWA));
  const set = _kelasSetFromCtx(userCtx);
  if (!set) return rows;
  return rows.filter(function(r){ return set.has(String(r[3]||'').trim()); });
}

function getJadwalDataForUser(userCtx) {
  const rows = _readSheet(SHEET_NAMES.JADWAL);
  const set = _kelasSetFromCtx(userCtx);
  if (!set) return rows;
  // Jadwal: [ID, Hari, JamMulai, JamSelesai, Mapel, Kelas]
  return rows.filter(function(r){ return set.has(String(r[5]||'').trim()); });
}

// Membungkus getAllInitialData supaya siswa & jadwal juga sudah difilter untuk pengampu.
function getAllInitialDataForUser(userCtx) {
  const base = getAllInitialData();
  const set = _kelasSetFromCtx(userCtx);
  if (!set) return base;
  return {
    kelas: (base.kelas||[]).filter(function(r){ return set.has(String(r[1]||'').trim()); }),
    mapel: base.mapel,
    setting: base.setting,
    appInfo: base.appInfo,
    pengampu: []  // pengampu tidak butuh daftar pengampu lain
  };
}

// ==========================================
// 23. REKAP CETAK ABSENSI PENGAMPU
// ==========================================
function cetakRekapAbsensiPengampu(tglMulai, tglAkhir, usernameFilter) {
  const data = getAbsensiPengampuData(tglMulai, tglAkhir, usernameFilter || '');
  const kop = _getSettingValues();
  const namaSekolah = kop[0] || DEFAULT_APP_NAME;
  const alamat = kop[1] || '';
  const kepsek = kop[4] || '';
  const nipKepsek = kop[5] || '';
  const logoUrl = kop[6] || '';
  const tgl = tglMulai + ' s.d. ' + tglAkhir;
  const total = data.length;
  const hitung = { Hadir:0, Sakit:0, Izin:0, Alpa:0, Terlambat:0 };
  data.forEach(d => { let s = _t(d.status); if (/^al/i.test(s)) s = 'Alpa'; if (hitung[s] !== undefined) hitung[s]++; });

  let rows = '';
  data.forEach((d,i) => {
    rows += '<tr><td>'+(i+1)+'</td><td>'+d.tanggal+'</td><td>'+d.nama+'</td><td>'+d.jamMasuk+'</td><td>'+d.jamPulang+'</td><td>'+d.status+'</td><td>'+d.keterangan+'</td></tr>';
  });

  const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Rekap Absensi Pengampu</title>' +
    '<style>body{font-family:Arial;padding:20px;color:#000;}h2,h3{margin:2px 0;text-align:center;}' +
    'table{width:100%;border-collapse:collapse;margin-top:15px;font-size:12px;}' +
    'th,td{border:1px solid #333;padding:6px 8px;text-align:left;}th{background:#eee;}' +
    '.kop{border-bottom:3px double #000;padding-bottom:10px;margin-bottom:15px;display:flex;align-items:center;gap:15px;}' +
    '.kop img{max-height:70px;}.kop .txt{flex:1;text-align:center;}' +
    '.ttd{margin-top:40px;text-align:right;padding-right:60px;}' +
    '.summary{margin-top:10px;font-size:13px;}' +
    '@media print{.no-print{display:none;}}</style></head><body>' +
    '<div class="kop">' +
    (logoUrl ? '<img src="'+logoUrl+'" alt="Logo">' : '') +
    '<div class="txt"><h2>'+namaSekolah+'</h2>'+(alamat?'<div>'+alamat+'</div>':'')+'</div></div>' +
    '<h3>REKAP ABSENSI PENGAMPU</h3>' +
    '<div class="summary"><b>Periode:</b> '+tgl+' | <b>Total Entri:</b> '+total+' | Hadir: '+hitung.Hadir+', Sakit: '+hitung.Sakit+', Izin: '+hitung.Izin+', Alpa: '+hitung.Alpa+', Terlambat: '+hitung.Terlambat+'</div>' +
    '<table><thead><tr><th>No</th><th>Tanggal</th><th>Nama</th><th>Jam Masuk</th><th>Jam Pulang</th><th>Status</th><th>Keterangan</th></tr></thead>' +
    '<tbody>'+(rows || '<tr><td colspan="7" style="text-align:center;">Tidak ada data</td></tr>')+'</tbody></table>' +
    '<div class="ttd">Kepala Sekolah,<br><br><br><b><u>'+kepsek+'</u></b><br>NIP. '+nipKepsek+'</div>' +
    '<div class="no-print" style="margin-top:20px;text-align:center;"><button onclick="window.print()" style="padding:8px 20px;font-size:14px;">Cetak / Simpan PDF</button></div>' +
    '<script>window.addEventListener("load",function(){setTimeout(function(){window.print();},500);});<\/script>' +
    '</body></html>';
  return html;
}

// Guard: hanya admin yang boleh mencetak.
function cetakRekapAbsensiPengampuGuarded(role, tglMulai, tglAkhir, usernameFilter) {
  if (_t(role).toLowerCase() !== 'admin') throw new Error('Akses ditolak: hanya Admin yang bisa mencetak rekap absensi pengampu.');
  return cetakRekapAbsensiPengampu(tglMulai, tglAkhir, usernameFilter);
}

// ==========================================
// 24. BUKU MINGGUAN SISWA
// Sheet BukuMingguan: [ID, Minggu, Tanggal, SiswaID, Nama, Kelas, Status, JenisIbadah]
// ==========================================
function _ensureBukuMingguanSheet() {
  const ss = _ss();
  let s = ss.getSheetByName('BukuMingguan');
  if (!s) s = ss.insertSheet('BukuMingguan');
  const headers = ['ID','Minggu','Tanggal','SiswaID','Nama','Kelas','Status','JenisIbadah'];
  s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  s.setFrozenRows(1);
  if (s.getMaxRows()>1) s.getRange(2,1,s.getMaxRows()-1,headers.length).setNumberFormat('@');
  return s;
}

function getBukuMingguanData(minggu, kelas, jenisIbadah, userCtx) {
  _ensureBukuMingguanSheet();
  const rows = _readSheet('BukuMingguan');
  const m = _t(minggu), k = _t(kelas), ji = _t(jenisIbadah);
  const set = _kelasSetFromCtx(userCtx);
  return rows.filter(function(r){
    if (m && _t(r[1]) !== m) return false;
    if (k && _t(r[5]) !== k) return false;
    if (ji && _t(r[7]) !== ji) return false;
    if (set && !set.has(_t(r[5]))) return false;
    return true;
  }).map(function(r){
    return { id:_t(r[0]), minggu:_t(r[1]), tanggal:_t(r[2]), siswaId:_t(r[3]),
             nama:_t(r[4]), kelas:_t(r[5]), status:_t(r[6]), jenisIbadah:_t(r[7]) };
  });
}

function simpanBukuMingguan(data) {
  const minggu = _t(data.minggu);
  const kelas = _t(data.kelas);
  const jenisIbadah = _t(data.jenisIbadah) || 'Ibadah Minggu';
  const tanggal = _t(data.tanggal) || _todayStr();
  const entries = Array.isArray(data.entries) ? data.entries : [];
  if (!minggu || !kelas) throw new Error('Minggu dan kelas wajib diisi!');
  return _withLock(function(){
    const sheet = _ensureBukuMingguanSheet();
    // Hapus data lama untuk minggu+kelas+jenis ibadah ini
    if (sheet.getLastRow() >= 2) {
      const all = sheet.getRange(2,1,sheet.getLastRow()-1,8).getDisplayValues();
      for (let i = all.length - 1; i >= 0; i--) {
        if (_t(all[i][1]) === minggu && _t(all[i][5]) === kelas && _t(all[i][7]) === jenisIbadah) sheet.deleteRow(i+2);
      }
    }
    entries.forEach(function(e){
      const st = _t(e.status);
      if (st !== 'Kumpul' && st !== 'Tidak Kumpul') return;
      sheet.appendRow([_newId('BM'), minggu, tanggal, _t(e.siswaId), _t(e.nama), kelas, st, jenisIbadah]);
    });
    _logActivity('Absen Buku Mingguan', kelas + ' - ' + minggu + ' - ' + jenisIbadah + ' (' + entries.length + ' siswa)', _t(data.actor));
    return 'Absen buku mingguan tersimpan!';
  });
}

function cetakRekapBukuMingguan(role, minggu, kelas, jenisIbadah) {
  if (_t(role).toLowerCase() !== 'admin') throw new Error('Hanya Admin yang bisa mencetak.');
  const ji = _t(jenisIbadah);
  const data = getBukuMingguanData(minggu, kelas, ji, null);
  const kop = _getSettingValues();
  const logo = kop[6] || '';
  let rows = ''; let kumpul=0, tdk=0;
  data.forEach(function(d,i){
    if (d.status==='Kumpul') kumpul++; else if (d.status==='Tidak Kumpul') tdk++;
    rows += '<tr><td>'+(i+1)+'</td><td>'+_esc(_formatTanggalId(d.tanggal))+'</td><td>'+_esc(d.nama)+'</td><td>'+_esc(d.status)+'</td></tr>';
  });
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Rekap Buku Mingguan</title>'+
    '<style>body{font-family:Arial;padding:20px;}table{width:100%;border-collapse:collapse;font-size:12px;}th,td{border:1px solid #333;padding:6px;}th{background:#eee;}h2,h3{text-align:center;margin:4px 0;}.kop{border-bottom:3px double #000;padding-bottom:10px;text-align:center;}@media print{.np{display:none;}}</style></head><body>'+
    '<div class="kop">'+(logo?'<img src="'+logo+'" style="max-height:70px;">':'')+'<h2>'+(kop[0]||DEFAULT_APP_NAME)+'</h2>'+(kop[1]?'<div>'+kop[1]+'</div>':'')+'</div>'+
    '<h3>REKAP ABSEN BUKU MINGGUAN</h3><p><b>Kelas:</b> '+_esc(kelas)+' | <b>Jenis Ibadah:</b> '+_esc(ji||'-')+' | <b>Tanggal:</b> '+_esc(_formatTanggalId(minggu))+' | <b>Total:</b> '+data.length+' | <b>Kumpul:</b> '+kumpul+' | <b>Tidak Kumpul:</b> '+tdk+'</p>'+
    '<table><thead><tr><th>No</th><th>Tanggal</th><th>Nama Siswa</th><th>Status</th></tr></thead><tbody>'+(rows||'<tr><td colspan=4 style="text-align:center">Tidak ada data</td></tr>')+'</tbody></table>'+
    '<div style="margin-top:40px;text-align:right;padding-right:60px;">Kepala Sekolah,<br><br><br><b><u>'+(kop[4]||'')+'</u></b><br>NIP. '+(kop[5]||'')+'</div>'+
    '<div class="np" style="text-align:center;margin-top:20px;"><button onclick="window.print()">Cetak</button></div>'+
    '<script>window.addEventListener("load",function(){setTimeout(function(){window.print();},500);});<\/script></body></html>';
}

// ==========================================
// 25. HARI LIBUR + PERINGATAN ABSEN MINGGU BERJALAN
// Sheet HariLibur: [ID, Tanggal, Keterangan]
// ==========================================
function _ensureHariLiburSheet() {
  const ss = _ss();
  let s = ss.getSheetByName('HariLibur');
  if (!s) s = ss.insertSheet('HariLibur');
  const headers = ['ID','Tanggal','Keterangan'];
  s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  s.setFrozenRows(1);
  if (s.getMaxRows()>1) s.getRange(2,1,s.getMaxRows()-1,headers.length).setNumberFormat('@');
  return s;
}

function getHariLibur() {
  _ensureHariLiburSheet();
  return _readSheet('HariLibur').map(function(r){
    return { id:_t(r[0]), tanggal:_t(r[1]), keterangan:_t(r[2]) };
  }).sort(function(a,b){ return a.tanggal < b.tanggal ? 1 : -1; });
}

function tambahHariLibur(role, tanggal, keterangan) {
  if (_t(role).toLowerCase() !== 'admin') throw new Error('Hanya Admin yang bisa menandai hari libur.');
  const tgl = _t(tanggal);
  if (!_isValidDate(tgl)) throw new Error('Tanggal tidak valid!');
  return _withLock(function(){
    const sheet = _ensureHariLiburSheet();
    const exists = _readSheet('HariLibur').some(function(r){ return _t(r[1]) === tgl; });
    if (exists) throw new Error('Tanggal ' + tgl + ' sudah ditandai libur.');
    sheet.appendRow([_newId('HL'), tgl, _t(keterangan)]);
    _logActivity('Tandai Hari Libur', tgl + (_t(keterangan) ? (' - ' + _t(keterangan)) : ''), 'admin');
    return 'Hari libur ' + tgl + ' ditambahkan.';
  });
}

function hapusHariLibur(role, id) {
  if (_t(role).toLowerCase() !== 'admin') throw new Error('Hanya Admin yang bisa menghapus hari libur.');
  return _withLock(function(){
    const found = _findRowById('HariLibur', id);
    if (!found) throw new Error('Data tidak ditemukan!');
    _sheet('HariLibur').deleteRow(found.rowIndex);
    _logActivity('Hapus Hari Libur', _t(found.values[1]), 'admin');
    return 'Hari libur dihapus.';
  });
}

// Kelas yang belum mengisi absen siswa sama sekali pada minggu berjalan (Senin-Minggu),
// mengecualikan hari libur. Jika seluruh hari minggu ini libur -> tidak ada peringatan.
function getKelasBelumAbsen(userCtx) {
  const range = _currentWeekRange();
  _ensureHariLiburSheet();
  const liburSet = {};
  _readSheet('HariLibur').forEach(function(r){ var t=_t(r[1]); if(t) liburSet[t]=true; });

  var workingDays = 0;
  var d = new Date(range.start + 'T00:00:00');
  var end = new Date(range.end + 'T00:00:00');
  for (; d <= end; d.setDate(d.getDate()+1)) {
    var ds = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!liburSet[ds]) workingDays++;
  }
  if (workingDays === 0) {
    return { start: range.start, end: range.end, libur: true, kelas: [] };
  }

  const set = _kelasSetFromCtx(userCtx);
  var kelasList = _readSheet(SHEET_NAMES.KELAS).map(function(r){ return _t(r[1]); }).filter(Boolean);
  if (set) kelasList = kelasList.filter(function(k){ return set.has(k); });

  const c = _absensiCols();
  const kelasSudah = {};
  _readSheet(SHEET_NAMES.ABSENSI).forEach(function(r){
    const tgl = _t(r[c.tanggal]);
    if (tgl < range.start || tgl > range.end) return;
    if (liburSet[tgl]) return;
    kelasSudah[_t(r[c.kelas])] = true;
  });

  const belum = kelasList.filter(function(k){ return !kelasSudah[k]; });
  return { start: range.start, end: range.end, libur: false, kelas: belum };
}


// ##########################################################
// ##########  TAMBAHAN FITUR BARU (REV 8) — DIGABUNG  #######
// ##########################################################

/**
 * ============================================================
 * FITUR BARU (REV 8) — TAMBAHAN, TIDAK MENGUBAH Code.gs
 * ------------------------------------------------------------
 * Cukup buat file baru di editor Apps Script (mis. "FiturBaru.gs")
 * lalu tempel seluruh isi file ini. Semua helper (_t, _ss, _sheet,
 * _readSheet, _newId, _withLock, _logActivity, _findUserByUsername,
 * _absensiCols, _isValidDate, _ensureHariLiburSheet, _todayStr,
 * _currentWeekRange, SHEET_NAMES, dsb) dipakai dari Code.gs.
 *
 * Fitur:
 *  1. Kelompok Kelas (grup, mis. "Kelas 11" berisi 11A, 11B)
 *  2. Role Moderator (admin terbatas per kelompok kelas)
 *  3. Permintaan akses kelompok lain + persetujuan Admin berbatas waktu
 *  4. Pengumuman running-text (durasi diatur pemosting)
 *  5. Peringatan absen siswa PERSISTEN (per tanggal, sampai diisi)
 *  6. Peringatan pengampu belum absen/jurnal/buku mingguan
 *  7. Kontak WhatsApp pengampu (untuk tombol kirim WA di klien)
 * ============================================================
 */

var FB_SHEETS = {
  KELOMPOK: 'KelompokKelas',        // [ID, Nama, Keterangan, Kelas(csv)]
  MOD_MAP: 'ModeratorKelompok',     // [ID, Username, KelompokID]
  PENGUMUMAN: 'Pengumuman',         // [ID, Judul, Isi, Pembuat, Role, KelompokID, MulaiISO, SelesaiISO, Active, Warna]
  AKSES: 'AksesRequest',            // [ID, Username, KelompokID, Status, Menit, DibuatISO, DiprosesISO, ExpiresISO, Catatan]
  DISMISS: 'PeringatanSiswaDismiss', // [ID, Kelas, Nama, DismissedCount, By, At]
  DISMISS2: 'PeringatanDismiss',    // [ID, Topik, Kelas, Pengampu, Tanggal, ByRole, By, At]
  BAHAN: 'BahanAjar'                // [ID, Judul, Deskripsi, FileID, NamaFile, Ukuran, Pengunggah, Role, KelompokID, DiunggahISO]
};

function _fbIso(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}
function _fbParse(s) {
  s = _t(s);
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function _fbRole(userCtx) { return String(userCtx && userCtx.role || '').toLowerCase(); }
function _fbAssertAdmin(userCtx) {
  if (_fbRole(userCtx) !== 'admin') throw new Error('Akses ditolak: hanya Admin.');
}

// Pastikan seluruh sheet fitur baru ada beserta headernya.
function _ensureFiturSheets() {
  var ss = _ss();
  var defs = {};
  defs[FB_SHEETS.KELOMPOK] = ['ID', 'Nama', 'Keterangan', 'Kelas'];
  defs[FB_SHEETS.MOD_MAP] = ['ID', 'Username', 'KelompokID'];
  defs[FB_SHEETS.PENGUMUMAN] = ['ID', 'Judul', 'Isi', 'Pembuat', 'Role', 'KelompokID', 'MulaiISO', 'SelesaiISO', 'Active', 'Warna'];
  defs[FB_SHEETS.AKSES] = ['ID', 'Username', 'KelompokID', 'Status', 'Menit', 'DibuatISO', 'DiprosesISO', 'ExpiresISO', 'Catatan'];
  defs[FB_SHEETS.DISMISS] = ['ID', 'Kelas', 'Nama', 'DismissedCount', 'By', 'At'];
  defs[FB_SHEETS.DISMISS2] = ['ID', 'Topik', 'KelompokID', 'Pengampu', 'Tanggal', 'ByRole', 'By', 'At'];
  defs[FB_SHEETS.BAHAN] = ['ID', 'Judul', 'Deskripsi', 'FileID', 'NamaFile', 'Ukuran', 'Pengunggah', 'Role', 'KelompokID', 'DiunggahISO'];
  Object.keys(defs).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    var headers = defs[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length).setNumberFormat('@');
  });
}

// Jalankan sekali dari editor bila ingin menyiapkan sheet manual.
function setupFiturBaru() {
  _ensureFiturSheets();
  return 'Sheet fitur baru siap: ' + Object.keys(FB_SHEETS).map(function (k) { return FB_SHEETS[k]; }).join(', ');
}

// ============================================================
// 1. KELOMPOK KELAS
// ============================================================
function getKelompokKelas() {
  _ensureFiturSheets();
  return _readSheet(FB_SHEETS.KELOMPOK).map(function (r) {
    return {
      id: _t(r[0]),
      nama: _t(r[1]),
      keterangan: _t(r[2]),
      kelas: _t(r[3]).split(',').map(_t).filter(Boolean)
    };
  });
}

function _kelompokById(id) {
  var found = getKelompokKelas().filter(function (g) { return g.id === _t(id); });
  return found.length ? found[0] : null;
}

function simpanKelompok(userCtx, data) {
  _fbAssertAdmin(userCtx);
  var nama = _t(data && data.nama);
  if (!nama) throw new Error('Nama kelompok wajib diisi!');
  var kelas = Array.isArray(data && data.kelas) ? data.kelas.map(_t).filter(Boolean) : [];
  return _withLock(function () {
    _ensureFiturSheets();
    var dup = getKelompokKelas().some(function (g) { return g.nama.toLowerCase() === nama.toLowerCase(); });
    if (dup) throw new Error('Kelompok "' + nama + '" sudah ada!');
    _sheet(FB_SHEETS.KELOMPOK).appendRow([_newId('KG'), nama, _t(data && data.keterangan), kelas.join(',')]);
    _logActivity('Tambah Kelompok Kelas', nama + ' (' + kelas.length + ' kelas)', userCtx && userCtx.username);
    return 'Kelompok "' + nama + '" berhasil dibuat!';
  });
}

function editKelompok(userCtx, id, data) {
  _fbAssertAdmin(userCtx);
  var nama = _t(data && data.nama);
  if (!nama) throw new Error('Nama kelompok wajib diisi!');
  var kelas = Array.isArray(data && data.kelas) ? data.kelas.map(_t).filter(Boolean) : [];
  return _withLock(function () {
    var found = _findRowById(FB_SHEETS.KELOMPOK, id);
    if (!found) throw new Error('Kelompok tidak ditemukan!');
    var dup = getKelompokKelas().some(function (g) { return g.id !== _t(id) && g.nama.toLowerCase() === nama.toLowerCase(); });
    if (dup) throw new Error('Nama kelompok "' + nama + '" sudah dipakai!');
    _sheet(FB_SHEETS.KELOMPOK).getRange(found.rowIndex, 2, 1, 3).setValues([[nama, _t(data && data.keterangan), kelas.join(',')]]);
    _logActivity('Edit Kelompok Kelas', nama, userCtx && userCtx.username);
    return 'Kelompok diperbarui!';
  });
}

function hapusKelompok(userCtx, id) {
  _fbAssertAdmin(userCtx);
  return _withLock(function () {
    var found = _findRowById(FB_SHEETS.KELOMPOK, id);
    if (!found) throw new Error('Kelompok tidak ditemukan!');
    _sheet(FB_SHEETS.KELOMPOK).deleteRow(found.rowIndex);
    // Bersihkan penugasan moderator ke kelompok ini
    var mapSheet = _sheet(FB_SHEETS.MOD_MAP);
    var rows = _readSheet(FB_SHEETS.MOD_MAP);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (_t(rows[i][2]) === _t(id)) mapSheet.deleteRow(i + 2);
    }
    _logActivity('Hapus Kelompok Kelas', _t(found.values[1]), userCtx && userCtx.username);
    return 'Kelompok dihapus!';
  });
}

// Auto-buat kelompok berdasarkan pola nama kelas (angka tingkat, mis. 11A/11B -> "Kelas 11").
function autoKelompokDariPola(userCtx) {
  _fbAssertAdmin(userCtx);
  return _withLock(function () {
    _ensureFiturSheets();
    var kelasList = _readSheet(SHEET_NAMES.KELAS).map(function (r) { return _t(r[1]); }).filter(Boolean);
    var groups = {}; // key tingkat -> [kelas]
    kelasList.forEach(function (k) {
      var m = k.match(/(\d+)/);
      var key = m ? m[1] : k.replace(/[^A-Za-z]/g, '').toUpperCase();
      if (!key) key = k;
      if (!groups[key]) groups[key] = [];
      groups[key].push(k);
    });
    var existing = getKelompokKelas();
    var existingNama = {};
    existing.forEach(function (g) { existingNama[g.nama.toLowerCase()] = g; });
    var dibuat = 0, diupdate = 0;
    var sheet = _sheet(FB_SHEETS.KELOMPOK);
    Object.keys(groups).forEach(function (key) {
      var nama = /^\d+$/.test(key) ? ('Kelas ' + key) : ('Kelompok ' + key);
      var anggota = groups[key];
      var ex = existingNama[nama.toLowerCase()];
      if (ex) {
        var merged = {};
        ex.kelas.concat(anggota).forEach(function (x) { merged[x] = true; });
        var newKelas = Object.keys(merged);
        if (newKelas.length !== ex.kelas.length) {
          var f = _findRowById(FB_SHEETS.KELOMPOK, ex.id);
          if (f) { sheet.getRange(f.rowIndex, 4).setValue(newKelas.join(',')); diupdate++; }
        }
      } else {
        sheet.appendRow([_newId('KG'), nama, 'Dibuat otomatis dari pola nama', anggota.join(',')]);
        dibuat++;
      }
    });
    _logActivity('Auto Kelompok Kelas', dibuat + ' dibuat, ' + diupdate + ' diperbarui', userCtx && userCtx.username);
    return 'Selesai. ' + dibuat + ' kelompok dibuat, ' + diupdate + ' diperbarui.';
  });
}

// ============================================================
// 2. MODERATOR (disimpan di sheet Auth, Role='Moderator')
// ============================================================
function _getModeratorGroupIds(username) {
  var u = _t(username).toLowerCase();
  return _readSheet(FB_SHEETS.MOD_MAP)
    .filter(function (r) { return _t(r[1]).toLowerCase() === u; })
    .map(function (r) { return _t(r[2]); })
    .filter(Boolean);
}

function _setModeratorGroups(username, kelompokIds) {
  var u = _t(username);
  var sheet = _sheet(FB_SHEETS.MOD_MAP);
  var rows = _readSheet(FB_SHEETS.MOD_MAP);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (_t(rows[i][1]).toLowerCase() === u.toLowerCase()) sheet.deleteRow(i + 2);
  }
  (kelompokIds || []).map(_t).filter(Boolean).forEach(function (kid) {
    sheet.appendRow([_newId('MK'), u, kid]);
  });
}

function getModeratorData() {
  _ensureFiturSheets();
  var groups = {};
  getKelompokKelas().forEach(function (g) { groups[g.id] = g; });
  return _readSheet(SHEET_NAMES.AUTH)
    .filter(function (r) { return _t(r[3]).toLowerCase() === 'moderator'; })
    .map(function (r) {
      var uname = _t(r[1]);
      var gids = _getModeratorGroupIds(uname);
      return {
        id: _t(r[0]),
        username: uname,
        nama: _t(r[4]),
        nip: _t(r[5]),
        noHp: _t(r[6]),
        active: (_t(r[8]).toLowerCase() !== 'tidak'),
        kelompokIds: gids,
        kelompokNama: gids.map(function (id) { return (groups[id] || {}).nama || id; })
      };
    });
}

function simpanModerator(userCtx, data) {
  _fbAssertAdmin(userCtx);
  var nama = _t(data && data.nama);
  var username = _t(data && data.username);
  var password = _t(data && data.password);
  var kelompokIds = Array.isArray(data && data.kelompokIds) ? data.kelompokIds.map(_t).filter(Boolean) : [];
  if (!nama) throw new Error('Nama moderator wajib diisi!');
  if (!username || username.length < 3) throw new Error('Username minimal 3 karakter!');
  if (!password || password.length < 4) throw new Error('Password minimal 4 karakter!');
  return _withLock(function () {
    _ensureAuthSchema();
    _ensureFiturSheets();
    if (_findUserByUsername(username)) throw new Error('Username "' + username + '" sudah dipakai!');
    _sheet(SHEET_NAMES.AUTH).appendRow([_newId('U'), username, password, 'Moderator', nama, _t(data && data.nip), _t(data && data.noHp), '', 'Ya']);
    _setModeratorGroups(username, kelompokIds);
    _logActivity('Tambah Moderator', nama + ' (' + username + ')', userCtx && userCtx.username);
    return 'Moderator "' + nama + '" berhasil ditambahkan!';
  });
}

function editModerator(userCtx, id, data) {
  _fbAssertAdmin(userCtx);
  var nama = _t(data && data.nama);
  var username = _t(data && data.username);
  var password = _t(data && data.password);
  var kelompokIds = Array.isArray(data && data.kelompokIds) ? data.kelompokIds.map(_t).filter(Boolean) : [];
  var active = data && data.active === false ? 'Tidak' : 'Ya';
  if (!nama) throw new Error('Nama moderator wajib diisi!');
  if (!username || username.length < 3) throw new Error('Username minimal 3 karakter!');
  return _withLock(function () {
    var found = _findRowById(SHEET_NAMES.AUTH, id);
    if (!found) throw new Error('Moderator tidak ditemukan!');
    if (_t(found.values[3]).toLowerCase() !== 'moderator') throw new Error('User ini bukan moderator.');
    var dup = _findUserByUsername(username);
    if (dup && _t(dup.values[0]) !== _t(id)) throw new Error('Username "' + username + '" sudah dipakai user lain!');
    var newPass = password ? password : _t(found.values[2]);
    var sheet = _sheet(SHEET_NAMES.AUTH);
    sheet.getRange(found.rowIndex, 1, 1, 9).setNumberFormat('@');
    sheet.getRange(found.rowIndex, 1, 1, 9).setValues([[
      _t(found.values[0]), username, newPass, 'Moderator', nama, _t(data && data.nip), _t(data && data.noHp), '', active
    ]]);
    _setModeratorGroups(username, kelompokIds);
    _logActivity('Edit Moderator', nama + ' (' + username + ')', userCtx && userCtx.username);
    return 'Data moderator diperbarui!';
  });
}

function hapusModerator(userCtx, id) {
  _fbAssertAdmin(userCtx);
  return _withLock(function () {
    var found = _findRowById(SHEET_NAMES.AUTH, id);
    if (!found) throw new Error('Moderator tidak ditemukan!');
    if (_t(found.values[3]).toLowerCase() !== 'moderator') throw new Error('User ini bukan moderator.');
    var uname = _t(found.values[1]);
    _sheet(SHEET_NAMES.AUTH).deleteRow(found.rowIndex);
    _setModeratorGroups(uname, []);
    _logActivity('Hapus Moderator', _t(found.values[4]) + ' (' + uname + ')', userCtx && userCtx.username);
    return 'Moderator dihapus!';
  });
}

// ============================================================
// 3. KONTEKS USER (dipanggil klien setelah login)
// Mengembalikan kelas efektif (untuk filter), kelompok, pengumuman aktif.
// ============================================================
function _kelasUntukUser(username) {
  var found = _findUserByUsername(username);
  if (!found) return { role: '', kelas: [] };
  var role = _t(found.values[3]);
  var roleLc = role.toLowerCase();
  if (roleLc === 'admin') return { role: role, kelas: [] }; // [] = semua (tak difilter)
  if (roleLc === 'moderator') {
    return { role: role, kelas: _effectiveKelasModerator(username).kelas };
  }
  // pengampu
  return { role: role, kelas: _t(found.values[7]).split(',').map(_t).filter(Boolean) };
}

function _effectiveKelasModerator(username) {
  _ensureFiturSheets();
  var groups = {};
  getKelompokKelas().forEach(function (g) { groups[g.id] = g; });
  var assigned = _getModeratorGroupIds(username);
  var now = new Date();
  var grantRows = _readSheet(FB_SHEETS.AKSES).filter(function (r) {
    if (_t(r[1]).toLowerCase() !== _t(username).toLowerCase()) return false;
    if (_t(r[3]) !== 'Approved') return false;
    var exp = _fbParse(r[7]);
    return exp && exp > now;
  });
  var grantIds = grantRows.map(function (r) { return _t(r[2]); });
  var kset = {};
  assigned.concat(grantIds).forEach(function (id) {
    if (groups[id]) groups[id].kelas.forEach(function (k) { kset[k] = true; });
  });
  return {
    kelas: Object.keys(kset),
    assignedGroups: assigned,
    grantedGroups: grantRows.map(function (r) {
      return { id: _t(r[2]), nama: (groups[_t(r[2])] || {}).nama || _t(r[2]), expires: _t(r[7]) };
    })
  };
}

function getFiturKonteks(username) {
  _ensureFiturSheets();
  var found = _findUserByUsername(username);
  var role = found ? _t(found.values[3]) : '';
  var roleLc = role.toLowerCase();
  var info = { role: role, username: _t(username), nama: found ? _t(found.values[4]) : _t(username) };
  var kelasEff = _kelasUntukUser(username);
  info.effectiveKelas = kelasEff.kelas;
  if (roleLc === 'moderator') {
    var em = _effectiveKelasModerator(username);
    info.managedGroupIds = em.assignedGroups;
    info.grantedGroups = em.grantedGroups;
    var gmap = {}; getKelompokKelas().forEach(function (g) { gmap[g.id] = g; });
    info.managedGroups = em.assignedGroups.map(function (id) { return gmap[id] || { id: id, nama: id, kelas: [] }; });
  } else {
    info.managedGroupIds = [];
    info.grantedGroups = [];
    info.managedGroups = [];
  }
  info.pengumuman = getPengumumanAktif({ role: role, username: _t(username), kelasDiampu: kelasEff.kelas });
  return info;
}

// ============================================================
// 4. PENGUMUMAN
// ============================================================
function getPengumumanAktif(userCtx) {
  _ensureFiturSheets();
  var now = new Date();
  var roleLc = _fbRole(userCtx);
  var userClasses = {};
  (userCtx && userCtx.kelasDiampu || []).forEach(function (k) { userClasses[_t(k)] = true; });
  var gmap = {}; getKelompokKelas().forEach(function (g) { gmap[g.id] = g; });
  return _readSheet(FB_SHEETS.PENGUMUMAN).filter(function (r) {
    if (_t(r[8]).toLowerCase() === 'tidak') return false;
    var mulai = _fbParse(r[6]); if (mulai && mulai > now) return false;
    var selesai = _fbParse(r[7]); if (selesai && selesai < now) return false;
    return true;
  }).filter(function (r) {
    var kelompokId = _t(r[5]);
    if (!kelompokId) return true;          // global
    if (roleLc === 'admin') return true;   // admin lihat semua
    var g = gmap[kelompokId];
    if (!g) return false;
    return g.kelas.some(function (k) { return userClasses[_t(k)]; });
  }).map(function (r) {
    return {
      id: _t(r[0]), judul: _t(r[1]), isi: _t(r[2]), pembuat: _t(r[3]), role: _t(r[4]),
      kelompokId: _t(r[5]), mulai: _t(r[6]), selesai: _t(r[7]), warna: _t(r[9]) || ''
    };
  });
}

function buatPengumuman(userCtx, data) {
  var roleLc = _fbRole(userCtx);
  if (roleLc !== 'admin' && roleLc !== 'moderator') throw new Error('Hanya Admin/Moderator yang dapat membuat pengumuman.');
  var isi = _t(data && data.isi);
  if (!isi) throw new Error('Isi pengumuman wajib diisi!');
  var menit = parseInt(data && data.durasiMenit, 10);
  if (!menit || menit <= 0) throw new Error('Durasi tampil (menit) wajib diisi dan lebih dari 0!');
  var kelompokId = _t(data && data.kelompokId);
  if (roleLc === 'moderator') {
    var mine = _getModeratorGroupIds(userCtx.username);
    if (!kelompokId) throw new Error('Moderator wajib memilih salah satu kelompok kelasnya.');
    if (mine.indexOf(kelompokId) < 0) throw new Error('Anda tidak mengelola kelompok tersebut.');
  }
  var now = new Date();
  var selesai = new Date(now.getTime() + menit * 60000);
  return _withLock(function () {
    _ensureFiturSheets();
    _sheet(FB_SHEETS.PENGUMUMAN).appendRow([
      _newId('PG'), _t(data && data.judul), isi, _t(userCtx.username), _t(userCtx.role),
      kelompokId, _fbIso(now), _fbIso(selesai), 'Ya', _t(data && data.warna)
    ]);
    _logActivity('Buat Pengumuman', (_t(data && data.judul) || isi.substring(0, 30)) + ' (' + menit + ' mnt)', userCtx.username);
    return 'Pengumuman dipublikasikan (tampil ' + menit + ' menit).';
  });
}

function getPengumumanKelola(userCtx) {
  _ensureFiturSheets();
  var roleLc = _fbRole(userCtx);
  var now = new Date();
  var gmap = {}; getKelompokKelas().forEach(function (g) { gmap[g.id] = g; });
  return _readSheet(FB_SHEETS.PENGUMUMAN).filter(function (r) {
    if (roleLc === 'admin') return true;
    return _t(r[3]).toLowerCase() === _t(userCtx.username).toLowerCase();
  }).map(function (r) {
    var selesai = _fbParse(r[7]);
    return {
      id: _t(r[0]), judul: _t(r[1]), isi: _t(r[2]), pembuat: _t(r[3]), role: _t(r[4]),
      kelompokId: _t(r[5]), kelompokNama: (gmap[_t(r[5])] || {}).nama || (_t(r[5]) ? _t(r[5]) : 'Semua (Global)'),
      mulai: _t(r[6]), selesai: _t(r[7]),
      aktif: (_t(r[8]).toLowerCase() !== 'tidak') && (!selesai || selesai >= now)
    };
  }).reverse();
}

function hapusPengumuman(userCtx, id) {
  var roleLc = _fbRole(userCtx);
  return _withLock(function () {
    var found = _findRowById(FB_SHEETS.PENGUMUMAN, id);
    if (!found) throw new Error('Pengumuman tidak ditemukan!');
    if (roleLc !== 'admin' && _t(found.values[3]).toLowerCase() !== _t(userCtx.username).toLowerCase()) {
      throw new Error('Anda hanya boleh menghapus pengumuman Anda sendiri.');
    }
    _sheet(FB_SHEETS.PENGUMUMAN).deleteRow(found.rowIndex);
    _logActivity('Hapus Pengumuman', _t(found.values[1]) || _t(id), userCtx.username);
    return 'Pengumuman dihapus.';
  });
}

// ============================================================
// 5. PERMINTAAN AKSES KELOMPOK LAIN (moderator -> admin)
// ============================================================
function buatPermintaanAkses(userCtx, kelompokId, menit, catatan) {
  if (_fbRole(userCtx) !== 'moderator') throw new Error('Hanya Moderator yang dapat mengajukan akses.');
  kelompokId = _t(kelompokId);
  if (!kelompokId) throw new Error('Pilih kelompok yang ingin diakses!');
  var m = parseInt(menit, 10);
  if (!m || m <= 0) throw new Error('Isi durasi akses (menit) yang diinginkan!');
  var mine = _getModeratorGroupIds(userCtx.username);
  if (mine.indexOf(kelompokId) >= 0) throw new Error('Kelompok ini sudah menjadi wilayah Anda.');
  return _withLock(function () {
    _ensureFiturSheets();
    var pending = _readSheet(FB_SHEETS.AKSES).some(function (r) {
      return _t(r[1]).toLowerCase() === _t(userCtx.username).toLowerCase() && _t(r[2]) === kelompokId && _t(r[3]) === 'Pending';
    });
    if (pending) throw new Error('Sudah ada permintaan Pending untuk kelompok ini.');
    _sheet(FB_SHEETS.AKSES).appendRow([
      _newId('AR'), _t(userCtx.username), kelompokId, 'Pending', String(m), _fbIso(new Date()), '', '', _t(catatan)
    ]);
    _logActivity('Ajukan Akses Kelompok', kelompokId + ' (' + m + ' mnt)', userCtx.username);
    return 'Permintaan akses terkirim. Menunggu persetujuan Admin.';
  });
}

function getPermintaanAkses(userCtx) {
  _ensureFiturSheets();
  var roleLc = _fbRole(userCtx);
  var now = new Date();
  var gmap = {}; getKelompokKelas().forEach(function (g) { gmap[g.id] = g; });
  var namaMap = {};
  getModeratorData().forEach(function (m) { namaMap[m.username.toLowerCase()] = m.nama; });
  return _readSheet(FB_SHEETS.AKSES).filter(function (r) {
    if (roleLc === 'admin') return true;
    return _t(r[1]).toLowerCase() === _t(userCtx.username).toLowerCase();
  }).map(function (r) {
    var exp = _fbParse(r[7]);
    var status = _t(r[3]);
    var aktif = status === 'Approved' && exp && exp > now;
    var sisaMenit = aktif ? Math.max(0, Math.round((exp - now) / 60000)) : 0;
    return {
      id: _t(r[0]), username: _t(r[1]), namaModerator: namaMap[_t(r[1]).toLowerCase()] || _t(r[1]),
      kelompokId: _t(r[2]), kelompokNama: (gmap[_t(r[2])] || {}).nama || _t(r[2]),
      status: status, menit: _t(r[4]), dibuat: _t(r[5]), diproses: _t(r[6]), expires: _t(r[7]),
      catatan: _t(r[8]), aktif: aktif, sisaMenit: sisaMenit
    };
  }).reverse();
}

function prosesPermintaanAkses(userCtx, id, aksi, menit) {
  _fbAssertAdmin(userCtx);
  aksi = _t(aksi).toLowerCase();
  return _withLock(function () {
    var found = _findRowById(FB_SHEETS.AKSES, id);
    if (!found) throw new Error('Permintaan tidak ditemukan!');
    var sheet = _sheet(FB_SHEETS.AKSES);
    var now = new Date();
    if (aksi === 'approve') {
      var m = parseInt(menit, 10) || parseInt(found.values[4], 10) || 60;
      var exp = new Date(now.getTime() + m * 60000);
      sheet.getRange(found.rowIndex, 4, 1, 5).setValues([['Approved', String(m), _t(found.values[5]), _fbIso(now), _fbIso(exp)]]);
      _logActivity('Setujui Akses', _t(found.values[1]) + ' -> ' + _t(found.values[2]) + ' (' + m + ' mnt)', userCtx.username);
      return 'Akses disetujui selama ' + m + ' menit.';
    } else if (aksi === 'reject') {
      sheet.getRange(found.rowIndex, 4).setValue('Rejected');
      sheet.getRange(found.rowIndex, 7).setValue(_fbIso(now));
      _logActivity('Tolak Akses', _t(found.values[1]) + ' -> ' + _t(found.values[2]), userCtx.username);
      return 'Permintaan ditolak.';
    } else if (aksi === 'revoke') {
      sheet.getRange(found.rowIndex, 8).setValue(_fbIso(now)); // set expired sekarang
      _logActivity('Cabut Akses', _t(found.values[1]) + ' -> ' + _t(found.values[2]), userCtx.username);
      return 'Akses dicabut.';
    }
    throw new Error('Aksi tidak dikenal.');
  });
}

// Notifikasi ringkas untuk moderator: akses yang baru disetujui & masih aktif.
function getNotifikasiAkses(username) {
  var ctx = { role: 'moderator', username: _t(username) };
  return getPermintaanAkses(ctx).filter(function (r) { return r.aktif; });
}

// ============================================================
// 6. PENGAMPU DALAM WILAYAH MODERATOR (scoped) + assignment
// ============================================================
function _managedKelasSet(username) {
  var set = {};
  _effectiveKelasModerator(username).kelas.forEach(function (k) { set[k] = true; });
  return set;
}

function getPengampuScoped(userCtx) {
  var roleLc = _fbRole(userCtx);
  var all = getPengampuData(); // dari Code.gs
  if (roleLc === 'admin') return all;
  if (roleLc !== 'moderator') return [];
  var set = _managedKelasSet(userCtx.username);
  return all.filter(function (p) {
    return (p.kelasDiampu || []).some(function (k) { return set[_t(k)]; }) || (p.kelasDiampu || []).length === 0;
  });
}

function simpanPengampuMod(userCtx, data) {
  var roleLc = _fbRole(userCtx);
  if (roleLc === 'admin') return simpanPengampuCepat(data);
  if (roleLc !== 'moderator') throw new Error('Akses ditolak.');
  var set = _managedKelasSet(userCtx.username);
  var kelas = Array.isArray(data && data.kelasDiampu) ? data.kelasDiampu.map(_t).filter(Boolean) : [];
  if (kelas.length === 0) throw new Error('Moderator wajib memilih minimal 1 kelas dalam wilayahnya.');
  var luar = kelas.filter(function (k) { return !set[k]; });
  if (luar.length) throw new Error('Tidak boleh menugaskan ke kelas di luar wilayah Anda: ' + luar.join(', '));
  var msg = simpanPengampu(data);
  _logActivity('Moderator Tambah Pengampu', _t(data.nama), userCtx.username);
  return { message: msg, pengampu: getPengampuScoped(userCtx) };
}

function editPengampuMod(userCtx, id, data) {
  var roleLc = _fbRole(userCtx);
  if (roleLc === 'admin') return editPengampuCepat(id, data);
  if (roleLc !== 'moderator') throw new Error('Akses ditolak.');
  var set = _managedKelasSet(userCtx.username);
  // Pengampu yang diedit harus punya minimal 1 kelas di wilayah moderator
  var target = getPengampuData().filter(function (p) { return p.id === _t(id); })[0];
  if (!target) throw new Error('Pengampu tidak ditemukan.');
  var inScope = (target.kelasDiampu || []).some(function (k) { return set[_t(k)]; }) || (target.kelasDiampu || []).length === 0;
  if (!inScope) throw new Error('Pengampu ini bukan dalam wilayah moderasi Anda.');
  var kelas = Array.isArray(data && data.kelasDiampu) ? data.kelasDiampu.map(_t).filter(Boolean) : [];
  var luar = kelas.filter(function (k) { return !set[k]; });
  if (luar.length) throw new Error('Tidak boleh menugaskan ke kelas di luar wilayah Anda: ' + luar.join(', '));
  var msg = editPengampu(id, data);
  _logActivity('Moderator Edit Pengampu', _t(data.nama), userCtx.username);
  return { message: msg, pengampu: getPengampuScoped(userCtx) };
}

// ============================================================
// 7. HARI LIBUR — versi Admin ATAU Moderator (req: keduanya boleh)
// ============================================================
function tambahHariLiburFB(userCtx, tanggal, keterangan, kelompokId) {
  var roleLc = _fbRole(userCtx);
  if (roleLc !== 'admin' && roleLc !== 'moderator') throw new Error('Hanya Admin/Moderator yang bisa menandai hari libur.');
  var tgl = _t(tanggal);
  if (!_isValidDate(tgl)) throw new Error('Tanggal tidak valid!');
  kelompokId = _t(kelompokId);
  if (roleLc === 'moderator') {
    var mine = _getModeratorGroupIds(userCtx.username);
    if (!kelompokId) throw new Error('Moderator harus memilih kelompok kelas untuk hari libur.');
    if (mine.indexOf(kelompokId) < 0) throw new Error('Anda tidak mengelola kelompok tersebut.');
  }
  return _withLock(function () {
    var sheet = _ensureHariLiburSheet();
    var exists = _readSheet('HariLibur').some(function (r) { return _t(r[1]) === tgl && _t(r[3]) === kelompokId; });
    if (exists) throw new Error('Tanggal ' + tgl + ' sudah ditandai libur untuk cakupan yang sama.');
    sheet.appendRow([_newId('HL'), tgl, _t(keterangan), kelompokId]);
    _logActivity('Tandai Hari Libur', tgl + (kelompokId ? ' [grup]' : ' [global]') + (_t(keterangan) ? (' - ' + _t(keterangan)) : ''), userCtx.username);
    return 'Hari libur ' + tgl + ' ditambahkan.';
  });
}

// Daftar hari libur beserta cakupan (global / nama kelompok).
function getHariLiburFB(userCtx) {
  _ensureHariLiburSheet();
  var gm = {}; getKelompokKelas().forEach(function (g) { gm[g.id] = g; });
  return _readSheet('HariLibur').map(function (r) {
    var kid = _t(r[3]);
    return { id: _t(r[0]), tanggal: _t(r[1]), keterangan: _t(r[2]), kelompokId: kid, kelompokNama: kid ? ((gm[kid] || {}).nama || kid) : 'Semua (Global)' };
  }).sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : -1; });
}

function hapusHariLiburFB(userCtx, id) {
  var roleLc = _fbRole(userCtx);
  if (roleLc !== 'admin' && roleLc !== 'moderator') throw new Error('Hanya Admin/Moderator yang bisa menghapus hari libur.');
  return _withLock(function () {
    var found = _findRowById('HariLibur', id);
    if (!found) throw new Error('Data tidak ditemukan!');
    if (roleLc === 'moderator') {
      var kid = _t(found.values[3]);
      if (!kid || _getModeratorGroupIds(userCtx.username).indexOf(kid) < 0) throw new Error('Hari libur ini di luar wilayah Anda.');
    }
    _sheet('HariLibur').deleteRow(found.rowIndex);
    _logActivity('Hapus Hari Libur', _t(found.values[1]), userCtx.username);
    return 'Hari libur dihapus.';
  });
}

// ============================================================
// 8. PERINGATAN PERSISTEN (per tanggal, sampai diisi)
// ============================================================
function _liburSet() {
  _ensureHariLiburSheet();
  var set = {};
  _readSheet('HariLibur').forEach(function (r) { var t = _t(r[1]); if (t && !_t(r[3])) set[t] = true; });
  return set;
}

// Info libur lengkap: global (semua kelas) + per-kelompok (per kelas).
function _liburInfo() {
  _ensureHariLiburSheet();
  var groups = {}; getKelompokKelas().forEach(function (g) { groups[g.id] = g; });
  var glob = {}, byDate = {};
  _readSheet('HariLibur').forEach(function (r) {
    var tgl = _t(r[1]); if (!tgl) return;
    var kid = _t(r[3]);
    if (!kid) { glob[tgl] = true; }
    else { if (!byDate[tgl]) byDate[tgl] = {}; var g = groups[kid]; if (g) g.kelas.forEach(function (k) { byDate[tgl][k] = true; }); }
  });
  return { global: glob, byDate: byDate };
}
function _isKelasLibur(info, tgl, kelas) { return !!info.global[tgl] || !!(info.byDate[tgl] && info.byDate[tgl][kelas]); }

// Hari aktif berbasis JADWAL MENGAJAR (BUKAN dari absensi yang sudah ada).
// Perbaikan bug kelas 11: sebelumnya hari aktif dihitung dari tanggal yang sudah
// ada di sheet Absensi, sehingga kelas yang belum pernah mengisi absen tidak
// pernah memunculkan peringatan. Sekarang: tanggal dianggap aktif bila nama
// harinya punya jadwal mengajar; setiap kelas wajib mengisi pada hari jadwalnya;
// kelas yang belum punya jadwal TETAP ikut dalam loop (wajib di semua hari aktif).
// Pengecualian hanya untuk hari libur bertanda (ditangani pemanggil via _isKelasLibur).
var HARI_INDONESIA = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
// Normalisasi nama hari agar ejaan di sheet Jadwal selalu cocok dengan nama
// hari hasil perhitungan tanggal: huruf besar/kecil diabaikan, tanda baca
// dibuang ("Jum'at" -> jumat), alias umum dikenali ("Ahad" -> minggu).
function _normHari(h) {
  var s = _t(h).toLowerCase().replace(/[''‘`.\s-]/g, '');
  var alias = { ahad: 'minggu', minggu: 'minggu', senin: 'senin', selasa: 'selasa', rabu: 'rabu', kamis: 'kamis', jumat: 'jumat', sabtu: 'sabtu' };
  return alias[s] || s;
}
function _hariDariTanggalStr(tgl) {
  var d = new Date(tgl + 'T00:00:00');
  return _normHari(HARI_INDONESIA[d.getDay()]);
}
function _hariAktifDariJadwal(maxHari) {
  var jadwal = _readSheet(SHEET_NAMES.JADWAL);
  var hariAny = {};   // nama hari yang punya jadwal (kelas apa pun)
  var kelasHari = {}; // kelas -> { hari: true }
  jadwal.forEach(function (r) {
    var hari = _normHari(r[1]), kls = _t(r[5]); // Jadwal: [ID, Hari, Jam Mulai, Jam Selesai, Mapel, Kelas]
    if (!hari) return;
    hariAny[hari] = true;
    if (kls) { if (!kelasHari[kls]) kelasHari[kls] = {}; kelasHari[kls][hari] = true; }
  });
  var today = _todayStr();
  var start = today;
  if (maxHari) {
    var d0 = new Date();
    d0.setDate(d0.getDate() - maxHari);
    start = Utilities.formatDate(d0, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var dates = [], wajib = {};
  var cur = new Date(start + 'T00:00:00');
  var end = new Date(today + 'T00:00:00');
  while (cur <= end) {
    var ds = Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var hari = _hariDariTanggalStr(ds);
    if (hariAny[hari]) {
      dates.push(ds);
      var m = {};
      Object.keys(kelasHari).forEach(function (k) { if (kelasHari[k][hari]) m[k] = true; });
      wajib[ds] = m;
    }
    cur.setDate(cur.getDate() + 1);
  }
  // Fallback: kelas yang punya entri jadwal tetapi nama harinya tidak pernah
  // cocok dengan tanggal nyata (ejaan tak dikenal) diperlakukan seperti kelas
  // tanpa jadwal -> wajib di semua hari aktif, sehingga TETAP diperingatkan.
  Object.keys(kelasHari).forEach(function (k) {
    var cocok = false;
    for (var i = 0; i < dates.length; i++) { if (wajib[dates[i]][k]) { cocok = true; break; } }
    if (!cocok) delete kelasHari[k];
  });
  return { dates: dates, wajib: wajib, kelasHari: kelasHari, hariAny: hariAny };
}
// Apakah kelas wajib mengisi pada tanggal tsb? Kelas tanpa jadwal -> wajib setiap hari aktif.
function _kelasWajibPada(aktif, tgl, kelas) {
  var kh = aktif.kelasHari[kelas];
  if (!kh) return true; // belum ada jadwal -> tetap masuk loop peringatan
  return !!(aktif.wajib[tgl] && aktif.wajib[tgl][kelas]);
}

// Req #6: absen siswa belum diisi, muncul per tanggal sampai diisi.
// Hari aktif diambil dari JADWAL MENGAJAR sehingga kelas yang belum mengisi
// absen sama sekali (mis. kelas 11) tetap memunculkan peringatan.
function getPeringatanAbsenSiswaPersisten(userCtx) {
  _ensureFiturSheets();
  var c = _absensiCols();
  var aktif = _hariAktifDariJadwal(KONFIG.PERINGATAN_MAX_HARI);
  var info = _liburInfo();
  var isDismissed = _makeDismissMatcher();
  var set = _kelasSetPeringatan(userCtx); // null = admin (semua); pengampu = satu kelompok
  var semuaKelas = _readSheet(SHEET_NAMES.KELAS).map(function (r) { return _t(r[1]); }).filter(Boolean);
  if (set) semuaKelas = semuaKelas.filter(function (k) { return set.has(k); });

  // Peta tanggal -> {kelas: true} yang SUDAH mengisi
  var sudah = {};
  _readSheet(SHEET_NAMES.ABSENSI).forEach(function (r) {
    var tgl = _t(r[c.tanggal]);
    if (!sudah[tgl]) sudah[tgl] = {};
    sudah[tgl][_t(r[c.kelas])] = true;
  });

  var hasil = [];
  aktif.dates.forEach(function (tgl) {
    if (info.global[tgl]) return; // libur global: tidak ada peringatan sama sekali
    var belum = semuaKelas.filter(function (k) {
      if (!_kelasWajibPada(aktif, tgl, k)) return false; // bukan hari jadwal kelas ini
      if (sudah[tgl] && sudah[tgl][k]) return false;
      if (_isKelasLibur(info, tgl, k)) return false;
      return !isDismissed('absenSiswa', k, '', tgl);
    });
    if (belum.length) hasil.push({ tanggal: tgl, kelas: belum });
  });
  return hasil.sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : -1; });
}

// Diagnosa peringatan absen siswa: menelusuri alasan setiap kelas diperingatkan
// atau tidak (jadwal, hari aktif, sudah isi, libur, disembunyikan). Khusus Admin —
// hasilnya bisa disalin dan dikirim ke pengembang bila ada kelas yang "hilang".
function getDiagnosaPeringatan(userCtx) {
  if (_fbRole(userCtx) !== 'admin') throw new Error('Hanya Admin yang dapat menjalankan diagnosa.');
  _ensureFiturSheets();
  var aktif = _hariAktifDariJadwal(KONFIG.PERINGATAN_MAX_HARI);
  var info = _liburInfo();
  var isDismissed = _makeDismissMatcher();
  var c = _absensiCols();
  var sudah = {};
  _readSheet(SHEET_NAMES.ABSENSI).forEach(function (r) {
    var tgl = _t(r[c.tanggal]);
    if (!sudah[tgl]) sudah[tgl] = {};
    sudah[tgl][_t(r[c.kelas])] = true;
  });
  var L = [];
  L.push('Hari yang punya jadwal (setelah normalisasi): ' + (Object.keys(aktif.hariAny).join(', ') || '(KOSONG — periksa kolom Hari di sheet Jadwal!)'));
  L.push('Tanggal aktif dlm ' + KONFIG.PERINGATAN_MAX_HARI + ' hari terakhir: ' + aktif.dates.length + (aktif.dates.length ? ' (' + aktif.dates[0] + ' s.d. ' + aktif.dates[aktif.dates.length - 1] + ')' : ' — TIDAK ADA, periksa sheet Jadwal!'));
  L.push('');
  var kelasList = _readSheet(SHEET_NAMES.KELAS).map(function (r) { return _t(r[1]); }).filter(Boolean);
  if (!kelasList.length) L.push('(Sheet Kelas kosong!)');
  kelasList.forEach(function (kls) {
    var kh = aktif.kelasHari[kls];
    var wajibTgl = aktif.dates.filter(function (tgl) {
      return _kelasWajibPada(aktif, tgl, kls) && !info.global[tgl] && !_isKelasLibur(info, tgl, kls);
    });
    var belum = wajibTgl.filter(function (tgl) { return !(sudah[tgl] && sudah[tgl][kls]); });
    var dism = belum.filter(function (tgl) { return isDismissed('absenSiswa', kls, '', tgl); });
    L.push('- ' + kls +
      ' | jadwal: ' + (kh ? Object.keys(kh).join(', ') : '(tidak ada -> wajib semua hari aktif)') +
      ' | wajib isi: ' + wajibTgl.length + ' tgl' +
      ' | belum isi: ' + belum.length + ' tgl' +
      ' | disembunyikan: ' + dism.length +
      (belum.length ? ' | contoh: ' + belum.slice(Math.max(0, belum.length - 3)).join(', ') : ''));
  });
  return L.join('\n');
}

// Req #7: pengampu belum mengisi absen pengampu / jurnal / buku mingguan.
function getPeringatanPengampuPersisten(userCtx) {
  _ensureFiturSheets();
  var info = _liburInfo();
  var isDismissed = _makeDismissMatcher();
  var today = _todayStr();
  var set = _kelasSetPeringatan(userCtx); // null = admin; pengampu = kelas satu kelompoknya

  // --- Absen pengampu belum diisi per tanggal aktif (berbasis Jadwal mengajar) ---
  var aktif = _hariAktifDariJadwal(KONFIG.PERINGATAN_MAX_HARI);
  var tanggalAktif = aktif.dates;
  var apRows = _readSheet(SHEET_NAMES.ABSENSI_PENGAMPU || 'AbsensiPengampu');
  var apSudah = {};
  apRows.forEach(function (r) {
    var tgl = _t(r[1]); var uname = _t(r[2]);
    var ada = _t(r[4]) || _t(r[6]); // jam masuk atau status terisi
    if (!apSudah[tgl]) apSudah[tgl] = {};
    if (ada) apSudah[tgl][uname.toLowerCase()] = true;
  });
  // pengampu dalam scope (untuk moderator: yang mengampu kelas wilayahnya)
  var pengampu = getPengampuData();
  if (set) {
    pengampu = pengampu.filter(function (p) {
      return (p.kelasDiampu || []).length === 0 || (p.kelasDiampu || []).some(function (k) { return set.has(_t(k)); });
    });
  }
  var absenBelum = [];
  tanggalAktif.forEach(function (tgl) {
    if (info.global[tgl]) return;
    var belum = pengampu.filter(function (p) {
      if (apSudah[tgl] && apSudah[tgl][p.username.toLowerCase()]) return false;
      var kd = p.kelasDiampu || [];
      if (kd.length > 0 && kd.every(function (k) { return _isKelasLibur(info, tgl, k); })) return false;
      if (isDismissed('absenPengampu', '', (p.nama || p.username), tgl)) return false;
      return true;
    }).map(function (p) { return p.nama || p.username; });
    if (belum.length) absenBelum.push({ tanggal: tgl, pengampu: belum });
  });

  // --- Jurnal belum diisi per tanggal aktif per kelas ---
  var jurnalSudah = {};
  _readSheet(SHEET_NAMES.JURNAL).forEach(function (r) {
    var tgl = _t(r[2]); var kls = _t(r[4]);
    if (!jurnalSudah[tgl]) jurnalSudah[tgl] = {};
    jurnalSudah[tgl][kls] = true;
  });
  var semuaKelas = _readSheet(SHEET_NAMES.KELAS).map(function (r) { return _t(r[1]); }).filter(Boolean);
  if (set) semuaKelas = semuaKelas.filter(function (k) { return set.has(k); });
  var jurnalBelum = [];
  tanggalAktif.forEach(function (tgl) {
    if (info.global[tgl]) return; // libur global: lewati
    var belum = semuaKelas.filter(function (k) { return _kelasWajibPada(aktif, tgl, k) && !(jurnalSudah[tgl] && jurnalSudah[tgl][k]) && !_isKelasLibur(info, tgl, k) && !isDismissed('jurnal', k, '', tgl); });
    if (belum.length) jurnalBelum.push({ tanggal: tgl, kelas: belum });
  });

  // --- Buku mingguan belum diisi (per minggu yang sudah ada entri) ---
  var bmRows = _readSheet('BukuMingguan');
  var mingguSet = {};
  var bmSudah = {};
  bmRows.forEach(function (r) {
    var mgg = _t(r[1]); var kls = _t(r[5]);
    if (mgg && !info.global[mgg] && mgg <= today) mingguSet[mgg] = true;
    if (!bmSudah[mgg]) bmSudah[mgg] = {};
    bmSudah[mgg][kls] = true;
  });
  var bukuBelum = [];
  Object.keys(mingguSet).sort().reverse().forEach(function (mgg) {
    var belum = semuaKelas.filter(function (k) { return !(bmSudah[mgg] && bmSudah[mgg][k]) && !_isKelasLibur(info, mgg, k) && !isDismissed('bukuMingguan', k, '', mgg); });
    if (belum.length) bukuBelum.push({ tanggal: mgg, kelas: belum });
  });

  return {
    absenPengampu: absenBelum.sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : -1; }),
    jurnal: jurnalBelum.sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : -1; }),
    bukuMingguan: bukuBelum
  };
}

// ============================================================
// 9. KONTAK WHATSAPP PENGAMPU (klien membuat link wa.me)
// ============================================================
function getPengampuKontak(userCtx) {
  var list = getPengampuScoped(userCtx);
  return list.map(function (p) {
    return { nama: p.nama, username: p.username, noHp: p.noHp || '', kelasDiampu: p.kelasDiampu || [], active: p.active };
  });
}

// ============================================================
// 10. UPGRADE PENGAMPU -> MODERATOR
// ============================================================
function upgradePengampuKeModerator(userCtx, pengampuId, kelompokIds) {
  _fbAssertAdmin(userCtx);
  var ids = Array.isArray(kelompokIds) ? kelompokIds.map(_t).filter(Boolean) : [];
  if (!ids.length) throw new Error('Pilih minimal 1 kelompok kelas untuk moderator.');
  return _withLock(function () {
    _ensureFiturSheets();
    var f = _findRowById(SHEET_NAMES.AUTH, pengampuId);
    if (!f) throw new Error('Pengampu tidak ditemukan!');
    if (_t(f.values[3]).toLowerCase() !== 'pengampu') throw new Error('User yang dipilih bukan pengampu.');
    _sheet(SHEET_NAMES.AUTH).getRange(f.rowIndex, 4).setValue('Moderator');
    _setModeratorGroups(_t(f.values[1]), ids);
    _logActivity('Upgrade Pengampu->Moderator', _t(f.values[4]) + ' (' + _t(f.values[1]) + ')', userCtx.username);
    return 'Pengampu "' + _t(f.values[4]) + '" berhasil diangkat menjadi Moderator!';
  });
}

// ============================================================
// 11. SISWA SERING TIDAK HADIR (>5x) — peringatan permanen
// Non-hadir = Sakit / Izin / Alpa. Dismiss hanya Admin/Moderator.
// ============================================================
function _hitungTidakHadir(userCtx) {
  var c = _absensiCols();
  var set = _kelasSetFromCtx(userCtx);
  var NON = { 'Sakit': true, 'Izin': true, 'Alpa': true };
  var map = {};
  _readSheet(SHEET_NAMES.ABSENSI).forEach(function (r) {
    var kelas = _t(r[c.kelas]), nama = _t(r[c.nama]), st = _t(r[c.status]);
    if (/^al/i.test(st)) st = 'Alpa'; // kompatibel ejaan lama pada data tersimpan
    if (!nama) return;
    if (set && !set.has(kelas)) return;
    if (!NON[st]) return;
    var key = kelas + '||' + nama;
    if (!map[key]) map[key] = { kelas: kelas, nama: nama, jumlah: 0 };
    map[key].jumlah++;
  });
  return map;
}

function getSiswaSeringAbsen(userCtx, kelasFilter) {
  _ensureFiturSheets();
  var map = _hitungTidakHadir(userCtx);
  var dis = {};
  _readSheet(FB_SHEETS.DISMISS).forEach(function (r) { dis[_t(r[1]) + '||' + _t(r[2])] = parseInt(_t(r[3]), 10) || 0; });
  kelasFilter = _t(kelasFilter);
  var res = [];
  Object.keys(map).forEach(function (key) {
    var x = map[key];
    if (x.jumlah <= 5) return;
    if (kelasFilter && x.kelas !== kelasFilter) return;
    if (x.jumlah <= (dis[key] || 0)) return;
    res.push({ kelas: x.kelas, nama: x.nama, jumlah: x.jumlah });
  });
  return res.sort(function (a, b) { return b.jumlah - a.jumlah; });
}

function dismissPeringatanSiswa(userCtx, kelas, nama) {
  var rl = _fbRole(userCtx);
  if (rl !== 'admin' && rl !== 'moderator') throw new Error('Hanya Admin/Moderator yang dapat menghapus peringatan.');
  kelas = _t(kelas); nama = _t(nama);
  if (rl === 'moderator' && !_managedKelasSet(userCtx.username)[kelas]) throw new Error('Kelas di luar wilayah moderasi Anda.');
  return _withLock(function () {
    _ensureFiturSheets();
    var map = _hitungTidakHadir({ role: 'admin' });
    var cur = (map[kelas + '||' + nama] || {}).jumlah || 0;
    var sh = _sheet(FB_SHEETS.DISMISS);
    var rows = _readSheet(FB_SHEETS.DISMISS);
    var idx = -1;
    for (var i = 0; i < rows.length; i++) { if (_t(rows[i][1]) === kelas && _t(rows[i][2]) === nama) { idx = i; break; } }
    if (idx >= 0) sh.getRange(idx + 2, 4, 1, 3).setValues([[String(cur), _t(userCtx.username), _fbIso(new Date())]]);
    else sh.appendRow([_newId('PD'), kelas, nama, String(cur), _t(userCtx.username), _fbIso(new Date())]);
    _logActivity('Hapus Peringatan Siswa', nama + ' (' + kelas + ')', userCtx.username);
    return 'Peringatan untuk ' + nama + ' dihapus.';
  });
}

// ============================================================
// 12. DETAIL PERINGATAN PER PENGAMPU (untuk pesan WA custom)
// ============================================================
function getPengampuPeringatanDetail(userCtx) {
  _ensureFiturSheets();
  var peng = getPengampuScoped(userCtx);
  var absSiswa = getPeringatanAbsenSiswaPersisten(userCtx);
  var pengWarn = getPeringatanPengampuPersisten(userCtx);
  var sering = getSiswaSeringAbsen(userCtx, '');
  var kelasTgl = {};
  absSiswa.forEach(function (d) { d.kelas.forEach(function (k) { if (!kelasTgl[k]) kelasTgl[k] = []; kelasTgl[k].push(d.tanggal); }); });
  var pengTgl = {};
  (pengWarn.absenPengampu || []).forEach(function (d) { d.pengampu.forEach(function (nm) { if (!pengTgl[nm]) pengTgl[nm] = []; pengTgl[nm].push(d.tanggal); }); });
  var seringByKelas = {};
  sering.forEach(function (s) { if (!seringByKelas[s.kelas]) seringByKelas[s.kelas] = []; seringByKelas[s.kelas].push(s); });
  return peng.map(function (p) {
    var kd = p.kelasDiampu || [];
    var absenKelas = kd.filter(function (k) { return kelasTgl[k]; }).map(function (k) { return { kelas: k, tanggal: kelasTgl[k] }; });
    var siswa = [];
    kd.forEach(function (k) { (seringByKelas[k] || []).forEach(function (s) { siswa.push(s); }); });
    return {
      username: p.username, nama: p.nama, noHp: p.noHp || '', kelas: kd,
      absenKelasBelum: absenKelas,
      absenPengampuBelum: pengTgl[p.nama || p.username] || [],
      siswaSeringAbsen: siswa
    };
  });
}

// ============================================================
// 13. DISMISS GRANULAR PERINGATAN DASHBOARD
// Kriteria bebas: topik / kelas / pengampu / tanggal (kosong = semua).
// Topik: absenSiswa | absenPengampu | jurnal | bukuMingguan
// ============================================================
function _pengampuKelasMap() {
  var m = {};
  getPengampuData().forEach(function (p) {
    var kl = p.kelasDiampu || [];
    m[p.nama || p.username] = kl;
    m[p.username] = kl;
  });
  return m;
}
function _makeDismissMatcher() {
  _ensureFiturSheets();
  var rules = _readSheet(FB_SHEETS.DISMISS2).map(function (r) {
    return { topik: _t(r[1]), kelompokId: _t(r[2]), pengampu: _t(r[3]).split(',').map(_t).filter(Boolean), tanggal: _t(r[4]), byRole: _t(r[5]).toLowerCase(), by: _t(r[6]) };
  });
  var gmap = {}; getKelompokKelas().forEach(function (g) { var s = {}; g.kelas.forEach(function (k) { s[k] = true; }); gmap[g.id] = s; });
  var modCache = {}, pm = null;
  function modClasses(u) { if (!modCache[u]) modCache[u] = _managedKelasSet(u); return modCache[u]; }
  function pmap() { if (!pm) pm = _pengampuKelasMap(); return pm; }
  return function (topik, kelas, pengampu, tanggal) {
    for (var i = 0; i < rules.length; i++) {
      var r = rules[i];
      if (r.topik && r.topik !== topik) continue;
      if (r.tanggal && r.tanggal !== tanggal) continue;
      if (r.kelompokId) { if (!kelas) continue; var gc = gmap[r.kelompokId]; if (!gc || !gc[kelas]) continue; }
      if (r.pengampu.length) { if (!pengampu) continue; if (r.pengampu.indexOf(pengampu) < 0) continue; }
      if (r.byRole === 'moderator') { // rule moderator hanya berlaku pada wilayahnya
        var set = modClasses(r.by);
        if (kelas) { if (!set[kelas]) continue; }
        else if (pengampu) { var kl = pmap()[pengampu] || []; if (!kl.some(function (k) { return set[k]; })) continue; }
        else continue;
      }
      return true;
    }
    return false;
  };
}
function dismissPeringatan(userCtx, data) {
  var rl = _fbRole(userCtx);
  if (rl !== 'admin' && rl !== 'moderator') throw new Error('Hanya Admin/Moderator yang dapat menyembunyikan peringatan.');
  var topik = _t(data && data.topik), kelompokId = _t(data && data.kelompokId), tanggal = _t(data && data.tanggal);
  var pengList = Array.isArray(data && data.pengampu) ? data.pengampu.map(_t).filter(Boolean) : (_t(data && data.pengampu) ? [_t(data.pengampu)] : []);
  if (!topik && !kelompokId && !pengList.length && !tanggal) throw new Error('Pilih minimal satu kriteria (topik/kelompok/pengampu/tanggal).');
  if (rl === 'moderator') {
    var mine = _getModeratorGroupIds(userCtx.username);
    if (kelompokId && mine.indexOf(kelompokId) < 0) throw new Error('Kelompok di luar wilayah moderasi Anda.');
    if (pengList.length) {
      var set = _managedKelasSet(userCtx.username), pm = _pengampuKelasMap();
      pengList.forEach(function (nm) { var kl = pm[nm] || []; if (!kl.some(function (k) { return set[k]; })) throw new Error('Pengampu "' + nm + '" di luar wilayah Anda.'); });
    }
  }
  _withLock(function () {
    _ensureFiturSheets();
    _sheet(FB_SHEETS.DISMISS2).appendRow([_newId('DM'), topik, kelompokId, pengList.join(','), tanggal, _t(userCtx.role), _t(userCtx.username), _fbIso(new Date())]);
  });
  _logActivity('Sembunyikan Peringatan', [topik, kelompokId, pengList.join('+'), tanggal].filter(Boolean).join(' / '), userCtx.username);
  return 'Peringatan disembunyikan.';
}
function getDismissRules(userCtx) {
  _ensureFiturSheets();
  var rl = _fbRole(userCtx);
  var gm = {}; getKelompokKelas().forEach(function (g) { gm[g.id] = g; });
  return _readSheet(FB_SHEETS.DISMISS2).filter(function (r) {
    if (rl === 'admin') return true;
    return _t(r[6]).toLowerCase() === _t(userCtx.username).toLowerCase();
  }).map(function (r) {
    var kid = _t(r[2]);
    return { id: _t(r[0]), topik: _t(r[1]), kelompokId: kid, kelompokNama: kid ? ((gm[kid] || {}).nama || kid) : '', pengampu: _t(r[3]).split(',').map(_t).filter(Boolean), tanggal: _t(r[4]), byRole: _t(r[5]), by: _t(r[6]), at: _t(r[7]) };
  }).reverse();
}
function dismissPeringatanBatch(userCtx, items) {
  var rl = _fbRole(userCtx);
  if (rl !== 'admin' && rl !== 'moderator') throw new Error('Hanya Admin/Moderator yang dapat menyembunyikan peringatan.');
  if (!Array.isArray(items) || !items.length) throw new Error('Tidak ada peringatan yang dipilih.');
  var mine = rl === 'moderator' ? _getModeratorGroupIds(userCtx.username) : null;
  var mset = rl === 'moderator' ? _managedKelasSet(userCtx.username) : null;
  var pm = rl === 'moderator' ? _pengampuKelasMap() : null;
  var n = 0;
  _withLock(function () {
    _ensureFiturSheets();
    var sh = _sheet(FB_SHEETS.DISMISS2);
    items.forEach(function (d) {
      var topik = _t(d && d.topik), kelompokId = _t(d && d.kelompokId), tanggal = _t(d && d.tanggal);
      var pengList = Array.isArray(d && d.pengampu) ? d.pengampu.map(_t).filter(Boolean) : (_t(d && d.pengampu) ? [_t(d.pengampu)] : []);
      if (!topik && !kelompokId && !pengList.length && !tanggal) return;
      if (rl === 'moderator') {
        if (kelompokId && mine.indexOf(kelompokId) < 0) return;
        if (pengList.length) { pengList = pengList.filter(function (nm) { var kl = pm[nm] || []; return kl.some(function (k) { return mset[k]; }); }); if (!pengList.length) return; }
      }
      sh.appendRow([_newId('DM'), topik, kelompokId, pengList.join(','), tanggal, _t(userCtx.role), _t(userCtx.username), _fbIso(new Date())]);
      n++;
    });
  });
  _logActivity('Sembunyikan Peringatan (batch)', n + ' aturan', userCtx.username);
  return 'Menyembunyikan ' + n + ' peringatan.';
}
function hapusDismiss(userCtx, id) {
  var rl = _fbRole(userCtx);
  if (rl !== 'admin' && rl !== 'moderator') throw new Error('Akses ditolak.');
  return _withLock(function () {
    var f = _findRowById(FB_SHEETS.DISMISS2, id);
    if (!f) throw new Error('Data tidak ditemukan!');
    if (rl === 'moderator' && _t(f.values[6]).toLowerCase() !== _t(userCtx.username).toLowerCase()) throw new Error('Ini bukan sembunyian Anda.');
    _sheet(FB_SHEETS.DISMISS2).deleteRow(f.rowIndex);
    _logActivity('Tampilkan Kembali Peringatan', _t(id), userCtx.username);
    return 'Peringatan ditampilkan kembali.';
  });
}

// ============================================================
// 15. RENTANG PERIODE STATISTIK (minggu/bulan/triwulan/semester)
// Patokan triwulan & semester mengikuti kalender akademik di KONFIG.
// ============================================================
function getPeriodeRange(jenis) {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var y = parseInt(Utilities.formatDate(now, tz, 'yyyy'), 10);
  var m = parseInt(Utilities.formatDate(now, tz, 'M'), 10); // 1-12
  function fmt(d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  function lastOf(yy, mm1) { return new Date(yy, mm1, 0); } // tanggal terakhir bulan mm1 (1-12)
  jenis = _t(jenis).toLowerCase();
  if (jenis === 'bulan') {
    return { start: fmt(new Date(y, m - 1, 1)), end: fmt(lastOf(y, m)), label: 'Bulan ini' };
  }
  var g = KONFIG.GANJIL_MULAI_BULAN, e = KONFIG.GENAP_MULAI_BULAN; // default 7 (Jul) & 1 (Jan)
  if (jenis === 'triwulan' || jenis === 'semester') {
    // Bulan >= g -> Semester Ganjil (g..Des); selain itu -> Semester Genap (e..g-1)
    var semStart, semEnd, semLabel, twOffset;
    if (m >= g) { semStart = g; semEnd = 12; semLabel = 'Semester Ganjil'; twOffset = 0; }
    else { semStart = e; semEnd = g - 1; semLabel = 'Semester Genap'; twOffset = 2; }
    if (jenis === 'semester') {
      return { start: fmt(new Date(y, semStart - 1, 1)), end: fmt(lastOf(y, semEnd)), label: semLabel + ' ' + y };
    }
    // Triwulan = blok 3 bulan dalam semester berjalan (Ganjil: TW1-2, Genap: TW3-4)
    var idx = Math.floor((m - semStart) / 3); if (idx > 1) idx = 1; if (idx < 0) idx = 0;
    var tStart = semStart + idx * 3, tEnd = tStart + 2;
    return { start: fmt(new Date(y, tStart - 1, 1)), end: fmt(lastOf(y, tEnd)), label: 'Triwulan ' + (twOffset + idx + 1) };
  }
  var w = _currentWeekRange(); // default: per minggu
  return { start: w.start, end: w.end, label: 'Minggu ini' };
}

// ============================================================
// 16. DASHBOARD PUBLIK (Pengunjung/Pengamat) — tanpa login.
// Hanya agregat + daftar nama & wilayah (tanpa NIP/NoHP/password).
// ============================================================
function getPublikDashboard() {
  _ensureFiturSheets();
  var c = _absensiCols();
  var week = _currentWeekRange();
  var sem = getPeriodeRange('semester');

  // --- Kehadiran siswa minggu berjalan ---
  // Jika minggu ini belum ada data sama sekali, mundur ke minggu terakhir yang
  // punya data (maks 12 minggu ke belakang) agar grafik publik tidak kosong.
  var KODE = { 'Hadir': 'H', 'H': 'H', 'Sakit': 'S', 'S': 'S', 'Izin': 'I', 'I': 'I', 'Terlambat': 'T', 'T': 'T' };
  var absRows = _readSheet(SHEET_NAMES.ABSENSI);
  function _hitungKehadiranMinggu(wk) {
    var k = { H: 0, S: 0, I: 0, A: 0, T: 0 }, tot = 0;
    absRows.forEach(function (r) {
      var tgl = _t(r[c.tanggal]);
      if (tgl < wk.start || tgl > wk.end) return;
      var st = _t(r[c.status]);
      if (/^al/i.test(st)) { k.A++; tot++; return; } // kompatibel ejaan lama
      var code = KODE[st];
      if (code) { k[code]++; tot++; }
    });
    return { counts: k, total: tot };
  }
  function _geserMinggu(wk, n) {
    var tz = Session.getScriptTimeZone();
    var s = new Date(wk.start + 'T00:00:00'); s.setDate(s.getDate() - 7 * n);
    var e = new Date(wk.end + 'T00:00:00'); e.setDate(e.getDate() - 7 * n);
    return { start: Utilities.formatDate(s, tz, 'yyyy-MM-dd'), end: Utilities.formatDate(e, tz, 'yyyy-MM-dd') };
  }
  var mingguDipakai = week, hk = _hitungKehadiranMinggu(week), mundur = 0;
  while (hk.total === 0 && mundur < 12) { mundur++; mingguDipakai = _geserMinggu(week, mundur); hk = _hitungKehadiranMinggu(mingguDipakai); }
  var kh = hk.counts;

  // --- Tingkat kehadiran ibadah (absen buku mingguan, semester berjalan) ---
  var ib = { kumpul: 0, tidak: 0 };
  _readSheet('BukuMingguan').forEach(function (r) {
    var tgl = _t(r[1]) || _t(r[2]); // kolom Minggu (fallback Tanggal)
    if (tgl && (tgl < sem.start || tgl > sem.end)) return;
    var st = _t(r[6]).toLowerCase();
    if (st === 'kumpul') ib.kumpul++;
    else if (st) ib.tidak++;
  });

  // --- Siswa: umum, per kelas, per kelompok ---
  var perKelas = {}, totalSiswa = 0;
  _readSheet(SHEET_NAMES.SISWA).forEach(function (r) {
    var k = _t(r[3]); // SISWA: [ID, No Urut, Nama Siswa, Kelas, ...]
    if (!k) return;
    perKelas[k] = (perKelas[k] || 0) + 1;
    totalSiswa++;
  });
  var kelasList = _readSheet(SHEET_NAMES.KELAS).map(function (r) { return _t(r[1]); }).filter(Boolean);
  var groups = getKelompokKelas();
  var inGroup = {};
  var perKelompok = groups.map(function (gr) {
    var siswa = 0;
    (gr.kelas || []).forEach(function (k) { inGroup[k] = true; siswa += perKelas[k] || 0; });
    return { nama: gr.nama, jumlahKelas: (gr.kelas || []).length, jumlahSiswa: siswa };
  });
  var tanpaKelompok = kelasList.filter(function (k) { return !inGroup[k]; });
  if (tanpaKelompok.length) {
    var sisa = 0; tanpaKelompok.forEach(function (k) { sisa += perKelas[k] || 0; });
    perKelompok.push({ nama: 'Belum Berkelompok', jumlahKelas: tanpaKelompok.length, jumlahSiswa: sisa });
  }

  // --- Daftar pengampu + moderator + admin (nama & kelas/wilayah saja) ---
  var gmap = {}; groups.forEach(function (gr) { gmap[gr.id] = gr.nama; });
  var guru = [];
  _readSheet(SHEET_NAMES.AUTH).forEach(function (r) {
    if (_t(r[8]).toLowerCase() === 'tidak') return; // akun nonaktif disembunyikan
    var role = _t(r[3]) || 'Pengampu';
    var rl = role.toLowerCase();
    var wilayah = '';
    if (rl === 'admin') wilayah = 'Semua kelas';
    else if (rl === 'moderator') wilayah = _getModeratorGroupIds(_t(r[1])).map(function (id) { return gmap[id] || id; }).join(', ') || '-';
    else wilayah = _t(r[7]).split(',').map(_t).filter(Boolean).join(', ') || 'Semua kelas';
    guru.push({ nama: _t(r[4]) || _t(r[1]), role: role, wilayah: wilayah });
  });

  return {
    appInfo: getPublicAppInfo(),
    kehadiran: { labels: ['Hadir', 'Sakit', 'Izin', 'Alpa', 'Terlambat'], keys: ['H', 'S', 'I', 'A', 'T'], counts: kh, tglMulai: mingguDipakai.start, tglAkhir: mingguDipakai.end, mingguIni: mundur === 0 },
    ibadah: { kumpul: ib.kumpul, tidak: ib.tidak, periode: sem.label },
    totalSiswa: totalSiswa,
    totalKelas: kelasList.length,
    totalKelompok: groups.length,
    perKelompok: perKelompok,
    guru: guru
  };
}

// ============================================================
// 17. BAHAN AJAR — upload PDF ke Google Drive, baca via Drive viewer.
// Upload: Admin & Moderator. Download & baca: semua (termasuk pengunjung).
// ============================================================
function _bahanAjarFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('BAHAN_AJAR_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* buat ulang */ } }
  var folder = DriveApp.createFolder(KONFIG.BAHAN_AJAR_FOLDER);
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  props.setProperty('BAHAN_AJAR_FOLDER_ID', folder.getId());
  return folder;
}

function uploadBahanAjar(userCtx, data) {
  var rl = _fbRole(userCtx);
  if (rl !== 'admin' && rl !== 'moderator') throw new Error('Hanya Admin/Moderator yang dapat mengunggah bahan ajar.');
  var judul = _t(data && data.judul);
  if (!judul) throw new Error('Judul bahan ajar wajib diisi!');
  var namaFile = _t(data && data.namaFile) || (judul + '.pdf');
  var mime = _t(data && data.mimeType);
  if (mime !== 'application/pdf' && !/\.pdf$/i.test(namaFile)) throw new Error('Hanya file PDF yang diperbolehkan.');
  var base64 = String(data && data.base64 || '');
  if (!base64) throw new Error('File belum dipilih.');
  var kelompokId = _t(data && data.kelompokId);
  if (rl === 'moderator') {
    var mine = _getModeratorGroupIds(userCtx.username);
    if (kelompokId && mine.indexOf(kelompokId) < 0) throw new Error('Kelompok di luar wilayah moderasi Anda.');
  }
  return _withLock(function () {
    _ensureFiturSheets();
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', namaFile);
    var file = _bahanAjarFolder().createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    _sheet(FB_SHEETS.BAHAN).appendRow([
      _newId('BA'), judul, _t(data && data.deskripsi), file.getId(), file.getName(),
      String(file.getSize()), _t(userCtx.username), _t(userCtx.role), kelompokId, _fbIso(new Date())
    ]);
    _logActivity('Unggah Bahan Ajar', judul + ' (' + namaFile + ')', userCtx.username);
    return 'Bahan ajar "' + judul + '" berhasil diunggah.';
  });
}

function getBahanAjarList(userCtx) {
  _ensureFiturSheets();
  var gm = {}; getKelompokKelas().forEach(function (gr) { gm[gr.id] = gr; });
  return _readSheet(FB_SHEETS.BAHAN).map(function (r) {
    var fid = _t(r[3]);
    var kid = _t(r[8]);
    return {
      id: _t(r[0]), judul: _t(r[1]), deskripsi: _t(r[2]), fileId: fid,
      namaFile: _t(r[4]), ukuran: parseInt(_t(r[5]), 10) || 0,
      pengunggah: _t(r[6]), role: _t(r[7]), kelompokId: kid,
      kelompokNama: kid ? ((gm[kid] || {}).nama || kid) : 'Semua',
      diunggah: _t(r[9]),
      urlView: 'https://drive.google.com/file/d/' + fid + '/preview',
      urlDownload: 'https://drive.google.com/uc?export=download&id=' + fid
    };
  }).reverse();
}

function hapusBahanAjar(userCtx, id) {
  var rl = _fbRole(userCtx);
  if (rl !== 'admin' && rl !== 'moderator') throw new Error('Akses ditolak.');
  return _withLock(function () {
    var f = _findRowById(FB_SHEETS.BAHAN, id);
    if (!f) throw new Error('Bahan ajar tidak ditemukan!');
    if (rl === 'moderator' && _t(f.values[6]).toLowerCase() !== _t(userCtx.username).toLowerCase()) {
      throw new Error('Anda hanya boleh menghapus berkas yang Anda unggah sendiri.');
    }
    try { DriveApp.getFileById(_t(f.values[3])).setTrashed(true); } catch (e) {}
    _sheet(FB_SHEETS.BAHAN).deleteRow(f.rowIndex);
    _logActivity('Hapus Bahan Ajar', _t(f.values[1]), userCtx.username);
    return 'Bahan ajar dihapus.';
  });
}
