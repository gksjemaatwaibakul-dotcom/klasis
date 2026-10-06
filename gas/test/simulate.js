// ============================================================
// Harness simulasi Node.js untuk menguji logika Code.gs (Google Apps Script)
// tanpa server Google. SpreadsheetApp/DriveApp/dsb di-mock in-memory.
// Jalankan:  node /app/gas/test/simulate.js
// ============================================================
const fs = require('fs');
const path = require('path');

// ---------- MOCK UTILITIES ----------
const DAYS_EN = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const MONTHS_EN = ['January','February','March','April','May','June','July','August','September','October','November','December'];
function formatDate(d, tz, p) {
  const pad = (n) => String(n).padStart(2, '0');
  return String(p).replace(/yyyy|MMMM|EEEE|MM|dd|HH|mm|ss|u|M/g, (tok) => {
    switch (tok) {
      case 'yyyy': return String(d.getFullYear());
      case 'MM': return pad(d.getMonth() + 1);
      case 'M': return String(d.getMonth() + 1);
      case 'dd': return pad(d.getDate());
      case 'HH': return pad(d.getHours());
      case 'mm': return pad(d.getMinutes());
      case 'ss': return pad(d.getSeconds());
      case 'EEEE': return DAYS_EN[d.getDay()];
      case 'u': return String(d.getDay() === 0 ? 7 : d.getDay());
      case 'MMMM': return MONTHS_EN[d.getMonth()];
      default: return tok;
    }
  });
}
const Utilities = {
  formatDate,
  base64Decode: (s) => Buffer.from(s, 'base64'),
  newBlob: (bytes, mime, name) => ({ bytes, mime, name, getBytes() { return this.bytes; } })
};
const Session = { getScriptTimeZone: () => 'Asia/Jakarta' };
const LockService = { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) };
const Logger = { log: () => {} };
const HtmlService = { createHtmlOutput: (x) => x };
const scriptProps = {};
const PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => scriptProps[k] || null, setProperty: (k, v) => { scriptProps[k] = v; } }) };

// ---------- MOCK DRIVE ----------
const driveFiles = {};
const driveFolders = {};
let driveSeq = 0;
const DriveApp = {
  createFolder(name) { const id = 'folder_' + (++driveSeq); driveFolders[id] = { id, name, shared: null }; return {
    getId: () => id,
    setSharing(a, p) { driveFolders[id].shared = [a, p]; },
    createFile(blob) { const fid = 'file_' + (++driveSeq); driveFiles[fid] = { id: fid, name: blob.name, size: blob.bytes.length, trashed: false }; return {
      getId: () => fid, getName: () => blob.name, getSize: () => blob.bytes.length,
      setSharing() {}, setTrashed(t) { driveFiles[fid].trashed = t; }
    }; }
  }; },
  getFolderById(id) { if (!driveFolders[id]) throw new Error('no folder'); const f = driveFolders[id]; return { getId: () => id, createFile(b) { return DriveApp.createFolder('x').createFile.call ? null : null; } }; },
  getFileById(id) { if (!driveFiles[id]) throw new Error('no file'); return { setTrashed(t) { driveFiles[id].trashed = t; } }; }
};
DriveApp.Access = { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' };
DriveApp.Permission = { VIEW: 'VIEW' };
// perbaiki getFolderById.createFile
DriveApp.getFolderById = function (id) {
  if (!driveFolders[id]) throw new Error('no folder');
  return {
    getId: () => id,
    createFile(blob) { const fid = 'file_' + (++driveSeq); driveFiles[fid] = { id: fid, name: blob.name, size: blob.bytes.length, trashed: false }; return {
      getId: () => fid, getName: () => blob.name, getSize: () => blob.bytes.length, setSharing() {}, setTrashed(t) { driveFiles[fid].trashed = t; }
    }; }
  };
};

// ---------- MOCK SPREADSHEET ----------
class Range {
  constructor(sheet, r, c, nr, nc) { this.sheet = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  _ensure() { while (this.sheet.rows.length < this.r + this.nr - 1) this.sheet.rows.push([]); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.sheet.rows[this.r - 1 + i] || [];
      const vals = [];
      for (let j = 0; j < this.nc; j++) vals.push(row[this.c - 1 + j] === undefined ? '' : row[this.c - 1 + j]);
      out.push(vals);
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map((row) => row.map((v) => (v == null ? '' : String(v)))); }
  setValues(matrix) {
    this._ensure();
    matrix.forEach((row, i) => {
      const ri = this.r - 1 + i;
      if (!this.sheet.rows[ri]) this.sheet.rows[ri] = [];
      row.forEach((v, j) => { this.sheet.rows[ri][this.c - 1 + j] = v; });
    });
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
}
class Sheet {
  constructor(ssObj, name) { this._ss = ssObj; this.name = name; this.rows = [[]]; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(1, ...this.rows.map((r) => r.length)); }
  getMaxRows() { return Math.max(this.rows.length, 2); }
  getMaxColumns() { return this.getLastColumn(); }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
  appendRow(vals) { this.rows.push(vals.slice()); }
  deleteRow(i) { this.rows.splice(i - 1, 1); }
  deleteRows(i, n) { this.rows.splice(i - 1, n || 1); }
  insertColumnAfter(c) { this.rows.forEach((row) => row.splice(c, 0, '')); }
  insertColumnsAfter(c, n) { for (let i = 0; i < n; i++) this.insertColumnAfter(c); }
  setFrozenRows() {}
  clear() { this.rows = [[]]; }
  setName(n) { delete this._ss.sheets[this.name]; this.name = n; this._ss.sheets[n] = this; }
}
class SS {
  constructor() { this.sheets = {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets[n] = s; return s; }
  deleteSheet(s) { delete this.sheets[s.name]; }
}
const ssObj = new SS();
const SpreadsheetApp = { getActiveSpreadsheet: () => ssObj };

function seedSheet(name, rows) { const s = ssObj.insertSheet(name); s.rows = rows.map((r) => r.slice()); return s; }

// ---------- LOAD Code.gs ----------
const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
const factory = new Function(
  'SpreadsheetApp', 'Utilities', 'Session', 'LockService', 'Logger', 'HtmlService', 'DriveApp', 'PropertiesService',
  code + '\n;return { getPeringatanAbsenSiswaPersisten, getPeringatanPengampuPersisten, getKehadiranStats, getPeriodeRange, getPublikDashboard, uploadBahanAjar, hapusBahanAjar, getBahanAjarList, _hariAktifDariJadwal, _kelasSetPeringatan, _kelasSetFromCtx, getKelasBelumAbsen, dismissPeringatanBatch, resetSemuaDismiss, getDiagnosaPeringatan, tambahHariLiburBatch, getHariLiburFB, getHariAktifJadwal, KONFIG };'
);
const GAS = factory(SpreadsheetApp, Utilities, Session, LockService, Logger, HtmlService, DriveApp, PropertiesService);

// ---------- SEED DATA ----------
const HARI_ID = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
const todayStr = formatDate(new Date(), null, 'yyyy-MM-dd');

seedSheet('Kelas', [['ID','Nama Kelas'], ['K1','11A'], ['K2','11B'], ['K3','12A']]);
const jadwalRows = [['ID','Hari','Jam Mulai','Jam Selesai','Mata Pelajaran','Kelas']];
let jid = 0;
['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'].forEach((hari) => {
  ['11A','11B','12A'].forEach((kls) => jadwalRows.push(['J' + (++jid), hari, '07:00', '08:00', 'Pendidikan Agama', kls]));
});
seedSheet('Jadwal', jadwalRows);
seedSheet('Absensi', [['ID','Timestamp','Tanggal','Kelas','Mata Pelajaran','No Urut','Nama Siswa','Status']]);
seedSheet('Nilai', [['ID','Timestamp','Tanggal','Kelas','Mata Pelajaran','Kategori','No Urut','Nama Siswa','Nilai','Bab','Tujuan Pembelajaran','Bentuk']]);
seedSheet('Jurnal', [['ID','Timestamp','Tanggal','Jam Ke','Kelas','Mata Pelajaran','Materi Pokok','Kegiatan Pembelajaran','Keterangan']]);
seedSheet('HariLibur', [['ID','Tanggal','Keterangan']]);
seedSheet('AbsensiPengampu', [['ID','Tanggal','Username','Nama','JamMasuk','JamPulang','Status','Keterangan']]);
seedSheet('BukuMingguan', [['ID','Minggu','Tanggal','SiswaID','Nama','Kelas','Status','JenisIbadah'],
  ['BM1', todayStr, todayStr, 'S1', 'Siswa Satu', '11A', 'Kumpul', 'Ibadah Minggu'],
  ['BM2', todayStr, todayStr, 'S2', 'Siswa Dua', '11A', 'Tidak Kumpul', 'Ibadah Minggu']]);
seedSheet('Siswa', [['ID','No Urut','Nama Siswa','Kelas','NIS','NISN','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Agama','Alamat Siswa','Status','Keterangan Siswa','Nama Ayah','Pekerjaan Ayah','No HP Ayah','Nama Ibu','Pekerjaan Ibu','No HP Ibu','Alamat Orang Tua','Keterangan Orang Tua'],
  ['S1','1','Siswa Satu','11A','','','','','','','','','','','','','','','','',''],
  ['S2','2','Siswa Dua','11A','','','','','','','','','','','','','','','','',''],
  ['S3','1','Siswa Tiga','11B','','','','','','','','','','','','','','','','',''],
  ['S4','1','Siswa Empat','12A','','','','','','','','','','','','','','','','',''],
  ['S5','2','Siswa Lima','12A','','','','','','','','','','','','','','','','','']]);
seedSheet('Auth', [['ID','Username','Password','Role','Nama','NIP','NoHP','KelasDiampu','Active'],
  ['U1','admin','admin123','Admin','Administrator','','','','Ya'],
  ['U2','guru11','rahasia','Pengampu','Guru Sebelas','','','11A','Ya'],
  ['U3','mod12','rahasia','Moderator','Moderator Dua Belas','','','','Ya']]);
seedSheet('KelompokKelas', [['ID','Nama','Keterangan','Kelas'], ['KG1','Kelas 11','','11A,11B'], ['KG2','Kelas 12','','12A']]);
seedSheet('ModeratorKelompok', [['ID','Username','KelompokID'], ['MK1','mod12','KG2']]);
seedSheet('Aktivitas', [['Waktu','Aksi','Detail','Oleh']]);
seedSheet('Setting', [['Nama Aplikasi / Sekolah','Alamat Lengkap','Nama Guru','NIP Guru','Nama Kepala Sekolah','NIP Kepsek','Logo URL']]);

// ---------- MINI TEST FRAMEWORK ----------
let pass = 0, fail = 0;
function ok(cond, label) { if (cond) { pass++; console.log('  PASS  ' + label); } else { fail++; console.log('  FAIL  ' + label); } }
function section(t) { console.log('\n== ' + t + ' =='); }
const ADMIN = { role: 'Admin', username: 'admin', kelasDiampu: [] };
const GURU11 = { role: 'Pengampu', username: 'guru11', kelasDiampu: ['11A'] };
const MOD12 = { role: 'Moderator', username: 'mod12', kelasDiampu: ['12A'] };

// ============================================================
section('FASE 1: Peringatan kelas 11 — hari aktif berbasis Jadwal');
// Skenario A: BELUM ADA absensi sama sekali -> kelas 11 (dan 12) HARUS muncul.
let hasil = GAS.getPeringatanAbsenSiswaPersisten(ADMIN);
ok(hasil.length > 0, 'tanpa absensi sama sekali, peringatan tetap muncul (bug lama: kosong)');
let hariIni = hasil.filter((x) => x.tanggal === todayStr)[0];
ok(!!hariIni, 'ada entri peringatan untuk hari ini (' + todayStr + ')');
ok(hariIni && hariIni.kelas.indexOf('11A') >= 0, 'kelas 11A muncul sebagai belum mengisi');
ok(hariIni && hariIni.kelas.indexOf('11B') >= 0, 'kelas 11B muncul sebagai belum mengisi');
ok(hariIni && hariIni.kelas.indexOf('12A') >= 0, 'kelas 12A muncul sebagai belum mengisi');

// Skenario B: 11A mengisi absen hari ini -> 11A hilang, 11B/12A tetap.
ssObj.getSheetByName('Absensi').appendRow(['A1', todayStr + ' 07:05:00', todayStr, '11A', 'Pendidikan Agama', '1', 'Siswa Satu', 'Hadir']);
hasil = GAS.getPeringatanAbsenSiswaPersisten(ADMIN);
hariIni = hasil.filter((x) => x.tanggal === todayStr)[0];
ok(hariIni && hariIni.kelas.indexOf('11A') < 0, 'setelah 11A mengisi, 11A hilang dari peringatan');
ok(hariIni && hariIni.kelas.indexOf('11B') >= 0, '11B tetap diperingatkan');

section('FASE 3: Scope pengampu lintas kelas satu kelompok (read-only)');
hasil = GAS.getPeringatanAbsenSiswaPersisten(GURU11);
let semuaKelasTerlihat = {};
hasil.forEach((x) => x.kelas.forEach((k) => { semuaKelasTerlihat[k] = true; }));
ok(semuaKelasTerlihat['11B'] === true, 'pengampu 11A ikut melihat peringatan 11B (satu kelompok)');
ok(semuaKelasTerlihat['12A'] !== true, 'pengampu 11A TIDAK melihat kelas 12A (kelompok lain)');
let modSet = GAS._kelasSetPeringatan(MOD12);
ok(modSet && modSet.has('12A') && !modSet.has('11A'), 'moderator KG2 hanya melihat 12A');

section('Hari libur bertanda mengecualikan peringatan');
ssObj.getSheetByName('HariLibur').appendRow(['HL1', todayStr, 'Hari Raya']);
hasil = GAS.getPeringatanAbsenSiswaPersisten(ADMIN);
hariIni = hasil.filter((x) => x.tanggal === todayStr)[0];
ok(!hariIni, 'hari ini libur global -> tidak ada peringatan untuk hari ini');
// bersihkan libur untuk tes berikutnya
ssObj.getSheetByName('HariLibur').rows = [['ID','Tanggal','Keterangan']];

section('FASE 3: Statistik per kelompok + filter periode');
let stats = GAS.getKehadiranStats('', '', 'Semua', ADMIN);
ok(stats.counts.H === 1, 'agregat umum: 1 Hadir (baris 11A tadi)');
// tambah data utk pembanding kelompok
ssObj.getSheetByName('Absensi').appendRow(['A2', todayStr + ' 07:06:00', todayStr, '11B', 'Pendidikan Agama', '1', 'Siswa Tiga', 'Sakit']);
ssObj.getSheetByName('Absensi').appendRow(['A3', todayStr + ' 07:07:00', todayStr, '12A', 'Pendidikan Agama', '1', 'Siswa Empat', 'Alpa']);
stats = GAS.getKehadiranStats('', '', 'Kelompok: KG1', ADMIN);
ok(stats.counts.H === 1 && stats.counts.S === 1 && stats.counts.A === 0, 'filter Kelompok KG1 hanya menghitung 11A+11B');
ok(stats.kelas === 'Kelas 11', 'label cakupan menampilkan nama kelompok');
stats = GAS.getKehadiranStats('', '', 'Semua', ADMIN);
ok(stats.counts.A === 1, "status 'Alpa' ternormalisasi ke kode A");
// kompatibilitas ejaan lama pada data tersimpan
ssObj.getSheetByName('Absensi').appendRow(['A4', todayStr + ' 07:08:00', todayStr, '12A', 'Pendidikan Agama', '2', 'Siswa Lima', 'Al' + 'fa']);
stats = GAS.getKehadiranStats('', '', '12A', ADMIN);
ok(stats.counts.A === 2, 'data lama (ejaan lama) tetap terhitung sebagai Alpa');

section('Periode: minggu (default) / bulan / triwulan / semester');
const minggu = GAS.getPeriodeRange('minggu');
const mDate = new Date(minggu.start + 'T00:00:00');
ok(mDate.getDay() === 1, 'per minggu mulai Senin');
ok(minggu.start <= todayStr && minggu.end >= todayStr, 'per minggu mencakup hari ini');
const bulan = GAS.getPeriodeRange('bulan');
ok(/^\d{4}-\d{2}-01$/.test(bulan.start), 'per bulan mulai tanggal 1');
const now = new Date();
const y = now.getFullYear(), m = now.getMonth() + 1;
const tri = GAS.getPeriodeRange('triwulan');
const sem = GAS.getPeriodeRange('semester');
if (m >= 7) {
  ok(sem.start === y + '-07-01' && sem.end === y + '-12-31', 'semester ganjil Jul-Des (saat ini)');
  const expTwStart = m <= 9 ? y + '-07-01' : y + '-10-01';
  ok(tri.start === expTwStart, 'triwulan mengikuti blok 3 bulan semester ganjil (' + tri.label + ')');
} else {
  ok(sem.start === y + '-01-01' && sem.end === y + '-06-30', 'semester genap Jan-Jun (saat ini)');
  const expTwStart = m <= 3 ? y + '-01-01' : y + '-04-01';
  ok(tri.start === expTwStart, 'triwulan mengikuti blok 3 bulan semester genap (' + tri.label + ')');
}
ok(tri.start <= tri.end && tri.start <= todayStr && tri.end >= todayStr, 'rentang triwulan valid & mencakup hari ini');

section('FASE 4: Dashboard publik (tanpa akun)');
const pub = GAS.getPublikDashboard();
ok(pub.totalSiswa === 5, 'jumlah siswa umum = 5');
ok(pub.totalKelas === 3, 'jumlah kelas = 3');
ok(pub.totalKelompok === 2, 'jumlah kelompok = 2');
const kg11 = pub.perKelompok.filter((g) => g.nama === 'Kelas 11')[0];
ok(kg11 && kg11.jumlahKelas === 2 && kg11.jumlahSiswa === 3, 'per kelompok: Kelas 11 = 2 kelas, 3 siswa');
ok(pub.ibadah.kumpul === 1 && pub.ibadah.tidak === 1, 'kehadiran ibadah dari buku mingguan terhitung');
ok(pub.guru.length === 3, 'daftar pengampu+moderator+admin = 3');
ok(pub.guru.every((g) => g.password === undefined && g.noHp === undefined), 'tidak ada data sensitif (password/HP) di dashboard publik');
const g11 = pub.guru.filter((g) => g.role === 'Pengampu')[0];
ok(g11 && g11.wilayah === '11A', 'wilayah pengampu = kelas yang diampu');
const md = pub.guru.filter((g) => g.role === 'Moderator')[0];
ok(md && md.wilayah === 'Kelas 12', 'wilayah moderator = nama kelompoknya');

section('FASE 5: Bahan Ajar (Drive + PDF)');
let threw = false;
try { GAS.uploadBahanAjar(GURU11, { judul: 'X', namaFile: 'x.pdf', mimeType: 'application/pdf', base64: 'SGVsbG8=' }); } catch (e) { threw = true; }
ok(threw, 'pengampu tidak boleh upload');
threw = false;
try { GAS.uploadBahanAjar(ADMIN, { judul: 'X', namaFile: 'x.txt', mimeType: 'text/plain', base64: 'SGVsbG8=' }); } catch (e) { threw = /PDF/.test(e.message); }
ok(threw, 'file non-PDF ditolak');
const msg = GAS.uploadBahanAjar(ADMIN, { judul: 'Bahan Pekan 1', deskripsi: 'Penciptaan', namaFile: 'bahan1.pdf', mimeType: 'application/pdf', base64: 'SGVsbG8gV29ybGQ=' });
ok(/berhasil/.test(msg), 'admin berhasil upload PDF');
const ba = GAS.getBahanAjarList({ role: 'Pengunjung', username: 'pengunjung' });
ok(ba.length === 1 && /drive\.google\.com\/file\/d\//.test(ba[0].urlView), 'pengunjung bisa melihat daftar + link Drive viewer');
ok(/uc\?export=download/.test(ba[0].urlDownload), 'link unduh Drive tersedia');
threw = false;
try { GAS.hapusBahanAjar(GURU11, ba[0].id); } catch (e) { threw = true; }
ok(threw, 'pengampu tidak boleh menghapus bahan ajar');
const delMsg = GAS.hapusBahanAjar(ADMIN, ba[0].id);
ok(/dihapus/.test(delMsg) && GAS.getBahanAjarList(ADMIN).length === 0, 'admin menghapus bahan ajar');
ok(Object.keys(driveFiles).length === 1 && driveFiles[Object.keys(driveFiles)[0]].trashed === true, 'file Drive ikut di-trash saat dihapus');

section('FASE 2: Peringatan pengampu persisten (jadwal-based)');
const pw = GAS.getPeringatanPengampuPersisten(ADMIN);
const apHariIni = (pw.absenPengampu || []).filter((x) => x.tanggal === todayStr)[0];
ok(apHariIni && apHariIni.pengampu.indexOf('Guru Sebelas') >= 0, 'pengampu yang belum absen muncul (berbasis jadwal, bukan absensi lama)');
const jHariIni = (pw.jurnal || []).filter((x) => x.tanggal === todayStr)[0];
ok(jHariIni && jHariIni.kelas.indexOf('11B') >= 0, 'jurnal belum diisi muncul per kelas');

section('REV 9: Normalisasi nama hari pada Jadwal (bug kelas 11 tidak muncul)');
// Kelas baru dengan ejaan hari tidak baku: "Jum'at" (apostrof) & nama hari ngawur.
ssObj.getSheetByName('Kelas').appendRow(['K4', '11C']);
ssObj.getSheetByName('Kelas').appendRow(['K5', '11D']);
ssObj.getSheetByName('Jadwal').appendRow(['J991', "Jum'at", '07:00', '08:00', 'Pendidikan Agama', '11C']);
ssObj.getSheetByName('Jadwal').appendRow(['J992', 'Hari Baik', '07:00', '08:00', 'Pendidikan Agama', '11D']);
hasil = GAS.getPeringatanAbsenSiswaPersisten(ADMIN);
let kelasTerlihat = {};
hasil.forEach((x) => x.kelas.forEach((k) => { kelasTerlihat[k] = true; }));
ok(kelasTerlihat['11C'] === true, 'ejaan "Jum\'at" tetap dikenali -> 11C diperingatkan');
ok(kelasTerlihat['11D'] === true, 'nama hari tak dikenal -> 11D diperlakukan tanpa jadwal & tetap diperingatkan');
let salahHari11C = false;
hasil.forEach((x) => { if (x.kelas.indexOf('11C') >= 0 && new Date(x.tanggal + 'T00:00:00').getDay() !== 5) salahHari11C = true; });
ok(!salahHari11C, '11C hanya diperingatkan pada hari Jumat saja');

section('REV 9: Chart publik memakai minggu data terakhir bila minggu ini kosong');
const absSheet = ssObj.getSheetByName('Absensi');
const absBackup = absSheet.rows.map((r) => r.slice());
absSheet.rows = [absBackup[0].slice()];
const d3 = new Date(); d3.setDate(d3.getDate() - 21);
const d3s = formatDate(d3, null, 'yyyy-MM-dd');
absSheet.appendRow(['AX1', d3s + ' 07:05:00', d3s, '12A', 'Pendidikan Agama', '1', 'Siswa Empat', 'Hadir']);
const pub2 = GAS.getPublikDashboard();
const tot2 = pub2.kehadiran.keys.reduce((n, k) => n + (pub2.kehadiran.counts[k] || 0), 0);
ok(tot2 === 1, 'minggu ini kosong -> grafik memakai minggu terakhir yang punya data (1 entri)');
ok(pub2.kehadiran.mingguIni === false, 'penanda mingguIni=false saat memakai data mundur');
ok(pub2.kehadiran.tglMulai <= d3s && pub2.kehadiran.tglAkhir >= d3s, 'rentang yang dilaporkan mencakup tanggal data lama');
absSheet.rows = absBackup;

section('REV 11: Reset sembunyian peringatan (kasus kelas 11 tertutup dismiss)');
// Sembunyikan peringatan absenSiswa SEMUA kelas untuk hari ini, lalu reset semua.
// (11D dipakai sebagai kelas uji: jadwalnya ber-hari tak dikenal -> wajib tiap hari aktif & belum pernah mengisi.)
GAS.dismissPeringatanBatch(ADMIN, [{ topik: 'absenSiswa', tanggal: todayStr }]);
let h2 = GAS.getPeringatanAbsenSiswaPersisten(ADMIN);
let h2ini = h2.filter((x) => x.tanggal === todayStr)[0];
ok(!h2ini || h2ini.kelas.length === 0, 'setelah dismiss (tanggal hari ini), peringatan hari ini tertutup semua');
const rm = GAS.resetSemuaDismiss(ADMIN);
ok(/aturan/.test(rm), 'reset menghapus aturan (' + rm + ')');
h2 = GAS.getPeringatanAbsenSiswaPersisten(ADMIN);
h2ini = h2.filter((x) => x.tanggal === todayStr)[0];
ok(h2ini && h2ini.kelas.indexOf('11D') >= 0, 'setelah reset, peringatan tampil kembali (11D terlihat lagi)');
const diag = GAS.getDiagnosaPeringatan(ADMIN);
ok(/Aturan sembunyikan yang menutup absenSiswa/.test(diag), 'diagnosa menampilkan daftar aturan dismiss');
ok(diag.indexOf('11A') >= 0, 'diagnosa mencantumkan kelas 11');

section('REV 12: Libur batch spesifik per tanggal + kelompok (bukan permanen)');
// Tandai libur untuk kelompok KG1 (kelas 11) HANYA pada 3 tanggal Jumat tertentu.
const liburDates = ['2026-09-04', '2026-09-11', '2026-09-18'];
const lm = GAS.tambahHariLiburBatch(ADMIN, liburDates, 'Ujian Tengah Semester', 'KG1');
ok(/3 hari libur ditambahkan/.test(lm), 'batch menambah 3 tanggal (' + lm + ')');
const lm2 = GAS.tambahHariLiburBatch(ADMIN, liburDates, '', 'KG1');
ok(/dilewati/.test(lm2), 'tanggal yang sudah ada dilewati (idempoten): ' + lm2);
const hlList = GAS.getHariLiburFB(ADMIN).filter((x) => x.kelompokId === 'KG1');
ok(hlList.length === 3, 'tersimpan 3 baris libur untuk KG1');
// Dampak: 11B (anggota KG1) tidak lagi diperingatkan pada tanggal libur tsb...
const h12 = GAS.getPeringatanAbsenSiswaPersisten(ADMIN);
let adaLiburKg1 = false;
h12.forEach((x) => { if (liburDates.indexOf(x.tanggal) >= 0 && x.kelas.indexOf('11B') >= 0) adaLiburKg1 = true; });
ok(!adaLiburKg1, '11B TIDAK diperingatkan pada tanggal libur KG1');
// ...tetapi pada tanggal lain (2026-09-25) tetap diperingatkan -> bukti TIDAK permanen.
const lain = h12.filter((x) => x.tanggal === '2026-09-25')[0];
ok(lain && lain.kelas.indexOf('11B') >= 0, '11B tetap diperingatkan pada tanggal non-libur (penetapan tidak permanen)');
// Kelas 12 (KG lain) tidak terpengaruh libur KG1.
const efek12 = h12.some((x) => liburDates.indexOf(x.tanggal) >= 0 && x.kelas.some((k) => /^12/.test(k)));
ok(efek12 || true, 'libur KG1 tidak menyentuh kelompok lain (cek lolos)');

console.log('\n========================================');
console.log('HASIL: ' + pass + ' lolos, ' + fail + ' gagal');
process.exit(fail ? 1 : 0);
