/**
 * =========================================================
 *  DISPATCHER SEJASA
 *  Backend Google Apps Script
 * =========================================================
 *  Struktur Google Sheet yang dibutuhkan (buat manual sekali):
 *
 *  Tab "Izin" - header di baris 1:
 *  ID | Timestamp Submit | Nama Mitra | Kategori | Tanggal Mulai |
 *  Tanggal Selesai | Jumlah Hari | Keterangan | Link Lampiran |
 *  Status | Diproses Oleh | Timestamp Diproses | Catatan Admin
 *
 *  Tab "Mitra" - header di baris 1:
 *  Nama Mitra | PIN
 *  (isi manual: kolom A nama mitra, kolom B PIN 4 digit bebas,
 *   satu baris per mitra. PIN ini yang dipakai mitra buat isi
 *   pengajuan izin & buka riwayat izin mereka sendiri di tab
 *   "Riwayat Saya" — mencegah orang lain isi izin atau intip
 *   riwayat pakai nama mitra lain.)
 *
 *  Tab "Kategori" - dibuat OTOMATIS sama script kalau belum ada
 *  (isi default: Sakit, Keperluan Keluarga, Cuti Tahunan, Lainnya).
 *  Admin bisa tambah/hapus kategori langsung dari panel Admin,
 *  gak perlu edit sheet manual.
 *
 *  Setelah itu jalankan fungsi setupAdmin() SEKALI dari editor
 *  (klik dropdown function -> pilih setupAdmin -> Run) untuk
 *  set username & password admin pertama kali.
 * =========================================================
 */

const SHEET_IZIN = 'Izin';
const SHEET_MITRA = 'Mitra';
const SHEET_KATEGORI = 'Kategori';
const DEFAULT_KATEGORI = ['Sakit', 'Keperluan Keluarga', 'Cuti Tahunan', 'Lainnya'];
const DRIVE_FOLDER_NAME = 'Lampiran Izin Mitra';
const SESSION_DURATION_SEC = 60 * 60 * 4; // token admin valid 4 jam

// ------------------- ROUTING HALAMAN -------------------

function doGet(e) {
  const page = e && e.parameter && e.parameter.page === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'Admin' ? 'Dispatcher Sejasa - Admin' : 'Dispatcher Sejasa - Form Pengajuan Izin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ------------------- API UNTUK FRONTEND EKSTERNAL (GitHub Pages, dsb) -------------------
//
// Dipanggil via fetch() POST dari index.html yang di-hosting di luar
// script.google.com. Body request harus JSON:
//   { "action": "namaFungsi", "params": [arg1, arg2, ...] }
// Hanya fungsi yang terdaftar di API_ACTIONS_ yang boleh dipanggil dari luar.

const API_ACTIONS_ = {
  getMitraList: getMitraList,
  getKategoriList: getKategoriList,
  getFormInitData: getFormInitData,
  submitIzin: submitIzin,
  getRiwayatMitra: getRiwayatMitra,
  updateIzin: updateIzin,
  adminLogin: adminLogin,
  getPendingIzin: getPendingIzin,
  getRiwayatIzin: getRiwayatIzin,
  processIzin: processIzin,
  getLaporan: getLaporan,
  tambahKategori: tambahKategori,
  hapusKategori: hapusKategori
};

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput_({ success: false, message: 'Request kosong.' });
    }
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const params = Array.isArray(body.params) ? body.params : [];

    const fn = API_ACTIONS_[action];
    if (!fn) {
      return jsonOutput_({ success: false, message: 'Action tidak dikenal: ' + action });
    }

    const result = fn.apply(null, params);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, message: 'Server error: ' + err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------- HELPER -------------------

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan. Cek nama tab.');
  return sheet;
}

function getOrCreateDriveFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function generateId_() {
  return Utilities.formatString('IZ-%s', Utilities.getUuid().split('-')[0].toUpperCase());
}

function countWeekdays_(startStr, endStr) {
  const start = new Date(startStr);
  const end = new Date(endStr);
  let count = 0;
  const cur = new Date(start);
  while (cur <= end) {
    count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count > 0 ? count : 1;
}

function sameDate_(a, b) {
  const da = new Date(a);
  const db = new Date(b);
  return da.getFullYear() === db.getFullYear() &&
    da.getMonth() === db.getMonth() &&
    da.getDate() === db.getDate();
}

/**
 * Cek apakah mitra ini udah punya pengajuan lain dengan kategori +
 * tanggal mulai + tanggal selesai yang sama persis. Pengajuan yang
 * sudah Rejected diabaikan (mitra boleh ajukan ulang kalau ditolak).
 * excludeId dipakai pas edit, biar gak ke-detect nabrak diri sendiri.
 * Return status pengajuan yang bentrok ('Pending'/'Approved'), atau
 * null kalau gak ada yang bentrok.
 */
function findKonflikIzin_(nama, kategori, tanggalMulai, tanggalSelesai, excludeId) {
  const sheet = getSheet_(SHEET_IZIN);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0] || r[0] === excludeId) continue;
    if (r[2] !== nama || r[3] !== kategori) continue;
    if (r[9] === 'Rejected') continue;
    if (sameDate_(r[4], tanggalMulai) && sameDate_(r[5], tanggalSelesai)) {
      return r[9];
    }
  }
  return null;
}

/**
 * Cek nama + PIN cocok dengan data di tab Mitra.
 * Dipakai baik saat submit izin (cegah orang isi atas nama orang lain)
 * maupun saat buka riwayat (cegah orang intip riwayat orang lain).
 */
function verifyMitraPin_(nama, pin) {
  if (!nama || !pin) return false;
  const mitraSheet = getSheet_(SHEET_MITRA);
  const mitraValues = mitraSheet.getDataRange().getValues();
  for (let i = 1; i < mitraValues.length; i++) {
    if (mitraValues[i][0] === nama) {
      const storedPin = String(mitraValues[i][1] || '').trim();
      return !!(storedPin && storedPin === String(pin).trim());
    }
  }
  return false;
}

// ------------------- SETUP ADMIN (jalankan manual sekali) -------------------

function setupAdmin() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_USERNAME', 'admin');
  props.setProperty('ADMIN_PASSWORD', 'ubahPasswordIni123');
  Logger.log('Admin default dibuat. Username: admin / Password: ubahPasswordIni123');
  Logger.log('SEGERA ganti password lewat fungsi ubahPasswordAdmin() setelah login pertama.');
}

function ubahPasswordAdmin(passwordBaru) {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', passwordBaru);
}

// ------------------- USER: AMBIL LIST MITRA -------------------

function getMitraList() {
  const sheet = getSheet_(SHEET_MITRA);
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) list.push(values[i][0]);
  }
  list.sort((a, b) => a.localeCompare(b, 'id', { sensitivity: 'base' }));
  return list;
}

// ------------------- KATEGORI IZIN (bisa diatur dari Admin) -------------------

function getOrCreateKategoriSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_KATEGORI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KATEGORI);
    sheet.appendRow(['Nama Kategori']);
    DEFAULT_KATEGORI.forEach(k => sheet.appendRow([k]));
  }
  return sheet;
}

function getKategoriList() {
  const sheet = getOrCreateKategoriSheet_();
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) list.push(String(values[i][0]).trim());
  }
  return list.length ? list : DEFAULT_KATEGORI.slice();
}

/**
 * Dipanggil sekali pas form user/admin dibuka, gabungin data awal
 * yang dibutuhkan biar gak perlu beberapa kali round-trip terpisah.
 */
function getFormInitData() {
  return {
    mitraList: getMitraList(),
    kategoriList: getKategoriList()
  };
}

function tambahKategori(token, nama) {
  verifyToken_(token);
  nama = String(nama || '').trim();
  if (!nama) return { success: false, message: 'Nama kategori tidak boleh kosong.' };

  const existing = getKategoriList();
  if (existing.some(k => k.toLowerCase() === nama.toLowerCase())) {
    return { success: false, message: 'Kategori "' + nama + '" sudah ada.' };
  }

  const sheet = getOrCreateKategoriSheet_();
  sheet.appendRow([nama]);
  return { success: true, message: 'Kategori "' + nama + '" berhasil ditambahkan.', list: getKategoriList() };
}

function hapusKategori(token, nama) {
  verifyToken_(token);
  const sheet = getOrCreateKategoriSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === nama) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Kategori "' + nama + '" berhasil dihapus.', list: getKategoriList() };
    }
  }
  return { success: false, message: 'Kategori tidak ditemukan.' };
}

// ------------------- USER: SUBMIT IZIN -------------------

/**
 * data = {
 *   nama, pin, kategori, tanggalMulai, tanggalSelesai, keterangan,
 *   fileBase64, fileName, fileMime  // fileBase64 boleh kosong kalau tanpa lampiran
 * }
 */
function submitIzin(data) {
  try {
    if (!data.nama || !data.kategori || !data.tanggalMulai || !data.tanggalSelesai) {
      return { success: false, message: 'Mohon lengkapi semua field wajib.' };
    }
    if (new Date(data.tanggalSelesai) < new Date(data.tanggalMulai)) {
      return { success: false, message: 'Tanggal Selesai tidak boleh sebelum Tanggal Mulai.' };
    }
    if (!data.pin) {
      return { success: false, message: 'PIN wajib diisi.' };
    }
    if (!verifyMitraPin_(data.nama, data.pin)) {
      return { success: false, message: 'PIN salah. Pastikan nama dan PIN sesuai data Anda.' };
    }
    const konflik = findKonflikIzin_(data.nama, data.kategori, data.tanggalMulai, data.tanggalSelesai, null);
    if (konflik) {
      return { success: false, message: 'Anda sudah punya pengajuan izin dengan kategori dan tanggal yang sama (status: ' + konflik + '). Cek tab Riwayat Saya, atau edit pengajuan yang sudah ada kalau statusnya masih Pending.' };
    }
    if (!data.fileBase64) {
      return { success: false, message: 'Lampiran wajib diupload.' };
    }
    // Base64 blows up ~33% dari ukuran asli; 7,000,000 karakter ~ 5MB file asli.
    if (data.fileBase64.length > 7000000) {
      return { success: false, message: 'Ukuran lampiran terlalu besar. Maksimal 5MB.' };
    }

    let linkLampiran = '';
    if (data.fileBase64) {
      const folder = getOrCreateDriveFolder_();
      const decoded = Utilities.base64Decode(data.fileBase64);
      const blob = Utilities.newBlob(decoded, data.fileMime, data.fileName);
      const safeName = Utilities.formatString('%s_%s_%s',
        data.nama.replace(/\s+/g, ''), new Date().getTime(), data.fileName);
      blob.setName(safeName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkLampiran = file.getUrl();
    }

    const sheet = getSheet_(SHEET_IZIN);
    const id = generateId_();
    const jumlahHari = countWeekdays_(data.tanggalMulai, data.tanggalSelesai);

    sheet.appendRow([
      id,
      new Date(),
      data.nama,
      data.kategori,
      data.tanggalMulai,
      data.tanggalSelesai,
      jumlahHari,
      data.keterangan || '',
      linkLampiran,
      'Pending',
      '',
      '',
      ''
    ]);

    return { success: true, message: 'Pengajuan izin berhasil dikirim. ID: ' + id };
  } catch (err) {
    return { success: false, message: 'Terjadi error: ' + err.message };
  }
}

// ------------------- USER: RIWAYAT IZIN SAYA -------------------

/**
 * Self-service: mitra lihat riwayat izin miliknya sendiri berdasarkan nama
 * + PIN yang cocok dengan data di tab Mitra. PIN mencegah orang lain
 * membuka riwayat izin (yang berisi alasan sakit/keluarga, dsb) atas
 * nama orang lain.
 */
function getRiwayatMitra(nama, pin) {
  if (!nama || !pin) return { success: false, message: 'Nama dan PIN wajib diisi.' };

  if (!verifyMitraPin_(nama, pin)) {
    return { success: false, message: 'PIN salah. Coba lagi atau hubungi admin kalau lupa PIN.' };
  }

  const sheet = getSheet_(SHEET_IZIN);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  let approved = 0, pending = 0, rejected = 0;
  const byKategori = {};

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0] || r[2] !== nama) continue;

    rows.push({
      id: r[0],
      kategori: r[3],
      tanggalMulai: r[4],
      tanggalSelesai: r[5],
      jumlahHari: r[6],
      keterangan: r[7],
      status: r[9],
      catatanAdmin: r[12]
    });

    if (r[9] === 'Approved') approved++;
    else if (r[9] === 'Pending') pending++;
    else if (r[9] === 'Rejected') rejected++;

    if (r[9] === 'Approved') {
      byKategori[r[3]] = (byKategori[r[3]] || 0) + 1;
    }
  }

  rows.reverse(); // terbaru duluan

  return {
    success: true,
    rows: rows,
    totalPengajuan: rows.length,
    approved: approved,
    pending: pending,
    rejected: rejected,
    byKategori: byKategori
  };
}

// ------------------- USER: EDIT PENGAJUAN (selama masih Pending) -------------------

/**
 * Mitra edit pengajuan izinnya sendiri, TAPI hanya kalau statusnya
 * masih Pending (belum diproses admin). Begitu admin approve/reject,
 * datanya terkunci dan gak bisa diubah lagi lewat sini.
 *
 * data = { kategori, tanggalMulai, tanggalSelesai, keterangan }
 */
function updateIzin(nama, pin, id, data) {
  if (!verifyMitraPin_(nama, pin)) {
    return { success: false, message: 'PIN salah. Pastikan nama dan PIN sesuai data Anda.' };
  }
  if (!data || !data.kategori || !data.tanggalMulai || !data.tanggalSelesai) {
    return { success: false, message: 'Mohon lengkapi semua field wajib.' };
  }
  if (new Date(data.tanggalSelesai) < new Date(data.tanggalMulai)) {
    return { success: false, message: 'Tanggal Selesai tidak boleh sebelum Tanggal Mulai.' };
  }

  const konflik = findKonflikIzin_(nama, data.kategori, data.tanggalMulai, data.tanggalSelesai, id);
  if (konflik) {
    return { success: false, message: 'Anda sudah punya pengajuan izin lain dengan kategori dan tanggal yang sama (status: ' + konflik + ').' };
  }

  const sheet = getSheet_(SHEET_IZIN);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r[0] !== id || r[2] !== nama) continue;

    if (r[9] !== 'Pending') {
      return { success: false, message: 'Pengajuan ini sudah diproses admin (' + r[9] + ') dan tidak bisa diedit lagi.' };
    }

    const rowIndex = i + 1;
    const jumlahHari = countWeekdays_(data.tanggalMulai, data.tanggalSelesai);
    sheet.getRange(rowIndex, 4).setValue(data.kategori);       // Kategori
    sheet.getRange(rowIndex, 5).setValue(data.tanggalMulai);   // Tanggal Mulai
    sheet.getRange(rowIndex, 6).setValue(data.tanggalSelesai); // Tanggal Selesai
    sheet.getRange(rowIndex, 7).setValue(jumlahHari);          // Jumlah Hari
    sheet.getRange(rowIndex, 8).setValue(data.keterangan || ''); // Keterangan

    return { success: true, message: 'Pengajuan berhasil diperbarui.' };
  }

  return { success: false, message: 'Pengajuan tidak ditemukan.' };
}

// ------------------- ADMIN: AUTH -------------------

function adminLogin(username, password) {
  const props = PropertiesService.getScriptProperties();
  const validUser = props.getProperty('ADMIN_USERNAME');
  const validPass = props.getProperty('ADMIN_PASSWORD');

  if (!validUser || !validPass) {
    return { success: false, message: 'Admin belum di-setup. Jalankan setupAdmin() dulu di editor.' };
  }
  if (username === validUser && password === validPass) {
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('session_' + token, username, SESSION_DURATION_SEC);
    return { success: true, token: token, username: username };
  }
  return { success: false, message: 'Username atau password salah.' };
}

function verifyToken_(token) {
  const cached = CacheService.getScriptCache().get('session_' + token);
  if (!cached) throw new Error('Sesi habis atau belum login. Silakan login ulang.');
  return cached;
}

// ------------------- ADMIN: DATA -------------------

function getAllIzinRows_() {
  const sheet = getSheet_(SHEET_IZIN);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    rows.push({
      rowIndex: i + 1,
      id: r[0],
      timestampSubmit: r[1],
      nama: r[2],
      kategori: r[3],
      tanggalMulai: r[4],
      tanggalSelesai: r[5],
      jumlahHari: r[6],
      keterangan: r[7],
      linkLampiran: r[8],
      status: r[9],
      diprosesOleh: r[10],
      timestampDiproses: r[11],
      catatanAdmin: r[12]
    });
  }
  return rows;
}

function getPendingIzin(token) {
  verifyToken_(token);
  return getAllIzinRows_().filter(r => r.status === 'Pending');
}

function getRiwayatIzin(token) {
  verifyToken_(token);
  return getAllIzinRows_().filter(r => r.status !== 'Pending');
}

function processIzin(token, rowIndex, action, catatan) {
  const username = verifyToken_(token);
  const sheet = getSheet_(SHEET_IZIN);
  const status = action === 'approve' ? 'Approved' : 'Rejected';

  sheet.getRange(rowIndex, 10).setValue(status);       // Status
  sheet.getRange(rowIndex, 11).setValue(username);      // Diproses Oleh
  sheet.getRange(rowIndex, 12).setValue(new Date());    // Timestamp Diproses
  sheet.getRange(rowIndex, 13).setValue(catatan || '');// Catatan Admin

  return { success: true };
}

// ------------------- ADMIN: LAPORAN -------------------

function getLaporan(token, periode, customFrom, customTo) {
  // periode: 'weekly', 'monthly', atau 'custom' (pakai customFrom & customTo, format YYYY-MM-DD)
  verifyToken_(token);
  const rows = getAllIzinRows_().filter(r => r.status === 'Approved');
  const now = new Date();
  let cutoffStart = new Date();
  let cutoffEnd = now;

  if (periode === 'weekly') {
    cutoffStart.setDate(now.getDate() - 7);
  } else if (periode === 'custom') {
    if (!customFrom || !customTo) {
      return { success: false, message: 'Pilih tanggal Dari dan Sampai dulu.' };
    }
    cutoffStart = new Date(customFrom);
    cutoffEnd = new Date(customTo);
    cutoffEnd.setHours(23, 59, 59, 999); // biar tanggal "Sampai" ikut kehitung penuh
    if (cutoffEnd < cutoffStart) {
      return { success: false, message: 'Tanggal "Sampai" tidak boleh sebelum tanggal "Dari".' };
    }
  } else {
    cutoffStart.setMonth(now.getMonth() - 1);
  }

  const filtered = rows.filter(r => {
    const tgl = new Date(r.tanggalMulai);
    return tgl >= cutoffStart && tgl <= cutoffEnd;
  });

  const byKategori = {};
  const byMitra = {};
  filtered.forEach(r => {
    byKategori[r.kategori] = (byKategori[r.kategori] || 0) + 1;
    byMitra[r.nama] = (byMitra[r.nama] || 0) + Number(r.jumlahHari || 1);
  });

  const topMitra = Object.keys(byMitra)
    .map(nama => ({ nama, totalHari: byMitra[nama] }))
    .sort((a, b) => b.totalHari - a.totalHari);

  return {
    success: true,
    totalPengajuan: filtered.length,
    byKategori,
    topMitra,
    periode,
    customFrom: customFrom || '',
    customTo: customTo || ''
  };
}
