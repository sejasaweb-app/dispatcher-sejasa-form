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
  hapusKategori: hapusKategori,
  getKategoriWilayahMap: getKategoriWilayahMap,
  getKotaKecamatanMap: getKotaKecamatanMap,
  getWilayahConfig: getWilayahConfig,
  getDataMitraList: getDataMitraList,
  getDataMitraInitData: getDataMitraInitData,
  tambahDataMitra: tambahDataMitra,
  updateDataMitra: updateDataMitra,
  hapusDataMitra: hapusDataMitra,
  importDataMitraBulk: importDataMitraBulk,
  hapusIzin: hapusIzin,
  getRingkasan: getRingkasan,
  getGajiTerbaru: getGajiTerbaru,
  getMitraPinList: getMitraPinList,
  tambahMitraPin: tambahMitraPin,
  updateMitraPin: updateMitraPin,
  hapusMitraPin: hapusMitraPin,
  getGajiList: getGajiList,
  tambahGaji: tambahGaji,
  hapusGaji: hapusGaji,
  importGajiBulk: importGajiBulk
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
  // periode: 'today', 'weekly', 'monthly', atau 'custom' (pakai customFrom & customTo, format YYYY-MM-DD)
  verifyToken_(token);
  const rows = getAllIzinRows_().filter(r => r.status === 'Approved');
  const now = new Date();
  let cutoffStart = new Date();
  let cutoffEnd = now;

  if (periode === 'today') {
    cutoffStart.setHours(0, 0, 0, 0);
    cutoffEnd = new Date();
    cutoffEnd.setHours(23, 59, 59, 999);
  } else if (periode === 'weekly') {
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

/**
 * Admin hapus satu baris riwayat izin (Approved/Rejected) -- dipakai
 * kalau pengajuannya perlu direvisi, jadi mitra bisa isi ulang dari
 * awal. Pengajuan yang masih Pending sengaja TIDAK bisa dihapus lewat
 * sini (harus di-approve/reject dulu lewat tab Menunggu Approval).
 */
function hapusIzin(token, id) {
  verifyToken_(token);
  const sheet = getSheet_(SHEET_IZIN);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      if (values[i][9] === 'Pending') {
        return { success: false, message: 'Pengajuan ini masih Pending, proses dulu (approve/reject) sebelum dihapus.' };
      }
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Riwayat izin berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Data izin tidak ditemukan.' };
}

// ------------------- ADMIN: RINGKASAN (DASHBOARD) -------------------

function getRingkasan(token) {
  verifyToken_(token);

  const izinRows = getAllIzinRows_();
  const pendingCount = izinRows.filter(r => r.status === 'Pending').length;

  const now = new Date();
  const isThisMonth = (d) => {
    const dt = new Date(d);
    return dt.getFullYear() === now.getFullYear() && dt.getMonth() === now.getMonth();
  };
  const approvedBulanIni = izinRows.filter(r => r.status === 'Approved' && r.timestampDiproses && isThisMonth(r.timestampDiproses)).length;
  const rejectedBulanIni = izinRows.filter(r => r.status === 'Rejected' && r.timestampDiproses && isThisMonth(r.timestampDiproses)).length;

  const recentActivity = izinRows
    .filter(r => r.status !== 'Pending' && r.timestampDiproses)
    .sort((a, b) => new Date(b.timestampDiproses) - new Date(a.timestampDiproses))
    .slice(0, 6)
    .map(r => ({ nama: r.nama, kategori: r.kategori, status: r.status, timestampDiproses: r.timestampDiproses }));

  const mitraList = getDataMitraListRows_();
  const genderBreakdown = { 'Laki-laki': 0, 'Perempuan': 0 };
  // Selalu tampilkan semua kategori dari config (termasuk yang masih 0
  // mitra, misal "Massage" yang baru mau jalan) biar konsisten sama tab Data Mitra.
  const kategoriMitraBreakdown = {};
  Object.keys(KATEGORI_WILAYAH_MAP).forEach(k => { kategoriMitraBreakdown[k] = 0; });
  mitraList.forEach(m => {
    if (genderBreakdown[m.gender] !== undefined) genderBreakdown[m.gender]++;
    kategoriMitraBreakdown[m.kategori] = (kategoriMitraBreakdown[m.kategori] || 0) + 1;
  });

  return {
    success: true,
    totalMitra: mitraList.length,
    genderBreakdown,
    kategoriMitraBreakdown,
    pendingCount,
    approvedBulanIni,
    rejectedBulanIni,
    recentActivity
  };
}

// ------------------- ADMIN: DATA MITRA (master data mitra aktif) -------------------
//
// Terpisah dari tab "Mitra" (yang cuma Nama + PIN buat form izin).
// Ini data master mitra aktif: kontak, kategori layanan, jangkauan
// wilayah, hari libur, gender. Kalau mau nambah kategori layanan atau
// kota baru di masa depan, tinggal edit KATEGORI_WILAYAH_MAP &
// KOTA_KECAMATAN_MAP di bawah.

const SHEET_DATA_MITRA = 'Data Mitra';
const HARI_OPTIONS = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];

const JABODETABEK_KOTA = [
  'Jakarta Barat', 'Jakarta Timur', 'Jakarta Pusat', 'Jakarta Utara', 'Jakarta Selatan',
  'Bekasi', 'Depok', 'Tangerang', 'Tangerang Selatan', 'Bogor'
];

const KATEGORI_WILAYAH_MAP = {
  'Daily Cleaning': ['Surabaya', 'Bandung'].concat(JABODETABEK_KOTA),
  'Massage': JABODETABEK_KOTA.slice()
};

// Daftar kecamatan resmi per kota, biar admin tinggal pilih pas nambah
// mitra (gak perlu ketik manual & rawan beda-beda ejaan).
const KOTA_KECAMATAN_MAP = {
  'Jakarta Pusat': ['Gambir', 'Sawah Besar', 'Kemayoran', 'Senen', 'Cempaka Putih', 'Menteng', 'Tanah Abang', 'Johar Baru'],
  'Jakarta Utara': ['Penjaringan', 'Pademangan', 'Tanjung Priok', 'Koja', 'Kelapa Gading', 'Cilincing'],
  'Jakarta Barat': ['Cengkareng', 'Grogol Petamburan', 'Taman Sari', 'Tambora', 'Kebon Jeruk', 'Kalideres', 'Palmerah', 'Kembangan'],
  'Jakarta Selatan': ['Kebayoran Baru', 'Kebayoran Lama', 'Pesanggrahan', 'Cilandak', 'Pasar Minggu', 'Jagakarsa', 'Mampang Prapatan', 'Pancoran', 'Tebet', 'Setiabudi'],
  'Jakarta Timur': ['Matraman', 'Pulogadung', 'Jatinegara', 'Cakung', 'Duren Sawit', 'Kramat Jati', 'Makasar', 'Pasar Rebo', 'Ciracas', 'Cipayung'],
  'Bekasi': ['Bekasi Timur', 'Bekasi Barat', 'Bekasi Utara', 'Bekasi Selatan', 'Rawalumbu', 'Bantargebang', 'Pondokgede', 'Jatiasih', 'Jatisampurna', 'Mustikajaya', 'Medan Satria', 'Pondok Melati'],
  'Depok': ['Beji', 'Pancoran Mas', 'Cipayung', 'Sukmajaya', 'Cilodong', 'Cimanggis', 'Tapos', 'Sawangan', 'Bojongsari', 'Limo', 'Cinere'],
  'Tangerang': ['Tangerang', 'Jatiuwung', 'Batuceper', 'Benda', 'Cipondoh', 'Ciledug', 'Karawaci', 'Cibodas', 'Periuk', 'Neglasari', 'Karang Tengah', 'Larangan', 'Pinang'],
  'Tangerang Selatan': ['Ciputat', 'Ciputat Timur', 'Pondok Aren', 'Serpong', 'Serpong Utara', 'Setu', 'Pamulang'],
  'Bogor': ['Bogor Selatan', 'Bogor Timur', 'Bogor Utara', 'Bogor Tengah', 'Bogor Barat', 'Tanah Sareal'],
  'Surabaya': ['Genteng', 'Bubutan', 'Simokerto', 'Pabean Cantikan', 'Semampir', 'Krembangan', 'Kenjeran', 'Bulak', 'Tambaksari', 'Gubeng', 'Rungkut', 'Tenggilis Mejoyo', 'Gunung Anyar', 'Sukolilo', 'Mulyorejo', 'Sawahan', 'Wonokromo', 'Karangpilang', 'Dukuh Pakis', 'Wiyung', 'Wonocolo', 'Gayungan', 'Jambangan', 'Tegalsari', 'Sukomanunggal', 'Tandes', 'Sambikerep', 'Benowo', 'Pakal', 'Asemrowo', 'Lakarsantri'],
  'Bandung': ['Sukasari', 'Sukajadi', 'Cicendo', 'Andir', 'Bandung Kulon', 'Babakan Ciparay', 'Bojongloa Kaler', 'Bojongloa Kidul', 'Astanaanyar', 'Regol', 'Lengkong', 'Bandung Kidul', 'Batununggal', 'Kiaracondong', 'Cibeunying Kidul', 'Cibeunying Kaler', 'Coblong', 'Sumur Bandung', 'Cidadap', 'Antapani', 'Mandalajati', 'Arcamanik', 'Cinambo', 'Ujung Berung', 'Panyileukan', 'Cibiru', 'Gedebage', 'Rancasari', 'Buah Batu', 'Bandung Wetan']
};

const DATA_MITRA_HEADERS = [
  'ID', 'Nama Mitra', 'Nomor Telpon', 'Email', 'Password',
  'Kategori', 'Kota', 'Kecamatan', 'Hari Libur', 'Gender', 'Tanggal Ditambahkan'
];

function getKategoriWilayahMap() {
  return KATEGORI_WILAYAH_MAP;
}

function getKotaKecamatanMap() {
  return KOTA_KECAMATAN_MAP;
}

/**
 * Satu kali panggil buat narik kategori+kota+kecamatan sekaligus,
 * biar modal Tambah/Edit Mitra gak perlu 2-3 round-trip terpisah.
 */
function getWilayahConfig() {
  return {
    kategoriWilayah: KATEGORI_WILAYAH_MAP,
    kotaKecamatan: KOTA_KECAMATAN_MAP
  };
}

function getOrCreateDataMitraSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DATA_MITRA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATA_MITRA);
    sheet.appendRow(DATA_MITRA_HEADERS);
  }
  return sheet;
}

function generateMitraId_() {
  return Utilities.formatString('MT-%s', Utilities.getUuid().split('-')[0].toUpperCase());
}

function normalizeTelpon_(telpon) {
  return String(telpon || '').replace(/\s+/g, '');
}

/**
 * data = { nama, telpon, email, password, kategori, kota:[], kecamatan:[], hariLibur:[], gender }
 */
function validateDataMitra_(data) {
  if (!data) return 'Data mitra kosong.';
  if (!data.nama) return 'Nama Mitra wajib diisi.';
  if (!data.telpon) return 'Nomor Telpon wajib diisi.';
  if (!data.email) return 'Email wajib diisi.';
  if (!data.password) return 'Password wajib diisi.';
  if (!data.kategori || !KATEGORI_WILAYAH_MAP[data.kategori]) {
    return 'Kategori harus salah satu dari: ' + Object.keys(KATEGORI_WILAYAH_MAP).join(', ') + '.';
  }
  const allowedKota = KATEGORI_WILAYAH_MAP[data.kategori];
  const kotaList = Array.isArray(data.kota) ? data.kota.filter(k => allowedKota.indexOf(k) !== -1) : [];
  if (!kotaList.length) {
    return 'Pilih minimal 1 kota yang sesuai kategori "' + data.kategori + '" (' + allowedKota.join(', ') + ').';
  }
  if (!data.gender || (data.gender !== 'Laki-laki' && data.gender !== 'Perempuan')) {
    return 'Gender harus "Laki-laki" atau "Perempuan".';
  }
  return null;
}

function mapDataMitraRow_(r, rowIndex) {
  return {
    rowIndex: rowIndex,
    id: r[0],
    nama: r[1],
    telpon: r[2],
    email: r[3],
    password: r[4],
    kategori: r[5],
    kota: r[6] ? String(r[6]).split(',').map(s => s.trim()).filter(Boolean) : [],
    kecamatan: r[7] ? String(r[7]).split(',').map(s => s.trim()).filter(Boolean) : [],
    hariLibur: r[8] ? String(r[8]).split(',').map(s => s.trim()).filter(Boolean) : [],
    gender: r[9],
    tanggalDitambahkan: r[10]
  };
}

function getDataMitraListRows_() {
  const sheet = getOrCreateDataMitraSheet_();
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    list.push(mapDataMitraRow_(values[i], i + 1));
  }
  list.sort((a, b) => a.nama.localeCompare(b.nama, 'id', { sensitivity: 'base' }));
  return list;
}

function getDataMitraList(token) {
  verifyToken_(token);
  return getDataMitraListRows_();
}

function getDataMitraInitData(token) {
  verifyToken_(token);
  return {
    kategoriWilayah: KATEGORI_WILAYAH_MAP,
    kotaKecamatan: KOTA_KECAMATAN_MAP,
    mitraList: getDataMitraListRows_()
  };
}

function tambahDataMitra(token, data) {
  verifyToken_(token);
  const err = validateDataMitra_(data);
  if (err) return { success: false, message: err };

  const sheet = getOrCreateDataMitraSheet_();
  const values = sheet.getDataRange().getValues();
  const telponBaru = normalizeTelpon_(data.telpon);
  for (let i = 1; i < values.length; i++) {
    if (normalizeTelpon_(values[i][2]) === telponBaru) {
      return { success: false, message: 'Nomor Telpon "' + data.telpon + '" sudah terdaftar atas nama ' + values[i][1] + '.' };
    }
  }

  const id = generateMitraId_();
  sheet.appendRow([
    id, data.nama, data.telpon, data.email, data.password, data.kategori,
    data.kota.join(', '), (data.kecamatan || []).join(', '),
    (data.hariLibur || []).join(', '), data.gender, new Date()
  ]);
  return { success: true, message: 'Mitra "' + data.nama + '" berhasil ditambahkan.', id: id };
}

function updateDataMitra(token, id, data) {
  verifyToken_(token);
  const err = validateDataMitra_(data);
  if (err) return { success: false, message: err };

  const sheet = getOrCreateDataMitraSheet_();
  const values = sheet.getDataRange().getValues();
  const telponBaru = normalizeTelpon_(data.telpon);

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] !== id && normalizeTelpon_(values[i][2]) === telponBaru) {
      return { success: false, message: 'Nomor Telpon "' + data.telpon + '" sudah dipakai mitra lain (' + values[i][1] + ').' };
    }
  }

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, 2, 1, 9).setValues([[
        data.nama, data.telpon, data.email, data.password, data.kategori,
        data.kota.join(', '), (data.kecamatan || []).join(', '),
        (data.hariLibur || []).join(', '), data.gender
      ]]);
      return { success: true, message: 'Data mitra berhasil diperbarui.' };
    }
  }
  return { success: false, message: 'Mitra tidak ditemukan.' };
}

function hapusDataMitra(token, id) {
  verifyToken_(token);
  const sheet = getOrCreateDataMitraSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Mitra berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Mitra tidak ditemukan.' };
}

/**
 * Import massal dari CSV yang di-upload di panel Admin.
 * rows = array data mitra (shape sama kayak tambahDataMitra), sudah
 * di-parse jadi objek di sisi client. Baris yang gagal validasi atau
 * nomor telponnya dobel (baik sama data lama maupun sesama baris di
 * file yang sama) di-skip, TAPI baris lain yang valid tetap kesimpen.
 */
function importDataMitraBulk(token, rows) {
  verifyToken_(token);
  if (!Array.isArray(rows) || !rows.length) {
    return { success: false, message: 'Tidak ada data untuk diimport.' };
  }

  const sheet = getOrCreateDataMitraSheet_();
  const existing = sheet.getDataRange().getValues();
  const existingPhones = {};
  for (let i = 1; i < existing.length; i++) {
    existingPhones[normalizeTelpon_(existing[i][2])] = existing[i][1];
  }

  const toAppend = [];
  const errors = [];
  const seenInBatch = {};

  rows.forEach(function(data, idx) {
    const rowNum = idx + 1;
    const err = validateDataMitra_(data);
    if (err) {
      errors.push({ row: rowNum, nama: data.nama || '(kosong)', message: err });
      return;
    }
    const telponBaru = normalizeTelpon_(data.telpon);
    if (existingPhones[telponBaru]) {
      errors.push({ row: rowNum, nama: data.nama, message: 'Nomor telpon sudah terdaftar atas nama ' + existingPhones[telponBaru] + '.' });
      return;
    }
    if (seenInBatch[telponBaru]) {
      errors.push({ row: rowNum, nama: data.nama, message: 'Nomor telpon duplikat di dalam file yang sama.' });
      return;
    }
    seenInBatch[telponBaru] = true;
    toAppend.push([
      generateMitraId_(), data.nama, data.telpon, data.email, data.password, data.kategori,
      data.kota.join(', '), (data.kecamatan || []).join(', '),
      (data.hariLibur || []).join(', '), data.gender, new Date()
    ]);
  });

  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, DATA_MITRA_HEADERS.length).setValues(toAppend);
  }

  return {
    success: true,
    added: toAppend.length,
    failed: errors,
    message: toAppend.length + ' mitra berhasil ditambahkan' + (errors.length ? ', ' + errors.length + ' baris gagal (lihat detail di bawah).' : '.')
  };
}

// ------------------- ADMIN: PIN MITRA (kelola sheet "Mitra" -- Nama + PIN) -------------------
//
// Ini sheet auth sederhana (Nama + PIN 4 digit) yang dipakai mitra buat
// submit izin, lihat riwayat, dan cek gaji di index.html. TERPISAH dari
// "Data Mitra" (profil lengkap) -- sengaja gak digabung sesuai keputusan.

function validatePin_(pin) {
  return /^[0-9]{4}$/.test(String(pin || '').trim());
}

function getMitraPinList(token) {
  verifyToken_(token);
  const sheet = getSheet_(SHEET_MITRA);
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) list.push({ nama: values[i][0], pin: String(values[i][1] || '') });
  }
  list.sort((a, b) => a.nama.localeCompare(b.nama, 'id', { sensitivity: 'base' }));
  return list;
}

function tambahMitraPin(token, nama, pin) {
  verifyToken_(token);
  nama = String(nama || '').trim();
  if (!nama) return { success: false, message: 'Nama Mitra wajib diisi.' };
  if (!validatePin_(pin)) return { success: false, message: 'PIN harus 4 digit angka.' };

  const sheet = getSheet_(SHEET_MITRA);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === nama.toLowerCase()) {
      return { success: false, message: 'Nama "' + nama + '" sudah terdaftar. Pakai menu edit kalau mau ubah PIN-nya.' };
    }
  }
  sheet.appendRow([nama, String(pin).trim()]);
  return { success: true, message: 'PIN untuk "' + nama + '" berhasil dibuat.' };
}

function updateMitraPin(token, namaLama, namaBaru, pinBaru) {
  verifyToken_(token);
  namaBaru = String(namaBaru || '').trim();
  if (!namaBaru) return { success: false, message: 'Nama Mitra wajib diisi.' };
  if (!validatePin_(pinBaru)) return { success: false, message: 'PIN harus 4 digit angka.' };

  const sheet = getSheet_(SHEET_MITRA);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (i > 0 && String(values[i][0]).trim().toLowerCase() === namaBaru.toLowerCase() && String(values[i][0]).trim() !== String(namaLama).trim()) {
      return { success: false, message: 'Nama "' + namaBaru + '" sudah dipakai mitra lain.' };
    }
  }

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(namaLama).trim()) {
      sheet.getRange(i + 1, 1, 1, 2).setValues([[namaBaru, String(pinBaru).trim()]]);
      return { success: true, message: 'Data PIN "' + namaBaru + '" berhasil diperbarui.' };
    }
  }
  return { success: false, message: 'Mitra tidak ditemukan.' };
}

function hapusMitraPin(token, nama) {
  verifyToken_(token);
  const sheet = getSheet_(SHEET_MITRA);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(nama).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'PIN "' + nama + '" berhasil dihapus. Mitra ini gak akan bisa submit izin/cek gaji sampai PIN-nya dibuat ulang.' };
    }
  }
  return { success: false, message: 'Mitra tidak ditemukan.' };
}

// ------------------- GAJI MITRA -------------------
//
// Sheet "Gaji" -- satu baris per mitra per periode (bulan). Admin input
// manual dari panel Admin; mitra cek lewat form (nama + PIN) di index.html.

const SHEET_GAJI = 'Gaji';
const BULAN_ID_ = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

function periodeLabel_(periode) {
  // periode format: "YYYY-MM"
  const parts = String(periode || '').split('-');
  const y = parts[0], m = parseInt(parts[1], 10);
  if (!y || !m || m < 1 || m > 12) return periode;
  return BULAN_ID_[m - 1] + ' ' + y;
}

/**
 * Google Sheets suka auto-convert string kayak "2026-09" jadi objek Date
 * beneran (karena mirip format tanggal) walau kita niatnya nyimpen teks
 * biasa. Ini normalize balik ke string "YYYY-MM" apapun bentuk aslinya,
 * biar aman dipakai .localeCompare()/perbandingan string di tempat lain.
 */
function normalizePeriode_(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    return y + '-' + m;
  }
  return String(v || '');
}

function getOrCreateGajiSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_GAJI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_GAJI);
    sheet.appendRow(['ID', 'Nama Mitra', 'Periode', 'Durasi Kerja (Jam)', 'Gaji Pokok', 'Pinalti', 'Kompensasi', 'Bonus', 'Gaji Total', 'Diinput Oleh', 'Tanggal Diinput']);
  }
  // Paksa kolom Periode (C) tetap format Plain Text biar "2026-09" gak
  // ke-auto-convert jadi Date lagi di masa depan.
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) sheet.getRange(2, 3, maxRows - 1, 1).setNumberFormat('@');
  return sheet;
}

function generateGajiId_() {
  return Utilities.formatString('GJ-%s', Utilities.getUuid().split('-')[0].toUpperCase());
}

function hitungGajiTotal_(gajiPokok, pinalti, kompensasi, bonus) {
  return Number(gajiPokok || 0) - Number(pinalti || 0) + Number(kompensasi || 0) + Number(bonus || 0);
}

function validateGajiData_(data) {
  if (!data) return 'Data gaji kosong.';
  if (!String(data.nama || '').trim()) return 'Nama Mitra wajib diisi.';
  if (!/^\d{4}-\d{2}$/.test(String(data.periode || ''))) return 'Periode wajib diisi (format bulan-tahun).';
  const numFields = ['durasiKerja', 'gajiPokok', 'pinalti', 'kompensasi', 'bonus'];
  for (const f of numFields) {
    if (data[f] !== undefined && data[f] !== '' && isNaN(Number(data[f]))) return 'Nilai "' + f + '" harus berupa angka.';
    if (Number(data[f]) < 0) return 'Nilai tidak boleh negatif.';
  }
  return null;
}

function mapGajiRow_(r) {
  const periode = normalizePeriode_(r[2]);
  return {
    id: r[0], nama: r[1], periode: periode, periodeLabel: periodeLabel_(periode),
    durasiKerja: r[3], gajiPokok: r[4], pinalti: r[5], kompensasi: r[6], bonus: r[7],
    gajiTotal: r[8], diinputOleh: r[9], tanggalDiinput: r[10]
  };
}

function getGajiListRows_() {
  const sheet = getOrCreateGajiSheet_();
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) list.push(mapGajiRow_(values[i]));
  }
  list.sort((a, b) => b.periode.localeCompare(a.periode) || a.nama.localeCompare(b.nama, 'id', { sensitivity: 'base' }));
  return list;
}

function getGajiList(token) {
  verifyToken_(token);
  return getGajiListRows_();
}

/**
 * Upsert: kalau mitra ini di periode ini udah pernah diinput sebelumnya,
 * data lama di-update (bukan duplikat baris baru) -- payroll biasanya
 * diproses sekali per bulan per mitra, dan admin mungkin perlu koreksi.
 */
function tambahGaji(token, data) {
  const username = verifyToken_(token);
  const err = validateGajiData_(data);
  if (err) return { success: false, message: err };

  const gajiTotal = hitungGajiTotal_(data.gajiPokok, data.pinalti, data.kompensasi, data.bonus);
  const sheet = getOrCreateGajiSheet_();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim().toLowerCase() === String(data.nama).trim().toLowerCase() && normalizePeriode_(values[i][2]) === data.periode) {
      sheet.getRange(i + 1, 4, 1, 7).setValues([[
        Number(data.durasiKerja) || 0, Number(data.gajiPokok) || 0, Number(data.pinalti) || 0,
        Number(data.kompensasi) || 0, Number(data.bonus) || 0, gajiTotal, username
      ]]);
      sheet.getRange(i + 1, 11).setValue(new Date());
      return { success: true, message: 'Gaji ' + data.nama + ' periode ' + periodeLabel_(data.periode) + ' berhasil diperbarui.', gajiTotal: gajiTotal };
    }
  }

  const id = generateGajiId_();
  sheet.appendRow([
    id, String(data.nama).trim(), data.periode,
    Number(data.durasiKerja) || 0, Number(data.gajiPokok) || 0, Number(data.pinalti) || 0,
    Number(data.kompensasi) || 0, Number(data.bonus) || 0, gajiTotal, username, new Date()
  ]);
  return { success: true, message: 'Gaji ' + data.nama + ' periode ' + periodeLabel_(data.periode) + ' berhasil disimpan.', gajiTotal: gajiTotal, id: id };
}

function hapusGaji(token, id) {
  verifyToken_(token);
  const sheet = getOrCreateGajiSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Data gaji berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Data gaji tidak ditemukan.' };
}

/**
 * Import massal dari CSV yang di-upload di panel Admin (menu Gaji Mitra).
 * rows = array data gaji (shape sama kayak tambahGaji), sudah di-parse
 * jadi objek di sisi client. Sama kayak tambahGaji, ini UPSERT per
 * (nama + periode): kalau kombinasi itu udah ada -- baik di sheet
 * (data lama) maupun di baris lain pada file CSV yang sama -- baris
 * terakhir yang menang (di-update), bukan jadi baris duplikat baru.
 * Baris yang gagal validasi di-skip, baris lain yang valid tetap kesimpen.
 */
function importGajiBulk(token, rows) {
  const username = verifyToken_(token);
  if (!Array.isArray(rows) || !rows.length) {
    return { success: false, message: 'Tidak ada data untuk diimport.' };
  }

  const sheet = getOrCreateGajiSheet_();
  const values = sheet.getDataRange().getValues();

  // key "nama_lower|periode" -> nomor baris sheet (1-based, buat setValues langsung)
  const existingRowIndex = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][1]).trim().toLowerCase() + '|' + normalizePeriode_(values[i][2]);
    existingRowIndex[key] = i + 1;
  }

  const toAppend = [];
  const errors = [];
  let updated = 0;
  const batchKeyToAppendIdx = {}; // key -> index di toAppend, buat handle duplikat DALAM file yang sama

  rows.forEach(function(data, idx) {
    const rowNum = idx + 1;
    const err = validateGajiData_(data);
    if (err) {
      errors.push({ row: rowNum, nama: data.nama || '(kosong)', message: err });
      return;
    }

    const gajiTotal = hitungGajiTotal_(data.gajiPokok, data.pinalti, data.kompensasi, data.bonus);
    const key = String(data.nama).trim().toLowerCase() + '|' + data.periode;
    const rowValues = [
      Number(data.durasiKerja) || 0, Number(data.gajiPokok) || 0, Number(data.pinalti) || 0,
      Number(data.kompensasi) || 0, Number(data.bonus) || 0, gajiTotal, username
    ];

    if (existingRowIndex[key]) {
      // udah ada di sheet -> update langsung
      sheet.getRange(existingRowIndex[key], 4, 1, 7).setValues([rowValues]);
      sheet.getRange(existingRowIndex[key], 11).setValue(new Date());
      updated++;
    } else if (batchKeyToAppendIdx[key] !== undefined) {
      // duplikat di dalam file yang sama -> timpa baris sebelumnya di batch, jangan dobel
      toAppend[batchKeyToAppendIdx[key]] = [
        toAppend[batchKeyToAppendIdx[key]][0], String(data.nama).trim(), data.periode
      ].concat(rowValues);
    } else {
      batchKeyToAppendIdx[key] = toAppend.length;
      toAppend.push([generateGajiId_(), String(data.nama).trim(), data.periode].concat(rowValues, [new Date()]));
    }
  });

  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 11).setValues(toAppend);
  }

  const added = toAppend.length;
  return {
    success: true,
    added: added,
    updated: updated,
    failed: errors,
    message: added + ' baris baru ditambahkan, ' + updated + ' baris diupdate' +
      (errors.length ? ', ' + errors.length + ' baris gagal (lihat detail di bawah).' : '.')
  };
}

/**
 * Mitra cek gaji periode TERBARU yang udah diinput admin (nama + PIN,
 * sama kayak pola verifikasi di getRiwayatMitra).
 */
function getGajiTerbaru(nama, pin) {
  if (!nama || !pin) return { success: false, message: 'Nama dan PIN wajib diisi.' };
  if (!verifyMitraPin_(nama, pin)) {
    return { success: false, message: 'PIN salah. Coba lagi atau hubungi admin kalau lupa PIN.' };
  }

  const sheet = getOrCreateGajiSheet_();
  const values = sheet.getDataRange().getValues();
  let latest = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim().toLowerCase() !== String(nama).trim().toLowerCase()) continue;
    const row = mapGajiRow_(values[i]);
    if (!latest || row.periode > latest.periode) latest = row;
  }

  if (!latest) {
    return { success: false, message: 'Belum ada data gaji untuk Anda. Coba lagi setelah admin memproses periode ini.' };
  }

  return {
    success: true,
    periodeLabel: latest.periodeLabel,
    durasiKerja: latest.durasiKerja,
    gajiPokok: latest.gajiPokok,
    pinalti: latest.pinalti,
    kompensasi: latest.kompensasi,
    bonus: latest.bonus,
    gajiTotal: latest.gajiTotal
  };
}
