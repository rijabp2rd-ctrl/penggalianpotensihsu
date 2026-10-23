/**
 * ============================================================================
 * SISTEM INFORMASI PENGGALIAN POTENSI & PELAYANAN PAJAK DAERAH (SIPP-PAJAK)
 * FILE 1: Code.gs (Backend Server-Side Logic)
 * ============================================================================
 */

// 1. GLOBAL CONFIG
const APP_CONFIG = {
  APP_NAME: 'Aplikasi Penggalian Potensi Pajak & IKM',
  VERSION: '1.0.0',
  SESSION_DURATION_HOURS: 24,
  DRIVE_FOLDER_NAME: 'SIPP_PAJAK_DOKUMENTASI',
  PASSWORD_SALT: 'SIPP_PAJAK_SECURE_SALT_2026',
  SHEETS: {
    USERS: 'USERS',
    SESSIONS: 'SESSIONS',
    PAJAK_DAERAH: 'PAJAK_DAERAH',
    PBB: 'PBB',
    IKM: 'IKM',
    WILAYAH: 'WILAYAH'
  },
  HEADERS: {
    USERS: ['USERNAME', 'PASSWORD_HASH', 'FULL_NAME', 'ROLE', 'STATUS', 'CREATED_AT'],
    SESSIONS: ['SESSION_ID', 'USERNAME', 'ROLE', 'FULL_NAME', 'EXPIRES_AT', 'CREATED_AT'],
    PAJAK_DAERAH: [
      'ID', 'NAMA', 'KECAMATAN', 'DESA', 'ALAMAT',
      'NILAI_OMZET', 'PERSENTASE', 'ESTIMASI_PAJAK', 'TAHUN',
      'DOKUMENTASI_URL', 'TIMESTAMP', 'INPUT_BY'
    ],
    PBB: [
      'ID', 'NAMA', 'NIK', 'STATUS_NOP', 'NOMOR_NOP',
      'KECAMATAN', 'DESA', 'ALAMAT', 'TAHUN',
      'DOKUMENTASI_URL', 'TIMESTAMP', 'INPUT_BY'
    ],
    IKM: [
      'ID', 'NAMA', 'ALAMAT', 'KEGIATAN',
      'P1_NILAI', 'P2_NILAI', 'P3_NILAI', 'P4_NILAI', 'P5_NILAI',
      'TOTAL_NILAI', 'KATEGORI_MUTU', 'SARAN', 'TIMESTAMP'
    ],
    WILAYAH: ['ID', 'KECAMATAN', 'DESA']
  }
};

// 2. doGet()
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// 3. DATABASE HELPERS
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetByNameSafe(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (typeof setupDatabase === 'function') {
      setupDatabase();
      sheet = ss.getSheetByName(sheetName);
    }
  }
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" belum tersedia. Silakan jalankan setupDatabase() di Setup.gs terlebih dahulu.');
  }
  return sheet;
}

function getSheetRowsAsObjects(sheetName) {
  const sheet = getSheetByNameSafe(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) {
    return [];
  }
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function(h) { return String(h).trim(); });
  const results = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const obj = { _rowIndex: r + 1 };
    for (let c = 0; c < headers.length; c++) {
      const val = row[c];
      if (val instanceof Date) {
        obj[headers[c]] = val.toISOString();
      } else {
        obj[headers[c]] = val;
      }
    }
    results.push(obj);
  }
  return results;
}

// 4. RESPONSE HELPERS
function createResponse(success, message, data) {
  return {
    success: Boolean(success),
    message: String(message || ''),
    data: data !== undefined ? data : null
  };
}

// 5. VALIDATION HELPERS
function sanitizeString(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

function parseNumberSafe(val) {
  if (typeof val === 'number' && !isNaN(val)) return val;
  if (!val) return 0;
  const cleaned = String(val).replace(/[^0-9.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

function hashPassword(plainText) {
  const raw = APP_CONFIG.PASSWORD_SALT + ':' + String(plainText);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(function(byte) {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

// 6. AUTHENTICATION & 7. SESSION
function loginUser(username, password) {
  try {
    const cleanUser = sanitizeString(username).toLowerCase();
    const cleanPass = String(password || '');
    if (!cleanUser || !cleanPass) {
      return createResponse(false, 'Username dan password wajib diisi.', null);
    }

    const users = getSheetRowsAsObjects(APP_CONFIG.SHEETS.USERS);
    const passwordHash = hashPassword(cleanPass);

    let matchedUser = null;
    for (let i = 0; i < users.length; i++) {
      const u = users[i];
      if (String(u.USERNAME).toLowerCase() === cleanUser && String(u.PASSWORD_HASH) === passwordHash) {
        matchedUser = u;
        break;
      }
    }

    if (!matchedUser) {
      return createResponse(false, 'Username atau password tidak valid.', null);
    }

    if (String(matchedUser.STATUS).toUpperCase() !== 'ACTIVE') {
      return createResponse(false, 'Akun Anda sedang dinonaktifkan. Hubungi Administrator.', null);
    }

    const sessionId = Utilities.getUuid();
    const now = new Date();
    const expiresAt = new Date(now.getTime() + (APP_CONFIG.SESSION_DURATION_HOURS * 60 * 60 * 1000));

    const sessionSheet = getSheetByNameSafe(APP_CONFIG.SHEETS.SESSIONS);
    sessionSheet.appendRow([
      sessionId,
      matchedUser.USERNAME,
      matchedUser.ROLE,
      matchedUser.FULL_NAME,
      expiresAt.toISOString(),
      now.toISOString()
    ]);

    return createResponse(true, 'Login berhasil. Selamat datang, ' + matchedUser.FULL_NAME + '!', {
      sessionId: sessionId,
      username: matchedUser.USERNAME,
      fullName: matchedUser.FULL_NAME,
      role: matchedUser.ROLE,
      expiresAt: expiresAt.toISOString()
    });
  } catch (err) {
    return createResponse(false, 'Gagal memproses login: ' + err.message, null);
  }
}

function verifySession(sessionId) {
  try {
    const cleanToken = sanitizeString(sessionId);
    if (!cleanToken) {
      return createResponse(false, 'Sesi tidak ditemukan.', null);
    }

    const sessions = getSheetRowsAsObjects(APP_CONFIG.SHEETS.SESSIONS);
    const now = new Date().getTime();

    for (let i = sessions.length - 1; i >= 0; i--) {
      const s = sessions[i];
      if (String(s.SESSION_ID) === cleanToken) {
        const expTime = new Date(s.EXPIRES_AT).getTime();
        if (isNaN(expTime) || now > expTime) {
          return createResponse(false, 'Sesi telah berakhir. Silakan login kembali.', null);
        }
        return createResponse(true, 'Sesi valid.', {
          sessionId: s.SESSION_ID,
          username: s.USERNAME,
          fullName: s.FULL_NAME,
          role: s.ROLE,
          expiresAt: s.EXPIRES_AT
        });
      }
    }

    return createResponse(false, 'Sesi tidak valid atau sudah logout.', null);
  } catch (err) {
    return createResponse(false, 'Gagal memverifikasi sesi: ' + err.message, null);
  }
}

function logoutUser(sessionId) {
  try {
    const cleanToken = sanitizeString(sessionId);
    if (!cleanToken) {
      return createResponse(true, 'Berhasil logout.', null);
    }

    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.SESSIONS);
    const sessions = getSheetRowsAsObjects(APP_CONFIG.SHEETS.SESSIONS);

    for (let i = sessions.length - 1; i >= 0; i--) {
      if (String(sessions[i].SESSION_ID) === cleanToken) {
        sheet.deleteRow(sessions[i]._rowIndex);
      }
    }
    return createResponse(true, 'Anda telah berhasil keluar dari sistem.', null);
  } catch (err) {
    return createResponse(false, 'Gagal melakukan logout: ' + err.message, null);
  }
}

function requireValidSession(sessionId) {
  const check = verifySession(sessionId);
  if (!check.success) {
    throw new Error(check.message || 'Akses ditolak. Silakan login terlebih dahulu.');
  }
  return check.data;
}

// 8. USER MANAGEMENT
function changeUserPassword(sessionId, oldPassword, newPassword) {
  try {
    const sessionUser = requireValidSession(sessionId);
    const cleanOld = String(oldPassword || '');
    const cleanNew = String(newPassword || '');

    if (!cleanOld || !cleanNew) {
      return createResponse(false, 'Password lama dan password baru wajib diisi.', null);
    }
    if (cleanNew.length < 6) {
      return createResponse(false, 'Password baru minimal harus terdiri dari 6 karakter.', null);
    }

    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.USERS);
    const users = getSheetRowsAsObjects(APP_CONFIG.SHEETS.USERS);
    const oldHash = hashPassword(cleanOld);
    const newHash = hashPassword(cleanNew);

    for (let i = 0; i < users.length; i++) {
      if (String(users[i].USERNAME).toLowerCase() === String(sessionUser.username).toLowerCase()) {
        if (String(users[i].PASSWORD_HASH) !== oldHash) {
          return createResponse(false, 'Password lama yang Anda masukkan salah.', null);
        }
        sheet.getRange(users[i]._rowIndex, 2).setValue(newHash);
        return createResponse(true, 'Password berhasil diperbarui.', null);
      }
    }

    return createResponse(false, 'Data pengguna tidak ditemukan.', null);
  } catch (err) {
    return createResponse(false, err.message, null);
  }
}

function addNewUser(sessionId, username, password, fullName, role) {
  try {
    const sessionUser = requireValidSession(sessionId);
    if (String(sessionUser.role).toUpperCase() !== 'ADMIN') {
      return createResponse(false, 'Hanya Administrator yang dapat menambahkan pengguna baru.', null);
    }

    const cleanUsername = sanitizeString(username).toLowerCase().replace(/\s+/g, '');
    const cleanPass = String(password || '');
    const cleanName = sanitizeString(fullName);
    const cleanRole = sanitizeString(role).toUpperCase() === 'ADMIN' ? 'ADMIN' : 'PETUGAS';

    if (!cleanUsername || !cleanPass || !cleanName) {
      return createResponse(false, 'Username, password, dan nama lengkap wajib diisi.', null);
    }
    if (cleanPass.length < 6) {
      return createResponse(false, 'Password minimal 6 karakter.', null);
    }

    const users = getSheetRowsAsObjects(APP_CONFIG.SHEETS.USERS);
    for (let i = 0; i < users.length; i++) {
      if (String(users[i].USERNAME).toLowerCase() === cleanUsername) {
        return createResponse(false, 'Username "' + cleanUsername + '" sudah digunakan.', null);
      }
    }

    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.USERS);
    const nowIso = new Date().toISOString();
    sheet.appendRow([
      cleanUsername,
      hashPassword(cleanPass),
      cleanName,
      cleanRole,
      'ACTIVE',
      nowIso
    ]);

    return createResponse(true, 'Pengguna "' + cleanName + '" berhasil ditambahkan.', {
      username: cleanUsername,
      fullName: cleanName,
      role: cleanRole,
      status: 'ACTIVE',
      createdAt: nowIso
    });
  } catch (err) {
    return createResponse(false, err.message, null);
  }
}

// 9. BUSINESS LOGIC & FILE UPLOAD HELPER
function saveDocumentationFile(base64Data, fileName, mimeType, prefix) {
  if (!base64Data || String(base64Data).trim() === '') {
    return 'Tidak Ada Dokumentasi';
  }
  try {
    let rawBase64 = String(base64Data);
    let detectedMime = mimeType || 'image/jpeg';

    if (rawBase64.indexOf('base64,') !== -1) {
      const parts = rawBase64.split('base64,');
      const header = parts[0];
      rawBase64 = parts[1];
      const mimeMatch = header.match(/data:([^;]+);/);
      if (mimeMatch && mimeMatch[1]) {
        detectedMime = mimeMatch[1];
      }
    }

    const decodedBytes = Utilities.base64Decode(rawBase64);
    const safeName = (prefix || 'DOK') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '_' + (fileName || 'foto.jpg').replace(/[^a-zA-Z0-9._-]/g, '');
    const blob = Utilities.newBlob(decodedBytes, detectedMime, safeName);

    const folders = DriveApp.getFoldersByName(APP_CONFIG.DRIVE_FOLDER_NAME);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(APP_CONFIG.DRIVE_FOLDER_NAME);

    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // Tetap lanjut jika kebijakan domain membatasi sharing publik
    }
    return file.getUrl();
  } catch (driveErr) {
    // Jika izin Drive belum diberikan atau gagal, simpan penanda terformat agar transaksi data tidak gagal
    return 'Terlampir (' + (fileName || 'Dokumentasi Digital') + ')';
  }
}

function getIkmGrade(totalScore) {
  const score = parseNumberSafe(totalScore);
  if (score >= 85) return 'A (Sangat Baik)';
  if (score >= 70) return 'B (Baik)';
  if (score >= 55) return 'C (Cukup)';
  return 'D (Kurang Baik)';
}

// 10. CRUD & PUBLIC FORM SUBMISSION FUNCTIONS
function getInitialPublicData() {
  try {
    const wilayahRows = getSheetRowsAsObjects(APP_CONFIG.SHEETS.WILAYAH);
    const pajakRows = getSheetRowsAsObjects(APP_CONFIG.SHEETS.PAJAK_DAERAH);
    const pbbRows = getSheetRowsAsObjects(APP_CONFIG.SHEETS.PBB);
    const ikmRows = getSheetRowsAsObjects(APP_CONFIG.SHEETS.IKM);

    let totalEstimasiPajak = 0;
    for (let i = 0; i < pajakRows.length; i++) {
      totalEstimasiPajak += parseNumberSafe(pajakRows[i].ESTIMASI_PAJAK);
    }

    let sumIkm = 0;
    for (let j = 0; j < ikmRows.length; j++) {
      sumIkm += parseNumberSafe(ikmRows[j].TOTAL_NILAI);
    }
    const avgIkm = ikmRows.length > 0 ? Number((sumIkm / ikmRows.length).toFixed(1)) : 0;

    const wilayahList = wilayahRows.map(function(w) {
      return {
        id: w.ID,
        kecamatan: w.KECAMATAN,
        desa: w.DESA
      };
    });

    return createResponse(true, 'Data publik berhasil dimuat.', {
      wilayah: wilayahList,
      summaryStats: {
        totalPajakDaerah: pajakRows.length,
        totalEstimasiPajak: totalEstimasiPajak,
        totalPbb: pbbRows.length,
        totalIkm: ikmRows.length,
        avgIkm: avgIkm
      }
    });
  } catch (err) {
    return createResponse(false, 'Gagal memuat data awal: ' + err.message, null);
  }
}

function submitPajakDaerah(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!payload || typeof payload !== 'object') {
      return createResponse(false, 'Data input tidak valid.', null);
    }

    const nama = sanitizeString(payload.nama);
    const kecamatan = sanitizeString(payload.kecamatan);
    const desa = sanitizeString(payload.desa);
    const alamat = sanitizeString(payload.alamat);
    const nilaiOmzet = parseNumberSafe(payload.nilaiOmzet);
    const persentase = parseNumberSafe(payload.persentase);
    const tahun = sanitizeString(payload.tahun);

    if (!nama || !kecamatan || !desa || !alamat || !tahun) {
      return createResponse(false, 'Mohon lengkapi Nama, Kecamatan, Desa, Alamat, dan Tahun.', null);
    }
    if (nilaiOmzet <= 0) {
      return createResponse(false, 'Nilai Omzet harus lebih besar dari 0.', null);
    }
    if (persentase <= 0 || persentase > 100) {
      return createResponse(false, 'Persentase pajak harus berada di antara 1% hingga 100%.', null);
    }

    let inputBy = 'Publik / Mandiri';
    if (payload.sessionId) {
      const check = verifySession(payload.sessionId);
      if (check.success && check.data) {
        inputBy = check.data.username + ' (' + check.data.fullName + ')';
      }
    }

    const estimasiPajak = Math.round(nilaiOmzet * (persentase / 100));
    const dokumentasiUrl = saveDocumentationFile(
      payload.dokumentasiBase64,
      payload.dokumentasiName,
      payload.dokumentasiMime,
      'PAJAK'
    );

    const id = 'PJK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(100 + Math.random() * 900);
    const nowIso = new Date().toISOString();

    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.PAJAK_DAERAH);
    sheet.appendRow([
      id,
      nama,
      kecamatan,
      desa,
      alamat,
      nilaiOmzet,
      persentase,
      estimasiPajak,
      tahun,
      dokumentasiUrl,
      nowIso,
      inputBy
    ]);

    return createResponse(true, 'Data Potensi Pajak Daerah berhasil disimpan!', {
      id: id,
      nama: nama,
      nilaiOmzet: nilaiOmzet,
      persentase: persentase,
      estimasiPajak: estimasiPajak,
      dokumentasiUrl: dokumentasiUrl,
      timestamp: nowIso
    });
  } catch (err) {
    return createResponse(false, 'Gagal menyimpan data Potensi Pajak Daerah: ' + err.message, null);
  } finally {
    lock.releaseLock();
  }
}

function submitPbb(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!payload || typeof payload !== 'object') {
      return createResponse(false, 'Data input PBB tidak valid.', null);
    }

    const nama = sanitizeString(payload.nama);
    const nik = sanitizeString(payload.nik).replace(/[^0-9]/g, '');
    const statusNop = sanitizeString(payload.statusNop);
    const nomorNop = statusNop === 'Ada' ? sanitizeString(payload.nomorNop || '-') : '-';
    const kecamatan = sanitizeString(payload.kecamatan);
    const desa = sanitizeString(payload.desa);
    const alamat = sanitizeString(payload.alamat);
    const tahun = sanitizeString(payload.tahun);

    if (!nama || !nik || !statusNop || !kecamatan || !desa || !alamat || !tahun) {
      return createResponse(false, 'Semua kolom wajib pada Form PBB harus diisi lengkap.', null);
    }
    if (nik.length !== 16) {
      return createResponse(false, 'NIK harus terdiri dari tepat 16 digit angka.', null);
    }

    let inputBy = 'Publik / Mandiri';
    if (payload.sessionId) {
      const check = verifySession(payload.sessionId);
      if (check.success && check.data) {
        inputBy = check.data.username + ' (' + check.data.fullName + ')';
      }
    }

    const dokumentasiUrl = saveDocumentationFile(
      payload.dokumentasiBase64,
      payload.dokumentasiName,
      payload.dokumentasiMime,
      'PBB'
    );

    const id = 'PBB-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(100 + Math.random() * 900);
    const nowIso = new Date().toISOString();

    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.PBB);
    sheet.appendRow([
      id,
      nama,
      "'" + nik,
      statusNop,
      nomorNop,
      kecamatan,
      desa,
      alamat,
      tahun,
      dokumentasiUrl,
      nowIso,
      inputBy
    ]);

    return createResponse(true, 'Data Pendataan PBB berhasil disimpan!', {
      id: id,
      nama: nama,
      nik: nik,
      statusNop: statusNop,
      dokumentasiUrl: dokumentasiUrl,
      timestamp: nowIso
    });
  } catch (err) {
    return createResponse(false, 'Gagal menyimpan data PBB: ' + err.message, null);
  } finally {
    lock.releaseLock();
  }
}

function submitIkm(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!payload || typeof payload !== 'object') {
      return createResponse(false, 'Data survei IKM tidak valid.', null);
    }

    const nama = sanitizeString(payload.nama);
    const alamat = sanitizeString(payload.alamat);
    const kegiatan = sanitizeString(payload.kegiatan);
    const p1 = parseNumberSafe(payload.p1Nilai);
    const p2 = parseNumberSafe(payload.p2Nilai);
    const p3 = parseNumberSafe(payload.p3Nilai);
    const p4 = parseNumberSafe(payload.p4Nilai);
    const p5 = parseNumberSafe(payload.p5Nilai);
    const saran = sanitizeString(payload.saran || '-');

    if (!nama || !alamat || !kegiatan) {
      return createResponse(false, 'Nama, Alamat, dan Kegiatan wajib diisi.', null);
    }
    if (p1 <= 0 || p2 <= 0 || p3 <= 0 || p4 <= 0 || p5 <= 0) {
      return createResponse(false, 'Mohon berikan penilaian Emote untuk kelima pernyataan IKM.', null);
    }

    const totalNilai = p1 + p2 + p3 + p4 + p5;
    const kategoriMutu = getIkmGrade(totalNilai);
    const id = 'IKM-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(100 + Math.random() * 900);
    const nowIso = new Date().toISOString();

    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.IKM);
    sheet.appendRow([
      id,
      nama,
      alamat,
      kegiatan,
      p1,
      p2,
      p3,
      p4,
      p5,
      totalNilai,
      kategoriMutu,
      saran,
      nowIso
    ]);

    return createResponse(true, 'Terima kasih! Penilaian Indeks Kepuasan Masyarakat (IKM) Anda berhasil disimpan.', {
      id: id,
      totalNilai: totalNilai,
      kategoriMutu: kategoriMutu,
      timestamp: nowIso
    });
  } catch (err) {
    return createResponse(false, 'Gagal menyimpan survei IKM: ' + err.message, null);
  } finally {
    lock.releaseLock();
  }
}

// 11. REPORT & ADMIN DASHBOARD FUNCTIONS
function isWithinReportPeriod(isoDateStr, period) {
  if (!period || period === 'ALL') return true;
  if (!isoDateStr) return false;

  const itemDate = new Date(isoDateStr);
  if (isNaN(itemDate.getTime())) return true;

  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  if (period === 'DAILY') {
    return itemDate >= startOfToday;
  }
  if (period === 'WEEKLY') {
    const sevenDaysAgo = new Date(startOfToday.getTime() - (7 * 24 * 60 * 60 * 1000));
    return itemDate >= sevenDaysAgo;
  }
  if (period === 'MONTHLY') {
    return itemDate.getFullYear() === now.getFullYear() && itemDate.getMonth() === now.getMonth();
  }
  if (period === 'YEARLY') {
    return itemDate.getFullYear() === now.getFullYear();
  }
  return true;
}

function getAdminDashboardData(sessionId, reportPeriod) {
  try {
    requireValidSession(sessionId);
    const period = sanitizeString(reportPeriod || 'ALL').toUpperCase();

    const rawPajak = getSheetRowsAsObjects(APP_CONFIG.SHEETS.PAJAK_DAERAH);
    const rawPbb = getSheetRowsAsObjects(APP_CONFIG.SHEETS.PBB);
    const rawIkm = getSheetRowsAsObjects(APP_CONFIG.SHEETS.IKM);
    const rawWilayah = getSheetRowsAsObjects(APP_CONFIG.SHEETS.WILAYAH);
    const rawUsers = getSheetRowsAsObjects(APP_CONFIG.SHEETS.USERS);

    const pajakList = rawPajak.filter(function(r) {
      return isWithinReportPeriod(r.TIMESTAMP, period);
    }).reverse();

    const pbbList = rawPbb.filter(function(r) {
      return isWithinReportPeriod(r.TIMESTAMP, period);
    }).reverse();

    const ikmList = rawIkm.filter(function(r) {
      return isWithinReportPeriod(r.TIMESTAMP, period);
    }).reverse();

    let totalOmzet = 0;
    let totalEstimasiPajak = 0;
    for (let i = 0; i < pajakList.length; i++) {
      totalOmzet += parseNumberSafe(pajakList[i].NILAI_OMZET);
      totalEstimasiPajak += parseNumberSafe(pajakList[i].ESTIMASI_PAJAK);
    }

    let nopAdaCount = 0;
    let nopTidakCount = 0;
    for (let j = 0; j < pbbList.length; j++) {
      if (String(pbbList[j].STATUS_NOP).toLowerCase() === 'ada') {
        nopAdaCount++;
      } else {
        nopTidakCount++;
      }
    }

    let totalIkmScore = 0;
    for (let k = 0; k < ikmList.length; k++) {
      totalIkmScore += parseNumberSafe(ikmList[k].TOTAL_NILAI);
    }
    const avgIkmScore = ikmList.length > 0 ? Number((totalIkmScore / ikmList.length).toFixed(1)) : 0;

    const usersList = rawUsers.map(function(u) {
      return {
        username: u.USERNAME,
        fullName: u.FULL_NAME,
        role: u.ROLE,
        status: u.STATUS,
        createdAt: u.CREATED_AT
      };
    });

    const wilayahList = rawWilayah.map(function(w) {
      return {
        id: w.ID,
        kecamatan: w.KECAMATAN,
        desa: w.DESA
      };
    });

    return createResponse(true, 'Data dashboard berhasil dimuat.', {
      pajakList: pajakList,
      pbbList: pbbList,
      ikmList: ikmList,
      wilayahList: wilayahList,
      usersList: usersList,
      stats: {
        period: period,
        totalPajakCount: pajakList.length,
        totalOmzet: totalOmzet,
        totalEstimasiPajak: totalEstimasiPajak,
        totalPbbCount: pbbList.length,
        nopAdaCount: nopAdaCount,
        nopTidakCount: nopTidakCount,
        totalIkmCount: ikmList.length,
        avgIkmScore: avgIkmScore,
        avgIkmGrade: getIkmGrade(avgIkmScore)
      }
    });
  } catch (err) {
    return createResponse(false, err.message, null);
  }
}

// 12. UTILITY & MASTER DATA MANAGEMENT FUNCTIONS
function updateRecordData(sessionId, type, payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    requireValidSession(sessionId);

    if (!payload || typeof payload !== 'object' || !payload.id) {
      return createResponse(false, 'Data pembaruan atau ID tidak valid.', null);
    }

    const cleanType = sanitizeString(type).toUpperCase();
    const cleanId = sanitizeString(payload.id);

    let targetSheetName = '';
    if (cleanType === 'PAJAK_DAERAH') targetSheetName = APP_CONFIG.SHEETS.PAJAK_DAERAH;
    else if (cleanType === 'PBB') targetSheetName = APP_CONFIG.SHEETS.PBB;
    else if (cleanType === 'IKM') targetSheetName = APP_CONFIG.SHEETS.IKM;
    else return createResponse(false, 'Tipe tabel tidak dikenali.', null);

    const sheet = getSheetByNameSafe(targetSheetName);
    const rows = getSheetRowsAsObjects(targetSheetName);

    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i].ID) === cleanId) {
        const rowIndex = rows[i]._rowIndex;

        if (cleanType === 'PAJAK_DAERAH') {
          const nama = sanitizeString(payload.nama);
          const kecamatan = sanitizeString(payload.kecamatan);
          const desa = sanitizeString(payload.desa);
          const alamat = sanitizeString(payload.alamat);
          const nilaiOmzet = parseNumberSafe(payload.nilaiOmzet);
          const persentase = parseNumberSafe(payload.persentase);
          const tahun = sanitizeString(payload.tahun);

          if (!nama || !kecamatan || !desa || !alamat || !tahun) {
            return createResponse(false, 'Nama, Kecamatan, Desa, Alamat, dan Tahun wajib diisi.', null);
          }
          if (nilaiOmzet <= 0) {
            return createResponse(false, 'Nilai Omzet harus lebih besar dari 0.', null);
          }

          const estimasiPajak = Math.round(nilaiOmzet * (persentase / 100));
          let dokumentasiUrl = rows[i].DOKUMENTASI_URL || 'Tidak Ada Dokumentasi';
          if (payload.dokumentasiBase64) {
            dokumentasiUrl = saveDocumentationFile(
              payload.dokumentasiBase64,
              payload.dokumentasiName,
              payload.dokumentasiMime,
              'PAJAK_EDIT'
            );
          }

          sheet.getRange(rowIndex, 1, 1, 12).setValues([[
            cleanId,
            nama,
            kecamatan,
            desa,
            alamat,
            nilaiOmzet,
            persentase,
            estimasiPajak,
            tahun,
            dokumentasiUrl,
            rows[i].TIMESTAMP || new Date().toISOString(),
            rows[i].INPUT_BY || 'Petugas'
          ]]);
        } else if (cleanType === 'PBB') {
          const nama = sanitizeString(payload.nama);
          const nik = sanitizeString(payload.nik).replace(/[^0-9]/g, '');
          const statusNop = sanitizeString(payload.statusNop);
          const nomorNop = statusNop === 'Ada' ? sanitizeString(payload.nomorNop || '-') : '-';
          const kecamatan = sanitizeString(payload.kecamatan);
          const desa = sanitizeString(payload.desa);
          const alamat = sanitizeString(payload.alamat);
          const tahun = sanitizeString(payload.tahun);

          if (!nama || !nik || !statusNop || !kecamatan || !desa || !alamat || !tahun) {
            return createResponse(false, 'Semua kolom wajib pada data PBB harus diisi.', null);
          }
          if (nik.length !== 16) {
            return createResponse(false, 'NIK harus terdiri dari tepat 16 digit angka.', null);
          }

          let dokumentasiUrl = rows[i].DOKUMENTASI_URL || 'Tidak Ada Dokumentasi';
          if (payload.dokumentasiBase64) {
            dokumentasiUrl = saveDocumentationFile(
              payload.dokumentasiBase64,
              payload.dokumentasiName,
              payload.dokumentasiMime,
              'PBB_EDIT'
            );
          }

          sheet.getRange(rowIndex, 1, 1, 12).setValues([[
            cleanId,
            nama,
            "'" + nik,
            statusNop,
            nomorNop,
            kecamatan,
            desa,
            alamat,
            tahun,
            dokumentasiUrl,
            rows[i].TIMESTAMP || new Date().toISOString(),
            rows[i].INPUT_BY || 'Petugas'
          ]]);
        } else if (cleanType === 'IKM') {
          const nama = sanitizeString(payload.nama);
          const alamat = sanitizeString(payload.alamat);
          const kegiatan = sanitizeString(payload.kegiatan);
          const p1 = parseNumberSafe(payload.p1Nilai);
          const p2 = parseNumberSafe(payload.p2Nilai);
          const p3 = parseNumberSafe(payload.p3Nilai);
          const p4 = parseNumberSafe(payload.p4Nilai);
          const p5 = parseNumberSafe(payload.p5Nilai);
          const saran = sanitizeString(payload.saran || '-');

          if (!nama || !alamat || !kegiatan) {
            return createResponse(false, 'Nama, Alamat, dan Kegiatan responden wajib diisi.', null);
          }

          const totalNilai = p1 + p2 + p3 + p4 + p5;
          const kategoriMutu = getIkmGrade(totalNilai);

          sheet.getRange(rowIndex, 1, 1, 13).setValues([[
            cleanId,
            nama,
            alamat,
            kegiatan,
            p1,
            p2,
            p3,
            p4,
            p5,
            totalNilai,
            kategoriMutu,
            saran,
            rows[i].TIMESTAMP || new Date().toISOString()
          ]]);
        }

        return createResponse(true, 'Data dengan ID ' + cleanId + ' berhasil diperbarui.', { updatedId: cleanId });
      }
    }

    return createResponse(false, 'Data dengan ID tersebut tidak ditemukan.', null);
  } catch (err) {
    return createResponse(false, 'Gagal memperbarui data: ' + err.message, null);
  } finally {
    lock.releaseLock();
  }
}

function deleteRecordData(sessionId, type, recordId) {
  try {
    requireValidSession(sessionId);
    const cleanType = sanitizeString(type).toUpperCase();
    const cleanId = sanitizeString(recordId);

    let targetSheetName = '';
    if (cleanType === 'PAJAK_DAERAH') targetSheetName = APP_CONFIG.SHEETS.PAJAK_DAERAH;
    else if (cleanType === 'PBB') targetSheetName = APP_CONFIG.SHEETS.PBB;
    else if (cleanType === 'IKM') targetSheetName = APP_CONFIG.SHEETS.IKM;
    else return createResponse(false, 'Tipe tabel tidak dikenali.', null);

    const sheet = getSheetByNameSafe(targetSheetName);
    const rows = getSheetRowsAsObjects(targetSheetName);

    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i].ID) === cleanId) {
        sheet.deleteRow(rows[i]._rowIndex);
        return createResponse(true, 'Data dengan ID ' + cleanId + ' berhasil dihapus.', { deletedId: cleanId });
      }
    }

    return createResponse(false, 'Data dengan ID tersebut tidak ditemukan.', null);
  } catch (err) {
    return createResponse(false, 'Gagal menghapus data: ' + err.message, null);
  }
}

function addWilayahData(sessionId, kecamatan, desa) {
  try {
    requireValidSession(sessionId);
    const cleanKec = sanitizeString(kecamatan);
    const cleanDesa = sanitizeString(desa);

    if (!cleanKec || !cleanDesa) {
      return createResponse(false, 'Nama Kecamatan dan Desa wajib diisi.', null);
    }

    const existing = getSheetRowsAsObjects(APP_CONFIG.SHEETS.WILAYAH);
    for (let i = 0; i < existing.length; i++) {
      if (
        String(existing[i].KECAMATAN).toLowerCase() === cleanKec.toLowerCase() &&
        String(existing[i].DESA).toLowerCase() === cleanDesa.toLowerCase()
      ) {
        return createResponse(false, 'Kombinasi Kecamatan dan Desa tersebut sudah terdaftar.', null);
      }
    }

    const id = 'WIL-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.WILAYAH);
    sheet.appendRow([id, cleanKec, cleanDesa]);

    return createResponse(true, 'Wilayah ' + cleanKec + ' - ' + cleanDesa + ' berhasil ditambahkan.', {
      id: id,
      kecamatan: cleanKec,
      desa: cleanDesa
    });
  } catch (err) {
    return createResponse(false, 'Gagal menambah wilayah: ' + err.message, null);
  }
}

function deleteWilayahData(sessionId, wilayahId) {
  try {
    requireValidSession(sessionId);
    const cleanId = sanitizeString(wilayahId);
    const sheet = getSheetByNameSafe(APP_CONFIG.SHEETS.WILAYAH);
    const rows = getSheetRowsAsObjects(APP_CONFIG.SHEETS.WILAYAH);

    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i].ID) === cleanId) {
        sheet.deleteRow(rows[i]._rowIndex);
        return createResponse(true, 'Data wilayah berhasil dihapus.', { deletedId: cleanId });
      }
    }

    return createResponse(false, 'Data wilayah tidak ditemukan.', null);
  } catch (err) {
    return createResponse(false, 'Gagal menghapus wilayah: ' + err.message, null);
  }
}