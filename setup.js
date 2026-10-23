/**
 * ============================================================================
 * SISTEM INFORMASI PENGGALIAN POTENSI & PELAYANAN PAJAK DAERAH (SIPP-PAJAK)
 * FILE 3: Setup.gs (Database Initialization & Idempotent Schema Setup)
 * ============================================================================
 */

// 1. DATABASE, SHEET & HEADER CONFIG
const SETUP_DB_SCHEMA = {
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

/**
 * 2. Fungsi Utama Setup Database (Idempotent & Aman Dijalankan Berulang Kali)
 * Jalankan fungsi setupDatabase() ini dari Editor Google Apps Script.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Pastikan script ini terhubung dengan Google Spreadsheet (Container-bound Script).');
  }

  createMissingSheets(ss);
  createMissingHeaders(ss);
  createDefaultData(ss);
  createDocumentationFolderSafe();

  SpreadsheetApp.flush();
  Logger.log('✅ Setup Database SIPP-PAJAK Selesai Tanpa Error.');
  return {
    success: true,
    message: 'Database berhasil diinisialisasi secara lengkap.'
  };
}

/**
 * 3. Membuat Sheet yang belum tersedia tanpa menghapus Sheet existing
 */
function createMissingSheets(ss) {
  const sheetKeys = Object.keys(SETUP_DB_SCHEMA.SHEETS);
  for (let i = 0; i < sheetKeys.length; i++) {
    const sheetName = SETUP_DB_SCHEMA.SHEETS[sheetKeys[i]];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log('Membuat Sheet baru: ' + sheetName);
    }
  }
}

/**
 * 4. Membuat dan memformat baris Header jika Sheet masih kosong
 */
function createMissingHeaders(ss) {
  const sheetKeys = Object.keys(SETUP_DB_SCHEMA.SHEETS);
  for (let i = 0; i < sheetKeys.length; i++) {
    const key = sheetKeys[i];
    const sheetName = SETUP_DB_SCHEMA.SHEETS[key];
    const expectedHeaders = SETUP_DB_SCHEMA.HEADERS[key];
    const sheet = ss.getSheetByName(sheetName);

    if (sheet.getLastRow() === 0) {
      const headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
      headerRange.setValues([expectedHeaders]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#1e1b4b');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    } else {
      const currentFirstRow = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
      let needsUpdate = false;
      for (let c = 0; c < expectedHeaders.length; c++) {
        if (String(currentFirstRow[c]).trim() === '') {
          currentFirstRow[c] = expectedHeaders[c];
          needsUpdate = true;
        }
      }
      if (needsUpdate) {
        sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([currentFirstRow]);
      }
    }
  }
}

/**
 * 5. Mengisi Data Awal (Akun Default Admin/Petugas & Wilayah Kecamatan/Desa)
 * Hanya mengisi apabila tabel masih kosong (Idempotent).
 */
function createDefaultData(ss) {
  const nowIso = new Date().toISOString();

  // A. Default Users (admin / admin123 dan petugas / petugas123)
  const userSheet = ss.getSheetByName(SETUP_DB_SCHEMA.SHEETS.USERS);
  if (userSheet && userSheet.getLastRow() <= 1) {
    const adminHash = computeSetupPasswordHash('admin123');
    const petugasHash = computeSetupPasswordHash('petugas123');

    userSheet.appendRow(['admin', adminHash, 'Administrator Utama', 'ADMIN', 'ACTIVE', nowIso]);
    userSheet.appendRow(['petugas', petugasHash, 'Petugas Pendataan Lapangan', 'PETUGAS', 'ACTIVE', nowIso]);
    Logger.log('Akun default (admin & petugas) berhasil dibuat.');
  }

  // B. Default Master Wilayah (Kecamatan & Desa)
  const wilayahSheet = ss.getSheetByName(SETUP_DB_SCHEMA.SHEETS.WILAYAH);
  if (wilayahSheet && wilayahSheet.getLastRow() <= 1) {
    const defaultWilayah = [
      ['WIL-001', 'Kecamatan Tengah', 'Desa Sukamaju'],
      ['WIL-002', 'Kecamatan Tengah', 'Desa Mekarsari'],
      ['WIL-003', 'Kecamatan Tengah', 'Kelurahan Kota Baru'],
      ['WIL-004', 'Kecamatan Utara', 'Desa Harapan Jaya'],
      ['WIL-005', 'Kecamatan Utara', 'Desa Karya Bakti'],
      ['WIL-006', 'Kecamatan Utara', 'Desa Sumber Makmur'],
      ['WIL-007', 'Kecamatan Selatan', 'Desa Pasir Putih'],
      ['WIL-008', 'Kecamatan Selatan', 'Desa Cempaka Indah'],
      ['WIL-009', 'Kecamatan Timur', 'Desa Karang Anyar'],
      ['WIL-010', 'Kecamatan Barat', 'Desa Sindang Sari']
    ];
    wilayahSheet.getRange(2, 1, defaultWilayah.length, 3).setValues(defaultWilayah);
    Logger.log('Data awal Master Wilayah Kecamatan & Desa berhasil dibuat.');
  }
}

/**
 * 6. Membuat Folder Dokumentasi di Google Drive jika belum ada
 */
function createDocumentationFolderSafe() {
  try {
    const folderName = 'SIPP_PAJAK_DOKUMENTASI';
    const folders = DriveApp.getFoldersByName(folderName);
    if (!folders.hasNext()) {
      DriveApp.createFolder(folderName);
      Logger.log('Folder Google Drive "' + folderName + '" berhasil dibuat.');
    }
  } catch (e) {
    Logger.log('Info Folder Drive: ' + e.message);
  }
}

/**
 * Helper Hash SHA-256 yang identik dengan Code.gs
 */
function computeSetupPasswordHash(plainText) {
  const salt = 'SIPP_PAJAK_SECURE_SALT_2026';
  const raw = salt + ':' + String(plainText);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(function(byte) {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}