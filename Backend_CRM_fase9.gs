/**
 * ============================================================================
 * UNGU LAUNDRY ERP - FASE 2: CRM & FINANCIAL CONTROL
 * Module: Database Architecture & Migration
 * Architect: Senior ERP Engineer
 * ============================================================================
 */

/**
 * TUGAS 1: Setup Database V2 (Idempotent)
 * Fungsi ini aman dijalankan berkali-kali. Dia akan mengecek struktur
 * dan hanya menambahkan apa yang kurang tanpa menghapus data lama.
 */
function setupDatabaseV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000); // Tunggu antrian lock max 30 detik
    
    // 1. Update Sheet Data_Pelanggan (Alter Table)
    _updateTableStructure(ss, CONF.PELANGGAN, HEADERS_V2.PELANGGAN_ADDONS);
    
    // 2. Buat Sheet Log_Deposit (Create Table if not exists)
    _ensureSheetExists(ss, 'Log_Deposit', HEADERS_V2.LOG_DEPOSIT);
    
    // 3. Buat Sheet Log_Poin (Create Table if not exists)
    _ensureSheetExists(ss, 'Log_Poin', HEADERS_V2.LOG_POIN);
    
    // 4. Buat Sheet Master_Membership & Isi Default Data
    const memberSheetCreated = _ensureSheetExists(ss, 'Master_Membership', HEADERS_V2.MEMBERSHIP);
    if (memberSheetCreated) {
      const shMember = ss.getSheetByName('Master_Membership');
      // Seed Data Default
      shMember.getRange(2, 1, 3, 4).setValues([
        ['Regular', 0, 0, 1],
        ['Silver', 1000000, 0.05, 1.2],
        ['Gold', 5000000, 0.10, 1.5]
      ]);
    }

    return { status: 'SUCCESS', msg: 'Database Fase 2 Berhasil Di-setup / Diperbarui.' };
    
  } catch (e) {
    Logger.log("Error Setup DB V2: " + e.message);
    return { status: 'ERROR', msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * TUGAS 2: Migrasi Data Hutang (One-Time Run)
 * Menghitung ulang hutang dari Data_Pesanan dan menyuntikkannya ke Data_Pelanggan.
 */
function migrateCustomerData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPesanan = ss.getSheetByName(CONF.PESANAN);
  const shPelanggan = ss.getSheetByName(CONF.PELANGGAN);
  
  if (!shPesanan || !shPelanggan) {
    return "Error: Sheet Pesanan atau Pelanggan tidak ditemukan.";
  }

  // 1. Ambil Data Pesanan (Invoice)
  const lastRowPes = shPesanan.getLastRow();
  const lastColPes = shPesanan.getLastColumn();
  if (lastRowPes < 2 || lastColPes < 1) {
    return "Error: Data pesanan kosong.";
  }

  const dataPesanan = shPesanan.getRange(1, 1, lastRowPes, lastColPes).getValues();
  const headersPesanan = dataPesanan[0];
  const idxMapPesanan = _mapHeaders(headersPesanan);
  const IDX_INV_CUST = safeIdx(idxMapPesanan, 'ID_Pelanggan', 'migrateCustomerData: Data_Pesanan');
  const IDX_INV_TOTAL = safeIdx(idxMapPesanan, 'Total', 'migrateCustomerData: Data_Pesanan');
  const IDX_INV_STATUS = safeIdx(idxMapPesanan, 'Status', 'migrateCustomerData: Data_Pesanan');

  // Map untuk menyimpan total hutang per pelanggan
  let hutangMap = {};
  let spendingMap = {}; // Sekalian hitung Total Spending untuk Membership

  // Skip Header (i=1)
  for (let i = 1; i < dataPesanan.length; i++) {
    const row = dataPesanan[i];
    const custId = String(row[IDX_INV_CUST]).trim();
    const total = Number(row[IDX_INV_TOTAL]) || 0;
    const status = String(row[IDX_INV_STATUS]).trim(); // 'Lunas' atau 'Belum Lunas'

    if (!custId) continue;

    // Init Map jika belum ada
    if (!hutangMap[custId]) hutangMap[custId] = 0;
    if (!spendingMap[custId]) spendingMap[custId] = 0;

    // Akumulasi Spending (Semua transaksi dihitung)
    spendingMap[custId] += total;

    // Akumulasi Hutang (Hanya yang Belum Lunas)
    if (status === 'Belum Lunas') {
      hutangMap[custId] += total;
    }
  }

  // 2. Update Data Pelanggan
  const lastRowPel = shPelanggan.getLastRow();
  const lastColPel = shPelanggan.getLastColumn();
  if (lastRowPel < 2 || lastColPel < 1) {
    return "Error: Data pelanggan kosong.";
  }

  const dataPelanggan = shPelanggan.getRange(1, 1, lastRowPel, lastColPel).getValues();
  const headers = dataPelanggan[0];
  
  // Cari Index Kolom Target (Dinamis, agar aman jika urutan berubah)
  const idxHutang = headers.indexOf('Hutang_Aktif');
  const idxSpending = headers.indexOf('Total_Spending');
  const idxMember = headers.indexOf('Status_Member');
  
  if (idxHutang === -1 || idxSpending === -1) {
    return "Error: Kolom baru belum dibuat. Jalankan setupDatabaseV2() terlebih dahulu.";
  }

  // Array untuk menyimpan update
  let updatesHutang = [];
  let updatesSpending = [];
  let updatesMember = [];

  for (let i = 1; i < dataPelanggan.length; i++) {
    const custId = String(dataPelanggan[i][0]).trim(); // Asumsi ID di kolom A
    
    // Ambil nilai kalkulasi
    const hutang = hutangMap[custId] || 0;
    const spending = spendingMap[custId] || 0;
    
    // Tentukan Level Member Sederhana (Bisa dipercanggih nanti)
    let level = 'Regular';
    if (spending >= 5000000) level = 'Gold';
    else if (spending >= 1000000) level = 'Silver';

    updatesHutang.push([hutang]);
    updatesSpending.push([spending]);
    updatesMember.push([level]);
  }

  // 3. Tulis ke Sheet (Batch Write untuk Performa)
  // Tulis Hutang
  shPelanggan.getRange(2, idxHutang + 1, updatesHutang.length, 1).setValues(updatesHutang);
  // Tulis Spending
  shPelanggan.getRange(2, idxSpending + 1, updatesSpending.length, 1).setValues(updatesSpending);
  // Tulis Level Member
  shPelanggan.getRange(2, idxMember + 1, updatesMember.length, 1).setValues(updatesMember);

  return `Sukses Migrasi: ${updatesHutang.length} data pelanggan diperbarui (Hutang & Spending).`;
}

// --- PRIVATE HELPER FUNCTIONS (INTERNAL ENGINE) ---

/**
 * Menambahkan kolom baru ke sheet jika belum ada (Safe Alter Table).
 */
function _updateTableStructure(ss, sheetName, newHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) return; // Skip jika sheet master belum ada (safety)

  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let lastCol = sheet.getLastColumn();

  newHeaders.forEach(header => {
    // Cek apakah header sudah ada (case-insensitive agar aman)
    if (!currentHeaders.includes(header)) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(header);
      // Opsional: Set format header agar seragam (Bold, Warna)
      sheet.getRange(1, lastCol).setFontWeight("bold").setBackground("#f3f4f6");
    }
  });
}

/**
 * Membuat sheet baru jika belum ada, dan set header.
 * Mengembalikan true jika sheet baru dibuat.
 */
function _ensureSheetExists(ss, sheetName, headersArray) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headersArray);
    // Formatting Header
    sheet.getRange(1, 1, 1, headersArray.length).setFontWeight("bold").setBackground("#4f46e5").setFontColor("white");
    sheet.setFrozenRows(1);
    return true;
  }
  return false;
}