/**
 * ============================================================================
 * UNGU LAUNDRY ERP - BACKEND INFRASTRUCTURE (PHASE 1: INTEGRITY & CONTROL)
 * ============================================================================
 * * Module: Database Setup, Security Helpers, & Core Calculation Engine
 * Architect: Senior ERP Engineer
 * Context: Foundation for Blind Count Shift & Cycle Counting Inventory
 */

// --- KONFIGURASI NAMA SHEET (Sesuai Global CONF) ---
// Kita mendefinisikan ini secara lokal untuk intellisense, 
// namun di production pastikan variabel global CONF sudah memuat key ini.
const CONF_PHASE1 = {
  LOG_SHIFT: 'Log_Shift',
  LOG_DROP: 'Log_Cash_Drop',
  LOG_OPNAME_HEAD: 'Log_Opname_Header',
  LOG_OPNAME_DET: 'Log_Opname_Detail',
  LOG_AUDIT: 'Log_Audit_Trail',
  // Sheet Eksisting yang dibutuhkan logic
  LOG_STOK: 'Log_Stok', 
  SALDO_AWAL: 'Saldo_Awal',
  JURNAL_UMUM: 'Jurnal_Transaksi_Umum',
  COA: 'COA_Master'
};

// ============================================================================
// TUGAS 1: DATABASE INITIALIZATION
// ============================================================================

/**
 * Menginisialisasi Database untuk Fase 1.
 * Mengecek ketersediaan sheet dan membuat header jika belum ada.
 * Idempotent: Aman dijalankan berkali-kali.
 */
function setupDatabaseFase1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Definisi Schema Database Baru
  const schemas = [
    {
      name: CONF_PHASE1.LOG_SHIFT,
      headers: [
        'ID_Shift', 'Cabang', 'User_ID', 'Waktu_Buka', 'Saldo_Awal_Sistem', 
        'Total_Penjualan_Tunai', 'Total_DP_Tunai', 'Total_Pengeluaran_Kas', 
        'Total_Cash_Drop', 'Waktu_Tutup', 'Saldo_Akhir_Fisik', 'Selisih', 'Status'
      ]
    },
    {
      name: CONF_PHASE1.LOG_DROP,
      headers: ['ID_Drop', 'ID_Shift', 'Waktu', 'Nominal', 'Penerima', 'Keterangan']
    },
    {
      name: CONF_PHASE1.LOG_OPNAME_HEAD,
      headers: [
        'ID_Opname', 'Tanggal', 'Cabang', 'User_Checker', 
        'Kategori_Filter', 'Status', 'Waktu_Mulai', 'Waktu_Selesai'
      ]
    },
    {
      name: CONF_PHASE1.LOG_OPNAME_DET,
      headers: [
        'ID_Opname', 'Kode_Barang', 'Nama_Barang', 
        'Stok_Sistem_Snapshot', 'Stok_Fisik_Input', 'Selisih', 'Nilai_Adjust_Rp'
      ]
    },
    {
      name: CONF_PHASE1.LOG_AUDIT,
      headers: ['Timestamp', 'User_ID', 'Action', 'Target_ID', 'Payload_Changes']
    }
  ];

  schemas.forEach(schema => {
    let sheet = ss.getSheetByName(schema.name);
    
    // Jika sheet belum ada, buat baru
    if (!sheet) {
      sheet = ss.insertSheet(schema.name);
      // Set Header
      sheet.getRange(1, 1, 1, schema.headers.length).setValues([schema.headers]);
      // Styling Header (Bold + Freeze)
      sheet.getRange(1, 1, 1, schema.headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log(`[INIT] Membuat sheet baru: ${schema.name}`);
    } else {
      // Validasi Header (Opsional: Cek apakah kolom sesuai, di fase ini kita skip untuk efisiensi)
      Logger.log(`[SKIP] Sheet sudah ada: ${schema.name}`);
    }
  });

  SpreadsheetApp.flush(); // Pastikan perubahan tersimpan
}

// ============================================================================
// TUGAS 2: SECURITY & UTILITY HELPERS
// ============================================================================

/**
 * Mendapatkan waktu server terkini dalam zona waktu WIB (Asia/Jakarta).
 * Mencegah manipulasi waktu dari sisi klien (Backdating).
 * @returns {string} Formatted Date String "YYYY-MM-DD HH:mm:ss"
 */
function getWIBTime() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
}

/**
 * Menghasilkan ID unik yang human-readable.
 * Format: PREFIX-YYMMDD-RANDOM
 * @param {string} prefix - Kode awalan (misal: SH, OP, TRX)
 * @returns {string} Unique ID
 */
function generateERPId(prefix) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, "Asia/Jakarta", "yyMMdd");
  
  // Random string 4 karakter (Base36: 0-9, a-z)
  const randomStr = Math.floor(Math.random() * 1679615).toString(36).toUpperCase().padStart(4, '0');
  
  return `${prefix}-${dateStr}-${randomStr}`;
}

/**
 * Memvalidasi token sesi pengguna dengan standar keamanan Enterprise.
 * - Menghapus Backdoor akses dev.
 * - Memvalidasi keberadaan sesi di Cache Server.
 * * @param {string} token - Token otentikasi (UUID) dari frontend.
 * @returns {Object} User Object {username, role, cabang} yang terverifikasi.
 * @throws {Error} Jika token tidak valid, kosong, atau kadaluarsa.
 */
function validateUserSession(token) {
  // 1. Validasi Input Awal (Sanitization)
  // Pastikan token ada dan bertipe string untuk mencegah injection object
  if (!token || typeof token !== 'string' || token.trim() === '') {
    throw new Error("Akses Ditolak: Token tidak valid.");
  }

  // 2. Cek Sesi di CacheService (Server Memory)
  // Menggunakan getScriptCache() agar sesi bisa diakses antar eksekusi script
  var cache = CacheService.getScriptCache();
  var sessionJson = cache.get(token);

  // 3. Validasi Keberadaan Sesi
  if (!sessionJson) {
    // Jika null, berarti token salah atau TTL (Time To Live) sudah habis
    throw new Error("Sesi Kadaluarsa atau Tidak Valid. Silakan Login Ulang.");
  }

  // 4. Parsing Data & Konstruksi Objek User
  try {
    var user = JSON.parse(sessionJson);
    
    // Pastikan properti penting ada untuk mencegah undefined error di hilir
    if (!user.username && !user.Username) throw new Error("Data sesi tidak lengkap.");

    return {
      username: user.Username || user.username, // Fallback untuk kompatibilitas
      role: user.Role || user.role || 'User',   // Default ke role terendah jika kosong
      cabang: user.Cabang || user.cabang || ""  // Cabang bisa kosong jika user pusat
    };
  } catch (e) {
    // Menangani potensi error JSON parse (Data corruption)
    Logger.log("[SECURITY ALERT] Corrupt Session Data for token: " + token);
    throw new Error("Kesalahan Sistem: Data sesi korup.");
  }
}

// ============================================================================
// TUGAS 3: SMART CALCULATION LOGIC (THE "BRAIN")
// ============================================================================

/**
 * Menghitung stok saat ini dengan metode Cycle Counting.
 * Logic: Saldo Akhir = Saldo Opname Terakhir + Mutasi (setelah opname).
 * * @param {string} kodeBarang - Kode barang yang akan dihitung (misal: BRG-001)
 * @returns {number} Sisa stok terkini
 */
/**
 * Menghitung Stok Terkini dengan dukungan Cycle Count Reset.
 * [UPDATED] Mendukung baseline 'OPNAME_RESET'.
 */
function calculateCurrentStock(kodeBarang) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStok = ss.getSheetByName(CONF_PHASE1.LOG_STOK);
  const dataStok = sheetStok.getDataRange().getValues();

  // Index Kolom (Sesuaikan dengan Log_Stok)
  const IDX_TIME = 0;
  const IDX_KODE = 4;
  const IDX_JENIS = 6;
  const IDX_IN = 7;
  const IDX_OUT = 8;
  
  let saldo = 0;
  let lastResetTime = 0; 

  // 1. Cari Titik Reset Terakhir (OPNAME_RESET)
  // Kita loop mundur dari bawah untuk mencari reset paling baru
  for (let i = dataStok.length - 1; i >= 1; i--) {
    const row = dataStok[i];
    if (String(row[IDX_KODE]) === String(kodeBarang)) {
      if (String(row[IDX_JENIS]) === 'OPNAME_RESET') {
        // [CRITICAL] Jadikan Qty Masuk (Fisik) sebagai saldo dasar
        saldo = Number(row[IDX_IN]) || 0;
        lastResetTime = new Date(row[IDX_TIME]).getTime();
        break; // Stop loop, kita sudah ketemu titik nol terbaru
      }
    }
  }

  // 2. Jika tidak ada reset opname, cari Saldo Awal Master (Fallback)
  if (lastResetTime === 0) {
    const sheetSaldo = ss.getSheetByName(CONF_PHASE1.SALDO_AWAL);
    if (sheetSaldo) {
      const dataSaldo = sheetSaldo.getDataRange().getValues();
      for (let j = 1; j < dataSaldo.length; j++) {
        if (String(dataSaldo[j][1]) === String(kodeBarang)) { 
           saldo = Number(dataSaldo[j][3]) || 0; 
           break;
        }
      }
    }
  }

  // 3. Hitung Mutasi (Masuk - Keluar) HANYA SETELAH Reset Terakhir
  for (let i = 1; i < dataStok.length; i++) {
    const row = dataStok[i];
    const rowTime = new Date(row[IDX_TIME]).getTime();
    const jenis = String(row[IDX_JENIS]);

    // Filter: Barang sama DAN Transaksi terjadi SETELAH reset terakhir
    // Jangan hitung baris OPNAME_RESET itu sendiri lagi (karena sudah diambil sebagai saldo awal)
    if (String(row[IDX_KODE]) === String(kodeBarang) && rowTime > lastResetTime && jenis !== 'OPNAME_RESET') {
      const qtyIn = Number(row[IDX_IN]) || 0;
      const qtyOut = Number(row[IDX_OUT]) || 0;
      saldo = saldo + qtyIn - qtyOut;
    }
  }
  
  return saldo;
}

/**
 * Menghitung posisi keuangan Shift untuk Blind Count.
 * * @param {string} idShift - ID Shift yang sedang berjalan
 * @param {string} userCabang - Filter cabang user
 * @returns {Object} {total_masuk, total_keluar, total_cash_drop, saldo_sistem_sekarang}
 */
function calculateShiftFinancials(idShift, userCabang) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Ambil Data Shift (Saldo Awal)
  const sheetShift = ss.getSheetByName(CONF_PHASE1.LOG_SHIFT);
  // ... (kode pencarian row shift sama seperti sebelumnya) ...
  // (Anggap code pencarian shiftRow sudah ada disini)
  // --- SHORTCUT CODE ---
  const dataShift = sheetShift.getDataRange().getValues();
  let shiftRow = null;
  for (let i = 1; i < dataShift.length; i++) {
    if (String(dataShift[i][0]) === String(idShift)) { shiftRow = dataShift[i]; break; }
  }
  if (!shiftRow) throw new Error("Shift ID not found");
  // ---------------------

  const waktuBuka = new Date(shiftRow[3]); 
  const saldoAwal = Number(shiftRow[4]) || 0;
  const timeStart = waktuBuka.getTime();
  
  // 2. Hitung Drop (Hanya untuk Info Display, tidak mengurangi saldo lagi karena sudah dijurnal)
  const sheetDrop = ss.getSheetByName(CONF_PHASE1.LOG_DROP);
  const dataDrop = sheetDrop.getDataRange().getValues();
  let totalDropHistory = 0;
  for (let i = 1; i < dataDrop.length; i++) {
    if (String(dataDrop[i][1]) === String(idShift)) {
      totalDropHistory += (Number(dataDrop[i][3]) || 0);
    }
  }
  
  // 3. Hitung Mutasi Jurnal (Single Source of Truth)
  const sheetJurnal = ss.getSheetByName(CONF_PHASE1.JURNAL_UMUM);
  const dataJurnal = sheetJurnal.getDataRange().getValues();
  
  // Index Kolom Jurnal
  const J_TIME = 19; // Waktu_Edit
  const J_CABANG = 4;
  const J_COA = 8;
  const J_DEBIT = 11;
  const J_KREDIT = 12;
  
  let totalMasuk = 0; 
  let totalKeluar = 0; 
  
  for (let i = 1; i < dataJurnal.length; i++) {
    const row = dataJurnal[i];
    if (row[J_CABANG] !== userCabang) continue;
    
    // Filter Waktu
    let rowDateObj = new Date(row[J_TIME] || row[0]); 
    if (rowDateObj.getTime() < timeStart) continue;
    
    // Filter Akun KAS (111x)
    // Logic: Semua Debit ke Kas = Uang Masuk
    //        Semua Kredit dari Kas = Uang Keluar (Termasuk Drop & Biaya)
    if (String(row[J_COA]).startsWith('111')) { 
      totalMasuk += (Number(row[J_DEBIT]) || 0);
      totalKeluar += (Number(row[J_KREDIT]) || 0);
    }
  }
  
  // 4. Kalkulasi Akhir
  // Saldo Sistem = Awal + Masuk - Keluar
  // Note: 'totalKeluar' sekarang SUDAH termasuk Cash Drop karena Drop dijurnal sebagai Kredit Kas.
  // Jadi variabel totalDropHistory TIDAK perlu dikurangkan lagi.
  
  const saldoSistem = saldoAwal + totalMasuk - totalKeluar;
  
  return {
    total_masuk: totalMasuk,
    total_keluar: totalKeluar,
    total_cash_drop: totalDropHistory, // Dikirim hanya untuk info UI
    saldo_sistem_sekarang: saldoSistem
  };
}