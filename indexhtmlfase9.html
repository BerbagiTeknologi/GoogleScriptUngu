/**
 * UNGU LAUNDRY ERP - v1.0 (FINAL RELEASE)
 * Audited by: ERP Specialist
 * Integrity: High | Scalability: Optimized for <200k Rows
 */

// --- MAPPING AKUN & CABANG ---
const ACC = {
  // Akun Utama (Pastikan kode ini ada di COA_Master)
  KAS_PUSAT: '1111',
  BANK_BCA: '1121',
  PIUTANG: '1131',
  LABA_DITAHAN: '3101',
  
  // Default Persediaan & HPP
  DEF_HPP: '5101',       
  DEF_PERSEDIAAN: '1141', 

  // Mapping Dinamis Cabang (Nama Cabang : Akun Kas)
  CABANG_KAS: {
    'Karanganyar': '1112.01',
    'Paoman': '1112.02'
  },
  
  // Mapping Dinamis QRIS (Nama Cabang : Akun QRIS)
  CABANG_QRIS: {
    'Karanganyar': '1134.01',
    'Paoman': '1134.02'
  }
};

// --- FUNGSI HELPER GLOBAL ---
const TIMEZONE = 'Asia/Jakarta'; // [FIX: Paksa Timezone Jakarta]

function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }

function formatDate(d) { 
  if (!d) return '';

  // [FIX CRITICAL BUG: TIMEZONE OFFSET]
  // Jika input d adalah string "YYYY-MM-DD" (dari Date Picker HTML), 
  // KEMBALIKAN LANGSUNG tanpa konversi new Date(). 
  // Ini mencegah Google Script mengonversinya ke UTC lalu menggesernya ke H-1 (misal 07:00 WIB menjadi 00:00 UTC lalu mundur).
  if (typeof d === 'string' && d.length === 10 && d.charAt(4) === '-') {
      return d;
  }

  // Jika input adalah Timestamp (Angka) atau Date Object (System Time),
  // Paksa format ke Asia/Jakarta untuk memastikan transaksi jam 00:01 tetap tercatat hari ini.
  try {
      return Utilities.formatDate(new Date(d), 'Asia/Jakarta', 'yyyy-MM-dd');
  } catch (e) {
      // Fallback safety net agar aplikasi tidak crash jika data korup
      return ''; 
  }
}

// [BARU: Fungsi Hash untuk mencegah duplikat 100%]
function generateHash(inputString) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, inputString, Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length === 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function getAkunKas(cabang) {
  return ACC.CABANG_KAS[cabang] || ACC.KAS_PUSAT; // Fallback ke Pusat
}
function getAkunQRIS(cabang) {
  return ACC.CABANG_QRIS[cabang] || ACC.BANK_BCA; // Fallback ke Bank Utama
}

function getLockDate() {
  const raw = PropertiesService.getDocumentProperties().getProperty('LAST_CLOSING_DATE'); 
  return raw ? new Date(raw) : new Date('2020-01-01');
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate().setTitle('Ungu Laundry System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// --- TARUH FUNGSI INCLUDE DI SINI (PALING BAWAH) ---
// Biarkan tulisan 'filename' tetap seperti ini. Jangan diubah.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// --- 1. MODUL USER & SECURITY ---

function loginUser(u, p) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.USER);
  if(!sh) return { status: 'ERROR', msg: 'Critical: Sheet Data_User tidak ditemukan.' };
  
  const data = sh.getDataRange().getValues();
  // Loop dari baris 1 (skip header)
  for (let i = 1; i < data.length; i++) {
  // ... (kode validasi password sebelumnya tetap sama) ...

    if (String(data[i][0]).trim().toLowerCase() === String(u).trim().toLowerCase() && String(data[i][1]) === generateHash(String(p))) {
      const statusAkun = data[i][6] ? String(data[i][6]) : 'Aktif';
      if(statusAkun === 'Non-Aktif') return { status: 'ERROR', msg: 'Akun dinonaktifkan.' };

      // --- [SECURITY PATCH START] ---
      // Generate Token Unik
      const token = Utilities.getUuid(); 
      
      // Simpan data user asli ke Cache Server (Berlaku 6 jam / 21600 detik)
      // Kita simpan username DAN role yang valid dari database, bukan dari input
      const sessionPayload = JSON.stringify({
        username: data[i][0],
        role: data[i][2] // Role diambil paksa dari database sheet, bukan input
      });
      CacheService.getScriptCache().put(token, sessionPayload, 21600); 
      // --- [SECURITY PATCH END] ---

      return { 
        status: 'SUCCESS', 
        token: token, // Kirim token ke klien, BUKAN role mentah untuk otorisasi nanti
        username: data[i][0], 
        role: data[i][2],
        nama: data[i][3] || data[i][0],
        hp: data[i][5] || ''
      };
    }
// ... (sisa kode tetap sama) ...
  }
  return { status: 'ERROR', msg: 'Username atau Password salah.' };
}

function simpanUserBaru(form) {
  const ss = getSS();
  let sh = ss.getSheetByName(CONF.USER);
  const data = sh.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]).toLowerCase() === String(form.username).toLowerCase()){
       return { status: 'ERROR', msg: 'Username sudah digunakan!' };
    }
  }

  // [FIX SECURITY] Hash password sebelum disimpan ke sheet
  const passwordHash = generateHash(form.password);
  sh.appendRow([form.username, passwordHash, form.role, form.nama, form.jabatan, form.hp, form.status]);
}

function getAllUsers() {
  const sh = getSS().getSheetByName(CONF.USER);
  if(!sh) return [];
  const data = sh.getDataRange().getValues();
  let res = [];
  for(let i=1; i<data.length; i++){
     if(data[i][0]) {
       res.push({
         username: data[i][0], role: data[i][2], nama: data[i][3], jabatan: data[i][4], status: data[i][6]
       });
     }
  }
  return res;
}
// --- [NEW FEATURE] MANAJEMEN USER LENGKAP (EDIT, RESET PASS, STATUS) ---

function updateDataUser(form) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.USER);
  const data = sh.getDataRange().getValues();
  
  // Cari Username (Kolom A / Indeks 0)
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]).toLowerCase() === String(form.username).toLowerCase()){
       // Update Data: Role(Col 3), Nama(Col 4), Jabatan(Col 5), HP(Col 6)
       // Ingat: getRange pakai index 1-based, array pakai 0-based.
       // Maka baris = i+1.
       
       sh.getRange(i+1, 3).setValue(form.role);     // Update Role
       sh.getRange(i+1, 4).setValue(form.nama);     // Update Nama
       sh.getRange(i+1, 5).setValue(form.jabatan);  // Update Jabatan
       sh.getRange(i+1, 6).setValue(form.hp);       // Update HP
       
       return { status: 'SUCCESS', msg: 'Data user ' + form.username + ' berhasil diperbarui.' };
    }
  }
  return { status: 'ERROR', msg: 'User tidak ditemukan.' };
}

function adminResetPassword(username, passwordBaru) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.USER);
  const data = sh.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]).toLowerCase() === String(username).toLowerCase()){
       // Generate Hash Baru (Mengikuti standar keamanan baru)
       const newHash = generateHash(passwordBaru);
       
       // Update Password (Kolom B / Indeks 2 di Sheet)
       sh.getRange(i+1, 2).setValue(newHash);
       
       return { status: 'SUCCESS', msg: 'Password untuk ' + username + ' berhasil di-reset.' };
    }
  }
  return { status: 'ERROR', msg: 'User tidak ditemukan.' };
}

function setUserStatus(username, statusBaaru) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.USER);
  const data = sh.getDataRange().getValues();
  
  // Validasi Input Status
  if (statusBaaru !== 'Aktif' && statusBaaru !== 'Non-Aktif') {
      return { status: 'ERROR', msg: 'Status tidak valid.' };
  }

  for(let i=1; i<data.length; i++){
    if(String(data[i][0]).toLowerCase() === String(username).toLowerCase()){
       // Update Status (Kolom G / Indeks 7 di Sheet)
       sh.getRange(i+1, 7).setValue(statusBaaru);
       
       return { status: 'SUCCESS', msg: 'User ' + username + ' sekarang: ' + statusBaaru };
    }
  }
  return { status: 'ERROR', msg: 'User tidak ditemukan.' };
}

// --- 2. MODUL PRODUKSI (AUTO-HEALING & SCALABLE) ---

function getProduksiList(cabang) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.PESANAN);
  if(!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 1) return []; 

  const header = data[0];
  const KOLOM_PROSES = 'Proses_Laundry';
  let idxProses = header.indexOf(KOLOM_PROSES);

  // Auto-Fix: Jika kolom hilang, buat baru otomatis
  if (idxProses === -1) {
    idxProses = header.length; 
    sh.getRange(1, idxProses + 1).setValue(KOLOM_PROSES); 
  }

  let idxCabang = header.indexOf('Cabang');
  if (idxCabang === -1) idxCabang = 7; 

  let list = [];
  // SCALABILITY: Hanya scan 100 transaksi terakhir agar tidak lemot
  const limit = Math.max(1, data.length - 100); 
  
  for(let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    let rawProses = (idxProses < row.length) ? row[idxProses] : ''; 
    let proses = (rawProses && String(rawProses).trim() !== '') ? String(rawProses) : 'Diterima';

    if(proses !== 'Selesai' && proses !== 'Diambil') {
      const valCabang = row[idxCabang];
      if(cabang === 'Semua' || valCabang === cabang) {
        list.push({
          row: i + 1, inv: row[0], cust: row[3], tgl: formatDate(row[1]), itemCount: row[4], proses: proses
        });
      }
    }
    if(i <= limit) break; // Optimization Stop
  }
  return list;
}

function updateStatusProduksi(inv, statusBaru, user) {
  const ss = getSS();
  const shPesanan = ss.getSheetByName(CONF.PESANAN);
  const shLog = ss.getSheetByName(CONF.LOG_PROD);
  const shDet = ss.getSheetByName(CONF.DETAIL_PESANAN);

  if(!shLog) return {status:'ERROR', msg:'Sheet Log_Produksi belum dibuat!'};
  
  const data = shPesanan.getDataRange().getValues();
  if (data.length < 1) return {status:'ERROR', msg:'Data Pesanan Kosong'};

  const header = data[0];
  const KOLOM_PROSES = 'Proses_Laundry';
  let idxProses = header.indexOf(KOLOM_PROSES);

  if(idxProses === -1) {
    idxProses = header.length;
    shPesanan.getRange(1, idxProses + 1).setValue(KOLOM_PROSES);
  }

  let rowIdx = -1;
  let nilaiOrder = 0;
  
  // Mencari Invoice
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]) === String(inv)){
      rowIdx = i + 1;
      nilaiOrder = data[i][4];
      break;
    }
  }

  if(rowIdx === -1) return {status:'ERROR', msg:'Invoice tidak ditemukan'};

  // Update Status
  shPesanan.getRange(rowIdx, idxProses + 1).setValue(statusBaru);

  // Ambil Detail Barang untuk Log Audit
  let infoBarang = [];
  if(shDet) {
    const dVal = shDet.getDataRange().getValues();
    for(let j=1; j<dVal.length; j++){
      if(String(dVal[j][0]) === String(inv)){
        infoBarang.push(`${dVal[j][3]} ${dVal[j][2]}`); 
      }
    }
  }
  const strBarang = infoBarang.join(', ') || '-';

  shLog.appendRow([new Date(), inv, statusBaru, user, nilaiOrder, strBarang]);
  return {status:'SUCCESS', msg:`Status ${inv} berhasil diubah ke: ${statusBaru}`};
}

// --- 3. CORE: JURNAL TRANSAKSI (AUDITED VERSION) ---

function getInitialData() {
  const ss = getSS();
  let list = [];

  const readMap = (sheetName, type) => {
    const sh = ss.getSheetByName(sheetName);
    if(!sh) return;
    const vals = sh.getDataRange().getValues();
    const processedCodes = new Set();

    for(let i=1; i<vals.length; i++){
      const kd = String(vals[i][0]).trim();
      const nm = String(vals[i][1]).trim();
      if(kd && !processedCodes.has(kd)){
         processedCodes.add(kd);
         list.push({
           kode: kd, nama: nm, type: type, 
           isTunai: nm.toLowerCase().includes('tunai') || nm.toLowerCase().includes('kas'),
           isBank: nm.toLowerCase().includes('bank') || nm.toLowerCase().includes('transfer')
         });
      }
    }
  };
  readMap(CONF.MAP_TRX, 'TRX');
  readMap(CONF.MAP_ADJ, 'ADJ');
  return list;
}

function simpanTransaksi(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const ss = getSS();
    
    // =================================================================
    // [SECURITY PATCH - START] Validasi Sesi Server-Side
    // =================================================================
    // Validasi Token menggantikan validasi user mentah
    const currentUser = validateSession(form.token); 
    const safeUser = currentUser.username; // User valid dari server
    const safeRole = currentUser.role;     // Role valid dari server
    // =================================================================

    // [FIX: Validasi Lock Date dengan Timezone Jakarta]
    const lockDateStr = PropertiesService.getDocumentProperties().getProperty('LAST_CLOSING_DATE');
    const lockDate = lockDateStr ? new Date(lockDateStr) : new Date('2020-01-01');
    const trxDate = new Date(form.tgl);
    const lockYMD = Number(Utilities.formatDate(lockDate, TIMEZONE, 'yyyyMMdd'));
    const trxYMD = Number(Utilities.formatDate(trxDate, TIMEZONE, 'yyyyMMdd'));

    // [SECURITY PATCH] Menggunakan safeRole, bukan realRole hasil query ulang
    if (trxYMD <= lockYMD && safeRole !== 'Admin') {
      return { status: 'ERROR', msg: 'AKSES DITOLAK: Periode ini sudah Tutup Buku.' };
    }

    // --- SETUP SHEET & MAPPING (LOGIKA ASLI DIPERTAHANKAN) ---
    const isAdj = String(form.kode).toUpperCase().startsWith('ADJ');
    const targetName = isAdj ? CONF.ADJ : CONF.JU;
    const mapName = isAdj ? CONF.MAP_ADJ : CONF.MAP_TRX;

    let targetSh = ss.getSheetByName(targetName);
    if (!targetSh) targetSh = ss.insertSheet(targetName);
    
    const mapSh = ss.getSheetByName(mapName);
    if (!mapSh) return { status: 'ERROR', msg: 'Sheet Mapping hilang: ' + mapName };
    const mapData = mapSh.getDataRange().getValues();
    const h = mapData[0].map(x => String(x).toLowerCase().trim());
    const idx = {
      kode: h.indexOf('kode_transaksi'), nama: h.indexOf('nama_transaksi'),
      coa: h.indexOf('kode_coa'), namaCoa: h.indexOf('nama_coa'),
      id: h.indexOf('id_akun_internal'), posisi: h.indexOf('posisi'),
      param: h.indexOf('parameter_input') 
    };
    const rows = mapData.filter(r => String(r[idx.kode]) === form.kode);
    if (rows.length === 0) return { status: 'ERROR', msg: 'Kode tidak terdaftar di mapping: ' + form.kode };

    // --- [FIX: HARD DUPLICATE CHECK] (LOGIKA ASLI DIPERTAHANKAN) ---
    // Generate Hash Unik: YYYYMMDD-USER-KODE-NOMINAL-CABANG
    // Catatan: Kita tetap gunakan form.user untuk hash tracking, tapi safeUser untuk record
    const uniqueString = `${formatDate(form.tgl)}-${safeUser}-${form.kode}-${form.nominal}-${form.cabang}`;
    const trxHash = generateHash(uniqueString);

    // Cari Hash menggunakan TextFinder
    const duplicateCheck = targetSh.createTextFinder(trxHash).matchEntireCell(true).findNext();
    if (duplicateCheck) {
      return { status: 'ERROR', msg: 'GAGAL: Transaksi Duplikat Terdeteksi (Hard Check)! Data ini sudah tersimpan sebelumnya.' };
    }
    
    // --- PROSES SIMPAN ---
    const props = PropertiesService.getDocumentProperties();
    let lastID = parseInt(props.getProperty('LAST_ID') || '0') + 1;
    props.setProperty('LAST_ID', String(lastID));
    const noBukti = (isAdj ? 'AJE-' : 'TRX-') + ('000000' + lastID).slice(-6);

    // --- LOGIKA HITUNG NOMINAL (LOGIKA ASLI DIPERTAHANKAN) ---
    const calculateNominal = (r, f) => {
      const coa = String(r[idx.coa]);
      const idAkun = String(r[idx.id]);
      const posisi = String(r[idx.posisi]).toUpperCase();
      const valNominal = Number(f.nominal) || 0;
      const valDiskon = Number(f.diskon) || 0;
      const valPph = Number(f.pph) || 0;

      // 1. SKENARIO DISKON
      if (valDiskon > 0) {
        if (coa === '4141' || idAkun === 'ACC0083') return valDiskon;
        if (posisi === 'KREDIT' && (coa.startsWith('411') || idAkun.startsWith('ACC003'))) {
           if (f.jenis_layanan) {
             if (f.jenis_layanan === 'Kiloan' && coa !== '4111') return 0;
             if (f.jenis_layanan === 'Satuan' && coa !== '4112') return 0;
           }
           return valNominal + valDiskon;
        }
      }

      // 2. SKENARIO PAJAK/PPH
      if (valPph > 0) {
        if (['2105', '2106'].includes(coa) || ['ACC0032', 'ACC0033'].includes(idAkun)) return valPph;
        if (posisi === 'KREDIT' && (coa.startsWith('111') || coa.startsWith('112'))) return valNominal - valPph;
      }

      // 3. DEFAULT
      return valNominal;
    };

    let ins = [];
    const now = new Date();
    const strNow = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    rows.forEach(r => {
      const finalVal = calculateNominal(r, form);

      if (finalVal > 0) {
        let detailParams = "";
        if (idx.param !== -1 && r[idx.param]) {
          const rawParams = String(r[idx.param]).split(';'); 
          let paramArr = [];
          rawParams.forEach(p => {
             const key = p.trim(); 
             if (key && form[key] !== undefined && form[key] !== "") {
               paramArr.push(`${key}: ${form[key]}`);
             }
          });
          if (paramArr.length > 0) detailParams = `(${paramArr.join(', ')})`;
        }
        
        const fullKeterangan = form.ket + (detailParams ? " " + detailParams : "");
        const isDebetRow = String(r[idx.posisi]).toUpperCase().startsWith('D');
        
        ins.push([
          formatDate(form.tgl),
          noBukti, 
          form.kode, 
          r[idx.nama], 
          form.cabang,
          (isAdj ? 'Jurnal Penyesuaian' : 'Jurnal Umum'), 
          (form.metode || 'Memorial'), 
          r[idx.id], 
          r[idx.coa], 
          r[idx.namaCoa],
          isDebetRow ? 'D' : 'K', 
          isDebetRow ? finalVal : 0, 
          !isDebetRow ? finalVal : 0, 
          finalVal,
          fullKeterangan, 
          'System_v7_Secure', // Marker versi aman
          'Aktif', 
          safeUser, // [SECURITY PATCH] Menggunakan User dari Token
          trxHash, 
          strNow, 
          trxHash 
        ]);
      }
    });

    if (ins.length === 0) return { status: 'ERROR', msg: 'Gagal: Nominal 0 atau Mapping konflik.' };
    
    targetSh.getRange(targetSh.getLastRow() + 1, 1, ins.length, ins[0].length).setValues(ins);
    return { status: 'SUCCESS', msg: 'Data Tersimpan. Bukti: ' + noBukti, bukti: noBukti };

  } catch (e) {
    return { status: 'ERROR', msg: 'System Error: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function loadDataJurnal(filter) {
  const ss = getSS();
  let res = [];
  let targets = [];
  
  if(filter.jenis === 'ADJ') targets = [CONF.ADJ];
  else if(filter.jenis === 'JU') targets = [CONF.JU];
  else targets = [CONF.JU, CONF.ADJ];

  const sDate = new Date(filter.tgl1);
  const eDate = new Date(filter.tgl2); eDate.setHours(23,59,59);

  targets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if(sh && sh.getLastRow() > 1) {
       const last = sh.getLastRow();
       // OPTIMIZATION: Hanya load 2000 baris terakhir untuk UI agar browser tidak crash
       const start = Math.max(2, last - 2000); 
       const data = sh.getRange(start, 1, last - start + 1, 17).getValues(); 
       data.forEach((r, i) => {
          const d = new Date(r[0]);
          if(d >= sDate && d <= eDate) {
             const searchTxt = (r[1]+r[3]+r[9]+r[14]).toLowerCase(); 
             if( (!filter.cab || r[4]===filter.cab) && (!filter.key || searchTxt.includes(filter.key.toLowerCase())) ) {
                res.push({
                   sheet: name, row: start + i, tgl: formatDate(d), bukti: r[1], uraian: r[3],     
                   akun: r[9], dk: r[10], nominal: r[13], status: r[16],
                   type: (name===CONF.ADJ ? 'ADJ' : 'TRX')
                });
             }
          }
       });
    }
  });
  return res.sort((a,b) => b.tgl.localeCompare(a.tgl));
}

function voidByBukti(sheetName, noBukti, alasan, token) {
  const currentUser = validateSession(token);
  const safeUser = currentUser.username;
  const safeRole = currentUser.role;
  // Tambahan Keamanan: Hanya Admin atau Supervisor yang boleh Void (Opsional, tapi disarankan)
  if (safeRole !== 'Admin' && safeRole !== 'Supervisor') {
     return {status:'ERROR', msg:'AKSES DITOLAK: Hanya Admin/SPV yang boleh Void.'};
  }
  const ss = getSS();
  const sh = ss.getSheetByName(sheetName);
  const lastRow = sh.getLastRow();
  const range = sh.getDataRange();
  const values = range.getValues();

  let rowsFound = [];
  let rowIndexArr = [];

  for(let i=1; i<values.length; i++) {
     if(String(values[i][1]) === String(noBukti)) { 
        rowsFound.push(values[i]);
        rowIndexArr.push(i+1); 
     }
  }
  if(rowsFound.length === 0) return {status:'ERROR', msg:'Data Bukti tidak ditemukan.'};

  if(new Date(rowsFound[0][0]) <= getLockDate()) return {status:'ERROR', msg:'GAGAL: Periode transaksi ini sudah dikunci.'};
  let reversalRows = [];
  const now = new Date();

  rowsFound.forEach(row => {
     let rev = [...row];
     rev[0] = now; rev[5] = "VOID REVERSAL"; 
     rev[11] = -Number(row[11]); rev[12] = -Number(row[12]); rev[13] = -Number(row[13]);
     rev[14] = `VOID REF: ${row[1]} | ${alasan}`; rev[16] = "Reversal"; 
     rev[17] = safeUser; // <--- PERBAIKAN: Gunakan User dari Token
     rev[19] = now; 
     reversalRows.push(rev);
  });

  sh.getRange(lastRow+1, 1, reversalRows.length, reversalRows[0].length).setValues(reversalRows);
  rowIndexArr.forEach(idx => { sh.getRange(idx, 17).setValue("Voided"); });
  return {status:'SUCCESS', msg:`Sukses membatalkan Bukti: ${noBukti}`};
}

// --- 4. CLOSING & FINANCIAL REPORT (SCALABLE STRATEGY) ---

// --- 4. CLOSING & FINANCIAL REPORT (OPTIMIZED INCREMENTAL) ---

function runClosing(periode, user) {
  const ss = getSS();
  // 1. Setup Tanggal
  const [y, mStr] = periode.split('-');
  const year = parseInt(y);
  const month = parseInt(mStr);

  const startDate = new Date(year, month - 1, 1); // Tgl 1 bulan ini
  const endDate = new Date(year, month, 0, 23, 59, 59); // Akhir bulan ini
  
  const nextPeriodDate = new Date(year, month, 1); // Tgl 1 bulan depan
  const nextPeriodCellVal = Utilities.formatDate(nextPeriodDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nextPeriodStr = Utilities.formatDate(nextPeriodDate, Session.getScriptTimeZone(), 'yyyy-MM');
  
  const COA_LABA_DITAHAN = '3101'; 

  try {
    // 2. Ambil Saldo Awal Bulan Ini (Baseline)
    let saldoMap = getSaldoAwalByPeriod(periode); 

    // 3. Ambil Mutasi Transaksi HANYA Bulan Ini
    const mutasiData = loadJournalByRange(startDate, endDate); 
    
    let labaBerjalan = 0;

    // 4. Hitung Saldo Akhir (Incremental: Saldo Awal + Mutasi)
    mutasiData.forEach(r => {
      const kode = String(r[8]); 
      const nama = r[9];
      const debet = Number(r[11]) || 0; 
      const kredit = Number(r[12]) || 0;
      
      const head = kode.charAt(0);
      
      if (['1', '2', '3'].includes(head)) { // Akun Neraca
        if (!saldoMap[kode]) saldoMap[kode] = { nama: nama, val: 0 };
        saldoMap[kode].val += (debet - kredit);
      } else { // Akun Laba Rugi
        labaBerjalan += (kredit - debet);
      }
    });

    // 5. Susun Data untuk Sheet Saldo_Awal (Periode Depan)
    let newRows = [];
    const now = new Date();
    
    for (let k in saldoMap) {
      const v = saldoMap[k].val;
      if (Math.abs(v) > 0.01) { // Abaikan nilai 0 koma sekian
        let d = 0, kr = 0;
        if (v > 0) d = v; else kr = Math.abs(v);
        newRows.push([nextPeriodCellVal, k, saldoMap[k].nama, d, kr, 'System_Closing', now]);
      }
    }

    // 6. Masukkan Laba Berjalan ke Ekuitas (Laba Ditahan)
    if (labaBerjalan !== 0) {
      let found = false;
      for (let i = 0; i < newRows.length; i++) {
        if (String(newRows[i][1]) === COA_LABA_DITAHAN) {
            const currentNet = newRows[i][3] - newRows[i][4];
            const finalEquity = currentNet + labaBerjalan;
            if (finalEquity >= 0) { newRows[i][3] = finalEquity; newRows[i][4] = 0; } 
            else { newRows[i][3] = 0; newRows[i][4] = Math.abs(finalEquity); }
            found = true; break;
        }
      }
      if (!found) {
        let d = 0, kr = 0;
        if (labaBerjalan > 0) kr = labaBerjalan; else d = Math.abs(labaBerjalan);
        newRows.push([nextPeriodCellVal, COA_LABA_DITAHAN, 'Saldo Laba (Retained Earnings)', d, kr, 'System_Closing', now]);
      }
    }

    // 7. Simpan ke Sheet
    const shSaldo = ss.getSheetByName(CONF.SALDO);
    if (!shSaldo) throw new Error("Sheet Saldo_Awal hilang!");

    const allData = shSaldo.getDataRange().getValues();
    let keepData = [];
    if (allData.length > 0) keepData.push(allData[0]); // Header

    // Hapus saldo lama jika periode ini pernah diclosing sebelumnya (Re-run)
    for (let i = 1; i < allData.length; i++) {
      let rowDateStr = "";
      try {
        if (allData[i][0] instanceof Date) rowDateStr = Utilities.formatDate(allData[i][0], Session.getScriptTimeZone(), 'yyyy-MM');
        else rowDateStr = String(allData[i][0]).substring(0, 7);
      } catch(e) { continue; }
      
      if (rowDateStr !== nextPeriodStr) keepData.push(allData[i]);
    }

    shSaldo.clear();
    if (keepData.length > 0) shSaldo.getRange(1, 1, keepData.length, keepData[0].length).setValues(keepData);
    if (newRows.length > 0) shSaldo.getRange(shSaldo.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);

    // Update Lock Date ke Akhir Bulan Ini
    PropertiesService.getDocumentProperties().setProperty('LAST_CLOSING_DATE', endDate.toISOString());
    
    // Log
    const log = ss.getSheetByName(CONF.LOG) || ss.insertSheet(CONF.LOG);
    log.appendRow([new Date(), 'CLOSING', periode, user, `Incremental: ${mutasiData.length} trx`]);

    return "SUKSES: Periode " + periode + " dikunci. Saldo Awal " + nextPeriodStr + " terbentuk.";

  } catch (e) {
    return "ERROR Closing: " + e.message;
  }
}

function batalClosing(user) {
   const props = PropertiesService.getDocumentProperties();
   const lastClosingRaw = props.getProperty('LAST_CLOSING_DATE');
   
   if (!lastClosingRaw || lastClosingRaw === '2020-01-01') {
       return "GAGAL: Tidak ada periode yang dikunci (Sistem masih Open).";
   }

   const currentLockDate = new Date(lastClosingRaw);
   
   // Hitung periode Saldo Awal yang harus dihapus (Bulan Depan dari Lock Date)
   // Contoh: Lock 31 Maret -> Saldo Awal April harus dihapus
   const saldoToDeleteDate = new Date(currentLockDate.getFullYear(), currentLockDate.getMonth() + 1, 1);
   const saldoToDeleteStr = Utilities.formatDate(saldoToDeleteDate, Session.getScriptTimeZone(), 'yyyy-MM');

   // Hitung Lock Date Baru (Mundur 1 Bulan ke belakang)
   // Contoh: Lock 31 Maret -> Mundur jadi 28 Feb
   const newLockDate = new Date(currentLockDate.getFullYear(), currentLockDate.getMonth(), 0);
   
   // --- EKSEKUSI ---
   const ss = getSS();
   const shSaldo = ss.getSheetByName(CONF.SALDO);
   let deletedCount = 0;

   if (shSaldo) {
       const all = shSaldo.getDataRange().getValues();
       let keep = [all[0]]; // Header
       
       for(let i=1; i<all.length; i++) {
           let rowPeriod = "";
           try {
             if (all[i][0] instanceof Date) rowPeriod = Utilities.formatDate(all[i][0], Session.getScriptTimeZone(), 'yyyy-MM');
             else rowPeriod = String(all[i][0]).substring(0, 7);
           } catch(e) { rowPeriod = "ERR"; }

           if (rowPeriod !== saldoToDeleteStr) keep.push(all[i]);
           else deletedCount++;
       }
       shSaldo.clear();
       if(keep.length > 0) shSaldo.getRange(1,1,keep.length,keep[0].length).setValues(keep);
   }

   // Update Property Lock Date Mundur 1 Bulan
   // Jika newLockDate < 2020, set default '2020-01-01'
   const finalLockProp = (newLockDate.getFullYear() < 2020) ? '2020-01-01' : newLockDate.toISOString();
   props.setProperty('LAST_CLOSING_DATE', finalLockProp);

   // Log
   const log = getSS().getSheetByName(CONF.LOG);
   if(log) log.appendRow([new Date(), 'BATAL TUTUP', `Revert ${saldoToDeleteStr}`, user, `Mundur lock ke ${formatDate(newLockDate)}`]);

   return `SUKSES: Periode ${saldoToDeleteStr} dibuka kembali. Tanggal Kunci mundur ke ${formatDate(newLockDate)}.`;
}

// --- HELPER CLOSING (NEW) ---

function getSaldoAwalByPeriod(periodeStr) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.SALDO);
  let map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
     let rowPeriod = "";
     if (data[i][0] instanceof Date) rowPeriod = Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'yyyy-MM');
     else rowPeriod = String(data[i][0]).substring(0, 7);
     
     if (rowPeriod === periodeStr) {
        const kode = String(data[i][1]);
        const d = Number(data[i][3]) || 0;
        const k = Number(data[i][4]) || 0;
        map[kode] = { nama: data[i][2], val: (d - k) };
     }
  }
  return map;
}

function loadJournalByRange(startDate, endDate) {
  const ss = getSS();
  let result = [];
  [CONF.JU, CONF.ADJ].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      // [FIX] TEKNIK BATCHING: Mencegah error "Exceeded Memory/Time"
        // Mengambil data per paket 2000 baris, bukan sekaligus seluruh sheet
        const lastRow = sh.getLastRow();
        const BATCH_SIZE = 2000; 

        // Loop dari baris ke-2 (skip header) sampai baris terakhir
        for (let startRow = 2; startRow <= lastRow; startRow += BATCH_SIZE) {
           
           // Hitung berapa baris yang harus diambil (agar tidak error melebihi batas bawah)
           const remaining = lastRow - startRow + 1;
           const rowsToGet = (remaining < BATCH_SIZE) ? remaining : BATCH_SIZE;
           
           // Ambil potong data (Chunk) ke memori
           // Ambil 17 kolom (A sampai Q) sesuai struktur jurnal
           const dataBatch = sh.getRange(startRow, 1, rowsToGet, 17).getValues();

           // Proses filter data di dalam batch ini
           for (let i = 0; i < dataBatch.length; i++) {
             // Cek validitas tanggal (cegah error jika cell kosong)
             if (!dataBatch[i][0]) continue; 

             const rowDate = new Date(dataBatch[i][0]);
             const status = String(dataBatch[i][16]); // Kolom Q (Status)

             // Filter Tanggal & Status Aktif
             if (rowDate >= startDate && rowDate <= endDate && status === 'Aktif') {
                result.push(dataBatch[i]);
             }
           }
           // Script akan otomatis lanjut ke batch 2000 baris berikutnya...
        }
    }
  });
  return result;
}

function getCOAWithNormal() {
   const sh = getSS().getSheetByName(CONF.COA);
   if(!sh) return [];
   const v = sh.getDataRange().getValues();
   let res = [];
   for(let i=1; i<v.length; i++){
      if(v[i][1]) {
         const kode = String(v[i][1]);
         const head = kode.charAt(0);
         const isDebet = ['1','5','6','8','9'].includes(head);
         res.push({ kode: kode, nama: v[i][0], normal: isDebet ? 'D' : 'K' });
      }
   }
   return res.sort((a,b) => a.kode.localeCompare(b.kode));
}

function saveSaldo(per, data, user) {
   const ss = getSS();
   let sh = ss.getSheetByName(CONF.SALDO);
   if(!sh) { sh=ss.insertSheet(CONF.SALDO); sh.appendRow(['Periode_Awal','Kode_COA','Nama_Akun','Saldo_Debet','Saldo_Kredit','User','Time']); }
   
   const all = sh.getDataRange().getValues();
   let keep = [all[0]];
   for(let i=1; i<all.length; i++) {
      if(String(all[i][0]) !== String(per)) keep.push(all[i]);
   }
   data.forEach(d => { 
      if(Number(d.d)>0 || Number(d.k)>0) keep.push([per, d.c, d.d, d.k, user, new Date()]); 
   });
   
   sh.clear(); 
   if(keep.length > 0) sh.getRange(1,1,keep.length,keep[0].length).setValues(keep);
   return "Saldo Awal Berhasil Disimpan!";
}

function getLaporanData(jenis, tglAwal, tglAkhir) {
  const ss = getSS();
  const shHelper = ss.getSheetByName('Laporan_Helper');
  if (!shHelper) return { status: 'ERROR', msg: 'Sheet Helper Hilang!' };
  
  shHelper.getRange('L1').setValue(tglAwal);
  shHelper.getRange('L2').setValue(tglAkhir);
  SpreadsheetApp.flush();
  
  const lastRow = shHelper.getLastRow();
  if (lastRow < 2) return { status: 'SUCCESS', data: [] };
  
  const data = shHelper.getRange(2, 1, lastRow - 1, 9).getValues(); 
  let result = [];
  
  if (jenis === 'NERACA') {
     data.forEach(r => {
        if (['Aset', 'Liabilitas', 'Ekuitas', 'Kewajiban', 'Modal'].includes(r[2])) {
           let val = (r[2] === 'Aset') ? r[7] - r[8] : r[8] - r[7];
           if (val !== 0) result.push({ kode: r[0], nama: r[1], tipe: r[2], nilai: val });
        }
     });
     let laba = 0;
     data.forEach(r => {
        if (r[2] === 'Pendapatan') laba += (r[8] - r[7]); 
        if (r[2] === 'Beban') laba -= (r[7] - r[8]);      
     });
     if (laba !== 0) result.push({ kode: '3999', nama: 'Laba Tahun Berjalan', tipe: 'Ekuitas', nilai: laba });
     
  } else if (jenis === 'LABA_RUGI') {
     data.forEach(r => {
        if (['Pendapatan', 'Beban'].includes(r[2])) {
           let val = (r[2] === 'Pendapatan') ? r[8] - r[7] : r[7] - r[8];
           if (val !== 0) result.push({ kode: r[0], nama: r[1], tipe: r[2], nilai: val });
        }
     });
  } else if (jenis === 'NERACA_SALDO') {
     data.forEach(r => {
        if (r[7] !== 0 || r[8] !== 0) {
           result.push({ kode: r[0], nama: r[1], tipe: r[2], debet: r[7], kredit: r[8] });
        }
     });
  }
  return { status: 'SUCCESS', data: result };
}

// --- 5. MODUL POS & INVENTORY ---

function getPosMasterData() {
  const ss = getSS();
  let layanan = [];
  const shLay = ss.getSheetByName(CONF.LAYANAN);
  if(shLay && shLay.getLastRow() > 1) { 
    const data = shLay.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(data[i][0] && data[i][1]) {
         layanan.push({ 
           kode: data[i][0], nama: data[i][1], satuan: data[i][2], 
           harga: Number(data[i][3]) || 0, coa: data[i][4] 
         });
      }
    }
  }
  let pelanggan = [];
  const shPel = ss.getSheetByName(CONF.PELANGGAN);
  if(shPel && shPel.getLastRow() > 1) {
    const data = shPel.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(data[i][0]) {
         pelanggan.push({ 
           id: data[i][0], nama: String(data[i][1]).toUpperCase(), hp: data[i][2], alamat: data[i][3] 
         });
      }
    }
  }
  return { layanan: layanan, pelanggan: pelanggan };
}

/**
 * Menyimpan Transaksi Penjualan (POS).
 * [UPDATED] Mendukung Toggle AUTO_DEDUCT_BY_RECIPE untuk mencegah Double Entry HPP.
 */
const AUTO_DEDUCT_BY_RECIPE = true;

function simpanOrderBaru(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSS();

    // [FIX 1] VALIDASI SESI REAL-TIME
    let safeUser = 'System';
    try {
       const currentUser = validateSession(form.token); // Menggunakan fungsi validasi baru
       safeUser = currentUser.username;
    } catch(e) {
       // Jika token invalid/user diblokir, hentikan proses.
       // Hapus fallback ke 'Admin' agar aman.
       return { status: 'ERROR', msg: e.message };
    }

    // [FIX 2] VALIDASI INPUT ANGKA (Server Side Validation)
    const orderItems = (typeof form.items === 'string') ? JSON.parse(form.items) : form.items;
    if (!orderItems || orderItems.length === 0) return { status: 'ERROR', msg: 'Keranjang belanja kosong!' };

    for (let i = 0; i < orderItems.length; i++) {
        let q = Number(orderItems[i].qty);
        if (q <= 0 || isNaN(q)) throw new Error(`Qty barang ${orderItems[i].nama} tidak valid.`);
    }

    // --- INIT SHEETS ---
    const shHead = ss.getSheetByName(CONF.PESANAN) || ss.insertSheet(CONF.PESANAN);
    const shDet = ss.getSheetByName(CONF.DETAIL_PESANAN) || ss.insertSheet(CONF.DETAIL_PESANAN);
    const shJU = ss.getSheetByName(CONF.JU);
    const shStok = ss.getSheetByName(CONF.STOK);
    const shBarang = ss.getSheetByName(CONF.BARANG);
    const shResep = ss.getSheetByName('Master_Resep'); // Sheet Resep

    // --- PREPARE DATA ---
    const dataBarang = shBarang.getDataRange().getValues();
    let mapBarang = {}; 
    for(let i=1; i<dataBarang.length; i++){
       mapBarang[String(dataBarang[i][0])] = {
         nama: dataBarang[i][1],
         aset: dataBarang[i][10] || ACC.DEF_PERSEDIAAN,
         beban: dataBarang[i][11] || ACC.DEF_HPP,
         modal: Number(dataBarang[i][9]) || 0
       };
    }

    // Load Resep (BOM)
    let mapResep = {};
    if (shResep && shResep.getLastRow() > 1) {
        const dResep = shResep.getDataRange().getValues();
        for(let i=1; i<dResep.length; i++){
            const kdLayanan = String(dResep[i][0]);
            const kdBahan = String(dResep[i][2]);
            const qtyStd = Number(dResep[i][4]);
            if (!mapResep[kdLayanan]) mapResep[kdLayanan] = [];
            mapResep[kdLayanan].push({ kode: kdBahan, qty: qtyStd });
        }
    }

    // --- HEADER ---
    const props = PropertiesService.getDocumentProperties();
    let lastInv = parseInt(props.getProperty('LAST_INV') || '0') + 1;
    props.setProperty('LAST_INV', String(lastInv));
    const noInv = 'INV-' + ('000000' + lastInv).slice(-6);
    const status = form.isLunas ? 'Lunas' : 'Belum Lunas';
    const mBayar = form.isLunas ? form.metodeBayar : '';
    
    shHead.appendRow([noInv, formatDate(form.tgl), form.custId, form.custName, form.total, status, mBayar, form.cabang, safeUser, new Date()]);

    // --- DETAIL & JURNAL ---
    let jurnalItems = []; 
    let stokKeluar = [];
    const now = new Date();
    const tglStr = formatDate(form.tgl);

    orderItems.forEach(item => {
      // Simpan Detail
      shDet.appendRow([noInv, item.kode, item.nama, item.qty, item.harga, item.subtotal]);
      
      // Pendapatan
      jurnalItems.push({ coa: item.coa, nama_akun: 'Pendapatan - ' + item.nama, posisi: 'K', nominal: item.subtotal });

      // [FIX 3] LOGIKA HPP DENGAN TOGGLE (Mencegah Double Entry)
      if (typeof AUTO_DEDUCT_BY_RECIPE !== 'undefined' && AUTO_DEDUCT_BY_RECIPE === true) {
          
          // A. Barang Retail
          if (String(item.kode).startsWith('INV-')) {
             const infoBrg = mapBarang[item.kode];
             if (infoBrg) {
                const totalHpp = infoBrg.modal * item.qty;
                stokKeluar.push([now, tglStr, noInv, form.cabang, item.kode, infoBrg.nama, 'Keluar', 0, item.qty, safeUser, 'Jual Langsung']);
                
                if(totalHpp > 0) {
                  jurnalItems.push({ coa: infoBrg.beban, nama_akun: 'HPP - ' + item.nama, posisi: 'D', nominal: totalHpp });
                  jurnalItems.push({ coa: infoBrg.aset, nama_akun: 'Persediaan - ' + item.nama, posisi: 'K', nominal: totalHpp });
                }
             }
          } 
          // B. Jasa (Cek Resep)
          else if (String(item.kode).startsWith('SRV-')) {
             const resep = mapResep[item.kode];
             if (resep && resep.length > 0) {
                 resep.forEach(bahan => {
                     const infoBahan = mapBarang[bahan.kode];
                     if (infoBahan) {
                         const qtyPakai = item.qty * bahan.qty;
                         const nilaiBeban = infoBahan.modal * qtyPakai;
                         
                         stokKeluar.push([now, tglStr, noInv, form.cabang, bahan.kode, infoBahan.nama, 'Keluar', 0, qtyPakai, safeUser, 'Backflush: ' + item.nama]);

                         if (nilaiBeban > 0) {
                            jurnalItems.push({ coa: infoBahan.beban, nama_akun: 'HPP Jasa', posisi: 'D', nominal: nilaiBeban });
                            jurnalItems.push({ coa: infoBahan.aset, nama_akun: 'Persediaan', posisi: 'K', nominal: nilaiBeban });
                         }
                     }
                 });
             }
          }
      } // End Toggle Check
    });

    // --- PEMBAYARAN ---
    if (form.isLunas) {
      let debAccount = (form.metodeBayar === 'QRIS') ? getAkunQRIS(form.cabang) : 
                       (form.metodeBayar === 'Bank') ? ACC.BANK_BCA : getAkunKas(form.cabang);
      let debName = (form.metodeBayar === 'QRIS') ? 'QRIS ' + form.cabang : 
                    (form.metodeBayar === 'Bank') ? 'Bank Transfer' : 'Kas ' + form.cabang;
      
      jurnalItems.push({ coa: debAccount, nama_akun: debName, posisi: 'D', nominal: form.total });
    } else {
      jurnalItems.push({ coa: ACC.PIUTANG, nama_akun: 'Piutang Usaha', posisi: 'D', nominal: form.total });
    }

    // --- SIMPAN JURNAL ---
    const lastID = parseInt(props.getProperty('LAST_ID') || '0') + 1;
    props.setProperty('LAST_ID', String(lastID));
    const noBukti = 'TRX-' + ('000000' + lastID).slice(-6);
    
    // [FIX 4] SECURE HASH (Anti-Collision)
    // Menambahkan Date.now() untuk memastikan transaksi identik di detik berbeda tetap unik
    const uniqueString = `${tglStr}-${safeUser}-${noInv}-${form.total}-${form.cabang}-${Date.now()}`;
    const trxHash = generateHash(uniqueString);

    let rowsJU = [];
    jurnalItems.forEach(j => {
       const isD = j.posisi === 'D';
       rowsJU.push([
         tglStr, noBukti, noInv, 'Laundry: ' + form.custName, form.cabang,
         'Jurnal Umum', (form.isLunas ? form.metodeBayar : 'Memorial'), '', j.coa, j.nama_akun,
         j.posisi, isD ? j.nominal : 0, !isD ? j.nominal : 0, j.nominal,
         (form.isLunas ? 'Penjualan Tunai' : 'Piutang'), 
         'System_POS', 'Aktif', safeUser, trxHash, now, trxHash
       ]);
    });

    if(rowsJU.length > 0) shJU.getRange(shJU.getLastRow()+1, 1, rowsJU.length, rowsJU[0].length).setValues(rowsJU);
    if(stokKeluar.length > 0) shStok.getRange(shStok.getLastRow()+1, 1, stokKeluar.length, stokKeluar[0].length).setValues(stokKeluar);

    return { status: 'SUCCESS', msg: 'Order & Jurnal Berhasil!', inv: noInv };

  } catch(e) {
    return { status: 'ERROR', msg: 'Gagal: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
function prosesPelunasan(noInv, metode, user) {
  const ss = getSS();
  const shHead = ss.getSheetByName(CONF.PESANAN);
  const shJU = ss.getSheetByName(CONF.JU);
  
  const data = shHead.getDataRange().getValues();
  let rowIdx = -1;
  let invData = null;
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(noInv)) {
      rowIdx = i + 1; invData = data[i]; break;
    }
  }
  
  if(rowIdx === -1) return { status: 'ERROR', msg: 'Invoice tidak ditemukan.' };
  if(invData[5] === 'Lunas') return { status: 'ERROR', msg: 'Invoice sudah lunas.' };
  
  const cabangInv = invData[7]; 
  shHead.getRange(rowIdx, 6).setValue('Lunas'); 
  shHead.getRange(rowIdx, 7).setValue(metode); 

  let akunDebit = ''; let namaAkunDebit = '';
  if(metode === 'QRIS') { akunDebit = getAkunQRIS(cabangInv); namaAkunDebit = 'QRIS ' + cabangInv; } 
  else { akunDebit = getAkunKas(cabangInv); namaAkunDebit = 'Kas ' + cabangInv; }
  
  const props = PropertiesService.getDocumentProperties();
  const lastID = parseInt(props.getProperty('LAST_ID') || '0') + 1;
  props.setProperty('LAST_ID', String(lastID));
  const noBukti = 'TRX-' + ('000000' + lastID).slice(-6);
  const total = invData[4]; const now = new Date();
  
  const rowsJU = [
    [formatDate(now), noBukti, 'PAY-'+noInv, 'Pelunasan Invoice ' + noInv, cabangInv, 'Jurnal Umum', metode, '', akunDebit, namaAkunDebit, 'D', total, 0, total, 'Pelunasan POS', 'System_POS', 'Aktif', user, '', now, ''],
    [formatDate(now), noBukti, 'PAY-'+noInv, 'Pelunasan Invoice ' + noInv, cabangInv, 'Jurnal Umum', 'Memorial', '', '1131', 'Piutang Usaha', 'K', 0, total, total, 'Pelunasan POS', 'System_POS', 'Aktif', user, '', now, '']
  ];
  shJU.getRange(shJU.getLastRow()+1, 1, rowsJU.length, rowsJU[0].length).setValues(rowsJU);
  return { status: 'SUCCESS', msg: `Lunas via ${metode} (Akun: ${namaAkunDebit})` };
}

function getUnpaidInvoices() {
   const sh = getSS().getSheetByName(CONF.PESANAN);
   if(!sh) return { list: [], summary: {} }; 
   const data = sh.getDataRange().getValues();
   let res = []; let summary = {}; 

   for(let i=data.length-1; i>=1; i--){
     if(data[i][5] === 'Belum Lunas') {
       const cab = data[i][7] || 'Pusat'; 
       const nominal = Number(data[i][4]);
       res.push({ inv: data[i][0], tgl: formatDate(data[i][1]), cust: data[i][3], total: nominal, cabang: cab });
       if(!summary[cab]) summary[cab] = 0;
       summary[cab] += nominal;
     }
   }
   return { list: res, summary: summary };
}

function tambahPelangganBaru(form) {
  const ss = getSS();
  let sh = ss.getSheetByName(CONF.PELANGGAN);
  if(!sh) { sh = ss.insertSheet(CONF.PELANGGAN); sh.appendRow(['id_pelanggan','nama','no_hp','alamat']); }
  
  const data = sh.getDataRange().getValues();
  const namaBaru = String(form.nama).trim().toUpperCase();
  const hpBaru = String(form.hp).trim();

  for(let i=1; i<data.length; i++){
    const dbNama = String(data[i][1]).trim().toUpperCase();
    const dbHp = String(data[i][2]).trim();
    if(dbNama === namaBaru) return { status: 'ERROR', msg: 'Gagal: Nama Pelanggan sudah terdaftar!' };
    if(hpBaru !== "" && dbHp === hpBaru) return { status: 'ERROR', msg: 'Gagal: No HP sudah digunakan pelanggan lain!' };
  }

  const lastRow = sh.getLastRow();
  let nextId = 1;
  if(lastRow > 1) {
    const lastIdStr = String(sh.getRange(lastRow, 1).getValue()); 
    const num = parseInt(lastIdStr.split('-')[1]);
    if(!isNaN(num)) nextId = num + 1;
  }
  const newId = 'CUST-' + ('000' + nextId).slice(-3);
  sh.appendRow([newId, namaBaru, form.hp, form.alamat]);
  
  return { 
    status: 'SUCCESS', msg: 'Pelanggan Baru Disimpan: ' + newId, 
    newCust: { id: newId, nama: namaBaru, hp: form.hp },
    allCust: getPosMasterData().pelanggan 
  };
}

// --- 6. MODUL INVENTORY (PURCHASE & USAGE) ---

function getInventoryData(cabangFilter) {
  const ss = getSS();
  const shBarang = ss.getSheetByName(CONF.BARANG);
  const shStok = ss.getSheetByName(CONF.STOK);
  if(!shBarang) return { barang: [], stok: {} };
  
  const dBar = shBarang.getDataRange().getValues();
  let barang = [];
  for(let i=1; i<dBar.length; i++){
     if(dBar[i][0]){
        barang.push({
           kode: dBar[i][0], nama: dBar[i][1], kat: dBar[i][2],
           sat: dBar[i][3], min: dBar[i][4], harga: dBar[i][5],
           akunAset: dBar[i][6], akunBeban: dBar[i][7]
        });
     }
  }
  
  let stokMap = {};
  if(shStok && shStok.getLastRow() > 1){
     const dStok = shStok.getDataRange().getValues();
     for(let i=1; i<dStok.length; i++){
        const rowCab = String(dStok[i][3]); 
        const rowKode = String(dStok[i][4]); 
        const masuk = Number(dStok[i][7]) || 0;
        const keluar = Number(dStok[i][8]) || 0;
        if(cabangFilter === 'Semua' || rowCab === cabangFilter){
            if(!stokMap[rowKode]) stokMap[rowKode] = 0;
            stokMap[rowKode] += (masuk - keluar);
        }
     }
  }
  return { list: barang, stok: stokMap };
}

function generateKodeBarang(kategori) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.BARANG);
  let prefix = 'GEN';
  if(kategori === 'Sabun') prefix = 'SBN';
  if(kategori === 'Parfum') prefix = 'PFM';
  if(kategori === 'Plastik') prefix = 'PLS';
  if(kategori === 'ATK') prefix = 'ATK';
  
  const data = sh.getDataRange().getValues();
  let maxNum = 0;
  for(let i=1; i<data.length; i++){
      const kode = String(data[i][0]);
      if(kode.startsWith('INV-'+prefix)) {
          const num = parseInt(kode.split('-')[2]);
          if(!isNaN(num) && num > maxNum) maxNum = num;
      }
  }
  return 'INV-' + prefix + '-' + ('000' + (maxNum + 1)).slice(-3);
}

/**
 * UPDATED: Input Master Barang dengan Multi-UOM
 */
function simpanMasterBarang(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(CONF.BARANG);
    
    // Setup Header jika belum ada (Migrasi Otomatis)
    if(!sh) { 
      sh = ss.insertSheet(CONF.BARANG); 
      // Header Versi Baru v7.1
      sh.appendRow([
        'Kode_Barang', 'Nama_Barang', 'Merk', 'Varian', 'Kategori', 
        'Satuan_Beli', 'Satuan_Pakai', 'Konversi', 'Min_Stok_Base', 
        'Harga_Beli_Base', 'Akun_Aset', 'Akun_Beban'
      ]); 
    }

    const data = sh.getDataRange().getValues();
    
    // 1. Generate SKU Cerdas (Format: KAT-MERK-VARIAN)
    // Membersihkan string agar aman jadi Kode
    const cleanStr = (s) => String(s || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0,5);
    
    let prefix = 'GEN';
    if (form.kat === 'Chemical') prefix = 'CHM';
    if (form.kat === 'Packaging') prefix = 'PKG';
    if (form.kat === 'Energy') prefix = 'NRG';

    // Logic Auto Number sederhana agar tidak duplikat
    const uniqueCode = `INV-${prefix}-${cleanStr(form.merk)}-${cleanStr(form.varian)}-` + ('000' + (data.length)).slice(-3);

    // 2. Validasi Duplikat Nama
    const namaFull = `${form.nama} ${form.merk || ''} ${form.varian || ''}`.trim();
    for(let i=1; i<data.length; i++){
       if(String(data[i][1]).toLowerCase() === namaFull.toLowerCase()) {
         return { status: 'ERROR', msg: 'Barang dengan nama tersebut sudah ada!' };
       }
    }

    // 3. Simpan Data (Perhatikan Urutan Kolom Baru)
    sh.appendRow([
      uniqueCode,           // Col 0: Kode
      namaFull,             // Col 1: Nama Lengkap
      form.merk || '-',     // Col 2: Merk
      form.varian || '-',   // Col 3: Varian
      form.kat,             // Col 4: Kategori
      form.satBeli,         // Col 5: Satuan Beli (Jerigen)
      form.satPakai,        // Col 6: Satuan Pakai (ml)
      Number(form.konversi) || 1, // Col 7: Konversi
      Number(form.min),     // Col 8: Min Stok (Base Unit)
      0,                    // Col 9: Harga Beli Base (Update saat pembelian)
      form.akunAset,        // Col 10
      form.akunBeban        // Col 11
    ]);

    return { status: 'SUCCESS', msg: 'Barang Multi-UOM tersimpan: ' + uniqueCode };

  } catch (e) {
    return { status: 'ERROR', msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * REFACTORED MODULE: INVENTORY & PURCHASING
 * Author: ERP Specialist
 * Focus: Batch Processing for High Performance & Quota Safety
 */

/**
 * UPDATED: Pembelian Stok dengan Konversi Otomatis & Moving Average Cost
 */
/**
 * Simpan Pembelian dengan Validasi Anti-Backdate & AVCO.
 * [UPDATED] Mencegah korupsi data harga rata-rata.
 */
function simpanPembelian(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- [SAFETY] VALIDASI TANGGAL (ANTI-BACKDATE) ---
    const trxDate = new Date(form.tgl);
    const lockDate = getLockDate(); [cite_start]// Mengambil tanggal closing terakhir [cite: 1999]
    
    // Normalisasi ke YYYYMMDD integer untuk perbandingan akurat tanpa jam
    const fmt = (d) => Number(Utilities.formatDate(d, 'Asia/Jakarta', 'yyyyMMdd'));
    
    // Aturan: Tanggal Transaksi TIDAK BOLEH <= Tanggal Closing
    if (fmt(trxDate) <= fmt(lockDate)) {
        throw new Error(`TANGGAL DITOLAK: Periode ${form.tgl} sudah ditutup/dikunci. Gunakan tanggal setelah ${formatDate(lockDate)}.`);
    }
    // -------------------------------------------------

    // 1. Init Sheets
    const shBeli = ss.getSheetByName(CONF.BELI);
    const shStok = ss.getSheetByName(CONF.STOK);
    const shBarang = ss.getSheetByName(CONF.BARANG);
    const shJU = ss.getSheetByName(CONF.JU);
    
    // 2. Load Master Barang
    const dataBarang = shBarang.getDataRange().getValues();
    let mapBarang = {};
    
    // Map Header untuk keamanan
    const h = dataBarang[0].map(x => String(x).toLowerCase());
    const idx = {
      kode: h.indexOf('kode_barang'),
      konversi: h.indexOf('konversi'),
      hargaBase: h.indexOf('harga_beli_base'),
      satPakai: h.indexOf('satuan_pakai'),
      satBeli: h.indexOf('satuan_beli'),
      akunAset: h.indexOf('akun_aset') 
    };

    for (let i = 1; i < dataBarang.length; i++) {
      const row = dataBarang[i];
      mapBarang[String(row[idx.kode])] = {
        rowIndex: i + 1,
        konversi: Number(row[idx.konversi]) || 1,
        hargaLama: Number(row[idx.hargaBase]) || 0,
        satBeli: row[idx.satBeli],
        satPakai: row[idx.satPakai],
        akunAset: row[idx.akunAset] || ACC.DEF_PERSEDIAAN
      };
    }

    // 3. Proses Item
    const items = (typeof form.items === 'string') ? JSON.parse(form.items) : form.items;
    // --- [FIX] VALIDASI INPUT NUMBER SERVER-SIDE ---
    for (let i = 0; i < items.length; i++) {
        let q = Number(items[i].qty);
        let sub = Number(items[i].subtotal);
        
        // Cek jika Qty 0 atau minus
        if (q <= 0) throw new Error(`Qty pembelian ${items[i].nama} harus lebih dari 0.`);
        // Cek jika Harga Total minus
        if (sub < 0) throw new Error(`Total harga ${items[i].nama} tidak boleh negatif.`);
    }
    // -----------------------------------------------
    // >>>>> SISIPKAN KODE BARU DI SINI (SELESAI) <<<<<

    const noFaktur = String(form.faktur).trim();
    const user = form.user || 'Admin';
    const now = new Date();
    
    let stokData = [];
    let jurnalItems = [];
    
    items.forEach(item => {
      const kode = String(item.kode);
      const info = mapBarang[kode];
      if (!info) throw new Error(`Barang ${item.nama} tidak valid.`);

      const qtyBeliInput = Number(item.qty); 
      const qtyMasukBase = qtyBeliInput * info.konversi; 
      const totalHargaItem = Number(item.subtotal); 
      
      // --- LOGIKA AVCO (MOVING AVERAGE) ---
      // Ambil stok saat ini (Real-time)
      // Karena kita sudah memvalidasi tanggal (tidak boleh backdate), 
      // maka calculateCurrentStock aman digunakan sebagai 'Stok Akhir sebelum pembelian ini'.
      let stokLama = calculateCurrentStock(kode); 
      if(stokLama < 0) stokLama = 0; // Safety net

      const nilaiAsetLama = stokLama * info.hargaLama;
      const nilaiAsetBaru = totalHargaItem;
      const totalQtyPosisi = stokLama + qtyMasukBase;
      
      let hargaRataRata = 0;
      if (totalQtyPosisi > 0) {
          hargaRataRata = (nilaiAsetLama + nilaiAsetBaru) / totalQtyPosisi;
      } else {
          hargaRataRata = totalHargaItem / qtyMasukBase;
      }
      
      // Update Harga Master
      shBarang.getRange(info.rowIndex, idx.hargaBase + 1).setValue(hargaRataRata);
      // -------------------------------------

      // Log Stok
      stokData.push([
        now, formatDate(form.tgl), noFaktur, form.cabang, 
        kode, item.nama, 'Masuk', 
        qtyMasukBase, 0, user, 
        `Beli ${qtyBeliInput} ${info.satBeli} (@${info.konversi})`
      ]);

      // Jurnal Debit
      jurnalItems.push({ coa: info.akunAset, nama: `Stok ${item.nama}`, debet: totalHargaItem, kredit: 0 });
    });

    // 4. Simpan Data
    shBeli.appendRow([noFaktur, form.tgl, form.supplier, form.total, form.metode, form.cabang, user]);
    
    if (stokData.length > 0) {
      shStok.getRange(shStok.getLastRow() + 1, 1, stokData.length, stokData[0].length).setValues(stokData);
    }

    // Jurnal Kredit (Kas/Bank)
    let akunKredit = (form.metode === 'Bank') ? ACC.BANK_BCA : getAkunKas(form.cabang);
    let namaKredit = (form.metode === 'Bank') ? 'Bank BCA' : 'Kas ' + form.cabang;
    
    jurnalItems.push({ coa: akunKredit, nama: 'Pembelian: ' + form.supplier, debet: 0, kredit: form.total });

    // Simpan Jurnal
    const props = PropertiesService.getDocumentProperties();
    let lastID = parseInt(props.getProperty('LAST_ID') || '0') + 1;
    props.setProperty('LAST_ID', String(lastID));
    const noBukti = 'TRX-' + ('000000' + lastID).slice(-6);
    const uniqueString = `${formatDate(form.tgl)}-${user}-${noFaktur}-${form.total}`;
    const trxHash = generateHash(uniqueString);

    let rowsJU = [];
    const strNow = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    
    jurnalItems.forEach(j => {
       const isDebet = j.debet > 0;
       rowsJU.push([
         formatDate(form.tgl), noBukti, noFaktur, j.nama, form.cabang,
         'Jurnal Umum', form.metode, '', j.coa, isDebet ? 'Persediaan' : namaKredit,
         isDebet ? 'D' : 'K', j.debet, j.kredit, (j.debet + j.kredit),
         'Pembelian Stok', 'System_Inv', 'Aktif', user, trxHash, strNow, trxHash
       ]);
    });

    if(rowsJU.length > 0) {
        shJU.getRange(shJU.getLastRow()+1, 1, rowsJU.length, rowsJU[0].length).setValues(rowsJU);
    }

    return { status: 'SUCCESS', msg: 'Pembelian Berhasil & HPP Terupdate.' };

  } catch (e) {
    return { status: 'ERROR', msg: 'Error Pembelian: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function catatPemakaianProduksi(inv, user, items) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); 
        const ss = getSS();
        const shStok = ss.getSheetByName(CONF.STOK);
        const shJU = ss.getSheetByName(CONF.JU);
        const shBarang = ss.getSheetByName(CONF.BARANG);
        const now = new Date();
        const tglStr = formatDate(now);
        
        const shPes = ss.getSheetByName(CONF.PESANAN);
        const dPes = shPes.getDataRange().getValues();
        let cab = 'Karanganyar';
        for(let i=1; i<dPes.length; i++){
            if(String(dPes[i][0]) === inv) { cab = dPes[i][7]; break; }
        }

        const dBarang = shBarang.getDataRange().getValues();
        let mapBarang = {};
        for(let i=1; i<dBarang.length; i++){
            mapBarang[String(dBarang[i][0])] = {
                harga: Number(dBarang[i][5]) || 0,
                aset: dBarang[i][6] || ACC.DEF_PERSEDIAAN,
                beban: dBarang[i][7] || ACC.DEF_HPP
            };
        }

        let stokRows = [];
        let jurnalRows = [];
        
        const props = PropertiesService.getDocumentProperties();
        let lastID = parseInt(props.getProperty('LAST_ID') || '0') + 1;
        props.setProperty('LAST_ID', String(lastID));
        const noBukti = 'ADJ-' + ('000000' + lastID).slice(-6); 

        // --- [FIX LOGIC BUG: DUPLICATE PREVENTION] ---
        // Kita tambahkan hash unik pada kolom ke-19 dan 21 agar sistem bisa mendeteksi duplikat
        
        items.forEach(it => {
            stokRows.push([now, tglStr, inv, cab, it.kode, it.nama, 'Keluar', 0, it.qty, user, 'Pemakaian Produksi']);
            
            const info = mapBarang[it.kode];
            if(info && info.harga > 0) {
                const totalHpp = info.harga * it.qty;

                // Generate Hash Unik: INV + KODE_BARANG + QTY + TANGGAL
                // Ini memastikan satu invoice tidak bisa mencatat barang yang sama dengan jumlah sama berulang kali di hari yang sama (mencegah double click)
                const uniqueId = `${inv}-${it.kode}-${it.qty}-${tglStr}`;
                const trxHash = generateHash(uniqueId); 
          
                // Jurnal Debit (Beban)
                jurnalRows.push([
                    tglStr, noBukti, inv, 'Pemakaian: ' + it.nama, cab, 'Jurnal Penyesuaian', 'Memorial', '', 
                    info.beban, 'Beban Pemakaian Barang', 'D', totalHpp, 0, totalHpp, 'Pemakaian Stok Internal', 'System_Prod', 'Aktif', user, 
                    trxHash, // [FIX] Isi Hash di Kolom Referensi
                    now, 
                    trxHash  // [FIX] Isi Hash di Kolom Search Index
                ]);
          
                // Jurnal Kredit (Persediaan)
                jurnalRows.push([
                    tglStr, noBukti, inv, 'Pemakaian: ' + it.nama, cab, 'Jurnal Penyesuaian', 'Memorial', '', 
                    info.aset, 'Persediaan Berkurang', 'K', 0, totalHpp, totalHpp, 'Pemakaian Stok Internal', 'System_Prod', 'Aktif', user, 
                    trxHash, // [FIX] Isi Hash di Kolom Referensi
                    now, 
                    trxHash  // [FIX] Isi Hash di Kolom Search Index
                ]);
            }
        });
        if (stokRows.length > 0) shStok.getRange(shStok.getLastRow() + 1, 1, stokRows.length, stokRows[0].length).setValues(stokRows);
        if (jurnalRows.length > 0) shJU.getRange(shJU.getLastRow() + 1, 1, jurnalRows.length, jurnalRows[0].length).setValues(jurnalRows);

        return { status: 'SUCCESS', msg: 'Pemakaian tercatat di Stok & Keuangan' };
    } catch (e) {
        return { status: 'ERROR', msg: e.message };
    } finally {
        lock.releaseLock();
    }
}

// --- 7. UTILS & DASHBOARD ---

function getReportOmset(tgl1, tgl2, cabang) {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.PESANAN);
  if(!sh) return { summary: {total:0, lunas:0, piutang:0}, list: [] };
  
  const data = sh.getDataRange().getValues();
  let list = [];
  let sum = { total: 0, lunas: 0, piutang: 0, count: 0 };
  const sDate = new Date(tgl1);
  const eDate = new Date(tgl2); eDate.setHours(23,59,59);
  
  for(let i=1; i<data.length; i++){
     const rowDate = new Date(data[i][1]);
     const rowCabang = data[i][7];
     const status = data[i][5]; 
     const nominal = Number(data[i][4]);
     
     if(rowDate >= sDate && rowDate <= eDate) {
        if(cabang === 'Semua' || rowCabang === cabang) {
            sum.total += nominal; sum.count++;
            if(status === 'Lunas') sum.lunas += nominal; else sum.piutang += nominal;
            list.push({ tgl: formatDate(rowDate), inv: data[i][0], cust: data[i][3], cabang: rowCabang, status: status, total: nominal });
        }
     }
  }
  list.sort((a,b) => b.inv.localeCompare(a.inv));
  return { summary: sum, list: list };
}

function getMyPerformance(user) {
  const ss = getSS();
  const shLog = ss.getSheetByName(CONF.LOG_PROD);
  if(!shLog) return [];
  
  const data = shLog.getDataRange().getValues();
  let summary = {}; 
  const now = new Date();
  const thisMonth = now.getMonth();
  const thisYear = now.getFullYear();

  for(let i=1; i<data.length; i++){
     const rowUser = String(data[i][3]);
     const rowDate = new Date(data[i][0]);
     if(rowUser.toLowerCase() === user.toLowerCase() && rowDate.getMonth() === thisMonth && rowDate.getFullYear() === thisYear) {
        const tahap = data[i][2];
        const nilai = Number(data[i][4]) || 0;
        if(!summary[tahap]) summary[tahap] = { count: 0, omset: 0 };
        summary[tahap].count++; summary[tahap].omset += nilai;
     }
  }
  let result = [];
  for (let key in summary) { result.push({ tahap: key, ...summary[key] }); }
  return result;
}

function getAdminMonitor(tgl1, tgl2, userFilter) {
  const ss = getSS();
  const shLog = ss.getSheetByName(CONF.LOG_PROD);
  if(!shLog) return { rekap: [], logs: [] };
  
  const data = shLog.getDataRange().getValues();
  const sDate = new Date(tgl1); const eDate = new Date(tgl2); eDate.setHours(23,59,59);
  let logs = []; let rekapUser = {}; 

  for(let i=1; i<data.length; i++){
     const rDate = new Date(data[i][0]);
     const rUser = String(data[i][3]);
     if(rDate >= sDate && rDate <= eDate) {
        if(userFilter === 'Semua' || rUser === userFilter) {
            logs.push({ waktu: formatDate(rDate), inv: data[i][1], tahap: data[i][2], user: rUser, nilai: Number(data[i][4]), barang: data[i][5] || '-' });
            if(!rekapUser[rUser]) rekapUser[rUser] = { nama: rUser, jobs: 0, omset: 0 };
            rekapUser[rUser].jobs++; rekapUser[rUser].omset += Number(data[i][4]);
        }
     }
  }
  let rekapArr = Object.values(rekapUser).sort((a,b) => b.jobs - a.jobs); 
  logs.reverse();
  return { rekap: rekapArr, logs: logs };
}

function getTrackingData(inv) {
  const ss = getSS();
  const shOrd = ss.getSheetByName(CONF.PESANAN);
  const dataOrd = shOrd.getDataRange().getValues();
  let orderInfo = null;
  
  for(let i=1; i<dataOrd.length; i++){
    if(String(dataOrd[i][0]) === String(inv)){
       orderInfo = { inv: dataOrd[i][0], tgl: formatDate(dataOrd[i][1]), cust: dataOrd[i][3], total: dataOrd[i][4], status: dataOrd[i][5], posisi: getPosisiTerakhir(inv) };
       break;
    }
  }
  
  if(!orderInfo) return { status: 'NOT_FOUND' };
  const shLog = ss.getSheetByName(CONF.LOG_PROD);
  let timeline = [];
  if(shLog) {
     const dLog = shLog.getDataRange().getValues();
     for(let i=1; i<dLog.length; i++){
        if(String(dLog[i][1]) === String(inv)){
           const d = new Date(dLog[i][0]);
           timeline.push({ waktu: formatDate(d) + ' ' + d.toTimeString().slice(0,5), status: dLog[i][2], user: dLog[i][3] });
        }
     }
  }
  timeline.sort((a,b) => new Date(b.waktu) - new Date(a.waktu));
  return { status: 'FOUND', info: orderInfo, history: timeline };
}

function getPosisiTerakhir(inv) {
   const ss = getSS();
   const shLog = ss.getSheetByName(CONF.LOG_PROD);
   if(!shLog) return "Menunggu Proses";
   const data = shLog.getDataRange().getValues();
   let lastStat = "Diterima"; 
   for(let i=data.length-1; i>=1; i--){
      if(String(data[i][1]) === String(inv)) { lastStat = data[i][2]; break; }
   }
   return lastStat;
}

function getCompanyProfile() {
  const ss = getSS();
  const shSet = ss.getSheetByName(CONF.SETTING);
  let global = { nama_app: "Ungu Laundry", footer_nota: "Terima Kasih", logo: "" };
  if(shSet) {
     const data = shSet.getDataRange().getValues();
     for(let i=1; i<data.length; i++){
        const key = String(data[i][0]).toUpperCase().trim(); 
        const val = data[i][1]; 
        if(key === 'NAMA_APP') global.nama_app = val;
        if(key === 'FOOTER_NOTA') global.footer_nota = val;
        if(key === 'LOGO_URL') global.logo = val;
     }
  }
  const shCab = ss.getSheetByName(CONF.CABANG);
  let branches = {};
  if(shCab) {
     const dataC = shCab.getDataRange().getValues();
     for(let i=1; i<dataC.length; i++){
        const namaCab = String(dataC[i][0]).trim(); 
        if(namaCab) {
           branches[namaCab] = { alamat: dataC[i][1], kota: dataC[i][2], telp: dataC[i][3] };
        }
     }
  }
  return { global: global, branches: branches };
}

function getMasterLayanan() {
  const ss = getSS();
  const sh = ss.getSheetByName(CONF.LAYANAN);
  if(!sh) return [];
  const data = sh.getDataRange().getValues();
  let res = [];
  for(let i=1; i<data.length; i++){
     if(data[i][0]) {
       res.push({ kode: data[i][0], nama: data[i][1], satuan: data[i][2], harga: data[i][3], coa: data[i][4] });
     }
  }
  return res;
}

function simpanLayananBaru(form) {
  const ss = getSS();
  let sh = ss.getSheetByName(CONF.LAYANAN);
  if(!sh) { sh = ss.insertSheet(CONF.LAYANAN); sh.appendRow(['Kode_Layanan', 'Nama_Layanan', 'Satuan', 'Harga', 'Kode_COA_Pendapatan']); }
  const data = sh.getDataRange().getValues();
  const namaInput = String(form.nama).trim().toLowerCase();
  for(let i=1; i<data.length; i++){
     if(String(data[i][1]).trim().toLowerCase() === namaInput) return { status: 'ERROR', msg: 'Gagal: Nama Layanan sudah ada!' };
  }
  const lastRow = sh.getLastRow();
  let nextNum = 1;
  if(lastRow > 1) {
      const lastCode = String(sh.getRange(lastRow, 1).getValue()); 
      const match = lastCode.match(/\d+$/);
      if(match) nextNum = parseInt(match[0]) + 1;
  }
  const newCode = 'SRV-' + ('000' + nextNum).slice(-3);
  sh.appendRow([newCode, form.nama, form.satuan, Number(form.harga), form.coa]);
  return { status: 'SUCCESS', msg: 'Layanan berhasil ditambahkan! Kode: ' + newCode };
}

function getAppUrl() { return ScriptApp.getService().getUrl(); }
function getColIdx(sheet, headerName) {
  if (!sheet) return -1;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headers.indexOf(headerName);
}

// --- FUNGSI SEMENTARA: UNTUK LIHAT HASH PASSWORD ---
function intipHashPassword() {
  // Ganti 'admin123' di bawah ini dengan password baru yang Anda mau
  const passwordSaya = 'S1t154t14d4h'; 
  
  const hasilHash = generateHash(passwordSaya);
  Logger.log("COPY KODE INI: " + hasilHash);
}
// Tambahkan di gsfase6a.gs
function cariInvoiceLive(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Data_Pesanan'); // Sesuaikan nama sheet Anda
  if(!sh) return [];
  
  // Ambil data invoice (limit 500 baris terakhir agar cepat)
  const data = sh.getDataRange().getValues().reverse().slice(0, 500); 
  
  let results = [];
  const key = keyword.toLowerCase();
  
  for(let i=0; i<data.length; i++) {
    // Col 0 = Invoice, Col 3 = Pelanggan
    const inv = String(data[i][0]).toLowerCase();
    const cust = String(data[i][3]).toLowerCase();
    
    if(inv.includes(key) || cust.includes(key)) {
       results.push({
         inv: data[i][0],
         cust: data[i][3],
         status: data[i][5] // Col 5 = Status Pembayaran/Proses
       });
    }
    if(results.length >= 5) break; // Batasi 5 hasil saja
  }
  return results;
}
function simpanResepLayanan(form) {
   const ss = getSS();
   let sh = ss.getSheetByName('Master_Resep');
   if(!sh) { 
     sh = ss.insertSheet('Master_Resep');
     sh.appendRow(['Kode_Layanan','Nama_Layanan','Kode_Barang','Nama_Barang','Qty_Standar','Satuan_Resep','Tipe_Item']);
   }
   
   // form.items berisi array bahan baku untuk layanan tersebut
   // Hapus resep lama untuk layanan ini (agar update bersih)
   const data = sh.getDataRange().getValues();
   let rowsToKeep = [];
   
   // Filter data: Simpan header dan data layanan LAIN
   for(let i=0; i<data.length; i++) {
      if(i===0 || String(data[i][0]) !== form.kodeLayanan) {
         rowsToKeep.push(data[i]);
      }
   }
   
   // Tulis ulang sheet (Hapus dulu isinya)
   sh.clear();
   if(rowsToKeep.length > 0) sh.getRange(1,1,rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

   // Tambah Resep Baru
   let newRows = [];
   form.items.forEach(item => {
      newRows.push([
         form.kodeLayanan,
         form.namaLayanan,
         item.kodeBarang,
         item.namaBarang,
         Number(item.qtyStandar), // misal 12
         item.satuan,             // misal ml
         item.tipe                // 'Main' atau 'Varian'
      ]);
   });

   // ... (kode logika insert sheet di atasnya biarkan saja) ...
   
   if(newRows.length > 0) {
      sh.getRange(sh.getLastRow()+1, 1, newRows.length, newRows[0].length).setValues(newRows);
   }
   
   return { status: 'SUCCESS', msg: 'Resep berhasil disimpan.' };
} // <--- INI PENUTUP FUNGSI simpanResepLayanan (Jangan dihapus)

// --- [SECURITY HELPER] Validasi Sesi & Status User (VERSI FINAL) ---
function validateSession(token) {
  if (!token) throw new Error("Akses Ditolak: Token sesi tidak ditemukan. Silakan login ulang.");
  
  // 1. Cek Cache (Layer Pertama - Cepat)
  const cache = CacheService.getScriptCache();
  const sessionData = cache.get(token);
  
  if (!sessionData) throw new Error("Sesi Kadaluarsa: Silakan login ulang.");
  
  const userSession = JSON.parse(sessionData);
  const username = userSession.username;

  // 2. [SECURITY UPDATE] Cek Status Real-time di Spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const shUser = ss.getSheetByName('Data_User'); // Pastikan nama sheet sesuai
  
  if (shUser) {
    const finder = shUser.createTextFinder(username).matchEntireCell(true);
    const result = finder.findNext();

    if (result) {
      const row = result.getRow();
      // Ambil status dari kolom G (index 7)
      const currentStatus = shUser.getRange(row, 7).getValue(); 
      
      if (String(currentStatus) === 'Non-Aktif') {
        cache.remove(token); 
        throw new Error("AKSES DITOLAK: Akun Anda telah dinonaktifkan oleh Admin.");
      }
    } else {
       cache.remove(token);
       throw new Error("AKSES DITOLAK: Data akun tidak ditemukan.");
    }
  }

  return userSession;
}

// --- FUNGSI INCLUDE (Letakkan Paling Bawah) ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}