// Global configuration for sheet names and headers
const CONF = {
  USER: 'Data_User',
  COA: 'COA_Master',
  JU: 'Jurnal_Transaksi_Umum',
  ADJ: 'Jurnal_Penyesuaian',
  SALDO: 'Saldo_Awal',
  LOG: 'Log_Tutup_Buku',
  LOG_PROD: 'Log_Produksi',
  MAP_TRX: 'Mapping_Jurnal_Detail',
  MAP_ADJ: 'Mapping_Jurnal_Penyesuaian',
  SETTING: 'Pengaturan_Global',
  CABANG: 'Data_Cabang',
  BARANG: 'Master_Barang',
  BELI: 'Data_Pembelian',
  STOK: 'Log_Stok',
  LAYANAN: 'Data_Layanan',
  PELANGGAN: 'Data_Pelanggan',
  PESANAN: 'Data_Pesanan',
  DETAIL_PESANAN: 'Detail_Pesanan',
  SHIFT: 'Log_Shift',
  DROP: 'Log_Cash_Drop',
  OPN_HEAD: 'Log_Opname_Head',
  OPN_DET: 'Log_Opname_Detail',
  LOG_DEPOSIT: 'Log_Deposit',
  LOG_POIN: 'Log_Poin',
  MEMBERSHIP: 'Master_Membership'
};

const HEADERS_V2 = {
  PELANGGAN_ADDONS: ['Total_Spending', 'Saldo_Deposit', 'Poin_Reward', 'Status_Member', 'Hutang_Aktif'],
  LOG_DEPOSIT: ['ID_Transaksi', 'Waktu', 'ID_Pelanggan', 'Nama_Pelanggan', 'Jenis_Mutasi', 'Nominal', 'Saldo_Awal', 'Saldo_Akhir', 'Petugas', 'Keterangan'],
  LOG_POIN: ['ID_Transaksi', 'Waktu', 'ID_Pelanggan', 'Poin_Masuk', 'Poin_Keluar', 'Saldo_Poin', 'Keterangan'],
  MEMBERSHIP: ['Level', 'Min_Spending', 'Diskon_Persen', 'Poin_Multiplier']
};
