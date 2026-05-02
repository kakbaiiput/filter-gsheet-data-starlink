// ========================================
// 🚀 STARLINK MANAGEMENT SYSTEM - CLEANED VERSION
// ========================================

// ========================================
// 🔘 MENU & TOMBOL FILTER (onOpen)
// ========================================

/**
 * Tambah menu custom "⚙️ Filter" ke toolbar spreadsheet saat dibuka
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Filter')
    .addItem('🔄 Reset & Aktifkan Filter — Semua Sheet', 'resetAndActivateAllFilters')
    .addItem('🔄 Reset & Aktifkan Filter — Client Aktif', 'resetAndActivateFilterClientAktif')
    .addSeparator()
    .addItem('🔶 Filter (SEGERA)', 'filterSegera')
    .addItem('🔵 Filter (PROSES)', 'filterProses')
    .addItem('🔶 Filter (SEGERA only)', 'filterSegeraOnly')
    .addItem('🔴 Filter (JATUH TEMPO)', 'filterJatuhTempo')
    .addSeparator()
    .addItem('🟢 Filter (READY)', 'filterReady')
    .addItem('✅ Filter (CONFIRMED)', 'filterConfirmed')
    .addItem('💰 Filter (PAID)', 'filterPaid')
    .addItem('📅 Filter (PLANNED)', 'filterPlanned')
    .addSeparator()
    .addItem('🔍 Filter (ID Transaksi)', 'filterIdTransaksi')
    .addSeparator()
    .addItem('🔄 Refresh WA Link', 'refreshWALink')
    .addToUi();
}

/**
 * Reset (clear) lalu aktifkan filter semua kolom pada semua sheet utama client
 */
function resetAndActivateAllFilters() {
  var sheetNames = ['Client Aktif', 'Client Non Aktif', 'Client Lepas', 'Client Tertagih', 'Akun Kosong'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activated = [];

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.getRange(1, 1, sheet.getLastRow(), COLUMNS.TOMBOL_WHATSAPP + 1).createFilter();
    if (name === 'Client Aktif') _sortByJatuhTempo(sheet);
    _applyZebraBanding(sheet);
    activated.push(name);
  });

  SpreadsheetApp.getUi().alert('🔄 Filter direset & diaktifkan pada:\n' + activated.join('\n'));
}

/**
 * Reset (clear) lalu aktifkan filter semua kolom khusus sheet Client Aktif
 */
function resetAndActivateFilterClientAktif() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Client Aktif');
  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet "Client Aktif" tidak ditemukan.'); return; }
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, sheet.getLastRow(), COLUMNS.TOMBOL_WHATSAPP + 1).createFilter();
  _sortByJatuhTempo(sheet);
  _applyZebraBanding(sheet);
  SpreadsheetApp.getUi().alert('🔄 Filter direset & diaktifkan pada sheet Client Aktif.');
}

// ========================================
// 🔍 COLUMN FILTER FUNCTIONS — Client Aktif
// ========================================

/**
 * Helper: sort A–Z berdasarkan Tanggal Jatuh Tempo (kolom K) pada data rows
 */
function _sortByJatuhTempo(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 1, lastRow - 1, COLUMNS.TOMBOL_WHATSAPP + 1)
    .sort({ column: COLUMNS.JATUH_TEMPO + 1, ascending: true }); // K = col 11
}

/**
 * Helper: format kolom S, U, V, W, X → font size 12, bold, warna dark blue 3 (#1f497d)
 */
function _formatColumnsStyle(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var darkBlue3 = '#1f497d';
  var cols = [
    COLUMNS.TIPE_PELANGGAN + 1,  // S = 19
    COLUMNS.LABEL_JATUH_TEMPO + 1, // U = 21
    COLUMNS.LABEL_PROSES + 1,    // V = 22
    COLUMNS.LABEL_SEGERA + 1,    // W = 23
    COLUMNS.LABEL_OBSERVASI + 1  // X = 24
  ];
  cols.forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1)
      .setFontSize(12)
      .setFontWeight('bold')
      .setFontColor(darkBlue3);
  });
  Logger.log('✅ Format kolom S,U,V,W,X: size 12, bold, dark blue 3 diterapkan.');
}

/**
 * Helper: terapkan zebra banding putih & Light Blue 3 (#cfe2f3) pada sheet
 */
function _applyZebraBanding(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = COLUMNS.TOMBOL_WHATSAPP + 1;
  if (lastRow < 2) return;

  // Hapus banding lama jika ada
  sheet.getBandings().forEach(function(b) { b.remove(); });

  // Terapkan banding baru: header dark blue 3, baris ganjil putih, baris genap light blue 3
  sheet.getRange(1, 1, lastRow, lastCol)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor('#1f497d')
    .setFirstRowColor('#ffffff')
    .setSecondRowColor('#cfe2f3');

  // Set eksplisit background header agar pasti teraplikasi
  sheet.getRange(1, 1, 1, lastCol).setBackground('#1f497d');

  Logger.log('✅ Zebra banding diterapkan pada sheet: ' + sheet.getName());
}

/**
 * Helper: siapkan sheet Client Aktif (reset filter + buat baru), return {sheet, filter}
 */
function _prepareClientAktifFilter() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Client Aktif');
  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet "Client Aktif" tidak ditemukan.'); return null; }
  if (sheet.getFilter()) sheet.getFilter().remove();
  var filter = sheet.getRange(1, 1, sheet.getLastRow(), COLUMNS.TOMBOL_WHATSAPP + 1).createFilter();
  return { sheet: sheet, filter: filter };
}

/**
 * 🔶 Filter SEGERA — tampilkan hanya baris dengan nilai "SEGERA" di kolom W + sort Jatuh Tempo
 */
function filterSegera() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('SEGERA').build();
  result.filter.setColumnFilterCriteria(COLUMNS.LABEL_SEGERA + 1, criteria); // W = col 23
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('🔶 Filter SEGERA aktif di kolom W + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 🔵 Filter PROSES — tampilkan hanya baris dengan nilai "PROSES" di kolom V + sort Jatuh Tempo
 */
function filterProses() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('PROSES').build();
  result.filter.setColumnFilterCriteria(COLUMNS.LABEL_PROSES + 1, criteria); // V = col 22
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('🔵 Filter PROSES aktif di kolom V + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 🔶 Filter SEGERA only — SEGERA di kolom W + sembunyikan PROSES di kolom V + sort Jatuh Tempo
 */
function filterSegeraOnly() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  // Tampilkan SEGERA di kolom W
  var criteriaW = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('SEGERA').build();
  result.filter.setColumnFilterCriteria(COLUMNS.LABEL_SEGERA + 1, criteriaW); // W = col 23
  // Sembunyikan PROSES di kolom V (uncheck PROSES)
  var criteriaV = SpreadsheetApp.newFilterCriteria().setHiddenValues(['PROSES']).build();
  result.filter.setColumnFilterCriteria(COLUMNS.LABEL_PROSES + 1, criteriaV); // V = col 22
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('🔶 Filter SEGERA only aktif (W=SEGERA, V tanpa PROSES) + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 🔴 Filter JATUH TEMPO — tampilkan hanya baris dengan nilai "JATUH TEMPO" di kolom U + sort Jatuh Tempo
 */
function filterJatuhTempo() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('JATUH TEMPO').build();
  result.filter.setColumnFilterCriteria(COLUMNS.LABEL_JATUH_TEMPO + 1, criteria); // U = col 21
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('🔴 Filter JATUH TEMPO aktif di kolom U + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 🟢 Filter READY — tampilkan hanya baris dengan nilai "READY" di kolom O + sort Jatuh Tempo
 */
function filterReady() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('READY').build();
  result.filter.setColumnFilterCriteria(COLUMNS.STATUS + 1, criteria); // O = col 15
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('🟢 Filter READY aktif di kolom O + diurutkan A–Z Jatuh Tempo.');
}

/**
 * ✅ Filter CONFIRMED — tampilkan hanya baris dengan nilai "CONFIRMED" di kolom O + sort Jatuh Tempo
 */
function filterConfirmed() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('CONFIRMED').build();
  result.filter.setColumnFilterCriteria(COLUMNS.STATUS + 1, criteria); // O = col 15
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('✅ Filter CONFIRMED aktif di kolom O + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 💰 Filter PAID — tampilkan hanya baris dengan nilai "PAID" di kolom O + sort Jatuh Tempo
 */
function filterPaid() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('PAID').build();
  result.filter.setColumnFilterCriteria(COLUMNS.STATUS + 1, criteria); // O = col 15
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('💰 Filter PAID aktif di kolom O + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 📅 Filter PLANNED — tampilkan hanya baris dengan nilai "PLANNED" di kolom O + sort Jatuh Tempo
 */
function filterPlanned() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('PLANNED').build();
  result.filter.setColumnFilterCriteria(COLUMNS.STATUS + 1, criteria); // O = col 15
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('📅 Filter PLANNED aktif di kolom O + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 🔍 Filter ID Transaksi — sembunyikan blank, --- 375.000 ---, --- 750.000 ---, HIGH DEMAND di kolom N
 */
function filterIdTransaksi() {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(['', '--- 375.000 ---', '--- 750.000 ---', 'HIGH DEMAND'])
    .build();
  result.filter.setColumnFilterCriteria(COLUMNS.TRANSACTION_ID + 1, criteria); // N = col 14
  _sortByJatuhTempo(result.sheet);
  SpreadsheetApp.getUi().alert('🔍 Filter ID Transaksi aktif di kolom N + diurutkan A–Z Jatuh Tempo.');
}

/**
 * 🔄 Refresh WA Link — sama seperti kirimWhatsApp tapi tanpa kirim email
 */
function refreshWALink() {
  kirimWhatsApp(true);
}

