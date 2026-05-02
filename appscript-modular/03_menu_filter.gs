[Resource from github at repo://kakbaiiput/filter-gsheet-data-starlink/sha/a8fb8015b1f821c873fe66a68cbd9e09eb249e71/contents/appscript-02052026] // ========================================
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
    .addItem('🔍 Filter Custom (Multiple)', 'showFilterCustomDialog')
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

  SpreadsheetApp.getActiveSpreadsheet().toast('🔄 Filter direset & diaktifkan pada: ' + activated.join(', '), 'Filter Reset', 4);
}

/**
 * Reset (clear) lalu aktifkan filter semua kolom khusus sheet Client Aktif
 */
function resetAndActivateFilterClientAktif() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Client Aktif');
  if (!sheet) { SpreadsheetApp.getActiveSpreadsheet().toast('Sheet "Client Aktif" tidak ditemukan.', 'Error', 5); return; }
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, sheet.getLastRow(), COLUMNS.TOMBOL_WHATSAPP + 1).createFilter();
  _sortByJatuhTempo(sheet);
  _applyZebraBanding(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast('🔄 Filter direset & diaktifkan pada sheet Client Aktif.', 'Filter Reset', 4);
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
  if (!sheet) { SpreadsheetApp.getActiveSpreadsheet().toast('Sheet "Client Aktif" tidak ditemukan.', 'Error', 5); return null; }
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
  SpreadsheetApp.getActiveSpreadsheet().toast('🔶 Filter SEGERA aktif di kolom W + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('🔵 Filter PROSES aktif di kolom V + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('🔶 Filter SEGERA only aktif (W=SEGERA, V tanpa PROSES) + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('🔴 Filter JATUH TEMPO aktif di kolom U + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('🟢 Filter READY aktif di kolom O + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Filter CONFIRMED aktif di kolom O + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('💰 Filter PAID aktif di kolom O + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Filter PLANNED aktif di kolom O + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
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
  SpreadsheetApp.getActiveSpreadsheet().toast('🔍 Filter ID Transaksi aktif di kolom N + diurutkan A–Z Jatuh Tempo.', 'Filter', 4);
}

/**
 * 🔍 Filter Custom (Multiple) — dialog pilih kolom + input nilai (1 per baris)
 */
function showFilterCustomDialog() {
  var cols = [
    [0,  'A', 'Nama'],
    [1,  'B', 'Login Gmail'],
    [2,  'C', 'Login Starlink'],
    [3,  'D', 'Login Alternatif'],
    [4,  'E', 'ACC No'],
    [5,  'F', 'Email Client'],
    [6,  'G', 'Nomor CS'],
    [7,  'H', 'Alamat'],
    [8,  'I', 'KIT Number'],
    [9,  'J', 'Serial Number'],
    [10, 'K', 'Jatuh Tempo'],
    [11, 'L', 'Kode'],
    [12, 'M', 'Payment'],
    [13, 'N', 'Transaction ID'],
    [14, 'O', 'Status'],
    [15, 'P', 'Paket'],
    [16, 'Q', 'Last 4 Digit'],
    [17, 'R', 'No Register'],
    [18, 'S', 'Tipe Pelanggan'],
    [19, 'T', '(Cadangan)'],
    [20, 'U', 'Label Jatuh Tempo'],
    [21, 'V', 'Label Proses'],
    [22, 'W', 'Label Segera'],
    [23, 'X', 'Label Observasi'],
    [24, 'Y', 'Tombol WhatsApp']
  ];

  var options = cols.map(function(c) {
    return '<option value="' + c[0] + '">' + c[1] + ' — ' + c[2] + '</option>';
  }).join('');

  var html = '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:16px;margin:0}' +
    'label{font-weight:bold;display:block;margin-top:14px;margin-bottom:4px}' +
    'select,textarea{width:100%;box-sizing:border-box;padding:6px;border:1px solid #ccc;border-radius:4px;font-size:13px}' +
    'textarea{height:140px;resize:vertical;font-family:monospace}' +
    '.hint{font-size:11px;color:#888;margin-top:3px}' +
    '.buttons{margin-top:18px;text-align:right}' +
    'button{padding:7px 18px;margin-left:8px;cursor:pointer;border-radius:4px;font-size:13px}' +
    '#btnOk{background:#1f497d;color:white;border:none}' +
    '#btnOk:disabled{background:#888;cursor:default}' +
    '#btnCancel{background:#f1f1f1;border:1px solid #ccc}' +
    '</style></head><body>' +
    '<label>Pilih Kolom:</label>' +
    '<select id="colSelect">' + options + '</select>' +
    '<label>Nilai yang dicari:</label>' +
    '<textarea id="values" placeholder="Ketik 1 nilai per baris, contoh:&#10;John Doe&#10;Jane Smith&#10;Ahmad"></textarea>' +
    '<div class="hint">Tidak case-sensitive. Nilai harus sesuai isi sel.</div>' +
    '<div class="buttons">' +
    '<button id="btnCancel" onclick="google.script.host.close()">Batal</button>' +
    '<button id="btnOk" onclick="applyFilter()">OK</button>' +
    '</div>' +
    '<script>' +
    'function applyFilter(){' +
    '  var colIndex=parseInt(document.getElementById("colSelect").value);' +
    '  var valuesStr=document.getElementById("values").value;' +
    '  if(!valuesStr.trim()){alert("Masukkan minimal 1 nilai.");return;}' +
    '  var btn=document.getElementById("btnOk");' +
    '  btn.disabled=true;btn.textContent="Memproses...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){google.script.host.close();})' +
    '    .withFailureHandler(function(e){alert("Error: "+e.message);btn.disabled=false;btn.textContent="OK";})' +
    '    .applyCustomFilter(colIndex,valuesStr);' +
    '}' +
    '<\/script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(390),
    '🔍 Filter Custom'
  );
}

function applyCustomFilter(colIndex, valuesStr) {
  var result = _prepareClientAktifFilter();
  if (!result) return;
  var sheet = result.sheet;
  var filter = result.filter;

  var inputValues = valuesStr.split('\n')
    .map(function(v) { return v.trim(); })
    .filter(function(v) { return v !== ''; });
  if (inputValues.length === 0) return;
  var inputLower = inputValues.map(function(v) { return v.toLowerCase(); });

  var lastRow = sheet.getLastRow();
  var colData = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  var allValues = {};
  colData.forEach(function(row) {
    var v = (row[0] !== null && row[0] !== undefined) ? row[0].toString().trim() : '';
    allValues[v] = true;
  });

  var hiddenValues = Object.keys(allValues).filter(function(v) {
    var vLower = v.toLowerCase();
    return !inputLower.some(function(input) {
      return vLower.indexOf(input) !== -1;
    });
  });

  filter.setColumnFilterCriteria(colIndex + 1,
    SpreadsheetApp.newFilterCriteria().setHiddenValues(hiddenValues).build()
  );
  _sortByJatuhTempo(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '🔍 Filter custom: ' + inputValues.length + ' nilai di kolom ' + String.fromCharCode(65 + colIndex),
    'Filter Custom', 4
  );
}

/**
 * 🔄 Refresh WA Link — sama seperti kirimWhatsApp tapi tanpa kirim email
 */
function refreshWALink() {
  kirimWhatsApp(true);
}

