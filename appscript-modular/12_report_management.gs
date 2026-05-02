// ========================================
// 📊 9. REPORT MANAGEMENT FUNCTIONS
// ========================================

/**
 * 🔧 Get Report Spreadsheet
 */
function getReportSpreadsheet() {
  try {
    var reportSpreadsheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID);
    Logger.log('✅ Report spreadsheet connected: ' + reportSpreadsheet.getName());
    return reportSpreadsheet;
  } catch (error) {
    Logger.log('❌ Error connecting to report spreadsheet: ' + error.message);
    throw new Error('Tidak dapat mengakses file laporan terpisah. ID: ' + REPORT_SPREADSHEET_ID);
  }
}

/**
 * 🔧 Get Report Sheet (FIXED: Better formatting for new sheets)
 */
function getReportSheet(month, year, createIfNotExists = false) {
  try {
    var reportSpreadsheet = getReportSpreadsheet();
    var sheetName = month.toString().padStart(2, '0') + '_' + year.toString();
    var sheet = reportSpreadsheet.getSheetByName(sheetName);
    
    if (!sheet && createIfNotExists) {
      Logger.log('Creating new report sheet with CORRECT formatting: ' + sheetName);
      
      sheet = reportSpreadsheet.insertSheet(sheetName);
      
      // Headers yang benar sesuai format konsisten
      var headers = [
        'Nomor',              // A
        'Tanggal Transaksi',  // B
        'ID Transaksi',       // C
        'Nama Client',        // D
        'KIT Number',         // E
        'Serial Number',      // F
        'Paket',              // G
        'Tipe Pembayaran',    // H
        'Periode',            // I
        'Nominal',            // J
        'Jumlah Total',       // K
        'Total Harian'        // L
      ];
      
      // Set headers
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // FIXED: Apply formatting khusus untuk sheet baru
      formatNewReportSheet(sheet);
      
      Logger.log('✅ New report sheet created with proper formatting: ' + sheetName);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error getting report sheet: ' + error.message);
    throw error;
  }
}

/**
 * 🎨 Format New Report Sheet (Khusus untuk sheet baru yang kosong)
 */
function formatNewReportSheet(sheet) {
  try {
    Logger.log('🎨 Applying formatting to new report sheet...');
    
    var headers = 12; // A-L columns
    
    // 1. HEADER FORMATTING (Row 1)
    var headerRange = sheet.getRange(1, 1, 1, headers);
    headerRange.setFontWeight('bold')
              .setBackground('#1e3a8a')        // Dark blue
              .setFontColor('white')
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle')
              .setFontSize(11)
              .setFontFamily('Roboto');
    
    // 2. COLUMN WIDTHS - Set specific widths for each column
    var columnWidths = {
      1: 60,   // A: Nomor
      2: 120,  // B: Tanggal  
      3: 160,  // C: Transaction ID
      4: 140,  // D: Nama Client
      5: 140,  // E: KIT Number
      6: 140,  // F: Serial Number
      7: 100,  // G: Paket
      8: 120,  // H: Tipe Pembayaran
      9: 80,   // I: Periode
      10: 110, // J: Nominal
      11: 130, // K: Jumlah Total
      12: 150  // L: Total Harian
    };
    
    for (var col = 1; col <= headers; col++) {
      sheet.setColumnWidth(col, columnWidths[col]);
    }
    
    // 3. HEADER BORDERS
    headerRange.setBorder(
      true, true, true, true, true, true,
      '#1e3a8a',
      SpreadsheetApp.BorderStyle.SOLID
    );
    
    // 4. SET ROW HEIGHT for header
    sheet.setRowHeight(1, 35);
    
    // 5. FREEZE HEADER ROW
    sheet.setFrozenRows(1);
    
    Logger.log('✅ New report sheet formatting completed successfully');
    
  } catch (error) {
    Logger.log('❌ Error formatting new report sheet: ' + error.message);
  }
}

/**
 * ✅ NEW: Remove existing filters from report sheet
 */
function removeExistingFiltersFromReportSheet(sheet) {
  try {
    Logger.log('🔄 Removing existing filters from report sheet...');
    
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
      Logger.log('✅ Existing filter removed from report sheet');
    } else {
      Logger.log('ℹ️ No existing filter found on report sheet');
    }
    
    // Clear any existing filter criteria
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow > 0 && lastCol > 0) {
      // Reset any hidden rows
      sheet.showRows(1, lastRow);
      Logger.log('✅ All rows shown in report sheet');
    }
    
  } catch (error) {
    Logger.log('⚠️ Error removing filters: ' + error.message);
  }
}

function applyDateFilterToReportSheet(sheet, filterDate) {
  try {
    // Safeguard: Pastikan hanya apply filter pada sheet yang benar
    var expectedSheetName = (filterDate.getMonth() + 1).toString().padStart(2, '0') + '_' + filterDate.getFullYear();
    var actualSheetName = sheet.getName();
    
    Logger.log('🔄 Applying filter to sheet: ' + actualSheetName);
    Logger.log('📅 Filter date: ' + filterDate.toDateString());
    Logger.log('📅 Expected sheet for filter date: ' + expectedSheetName);
    
    // Check if this is the correct sheet for the filter date
    if (actualSheetName !== expectedSheetName) {
      Logger.log('⚠️ SKIP: Filter not applied - sheet is not for the filter date month/year');
      Logger.log('   Current sheet: ' + actualSheetName);
      Logger.log('   Filter date requires: ' + expectedSheetName);
      return; // Exit early - don't apply filter to wrong month/year sheets
    }
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      Logger.log('ℹ️ No data to filter in sheet: ' + actualSheetName);
      return;
    }
    
    // ✅ MAJOR CHANGE: Use filterDate instead of today
    var targetDate = new Date(filterDate);
    targetDate.setHours(0, 0, 0, 0);
    
    var targetDateFormatted = targetDate.getDate().toString().padStart(2, '0') + '/' + 
                             (targetDate.getMonth() + 1).toString().padStart(2, '0') + '/' + 
                             targetDate.getFullYear();
    
    Logger.log('📅 Filtering sheet "' + actualSheetName + '" for date: ' + targetDateFormatted);
    
    // Remove any existing filter first on THIS SHEET ONLY
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
      Logger.log('🗑️ Removed existing filter from sheet: ' + actualSheetName);
    }
    
    // Create filter range for THIS SHEET ONLY
    var filterRange = sheet.getRange(1, 1, lastRow, lastCol);
    var filter = filterRange.createFilter();
    
    // Try Date format first
    try {
      var dateCriteria = SpreadsheetApp.newFilterCriteria()
        .whenDateEqualTo(targetDate)
        .build();
      
      filter.setColumnFilterCriteria(2, dateCriteria);
      Logger.log('✅ Date filter applied successfully to sheet: ' + actualSheetName + ' for date: ' + targetDateFormatted);
      return;
      
    } catch (dateError) {
      Logger.log('⚠️ Date filter failed for sheet: ' + actualSheetName + ', trying text format...');
      
      // Fallback to text format
      try {
        var textCriteria = SpreadsheetApp.newFilterCriteria()
          .whenTextEqualTo(targetDateFormatted)
          .build();
        
        filter.setColumnFilterCriteria(2, textCriteria);
        Logger.log('✅ Text filter applied successfully to sheet: ' + actualSheetName + ' for date: ' + targetDateFormatted);
        return;
        
      } catch (textError) {
        Logger.log('⚠️ Text filter also failed for sheet: ' + actualSheetName + ', keeping basic filter');
      }
    }
    
  } catch (error) {
    Logger.log('❌ Filter error for sheet ' + sheet.getName() + ': ' + error.message);
    
    // Final fallback - just create basic filter on THIS SHEET ONLY
    try {
      if (sheet.getFilter()) {
        sheet.getFilter().remove();
      }
      var filterRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
      filterRange.createFilter();
      Logger.log('✅ Basic filter created as fallback for sheet: ' + sheet.getName());
    } catch (fallbackError) {
      Logger.log('❌ Even basic filter failed for sheet ' + sheet.getName() + ': ' + fallbackError.message);
    }
  }
}

/**
 * Check if requestId has already been submitted
 * @param {string} requestId - Unique request identifier
 * @return {boolean} - true if already submitted, false otherwise
 */
function isRequestIdAlreadySubmitted(requestId) {
  try {
    const cacheKey = REQUEST_ID_CACHE_KEY_PREFIX + requestId;
    const cached = SCRIPT_PROPERTIES.getProperty(cacheKey);
    
    if (cached) {
      const cacheData = JSON.parse(cached);
      const now = new Date().getTime();
      
      // Check if cache is still valid (within 5 minutes)
      if (now - cacheData.timestamp < REQUEST_ID_CACHE_DURATION) {
        Logger.log('🚫 Duplicate requestId detected: ' + requestId);
        Logger.log('Original submission time: ' + new Date(cacheData.timestamp).toISOString());
        return true;
      } else {
        // Cache expired, remove it
        SCRIPT_PROPERTIES.deleteProperty(cacheKey);
      }
    }
    
    return false;
  } catch (error) {
    Logger.log('⚠️ Error checking requestId cache: ' + error.message);
    // On error, allow submission (fail-safe approach)
    return false;
  }
}

/**
 * Mark requestId as submitted
 * @param {string} requestId - Unique request identifier
 */
function markRequestIdAsSubmitted(requestId) {
  try {
    const cacheKey = REQUEST_ID_CACHE_KEY_PREFIX + requestId;
    const cacheData = {
      timestamp: new Date().getTime(),
      requestId: requestId
    };
    
    SCRIPT_PROPERTIES.setProperty(cacheKey, JSON.stringify(cacheData));
    Logger.log('✅ RequestId marked as submitted: ' + requestId);
    
    // Note: Apps Script doesn't support setTimeout, so cleanup happens on next check
    
  } catch (error) {
    Logger.log('⚠️ Error marking requestId: ' + error.message);
  }
}

/**
 * Cleanup expired requestId cache entries (call this periodically)
 * Can be run as a time-based trigger (e.g., every hour)
 */
function cleanupExpiredRequestIds() {
  try {
    const allProperties = SCRIPT_PROPERTIES.getProperties();
    const now = new Date().getTime();
    let cleanedCount = 0;
    
    for (var key in allProperties) {
      if (key.startsWith(REQUEST_ID_CACHE_KEY_PREFIX)) {
        try {
          const cacheData = JSON.parse(allProperties[key]);
          
          // If older than cache duration, delete
          if (now - cacheData.timestamp >= REQUEST_ID_CACHE_DURATION) {
            SCRIPT_PROPERTIES.deleteProperty(key);
            cleanedCount++;
          }
        } catch (parseError) {
          // If can't parse, delete it
          SCRIPT_PROPERTIES.deleteProperty(key);
          cleanedCount++;
        }
      }
    }
    
    Logger.log('🧹 Cleaned up ' + cleanedCount + ' expired requestId cache entries');
    return {
      status: 'success',
      cleanedCount: cleanedCount
    };
    
  } catch (error) {
    Logger.log('⚠️ Error in cleanup: ' + error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

function submitPaymentDataMultiKit(formData) {
  try {
    Logger.log('=== SUBMIT WITH DUPLICATE PREVENTION START ===');
    Logger.log('Form data received: ' + JSON.stringify(formData));
    
    // 🚫 CRITICAL: Check for duplicate requestId FIRST
    if (formData.requestId) {
      Logger.log('🔍 Checking requestId: ' + formData.requestId);
      
      if (isRequestIdAlreadySubmitted(formData.requestId)) {
        Logger.log('🚫 DUPLICATE SUBMISSION BLOCKED: ' + formData.requestId);
        return {
          status: 'error',
          message: '🚫 Duplicate submission detected. This request has already been processed.',
          isDuplicate: true,
          requestId: formData.requestId
        };
      }
      
      // Mark this requestId as being processed
      markRequestIdAsSubmitted(formData.requestId);
    } else {
      Logger.log('⚠️ No requestId provided - duplicate prevention disabled for this request');
    }
    
    if (!formData || !formData.selectedKits || formData.selectedKits.length === 0) {
      Logger.log('Error: No selectedKits found');
      return {
        status: 'error',
        message: 'Tidak ada KIT yang dipilih'
      };
    }
    
    // ✅ IMPORTANT: Extract transaction date from form data
    var tanggal = new Date(formData.tanggal);
    var month = tanggal.getMonth() + 1;
    var year = tanggal.getFullYear();
    
    Logger.log('📅 Transaction date from form: ' + tanggal.toDateString());
    Logger.log('📅 Will filter for: ' + tanggal.getDate() + '/' + month + '/' + year);
    
    try {
      var reportSpreadsheet = getReportSpreadsheet();
      Logger.log('✅ Connection to separate file successful');
    } catch (connectionError) {
      return {
        status: 'error',
        message: 'Gagal terhubung ke file laporan terpisah: ' + connectionError.message
      };
    }
    
    var sheet = getReportSheet(month, year, true);

    // ✅ NEW: Remove existing filters BEFORE processing
    removeExistingFiltersFromReportSheet(sheet);

    var lastRow = sheet.getLastRow();

    var day = tanggal.getDate().toString().padStart(2, '0');
    var monthStr = month.toString().padStart(2, '0');
    var yearStr = year.toString();
    var formattedDate = `${day}/${monthStr}/${yearStr}`;

    // 🆕 NEW LOGIC: Check if should submit as multiple entries (one row per KIT)
    // 🔧 FIX: Changed > 1 to >= 1 to include single KIT with per-KIT payment type
    if (formData.submitAsMultipleEntries && formData.selectedKits.length >= 1) {
      Logger.log('🔄 Processing as MULTIPLE SEPARATE ENTRIES (one row per KIT)');
      Logger.log('📊 Total KITs to insert: ' + formData.selectedKits.length);

      var insertedTransactionIDs = [];
      var insertedRows = [];

      // Pre-generate all transaction IDs: reads sheet ONCE instead of N openById calls
      // (Each generateTransactionID call was opening the report spreadsheet separately)
      var datePrefix = day + '-' + monthStr + '-' + yearStr;
      var nextIncrement = 1;
      try {
        var existingSheetData = sheet.getDataRange().getValues();
        var oldPat = new RegExp('^SL\\/\\d{2}' + monthStr + yearStr + '\\/\\d{3}$');
        var newPat = new RegExp('^SL\\/\\d{2}-' + monthStr + '-' + yearStr + '\\/\\d{3}$');
        var maxInc = 0;
        for (var ei = 1; ei < existingSheetData.length; ei++) {
          var idStr = (existingSheetData[ei][2] || '').toString().trim();
          if (oldPat.test(idStr) || newPat.test(idStr)) {
            var idParts = idStr.split('/');
            var idNum = idParts.length === 3 ? parseInt(idParts[2]) : NaN;
            if (!isNaN(idNum) && idNum > maxInc) maxInc = idNum;
          }
        }
        nextIncrement = maxInc + 1;
        Logger.log('✅ Pre-read sheet for ID generation: next increment = ' + nextIncrement);
      } catch (preReadError) {
        Logger.log('⚠️ Pre-read fallback, will use increment 1: ' + preReadError.message);
      }

      // Set seluruh kolom I (Periode) sebagai plain text SEKALI di sini,
      // sebelum loop, agar Google Sheets tidak auto-konversi "April 2026" ke tanggal.
      var maxRows = sheet.getMaxRows();
      if (maxRows > 1) {
        sheet.getRange(2, 9, maxRows - 1, 1).setNumberFormat('@STRING@');
      }

      // Loop each KIT and insert as separate entry
      for (var kitIndex = 0; kitIndex < formData.selectedKits.length; kitIndex++) {
        var kit = formData.selectedKits[kitIndex];

        // Generate transaction ID from pre-read data — no extra openById needed
        var transactionID = 'SL/' + datePrefix + '/' + (nextIncrement + kitIndex).toString().padStart(3, '0');
        insertedTransactionIDs.push(transactionID);

        Logger.log('Processing KIT ' + (kitIndex + 1) + '/' + formData.selectedKits.length + ': ' + kit.kitNumber);
        Logger.log('  - Transaction ID: ' + transactionID);

        // Get current last row (updates after each insert)
        lastRow = sheet.getLastRow();
        var newRowNumber = lastRow;

        // Create row data for this KIT
        var rowData = [
          newRowNumber,               // A: Nomor
          formattedDate,              // B: Tanggal
          transactionID,              // C: ID Transaksi
          kit.clientName || formData.nama,  // D: Nama Client
          kit.kitNumber,              // E: KIT Number
          kit.serialNumber || '',     // F: Serial Number
          kit.paket,                  // G: Paket
          kit.tipePembayaran || '',   // H: Tipe Pembayaran
          kit.periodeP || '',         // I: Periode
          kit.nominal                 // J: Nominal
        ];

        sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
        insertedRows.push(rowData);

        Logger.log('✅ Inserted row ' + (kitIndex + 1) + ': ' + kit.kitNumber);
      }

      Logger.log('✅ All ' + formData.selectedKits.length + ' KITs inserted as separate entries');
      Logger.log('📋 Transaction IDs: ' + insertedTransactionIDs.join(', '));

      // Sort and organize after all inserts
      sortAndOrganizeWithNew11ColumnStructure(sheet);

      // Apply filter based on transaction date
      applyDateFilterToReportSheet(sheet, tanggal);

      // Update client data with transaction IDs - each KIT gets its own transaction ID
      var allUpdateResults = { updated: [], notFound: [] };
      for (var updateIdx = 0; updateIdx < formData.selectedKits.length; updateIdx++) {
        var updateKit = formData.selectedKits[updateIdx];
        var updateTxID = insertedTransactionIDs[updateIdx];
        var updateResult = updateClientDataWithTransactionID([updateKit], updateTxID);
        allUpdateResults.updated = allUpdateResults.updated.concat(updateResult.updated);
        allUpdateResults.notFound = allUpdateResults.notFound.concat(updateResult.notFound);
      }

      Logger.log('✅ Client data updates completed: ' + allUpdateResults.updated.length + ' updated, ' + allUpdateResults.notFound.length + ' not found');

      // Send email notification
      try {
        sendEmailNotificationMultiKitWithTransactionID(formData, sheet.getName(), insertedTransactionIDs.join(', '), allUpdateResults);
      } catch (emailError) {
        Logger.log('⚠️ Email notification failed but submission successful: ' + emailError.message);
      }

      return {
        status: 'success',
        message: 'Data berhasil disubmit! ' + formData.selectedKits.length + ' KIT diproses sebagai entry terpisah.',
        transactionIDs: insertedTransactionIDs,
        totalEntries: formData.selectedKits.length,
        requestId: formData.requestId,
        sheetName: sheet.getName(),
        clientUpdates: allUpdateResults,
        timestamp: new Date().toISOString()
      };

    } else {
      // Original logic for single entry (backward compatibility)
      Logger.log('📝 Processing as SINGLE ENTRY (original behavior)');

      var transactionID = generateTransactionID(tanggal, month, year);
      Logger.log('Generated Transaction ID: ' + transactionID);

      var newRowNumber = lastRow;

      // ✅ CRITICAL: Prepare serialNumbers untuk multiple kit
      var serialNumbers = '';
      if (formData.isMultipleKit && formData.kitCount > 1) {
        var serialNumArray = formData.selectedKits.map(function(kit) {
          return kit.serialNumber || '';
        });
        serialNumbers = serialNumArray.join('\n');
        Logger.log('Prepared serial numbers for multiple kit: ' + serialNumbers);
      } else if (formData.selectedKits.length === 1) {
        serialNumbers = formData.selectedKits[0].serialNumber || '';
        Logger.log('Single kit serial number: ' + serialNumbers);
      }

      var rowData;
      if (formData.isMultipleKit && formData.kitCount > 1) {
        // ✅ NEW: Struktur 12 kolom dengan Serial Number dan Periode terpisah
        rowData = [
          newRowNumber,         // A: Nomor
          formattedDate,        // B: Tanggal
          transactionID,        // C: ID Transaksi
          formData.nama,        // D: Nama Client
          formData.kitNumbers,  // E: KIT Number
          serialNumbers,        // F: Serial Number
          formData.kitPackages, // G: Paket
          formData.selectedKits[0].tipePembayaran || formData.tipe || '',  // H: Tipe Pembayaran
          formData.selectedKits[0].periodeP || '',  // I: Periode
          formData.nominal      // J: Nominal
          // K: Jumlah Total akan dihitung running total
          // L: Total Harian akan di-merge
        ];
      } else {
        var kit = formData.selectedKits[0];
        // ✅ NEW: Struktur 12 kolom untuk single kit
        rowData = [
          newRowNumber,         // A: Nomor
          formattedDate,        // B: Tanggal
          transactionID,        // C: ID Transaksi
          formData.nama,        // D: Nama Client
          kit.kitNumber,        // E: KIT Number
          kit.serialNumber || '', // F: Serial Number
          kit.paket,           // G: Paket
          kit.tipePembayaran || formData.tipe || '',  // H: Tipe Pembayaran
          kit.periodeP || '',  // I: Periode
          kit.nominal || formData.nominal // J: Nominal
          // K: Jumlah Total akan dihitung running total
          // L: Total Harian akan di-merge
        ];
      }

      // Insert data (format kolom I sebagai plain text agar tidak jadi tanggal)
      var legacyMaxRows = sheet.getMaxRows();
      if (legacyMaxRows > 1) {
        sheet.getRange(2, 9, legacyMaxRows - 1, 1).setNumberFormat('@STRING@');
      }
      sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
      Logger.log('✅ Data inserted with NEW 11-column structure');
      Logger.log('Row data: ' + JSON.stringify(rowData));
    }
    
    // ✅ CRITICAL: Panggil fungsi dengan nama BARU
    sortAndOrganizeWithNew11ColumnStructure(sheet);
    
    // ✅ MAJOR CHANGE: Apply filter based on transaction date, not today
    applyDateFilterToReportSheet(sheet, tanggal);
    
    var updateResult = updateClientDataWithTransactionID(formData.selectedKits, transactionID);
    
    try {
      sendEmailNotificationMultiKitWithTransactionID(formData, sheet.getName(), transactionID, updateResult);
    } catch (emailError) {
      Logger.log('⚠️ Email notification failed but submission successful: ' + emailError.message);
    }
    
    // ✅ SUCCESS: Return dengan requestId untuk tracking
    return {
      status: 'success',
      message: 'Data berhasil disubmit!',
      transactionID: transactionID,
      requestId: formData.requestId,
      sheetName: sheet.getName(),
      clientUpdates: updateResult,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('❌ Error in submitPaymentDataMultiKit: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    
    // On error, try to cleanup the requestId cache so user can retry
    if (formData && formData.requestId) {
      try {
        const cacheKey = REQUEST_ID_CACHE_KEY_PREFIX + formData.requestId;
        SCRIPT_PROPERTIES.deleteProperty(cacheKey);
        Logger.log('🧹 Cleaned up failed requestId from cache: ' + formData.requestId);
      } catch (cleanupError) {
        Logger.log('⚠️ Error cleaning up requestId cache: ' + cleanupError.message);
      }
    }
    
    return {
      status: 'error',
      message: 'Error: ' + error.message,
      requestId: formData?.requestId,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🔧 Check Duplicate dengan kolom mapping baru
 */
function checkDuplicateInReportMultiKit(kitNumbers, month, year) {
  try {
    if (!month || !year) {
      var currentDate = new Date();
      month = currentDate.getMonth() + 1;
      year = currentDate.getFullYear();
    }
    
    var sheet = getReportSheet(month, year, false);
    
    if (!sheet) {
      return {
        status: 'no_duplicate',
        hasDuplicate: false,
        duplicateData: [],
        duplicateKits: []
      };
    }
    
    var data = sheet.getDataRange().getValues();
    var duplicates = [];
    var duplicateKits = [];
    
    // ✅ UNCHANGED: Masih cek duplicate di kolom E (Serial Number KIT)
    for (var i = 1; i < data.length; i++) {
      var reportKitCell = data[i][4]; // Kolom E: Serial Number (KIT) ← masih sama
      
      if (reportKitCell) {
        var reportKit = reportKitCell.toString().trim().toUpperCase();
        var reportKitArray = reportKit.includes('\n') ? reportKit.split('\n') : [reportKit];
        
        for (var j = 0; j < kitNumbers.length; j++) {
          var searchKit = kitNumbers[j].toString().trim().toUpperCase();
          
          for (var k = 0; k < reportKitArray.length; k++) {
            if (reportKitArray[k].trim() === searchKit) {
              
              var tanggalRaw = data[i][1] || '';
              var tanggalFormatted = '';
              
              if (tanggalRaw instanceof Date) {
                var day = tanggalRaw.getDate().toString().padStart(2, '0');
                var month = (tanggalRaw.getMonth() + 1).toString().padStart(2, '0');
                var year = tanggalRaw.getFullYear();
                tanggalFormatted = `${day}/${month}/${year}`;
              } else if (tanggalRaw) {
                var tempDate = new Date(tanggalRaw);
                if (!isNaN(tempDate.getTime())) {
                  var day = tempDate.getDate().toString().padStart(2, '0');
                  var month = (tempDate.getMonth() + 1).toString().padStart(2, '0');
                  var year = tempDate.getFullYear();
                  tanggalFormatted = `${day}/${month}/${year}`;
                } else {
                  tanggalFormatted = tanggalRaw.toString();
                }
              }
              
              // ✅ MINOR UPDATE: Struktur data dengan kolom baru
              duplicates.push({
                rowNumber: i + 1,
                transactionID: data[i][2] || 'No ID',        // C: Transaction ID
                namaClient: data[i][3] || 'No Name',         // D: Nama Client
                tanggal: tanggalFormatted,                   
                kitNumber: data[i][4] || '',                 // E: Serial Number (KIT) ← masih sama
                serialNumber: data[i][5] || '',              // F: Serial Number ← BARU
                paket: data[i][6] || '',                     // G: Paket ← GESER dari F
                tipe: data[i][7] || '',                      // H: Tipe Pembayaran ← GESER dari G
                periode: data[i][8] || '',                   // I: Periode
                nominal: data[i][9] || ''                    // J: Nominal
              });
              duplicateKits.push(kitNumbers[j]);
              break;
            }
          }
        }
      }
    }
    
    return {
      status: duplicates.length > 0 ? 'duplicate_found' : 'no_duplicate',
      hasDuplicate: duplicates.length > 0,
      duplicateData: duplicates,
      duplicateKits: duplicateKits,
      checkedSheet: sheet.getName(),
      reportStructure: 'Updated to 12 columns (A-L)'
    };
    
  } catch (error) {
    return {
      status: 'error',
      hasDuplicate: false,
      duplicateData: [],
      duplicateKits: [],
      error: error.message
    };
  }
}

