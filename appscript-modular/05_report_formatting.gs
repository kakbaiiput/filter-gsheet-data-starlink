// ========================================
// 📊 15. REPORT FORMATTING & ORGANIZATION FUNCTIONS
// ========================================

/**
 * 🔧 Calculate Running Total
 */
function calculateRunningTotalNew11Column(sheet) {
  try {
    Logger.log('📊 Calculating running total for column J (new structure)...');
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No data to process');
      return;
    }
    
    // Get nominal values dari kolom J
    var nominalRange = sheet.getRange(2, 10, lastRow - 1, 1); // J2:J(lastRow)
    var nominalValues = nominalRange.getValues();

    Logger.log(`Processing ${nominalValues.length} rows for running total`);

    var runningTotal = 0;
    var runningTotalValues = [];

    for (var i = 0; i < nominalValues.length; i++) {
      var nominal = parseFloat(nominalValues[i][0]) || 0;
      runningTotal += nominal;
      runningTotalValues.push([runningTotal]);

      Logger.log(`Row ${i + 2}: Nominal=${nominal} | Running Total=${runningTotal}`);
    }

    // Set running total values ke kolom K
    var totalRange = sheet.getRange(2, 11, lastRow - 1, 1); // K2:K(lastRow)
    totalRange.setValues(runningTotalValues);

    Logger.log('✅ Running total calculation completed (J→K)');
    Logger.log(`Final running total: ${runningTotal}`);
    
    return runningTotal;
    
  } catch (error) {
    Logger.log('❌ Running total calculation error: ' + error.message);
    throw error;
  }
}

/**
 * 🔧 Sort and Organize dengan kolom mapping baru
 */
function sortAndOrganizeWithNew11ColumnStructure(sheet) {
  try {
    Logger.log('🔧 Sort and organize with NEW 11-column structure...');
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 2) {
      Logger.log('Insufficient data to sort');
      return;
    }
    
    // Clear dan unmerge kolom L dulu sebelum sort
    if (lastRow > 1) {
      var clearRange = sheet.getRange(2, 12, lastRow - 1, 1); // L column
      clearRange.clearContent();
      clearRange.clearFormat();

      try {
        clearRange.breakApart();
        Logger.log('✅ Unmerged column L before sorting');
      } catch (e) {
        Logger.log('No merges to break apart');
      }
    }

    // Ensure column L header exists
    sheet.getRange(1, 12).setValue('Total Harian');

    // Sort hanya kolom A-K (tanpa kolom L)
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 11); // A2:K(lastRow) only
    var data = dataRange.getValues();
    
    Logger.log('Data length before sort: ' + data.length);
    
    // Sort data by date (column B - index 1)
    data.sort(function(a, b) {
      var dateA = new Date(a[1]);
      var dateB = new Date(b[1]);
      return dateA.getTime() - dateB.getTime();
    });
    
    Logger.log('✅ Data sorted by date');
    
    // Update nomor urut
    for (var i = 0; i < data.length; i++) {
      data[i][0] = i + 1;
    }
    
    Logger.log('✅ Sequential numbers updated');

    // Kolom I (Periode, index 8) bisa berubah menjadi Date object jika Google Sheets
    // sempat mengkonversi string bulan ("April 2026") sebelumnya.
    // Konversi kembali ke string, lalu paksa format plain text sebelum tulis ulang.
    var BULAN_ID = ['Januari','Februari','Maret','April','Mei','Juni',
                    'Juli','Agustus','September','Oktober','November','Desember'];
    for (var i = 0; i < data.length; i++) {
      if (data[i][8] instanceof Date) {
        var d = data[i][8];
        data[i][8] = BULAN_ID[d.getMonth()] + ' ' + d.getFullYear();
      }
    }
    sheet.getRange(2, 9, lastRow - 1, 1).setNumberFormat('@STRING@');

    // Write sorted data back ke sheet (kolom A-K)
    dataRange.setValues(data);
    
    // Calculate RUNNING TOTAL untuk kolom K
    calculateRunningTotalNew11Column(sheet);

    // Calculate daily summary dengan merge untuk kolom L
    calculateDailySummaryWithMergeNew11Column(sheet);
    
    // ✅ NEW: Apply formatting terakhir
    applyFormattingWithMergeNew11Column(sheet);
    
    Logger.log('✅ Sort and organize with NEW 11-column structure completed');
    
  } catch (error) {
    Logger.log('❌ Sort and organize error: ' + error.message);
    throw error;
  }
}

/**
 * 🔧 Calculate Daily Summary With Merge (Updated)
 */
function calculateDailySummaryWithMergeNew11Column(sheet) {
  try {
    Logger.log('📊 Calculating daily summary for column K (new structure)...');
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No data to process');
      return;
    }
    
    // Get all data setelah sorting (A2:K struktur baru)
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // A2:K(lastRow)

    if (data.length === 0) {
      Logger.log('No data found');
      return;
    }

    var dailyGroups = {};
    var rowsByDate = {};

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var tanggal = row[1]; // Kolom B: Tanggal
      var nominal = parseFloat(row[9]) || 0; // Kolom J: Nominal (index 9)
      
      var dateString = '';
      if (tanggal instanceof Date) {
        var day = tanggal.getDate().toString().padStart(2, '0');
        var month = (tanggal.getMonth() + 1).toString().padStart(2, '0');
        var year = tanggal.getFullYear();
        dateString = `${day}/${month}/${year}`;
      } else if (typeof tanggal === 'string') {
        dateString = tanggal;
      } else {
        continue;
      }
      
      if (!dailyGroups[dateString]) {
        dailyGroups[dateString] = 0;
        rowsByDate[dateString] = [];
      }
      
      dailyGroups[dateString] += nominal;
      rowsByDate[dateString].push(i + 2);
    }
    
    Logger.log('📊 Daily summary calculated for NEW structure:');
    for (var date in dailyGroups) {
      Logger.log(`  ${date}: Rp ${dailyGroups[date].toLocaleString('id-ID')} (${rowsByDate[date].length} transaksi)`);
    }
    
    // Apply daily summary dengan MERGE CELLS di kolom L
    for (var dateString in dailyGroups) {
      var dailyTotal = dailyGroups[dateString];
      var rows = rowsByDate[dateString];
      
      if (rows.length > 0) {
        var formattedTotal = 'Rp ' + dailyTotal.toLocaleString('id-ID') + '\n' + rows.length + ' Starlink';

        var firstRow = Math.min(...rows);
        var lastRowOfGroup = Math.max(...rows);

        // Set value di baris pertama di kolom L
        sheet.getRange(firstRow, 12).setValue(formattedTotal);

        if (rows.length > 1) {
          var mergeRange = sheet.getRange(firstRow, 12, lastRowOfGroup - firstRow + 1, 1);
          mergeRange.merge();
          Logger.log(`✅ Merged L${firstRow}:L${lastRowOfGroup} for ${dateString}`);
        }

        // Apply formatting untuk merged range
        var summaryRange = sheet.getRange(firstRow, 12, lastRowOfGroup - firstRow + 1, 1);
        summaryRange.setBackground('#fef3c7')
                   .setFontWeight('bold')
                   .setFontSize(10)
                   .setFontFamily('Roboto')
                   .setHorizontalAlignment('center')
                   .setVerticalAlignment('middle')
                   .setNumberFormat('@STRING@')
                   .setWrap(true)
                   .setBorder(true, true, true, true, false, false, '#d97706', SpreadsheetApp.BorderStyle.SOLID);
        
        Logger.log(`✅ Applied daily summary for ${dateString}: ${formattedTotal} (${rows.length} transaksi)`);
      }
    }
    
    Logger.log('✅ Daily summary calculation completed for NEW 11-column structure');
    
  } catch (error) {
    Logger.log('❌ Daily summary calculation error: ' + error.message);
    throw error;
  }
}

/**
 * 🔧 Apply Formatting With Merge (Updated)
 */
function applyFormattingWithMergeNew11Column(sheet) {
  try {
    Logger.log('🎨 Applying complete formatting with NEW 11-column structure...');
    
    var lastRow = sheet.getLastRow();
    var lastCol = 12; // Fixed to column L

    if (lastRow <= 1) return;

    Logger.log(`Applying improved formatting to report sheet: ${sheet.getName()} (${lastRow} rows, ${lastCol} columns)`);

    // Header formatting (Row 1)
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight('bold')
              .setBackground('#1e3a8a')
              .setFontColor('white')
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle')
              .setFontSize(11);

    // Data formatting (Row 2 onwards) - Format kolom A-K saja
    if (lastRow > 1) {
      var dataRange = sheet.getRange(2, 1, lastRow - 1, 11); // A2:K(lastRow)
      dataRange.setHorizontalAlignment('center')
               .setVerticalAlignment('middle')
               .setFontSize(10);

      // A: Nomor urut
      sheet.getRange(2, 1, lastRow - 1, 1)
           .setFontWeight('bold')
           .setNumberFormat('0')
           .setHorizontalAlignment('center');

      // B: Tanggal
      sheet.getRange(2, 2, lastRow - 1, 1)
           .setNumberFormat('dd/mm/yyyy')
           .setHorizontalAlignment('center');

      // C: Transaction ID
      sheet.getRange(2, 3, lastRow - 1, 1)
           .setFontFamily('Roboto Mono')
           .setFontWeight('bold')
           .setFontSize(9)
           .setBackground('#f8fafc')
           .setFontColor('#6366f1')
           .setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

      // D: Nama Client
      sheet.getRange(2, 4, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setFontWeight('normal')
           .setHorizontalAlignment('left')
           .setVerticalAlignment('middle');

      // E: KIT Number
      sheet.getRange(2, 5, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setFontWeight('normal')
           .setFontSize(10)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');

      // F: Serial Number
      sheet.getRange(2, 6, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setFontWeight('normal')
           .setFontSize(10)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle')
           .setBackground('#f0f9ff')
           .setFontColor('#0369a1');

      // G: Paket
      sheet.getRange(2, 7, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setHorizontalAlignment('center');

      // H: Tipe Pembayaran
      sheet.getRange(2, 8, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setHorizontalAlignment('center');

      // I: Periode
      sheet.getRange(2, 9, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setHorizontalAlignment('center')
           .setNumberFormat('@STRING@')
           .setFontWeight('bold')
           .setFontColor('#065f46');

      // J: Nominal
      sheet.getRange(2, 10, lastRow - 1, 1)
           .setNumberFormat('"Rp"#,##0')
           .setHorizontalAlignment('center')
           .setFontWeight('bold')
           .setFontColor('#7c3aed');

      // K: Jumlah Total
      sheet.getRange(2, 11, lastRow - 1, 1)
           .setNumberFormat('"Rp"#,##0')
           .setHorizontalAlignment('center')
           .setFontWeight('bold')
           .setBackground('#dcfdf7')
           .setFontColor('#047857');

      // Alternating row colors: batch via RangeList (1 call per warna, bukan 1 call per baris)
      var evenRows = [], oddRows = [];
      for (var row = 2; row <= lastRow; row++) {
        var n = 'A' + row + ':K' + row;
        if (row % 2 === 0) evenRows.push(n); else oddRows.push(n);
      }
      if (evenRows.length > 0) sheet.getRangeList(evenRows).setBackground('#f8f9fa');
      if (oddRows.length > 0)  sheet.getRangeList(oddRows).setBackground('#ffffff');

      // Override backgrounds untuk kolom khusus (satu panggilan per kolom, bukan per baris)
      sheet.getRange(2, 3, lastRow - 1, 1).setBackground('#f8fafc');  // Transaction ID
      sheet.getRange(2, 6, lastRow - 1, 1).setBackground('#f0f9ff');  // Serial Number
      sheet.getRange(2, 11, lastRow - 1, 1).setBackground('#dcfdf7'); // Jumlah Total
    }

    // Auto-resize to fit data content, then enforce minimum widths per column
    sheet.autoResizeColumns(1, lastCol);
    var minWidths = {
      1: 50,   // A: Nomor
      2: 90,   // B: Tanggal
      3: 140,  // C: Transaction ID
      4: 120,  // D: Nama Client
      5: 130,  // E: KIT Number
      6: 130,  // F: Serial Number
      7: 80,   // G: Paket
      8: 90,   // H: Tipe Pembayaran
      9: 80,   // I: Periode
      10: 90,  // J: Nominal
      11: 110, // K: Jumlah Total
      12: 140  // L: Total Harian
    };
    for (var col = 1; col <= Math.min(lastCol, 12); col++) {
      var currentWidth = sheet.getColumnWidth(col);
      if (currentWidth < (minWidths[col] || 80)) {
        sheet.setColumnWidth(col, minWidths[col] || 80);
      }
    }

    Logger.log('✅ Complete formatting with 12-column structure applied');
    
  } catch (error) {
    Logger.log('❌ Complete formatting error: ' + error.message);
  }
}

/**
 * 🔧 Apply Master Formatting To Report Sheet (Updated)
 */
function applyMasterFormattingToReportSheetNew11Column(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    // Pastikan ada kolom L untuk Daily Summary
    if (lastCol < 12) {
      sheet.getRange(1, 12).setValue('Total Harian');
      lastCol = 12;
    }
    
    if (lastRow <= 1) return;
    
    Logger.log(`Applying improved formatting to report sheet: ${sheet.getName()} (${lastRow} rows, ${lastCol} columns) with NEW 11-column structure`);
    
    // Header formatting (Row 1)
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight('bold')
              .setBackground('#1e3a8a')
              .setFontColor('white')
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle')
              .setFontSize(11);
    
    // Data formatting (Row 2 onwards)
    if (lastRow > 1) {
      var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      dataRange.setHorizontalAlignment('center')
               .setVerticalAlignment('middle')
               .setFontSize(10);
      
      // A: Nomor urut
      sheet.getRange(2, 1, lastRow - 1, 1)
           .setFontWeight('bold')
           .setNumberFormat('0')
           .setHorizontalAlignment('center');

      // B: Tanggal
      sheet.getRange(2, 2, lastRow - 1, 1)
           .setNumberFormat('dd/mm/yyyy')
           .setHorizontalAlignment('center');

      // C: Transaction ID
      sheet.getRange(2, 3, lastRow - 1, 1)
           .setFontFamily('Roboto Mono')
           .setFontWeight('bold')
           .setFontSize(9)
           .setBackground('#f8fafc')
           .setFontColor('#6366f1')
           .setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

      // D: Nama Client
      sheet.getRange(2, 4, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setFontWeight('normal')
           .setHorizontalAlignment('left')
           .setVerticalAlignment('middle');

      // E: KIT Number
      sheet.getRange(2, 5, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setFontWeight('normal')
           .setFontSize(10)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');

      // F: Serial Number
      sheet.getRange(2, 6, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setFontWeight('normal')
           .setFontSize(10)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle')
           .setBackground('#f0f9ff')
           .setFontColor('#0369a1');

      // G: Paket
      sheet.getRange(2, 7, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setHorizontalAlignment('center');

      // H: Tipe Pembayaran
      sheet.getRange(2, 8, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setHorizontalAlignment('center');

      // I: Periode
      sheet.getRange(2, 9, lastRow - 1, 1)
           .setFontFamily('Roboto')
           .setHorizontalAlignment('center')
           .setNumberFormat('@STRING@')
           .setFontWeight('bold')
           .setFontColor('#065f46');

      // J: Nominal
      sheet.getRange(2, 10, lastRow - 1, 1)
           .setNumberFormat('"Rp"#,##0')
           .setHorizontalAlignment('center')
           .setFontWeight('bold')
           .setFontColor('#7c3aed');

      // K: Jumlah Total
      sheet.getRange(2, 11, lastRow - 1, 1)
           .setNumberFormat('"Rp"#,##0')
           .setHorizontalAlignment('center')
           .setFontWeight('bold')
           .setBackground('#dcfdf7')
           .setFontColor('#047857');

      // Alternating row colors: batch via RangeList
      var evenRows2 = [], oddRows2 = [];
      for (var row = 2; row <= lastRow; row++) {
        var n = 'A' + row + ':K' + row;
        if (row % 2 === 0) evenRows2.push(n); else oddRows2.push(n);
      }
      if (evenRows2.length > 0) sheet.getRangeList(evenRows2).setBackground('#f8f9fa');
      if (oddRows2.length > 0)  sheet.getRangeList(oddRows2).setBackground('#ffffff');

      // Override backgrounds untuk kolom khusus
      sheet.getRange(2, 3, lastRow - 1, 1).setBackground('#f8fafc');  // Transaction ID
      sheet.getRange(2, 6, lastRow - 1, 1).setBackground('#f0f9ff');  // Serial Number
      sheet.getRange(2, 11, lastRow - 1, 1).setBackground('#dcfdf7'); // Jumlah Total
    }

    // Auto-resize columns dengan minimum widths untuk 12 kolom
    sheet.autoResizeColumns(1, lastCol);

    var minWidths = {
      1: 50,   // A: Nomor
      2: 90,   // B: Tanggal
      3: 140,  // C: Transaction ID
      4: 120,  // D: Nama Client
      5: 130,  // E: KIT Number
      6: 130,  // F: Serial Number
      7: 80,   // G: Paket
      8: 90,   // H: Tipe Pembayaran
      9: 80,   // I: Periode
      10: 90,  // J: Nominal
      11: 110, // K: Jumlah Total
      12: 140  // L: Total Harian
    };

    for (var col = 1; col <= Math.min(lastCol, 12); col++) {
      var currentWidth = sheet.getColumnWidth(col);
      var minWidth = minWidths[col] || 80;

      if (currentWidth < minWidth) {
        sheet.setColumnWidth(col, minWidth);
      }
    }

    Logger.log('✅ Report sheet formatting applied with 12-column structure');
    
  } catch (error) {
    Logger.log('Error formatting report sheet: ' + error.message);
  }
}

/**
 * 📧 Send Email Notification Multi Kit dengan kolom mapping baru
 */
function sendEmailNotificationMultiKitWithTransactionID(formData, sheetName, transactionID, updateResult) {
  try {
    Logger.log('Preparing email notification with transaction date filter info...');
    
    var selectedKitsList = formData.selectedKits.map(function(kit, index) {
      var nominal = kit.nominal || formData.nominal || 0;
      var paymentType = kit.tipePembayaran || '-';
      var paymentIcon = paymentType === 'Aktivasi' ? '🚀' : paymentType === 'Perpanjangan' ? '🔄' : '❓';
      // 🆕 UPDATED: Tampilkan KIT dengan payment type per-KIT
      var kitInfo = `${index + 1}. KIT: ${kit.kitNumber}`;
      if (kit.serialNumber) {
        kitInfo += ` | SN: ${kit.serialNumber}`;
      }
      kitInfo += ` (${kit.paket}) - ${paymentIcon} ${paymentType} - ${formatRupiah(nominal)}`;
      return kitInfo;
    }).join('\n');
    
    var totalNominal = formData.selectedKits.reduce(function(sum, kit) {
      return sum + (kit.nominal || formData.nominal || 0);
    }, 0);
    
    var clientUpdateSummary = '';
    if (updateResult.updated.length > 0) {
      clientUpdateSummary = `\n📋 Client Data Updated (${updateResult.updated.length}):\n`;
      updateResult.updated.forEach(client => {
        clientUpdateSummary += `   ✅ ${client.nama} (${client.kit}) - ${client.sheet}\n`;
      });
    }
    
    if (updateResult.notFound.length > 0) {
      clientUpdateSummary += `\n⚠️ KIT Not Found in Client Data (${updateResult.notFound.length}):\n`;
      updateResult.notFound.forEach(kit => {
        clientUpdateSummary += `   ❌ ${kit}\n`;
      });
    }
    
    // ✅ UPDATED: Format tanggal untuk email
    var transactionDate = new Date(formData.tanggal);
    var formattedTransactionDate = transactionDate.getDate().toString().padStart(2, '0') + '/' + 
                                  (transactionDate.getMonth() + 1).toString().padStart(2, '0') + '/' + 
                                  transactionDate.getFullYear();

var emailBody = `
🆔 TRANSAKSI PEMBAYARAN STARLINK BARU
Transaction ID: ${transactionID}

👤 Client: ${formData.nama}
📅 Tanggal Transaksi: ${formattedTransactionDate}
📊 Sheet: ${sheetName}

🛰️ KIT yang Diperpanjang (${formData.selectedKits.length} KIT):
${selectedKitsList}

💰 Total Nominal: ${formatRupiah(totalNominal)}
${clientUpdateSummary}

✅ Filter otomatis telah diterapkan untuk tanggal transaksi: ${formattedTransactionDate}
📋 Laporan akan menampilkan semua transaksi pada tanggal tersebut.

---
Sistem Form Pembayaran Starlink dengan Transaction ID + Date-Based Filter
`;
    
    Logger.log('Email body prepared with transaction date filter info, sending to baiiput@gmail.com');
    
    MailApp.sendEmail({
      to: 'baiiput@gmail.com',
      subject: `[Starlink] ${transactionID} - ${formData.nama} (${formData.selectedKits.length} KIT) - Filter: ${formattedTransactionDate}`,
      body: emailBody
    });
    
    Logger.log('Email sent successfully with transaction date filter info');
    
  } catch (emailError) {
    Logger.log('Error in sendEmailNotificationMultiKitWithTransactionID: ' + emailError.message);
  }
}

