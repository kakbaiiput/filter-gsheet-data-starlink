// ========================================
// 🆔 3. TRANSACTION ID FUNCTIONS
// ========================================

function generateTransactionID(tanggal, month, year) {
  try {
    var day = tanggal.getDate().toString().padStart(2, '0');
    var monthStr = month.toString().padStart(2, '0');
    var yearStr = year.toString();
    // 🔧 FIX: Changed format from DDMMYYYY to DD-MM-YYYY
    var datePrefix = day + '-' + monthStr + '-' + yearStr;

    Logger.log('🔍 Looking for existing Transaction IDs in month: ' + monthStr + '/' + yearStr + ' (resets monthly)');

    var reportSpreadsheet = getReportSpreadsheet();
    var sheetName = monthStr + '_' + yearStr;
    var sheet = reportSpreadsheet.getSheetByName(sheetName);
    var increment = 1;

    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var sameDateTransactions = [];

      for (var i = 1; i < data.length; i++) {
        var existingTransactionID = data[i][2]; // Kolom C: Transaction ID

        if (existingTransactionID) {
          var idString = existingTransactionID.toString().trim();

          // 🔧 FIX: Match BOTH old and new format for backward compatibility
          // Old format: SL/DDMMYYYY/XXX (e.g., SL/20012026/241)
          // New format: SL/DD-MM-YYYY/XXX (e.g., SL/20-01-2026/241)
          // Both reset every month (not every day)
          var oldFormatPattern = '^SL\\/\\d{2}' + monthStr + yearStr + '\\/\\d{3}$';
          var newFormatPattern = '^SL\\/\\d{2}-' + monthStr + '-' + yearStr + '\\/\\d{3}$';
          var monthPattern = new RegExp('(' + oldFormatPattern + ')|(' + newFormatPattern + ')');

          if (monthPattern.test(idString)) {
            var parts = idString.split('/');
            if (parts.length === 3) {
              var existingIncrement = parseInt(parts[2]);
              if (!isNaN(existingIncrement)) {
                sameDateTransactions.push(existingIncrement);
              }
            }
          }
        }
      }

      if (sameDateTransactions.length > 0) {
        var maxIncrement = Math.max.apply(Math, sameDateTransactions);
        increment = maxIncrement + 1;
        Logger.log('📊 Found ' + sameDateTransactions.length + ' existing transactions in this month. Next increment: ' + increment);
      } else {
        Logger.log('📊 No existing transactions found in this month. Starting from: 001');
      }
    }

    var transactionID = 'SL/' + datePrefix + '/' + increment.toString().padStart(3, '0');
    Logger.log('✅ Generated Transaction ID: ' + transactionID);
    return transactionID;

  } catch (error) {
    Logger.log('❌ Error generating transaction ID: ' + error.message);
    var timestamp = new Date().getTime().toString().slice(-6);
    return 'SL/' + timestamp + '/001';
  }
}

// Batch version: reads sheet once, generates 'count' sequential IDs
function generateBatchTransactionIDs(tanggal, month, year, count) {
  try {
    var day = tanggal.getDate().toString().padStart(2, '0');
    var monthStr = month.toString().padStart(2, '0');
    var yearStr = year.toString();
    var datePrefix = day + '-' + monthStr + '-' + yearStr;

    var reportSpreadsheet = getReportSpreadsheet();
    var sheetName = monthStr + '_' + yearStr;
    var sheet = reportSpreadsheet.getSheetByName(sheetName);
    var nextIncrement = 1;

    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var oldFormatPattern = '^SL\\/\\d{2}' + monthStr + yearStr + '\\/\\d{3}$';
      var newFormatPattern = '^SL\\/\\d{2}-' + monthStr + '-' + yearStr + '\\/\\d{3}$';
      var monthPattern = new RegExp('(' + oldFormatPattern + ')|(' + newFormatPattern + ')');
      var maxIncrement = 0;
      for (var i = 1; i < data.length; i++) {
        var idString = (data[i][2] || '').toString().trim();
        if (monthPattern.test(idString)) {
          var parts = idString.split('/');
          if (parts.length === 3) {
            var n = parseInt(parts[2]);
            if (!isNaN(n) && n > maxIncrement) maxIncrement = n;
          }
        }
      }
      nextIncrement = maxIncrement + 1;
    }

    var ids = [];
    for (var k = 0; k < count; k++) {
      ids.push('SL/' + datePrefix + '/' + (nextIncrement + k).toString().padStart(3, '0'));
    }
    Logger.log('✅ Generated ' + count + ' batch Transaction IDs: ' + ids[0] + ' ... ' + ids[ids.length - 1]);
    return ids;
  } catch (error) {
    Logger.log('❌ Error generating batch transaction IDs: ' + error.message);
    var ids = [];
    var ts = new Date().getTime().toString().slice(-6);
    for (var k = 0; k < count; k++) ids.push('SL/' + ts + '/' + (k + 1).toString().padStart(3, '0'));
    return ids;
  }
}

// Batch version: reads client sheet once, writes all KIT transaction IDs in one pass
function updateClientDataBatch(kitsWithTransactionIDs) {
  try {
    var updateResult = { updated: [], notFound: [] };
    var clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Client Aktif");
    if (!clientSheet) {
      Logger.log('❌ Client Aktif sheet not found');
      return updateResult;
    }

    var data = clientSheet.getDataRange().getValues();
    // Build map: rowIndex → accumulated new transactionID string
    var rowUpdates = {}; // key: row index (0-based in data), value: new transactionID string

    kitsWithTransactionIDs.forEach(function(item) {
      var kitNumber = item.kit.kitNumber;
      var transactionID = item.transactionID;
      var found = false;

      for (var i = 1; i < data.length; i++) {
        var kitCell = data[i][COLUMNS.KIT_NUMBER];
        if (!kitCell) continue;
        var kitArray = kitCell.toString().trim().split('\n');
        for (var j = 0; j < kitArray.length; j++) {
          if (kitArray[j].trim().toUpperCase() === kitNumber.toUpperCase()) {
            // Accumulate on the in-memory data so multiple KITs on same row work
            var current = (data[i][COLUMNS.TRANSACTION_ID] || '').toString().trim();
            var pending = rowUpdates[i] !== undefined ? rowUpdates[i] : current;
            rowUpdates[i] = pending ? pending + '\n' + transactionID : transactionID;
            data[i][COLUMNS.TRANSACTION_ID] = rowUpdates[i]; // keep in-memory up to date
            updateResult.updated.push({ kit: kitNumber, row: i + 1, transactionID: transactionID });
            found = true;
            break;
          }
        }
        if (found) break;
      }
      if (!found) {
        updateResult.notFound.push(kitNumber);
        Logger.log('⚠️ KIT not found in Client Aktif: ' + kitNumber);
      }
    });

    // Write all row updates — one setValue per unique row (reads already batched above)
    var rowKeys = Object.keys(rowUpdates);
    for (var ri = 0; ri < rowKeys.length; ri++) {
      var row = parseInt(rowKeys[ri]);
      clientSheet.getRange(row + 1, COLUMNS.TRANSACTION_ID + 1).setValue(rowUpdates[rowKeys[ri]]);
    }

    Logger.log('✅ Batch client update: ' + updateResult.updated.length + ' updated, ' + updateResult.notFound.length + ' not found');
    return updateResult;
  } catch (error) {
    Logger.log('❌ Error in updateClientDataBatch: ' + error.message);
    return { updated: [], notFound: kitsWithTransactionIDs.map(function(i) { return i.kit.kitNumber; }), error: error.message };
  }
}

