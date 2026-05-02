// ========================================
// 👥 13. CLIENT DATA MANAGEMENT FUNCTIONS
// ========================================

/**
 * 🔄 UPDATED: Submit Client Data dengan kolom mapping baru
 */
function submitClientData(clientData) {
  try {
    Logger.log('=== SUBMIT CLIENT DATA START (UPDATED MAPPING) ===');
    Logger.log('Client data received: ' + JSON.stringify(clientData));
    
    if (!clientData || !clientData.targetSheet) {
      return {
        status: 'error',
        message: 'Target sheet tidak ditentukan'
      };
    }
    
    const validSheets = ["Client Aktif", "Client Non Aktif", "Client Lepas"];
    if (!validSheets.includes(clientData.targetSheet)) {
      return {
        status: 'error',
        message: 'Target sheet tidak valid. Pilihan: ' + validSheets.join(', ')
      };
    }
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var targetSheet = spreadsheet.getSheetByName(clientData.targetSheet);
    
    if (!targetSheet) {
      return {
        status: 'error',
        message: `Sheet "${clientData.targetSheet}" tidak ditemukan`
      };
    }
    
    var lastRow = targetSheet.getLastRow();
    var newRow = lastRow + 1;
    
    Logger.log(`Target sheet: ${clientData.targetSheet}, New row: ${newRow}`);
    
    // ✅ UPDATED: Data sesuai kolom mapping baru A-Y (25 kolom)
    var rowData = [
      clientData.nama || '',                    // A: Nama
      clientData.loginGmail || '',              // B: Login GMAIL
      clientData.loginStarlink || '',           // C: Login Starlink
      clientData.loginAlternatif || '',         // D: Login Alternatif
      clientData.accNo || '',                   // E: ACC No.
      clientData.emailClient || '',             // F: Email Client
      clientData.nomorCS || '',                 // G: Nomor CS
      clientData.alamat || '',                  // H: Alamat
      clientData.kitNumber || '',               // I: KIT Number ✅ NEW
      clientData.serialNumber || '',            // J: Serial Number ✅ NEW
      clientData.tanggalJatuhTempo || '',       // K: Tanggal Jatuh Tempo ✅ MOVED
      clientData.kode || '',                    // L: Kode ✅ MOVED
      clientData.payment || '',                 // M: Payment ✅ MOVED
      clientData.idTransaksi || '',             // N: ID Transaksi ✅ MOVED
      clientData.status || '',                  // O: STATUS ✅ MOVED
      clientData.paket || '',                   // P: Paket ✅ MOVED
      clientData.last4Digit || '',              // Q: Last 4 Digit ✅ MOVED
      clientData.noRegister || '',              // R: No Register ✅ MOVED
      clientData.tipePelanggan || '',           // S: Tipe Pelanggan ✅ NEW (Full Support, Small Support, Non Support)
      '',                                       // T: Cadangan ✅ NEW
      '',                                       // U: JATUH TEMPO ✅ MOVED
      '',                                       // V: PROSES ✅ MOVED
      '',                                       // W: SEGERA ✅ MOVED
      '',                                       // X: OBSERVASI ✅ MOVED
      ''                                        // Y: Tombol WhatsApp ✅ MOVED
    ];
    
    // Insert data ke sheet (kolom A sampai Y = 25 kolom)
    targetSheet.getRange(newRow, 1, 1, 25).setValues([rowData]);
    
    Logger.log(`✅ Data inserted to ${clientData.targetSheet} at row ${newRow} with updated column mapping`);
    
    // Apply enhanced formatting untuk row baru
    try {
      var newRowRange = targetSheet.getRange(newRow, 1, 1, 25);
      
      // Set alignment untuk SEMUA kolom - center dan middle
      newRowRange.setHorizontalAlignment('center')
                 .setVerticalAlignment('middle')
                 .setFontSize(10);
      
      // KHUSUS untuk kolom A (Nama) - Bold dan font size 14
      var namaCell = targetSheet.getRange(newRow, COLUMNS.NAMA + 1);
      namaCell.setFontWeight('bold')
              .setFontSize(14)
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');
      
      // Format tanggal jika ada (kolom K) - tetap center
      if (clientData.tanggalJatuhTempo) {
        var tanggalCell = targetSheet.getRange(newRow, COLUMNS.JATUH_TEMPO + 1);
        tanggalCell.setNumberFormat('dd/mm/yyyy')
                   .setHorizontalAlignment('center')
                   .setVerticalAlignment('middle');
      }
      
      // Format kolom multi-line - center align juga
      var multiLineColumns = [
        COLUMNS.LOGIN_GMAIL + 1,     // B: Login Gmail
        COLUMNS.LOGIN_STARLINK + 1,  // C: Login Starlink
        COLUMNS.LOGIN_ALTERNATIF + 1, // D: Login Alternatif
        COLUMNS.ALAMAT + 1,          // H: Alamat
        COLUMNS.KIT_NUMBER + 1,      // I: KIT Number ✅ NEW
        COLUMNS.SERIAL_NUMBER + 1    // J: Serial Number ✅ NEW
      ];
      
      multiLineColumns.forEach(function(col) {
        targetSheet.getRange(newRow, col)
                   .setHorizontalAlignment('center')
                   .setVerticalAlignment('middle')
                   .setWrap(true);
      });
      
      // Set border untuk row baru
      newRowRange.setBorder(
        true, true, true, true, true, true,
        '#d1d5db',
        SpreadsheetApp.BorderStyle.SOLID
      );
      
      Logger.log('✅ Enhanced formatting applied with updated column mapping');
      
    } catch (formatError) {
      Logger.log('⚠️ Formatting error (non-critical): ' + formatError.message);
    }
    
    // Send email notification
    try {
      sendClientDataNotification(clientData, clientData.targetSheet, newRow);
      Logger.log('✅ Email notification sent');
    } catch (emailError) {
      Logger.log('⚠️ Email notification error: ' + emailError.message);
    }
    
    var result = {
      status: 'success',
      message: `Data client berhasil disimpan ke "${clientData.targetSheet}" dengan kolom mapping baru`,
      targetSheet: clientData.targetSheet,
      rowNumber: newRow,
      timestamp: new Date().toISOString()
    };
    
    Logger.log('=== SUBMIT CLIENT DATA SUCCESS (UPDATED MAPPING) ===');
    return result;
    
  } catch (error) {
    Logger.log('=== SUBMIT CLIENT DATA ERROR ===');
    Logger.log('Error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    return {
      status: 'error',
      message: 'Error menyimpan data client: ' + error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 📧 Send Email Notification untuk client data baru
 */
function sendClientDataNotification(clientData, sheetName, rowNumber) {
  try {
    Logger.log('Sending client data notification email...');
    
    var emailBody = `
📋 DATA CLIENT BARU DITAMBAHKAN (UPDATED COLUMN MAPPING)

🎯 Target Sheet: ${sheetName}
📍 Row Number: ${rowNumber}
⏰ Timestamp: ${new Date().toLocaleString('id-ID')}

👤 INFORMASI CLIENT:
• Nama: ${clientData.nama || '-'}
• Email Client: ${clientData.emailClient || '-'}
• Nomor CS: ${clientData.nomorCS || '-'}
• KIT Number: ${clientData.kitNumber || '-'} ✅ NEW
• Serial Number: ${clientData.serialNumber || '-'} ✅ NEW
• ACC No: ${clientData.accNo || '-'}

📦 PAKET & STATUS:
• Paket: ${clientData.paket || '-'}
• Status: ${clientData.status || '-'}
• Tanggal Jatuh Tempo: ${clientData.tanggalJatuhTempo || '-'}

🔑 LOGIN INFORMATION:
• Login Gmail: ${clientData.loginGmail || '-'}
• Login Starlink: ${clientData.loginStarlink || '-'}
• Login Alternatif: ${clientData.loginAlternatif || '-'}

💳 PAYMENT INFO:
• Payment: ${clientData.payment || '-'}
• ID Transaksi: ${clientData.idTransaksi || '-'}
• Last 4 Digit: ${clientData.last4Digit || '-'}
• Kode: ${clientData.kode || '-'}

🏠 ALAMAT:
${clientData.alamat || '-'}

📝 REGISTRASI:
• No Register: ${clientData.noRegister || '-'}

---
Form Submission System - Starlink Management (Updated Column Mapping)
`;
    
    MailApp.sendEmail({
      to: 'baiiput@gmail.com',
      subject: `[Client Data] ${clientData.nama || 'New Client'} - ${sheetName} (Updated Mapping)`,
      body: emailBody
    });
    
    Logger.log('Client data notification email sent successfully');
    
  } catch (emailError) {
    Logger.log('Error sending client data notification: ' + emailError.message);
  }
}

/**
 * 🔍 Get Clients By Sheet for dropdown
 */
function getClientsBySheet(sheetName = 'Client Aktif') {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return {
        status: 'error',
        message: `Sheet "${sheetName}" tidak ditemukan`
      };
    }

    var data = sheet.getDataRange().getValues();
    var clientNames = [];

    for (var i = 1; i < data.length; i++) {
      var nama = data[i][COLUMNS.NAMA]; // A: Nama
      if (nama && nama.toString().trim()) {
        clientNames.push({
          nama: nama.toString().trim(),
          rowNumber: i + 1
        });
      }
    }

    var uniqueNames = [];
    var seen = {};
    for (var k = 0; k < clientNames.length; k++) {
      var key = clientNames[k].nama.toLowerCase();
      if (!seen[key]) {
        seen[key] = true;
        uniqueNames.push(clientNames[k]);
      }
    }
    
    uniqueNames.sort(function(a, b) {
      return a.nama.localeCompare(b.nama);
    });

    return {
      status: 'success',
      data: uniqueNames,
      total: uniqueNames.length,
      sheetName: sheetName
    };

  } catch (error) {
    return {
      status: 'error',
      message: 'Error mengambil data client: ' + error.message
    };
  }
}

// ========================================
// 📊 14. DATA RETRIEVAL FUNCTIONS
// ========================================

/**
 * 🔄 UPDATED: Get All Client Data Complete dengan kolom mapping baru
 */
function getAllClientDataComplete() {
  try {
    Logger.log('=== GET ALL CLIENT DATA COMPLETE START (UPDATED MAPPING) ===');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Client Aktif");
    
    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet "Client Aktif" tidak ditemukan'
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        status: 'success',
        data: [],
        message: 'Sheet kosong atau hanya ada header',
        totalRecords: 0
      };
    }
    
    var clientData = [];
    
    Logger.log(`Processing ${data.length - 1} total rows from Client Aktif sheet...`);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var nama = (row[COLUMNS.NAMA] || '').toString().trim();
      
      if (!nama) {
        continue;
      }
      
      var rowData = {};
      
      // ✅ UPDATED: Map semua kolom A-Y (25 kolom) sesuai mapping baru
      var columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y'];
      
      for (var j = 0; j < Math.min(columns.length, row.length); j++) {
        var cellValue = row[j] || '';
        
        // Format tanggal khusus untuk kolom K (Jatuh Tempo) ✅ UPDATED
        if (j === COLUMNS.JATUH_TEMPO && cellValue instanceof Date) {
          var day = cellValue.getDate().toString().padStart(2, '0');
          var month = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          var year = cellValue.getFullYear();
          rowData[columns[j]] = day + '/' + month + '/' + year;
        } else if (cellValue instanceof Date) {
          var day = cellValue.getDate().toString().padStart(2, '0');
          var month = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          var year = cellValue.getFullYear();
          rowData[columns[j]] = day + '/' + month + '/' + year;
        } else {
          rowData[columns[j]] = cellValue.toString().trim();
        }
      }
      
      // Tambahkan metadata
      rowData.rowNumber = i + 1;
      rowData.sheetName = 'Client Aktif';
      
      clientData.push(rowData);
    }
    
    Logger.log(`✅ Successfully processed ${clientData.length} client records (ALL DATA from Client Aktif with updated mapping)`);
    
    return {
      status: 'success',
      data: clientData,
      totalRecords: clientData.length,
      lastUpdate: new Date().toISOString(),
      sheetName: 'Client Aktif',
      columnMapping: 'Updated to new structure with KIT Number (I) and Serial Number (J)'
    };
    
  } catch (error) {
    Logger.log('❌ Error in getAllClientDataComplete: ' + error.message);
    return {
      status: 'error',
      message: 'Error mengambil data: ' + error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🔄 UPDATED: Get Client Lepas Data dengan kolom mapping baru
 */
function getClientLepasData() {
  try {
    Logger.log('=== GET CLIENT LEPAS DATA START (UPDATED MAPPING) ===');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Client Lepas");
    
    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet "Client Lepas" tidak ditemukan'
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        status: 'success',
        data: [],
        message: 'Sheet kosong atau hanya ada header',
        totalRecords: 0
      };
    }
    
    var clientData = [];
    
    Logger.log(`Processing ${data.length - 1} total rows from Client Lepas sheet...`);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var nama = (row[COLUMNS.NAMA] || '').toString().trim();
      
      if (!nama) {
        continue;
      }
      
      var rowData = {};
      
      // ✅ UPDATED: Map semua kolom A-Y (25 kolom) sesuai mapping baru
      var columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y'];
      
      for (var j = 0; j < Math.min(columns.length, row.length); j++) {
        var cellValue = row[j] || '';
        
        // Format tanggal khusus untuk kolom K (Jatuh Tempo) ✅ UPDATED
        if (j === COLUMNS.JATUH_TEMPO && cellValue instanceof Date) {
          var day = cellValue.getDate().toString().padStart(2, '0');
          var month = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          var year = cellValue.getFullYear();
          rowData[columns[j]] = day + '/' + month + '/' + year;
        } else if (cellValue instanceof Date) {
          var day = cellValue.getDate().toString().padStart(2, '0');
          var month = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          var year = cellValue.getFullYear();
          rowData[columns[j]] = day + '/' + month + '/' + year;
        } else {
          rowData[columns[j]] = cellValue.toString().trim();
        }
      }
      
      // Tambahkan metadata
      rowData.rowNumber = i + 1;
      rowData.sheetName = 'Client Lepas';
      
      clientData.push(rowData);
    }
    
    Logger.log(`✅ Successfully processed ${clientData.length} client records (ALL DATA from Client Lepas with updated mapping)`);
    
    return {
      status: 'success',
      data: clientData,
      totalRecords: clientData.length,
      lastUpdate: new Date().toISOString(),
      sheetName: 'Client Lepas',
      columnMapping: 'Updated to new structure with KIT Number (I) and Serial Number (J)'
    };
    
  } catch (error) {
    Logger.log('❌ Error in getClientLepasData: ' + error.message);
    return {
      status: 'error',
      message: 'Error mengambil data: ' + error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🔄 UPDATED: Get Client Non Aktif Data dengan kolom mapping baru
 */
function getClientNonAktifData() {
  try {
    Logger.log('=== GET CLIENT NON AKTIF DATA START (UPDATED MAPPING) ===');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Client Non Aktif");
    
    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet "Client Non Aktif" tidak ditemukan'
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        status: 'success',
        data: [],
        message: 'Sheet kosong atau hanya ada header',
        totalRecords: 0
      };
    }
    
    var clientData = [];
    
    Logger.log(`Processing ${data.length - 1} total rows from Client Non Aktif sheet...`);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var nama = (row[COLUMNS.NAMA] || '').toString().trim();
      
      if (!nama) {
        continue;
      }
      
      var rowData = {};
      
      // ✅ UPDATED: Map semua kolom A-Y (25 kolom) sesuai mapping baru
      var columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y'];
      
      for (var j = 0; j < Math.min(columns.length, row.length); j++) {
        var cellValue = row[j] || '';
        
        // Format tanggal khusus untuk kolom K (Jatuh Tempo) ✅ UPDATED
        if (j === COLUMNS.JATUH_TEMPO && cellValue instanceof Date) {
          var day = cellValue.getDate().toString().padStart(2, '0');
          var month = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          var year = cellValue.getFullYear();
          rowData[columns[j]] = day + '/' + month + '/' + year;
        } else if (cellValue instanceof Date) {
          var day = cellValue.getDate().toString().padStart(2, '0');
          var month = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          var year = cellValue.getFullYear();
          rowData[columns[j]] = day + '/' + month + '/' + year;
        } else {
          rowData[columns[j]] = cellValue.toString().trim();
        }
      }
      
      // Tambahkan metadata
      rowData.rowNumber = i + 1;
      rowData.sheetName = 'Client Non Aktif';
      
      clientData.push(rowData);
    }
    
    Logger.log(`✅ Successfully processed ${clientData.length} client records (ALL DATA from Client Non Aktif with updated mapping)`);
    
    return {
      status: 'success',
      data: clientData,
      totalRecords: clientData.length,
      lastUpdate: new Date().toISOString(),
      sheetName: 'Client Non Aktif',
      columnMapping: 'Updated to new structure with KIT Number (I) and Serial Number (J)'
    };
    
  } catch (error) {
    Logger.log('❌ Error in getClientNonAktifData: ' + error.message);
    return {
      status: 'error',
      message: 'Error mengambil data: ' + error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🔄 Get Client Tertagih Data
 */
function getClientTertagihData() {
  try {
    Logger.log('=== GET CLIENT TERTAGIH DATA START ===');

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Client Tertagih");

    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet "Client Tertagih" tidak ditemukan'
      };
    }

    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        status: 'success',
        data: [],
        message: 'Sheet kosong atau hanya ada header',
        totalRecords: 0
      };
    }

    var clientData = [];

    Logger.log('Processing ' + (data.length - 1) + ' total rows from Client Tertagih sheet...');

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nama = (row[COLUMNS.NAMA] || '').toString().trim();
      if (!nama) continue;

      var rowData = {};
      var columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y'];

      for (var j = 0; j < Math.min(columns.length, row.length); j++) {
        var cellValue = row[j] || '';
        if (j === COLUMNS.JATUH_TEMPO && cellValue instanceof Date) {
          rowData[columns[j]] = cellValue.getDate().toString().padStart(2,'0') + '/' + (cellValue.getMonth()+1).toString().padStart(2,'0') + '/' + cellValue.getFullYear();
        } else if (cellValue instanceof Date) {
          rowData[columns[j]] = cellValue.getDate().toString().padStart(2,'0') + '/' + (cellValue.getMonth()+1).toString().padStart(2,'0') + '/' + cellValue.getFullYear();
        } else {
          rowData[columns[j]] = cellValue.toString().trim();
        }
      }

      rowData.rowNumber = i + 1;
      rowData.sheetName = 'Client Tertagih';
      clientData.push(rowData);
    }

    Logger.log('✅ Successfully processed ' + clientData.length + ' records from Client Tertagih');

    return {
      status: 'success',
      data: clientData,
      totalRecords: clientData.length,
      lastUpdate: new Date().toISOString(),
      sheetName: 'Client Tertagih'
    };

  } catch (error) {
    Logger.log('❌ Error in getClientTertagihData: ' + error.message);
    return {
      status: 'error',
      message: 'Error mengambil data: ' + error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🔄 Get Akun Kosong Data
 */
function getAkunKosongData() {
  try {
    Logger.log('=== GET AKUN KOSONG DATA START ===');

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Akun Kosong");

    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet "Akun Kosong" tidak ditemukan'
      };
    }

    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        status: 'success',
        data: [],
        message: 'Sheet kosong atau hanya ada header',
        totalRecords: 0
      };
    }

    var clientData = [];

    Logger.log('Processing ' + (data.length - 1) + ' total rows from Akun Kosong sheet...');

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nama = (row[COLUMNS.NAMA] || '').toString().trim();
      if (!nama) continue;

      var rowData = {};
      var columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y'];

      for (var j = 0; j < Math.min(columns.length, row.length); j++) {
        var cellValue = row[j] || '';
        if (j === COLUMNS.JATUH_TEMPO && cellValue instanceof Date) {
          rowData[columns[j]] = cellValue.getDate().toString().padStart(2,'0') + '/' + (cellValue.getMonth()+1).toString().padStart(2,'0') + '/' + cellValue.getFullYear();
        } else if (cellValue instanceof Date) {
          rowData[columns[j]] = cellValue.getDate().toString().padStart(2,'0') + '/' + (cellValue.getMonth()+1).toString().padStart(2,'0') + '/' + cellValue.getFullYear();
        } else {
          rowData[columns[j]] = cellValue.toString().trim();
        }
      }

      rowData.rowNumber = i + 1;
      rowData.sheetName = 'Akun Kosong';
      clientData.push(rowData);
    }

    Logger.log('✅ Successfully processed ' + clientData.length + ' records from Akun Kosong');

    return {
      status: 'success',
      data: clientData,
      totalRecords: clientData.length,
      lastUpdate: new Date().toISOString(),
      sheetName: 'Akun Kosong'
    };

  } catch (error) {
    Logger.log('❌ Error in getAkunKosongData: ' + error.message);
    return {
      status: 'error',
      message: 'Error mengambil data: ' + error.message,
      timestamp: new Date().toISOString()
    };
  }
}

function updateClientDataWithTransactionID(selectedKits, transactionID) {
  try {
    Logger.log('Updating client data with Transaction ID: ' + transactionID);
    
    var updateResult = { updated: [], notFound: [] };
    var clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Client Aktif");
    
    if (!clientSheet) {
      Logger.log('❌ Client Aktif sheet not found');
      return updateResult;
    }
    
    var data = clientSheet.getDataRange().getValues();
    
    selectedKits.forEach(function(kitInfo) {
      var kitNumber = kitInfo.kitNumber;
      var updated = false;
      
      for (var i = 1; i < data.length; i++) {
        var kitCell = data[i][COLUMNS.KIT_NUMBER]; // ✅ NEW: Kolom I (KIT Number)
        
        if (kitCell) {
          var kitString = kitCell.toString().trim();
          var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];
          
          for (var j = 0; j < kitArray.length; j++) {
            if (kitArray[j].trim().toUpperCase() === kitNumber.toUpperCase()) {
              
              var currentTransactionID = data[i][COLUMNS.TRANSACTION_ID] || ''; // ✅ NEW: Kolom N
              var newTransactionID;
              
              if (currentTransactionID && currentTransactionID.toString().trim()) {
                newTransactionID = currentTransactionID.toString().trim() + '\n' + transactionID;
              } else {
                newTransactionID = transactionID;
              }
              
              clientSheet.getRange(i + 1, COLUMNS.TRANSACTION_ID + 1).setValue(newTransactionID); // ✅ NEW: Kolom N
              
              updateResult.updated.push({
                nama: data[i][COLUMNS.NAMA] || '',
                kit: kitNumber,
                sheet: 'Client Aktif',
                row: i + 1,
                transactionID: transactionID
              });
              
              updated = true;
              Logger.log(`✅ Updated ${kitNumber} with Transaction ID: ${transactionID}`);
              break;
            }
          }
          
          if (updated) break;
        }
      }
      
      if (!updated) {
        updateResult.notFound.push(kitNumber);
        Logger.log(`⚠️ KIT not found in Client Aktif: ${kitNumber}`);
      }
    });
    
    Logger.log(`Client data update completed: ${updateResult.updated.length} updated, ${updateResult.notFound.length} not found`);
    return updateResult;
    
  } catch (error) {
    Logger.log('Error updating client data: ' + error.message);
    return {
      updated: [],
      notFound: selectedKits.map(kit => kit.kitNumber),
      error: error.message
    };
  }
}

