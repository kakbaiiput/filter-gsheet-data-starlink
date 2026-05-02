// ========================================
// 👥 10. CLIENT MANAGEMENT FUNCTIONS
// ========================================

/**
 * 🔄 UPDATED: Validate KIT dengan kolom mapping baru + Serial Number validation
 */
function validateKitForFormMultiKit(kitNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];

    Logger.log('🔍 Validating input: ' + kitNumber);
    Logger.log('   - Searching in both KIT Number and Serial Number columns');

    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) continue;

      var data = sheet.getDataRange().getValues();

      for (var i = 1; i < data.length; i++) {
        var kitCell = data[i][COLUMNS.KIT_NUMBER]; // I: KIT Number ✅ NEW
        var serialCell = data[i][COLUMNS.SERIAL_NUMBER] || ''; // J: Serial Number ✅ NEW

        // 🔧 FIX: Search in KIT Number column
        if (kitCell) {
          var kitString = kitCell.toString().trim();
          var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];

          for (var j = 0; j < kitArray.length; j++) {
            if (kitArray[j].trim().toUpperCase() === kitNumber.toUpperCase()) {
              Logger.log('✅ Found match in KIT Number column');
              return buildKitValidationResult(data, i, sheetName, kitNumber, 'KIT Number');
            }
          }
        }

        // 🔧 FIX: Search in Serial Number column (ALWAYS, not just if starts with UT)
        if (serialCell) {
          var serialString = serialCell.toString().trim();
          var serialArray = serialString.includes('\n') ? serialString.split('\n') : [serialString];

          for (var j = 0; j < serialArray.length; j++) {
            if (serialArray[j].trim().toUpperCase() === kitNumber.toUpperCase()) {
              Logger.log('✅ Found match in Serial Number column');
              return buildKitValidationResult(data, i, sheetName, kitNumber, 'Serial Number');
            }
          }
        }
      }
    }

    return {
      status: 'not_found',
      clientStatus: null,
      isActive: false,
      data: null,
      warning: "❌ KIT/Serial Number tidak ditemukan dalam database!"
    };

  } catch (error) {
    return {
      status: 'error',
      clientStatus: null,
      isActive: false,
      data: null,
      warning: "❌ Error dalam validasi: " + error.message
    };
  }
}

/**
 * 🆕 Helper function to build validation result
 */
function buildKitValidationResult(data, rowIndex, sheetName, searchValue, foundIn) {
  var nama = data[rowIndex][COLUMNS.NAMA] || '';
  var paketCell = data[rowIndex][COLUMNS.PAKET] || ''; // P: Paket
  var kitCell = data[rowIndex][COLUMNS.KIT_NUMBER] || ''; // I: KIT Number
  var serialCell = data[rowIndex][COLUMNS.SERIAL_NUMBER] || ''; // J: Serial Number

  Logger.log('🔍 DEBUG - Data yang diambil (found by ' + foundIn + '):');
  Logger.log('   - Nama: ' + nama);
  Logger.log('   - KIT Cell (col I): ' + kitCell);
  Logger.log('   - Serial Cell (col J): ' + serialCell);
  Logger.log('   - Paket Cell (col P): ' + paketCell);

  var kitString = kitCell.toString().trim();
  var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];

  var paketString = paketCell.toString().trim();
  var paketArray = paketString.includes('\n') ? paketString.split('\n') : [paketString];

  var serialString = serialCell.toString().trim();
  var serialArray = serialString.includes('\n') ? serialString.split('\n') : [serialString];

  var allKits = [];

  for (var k = 0; k < kitArray.length; k++) {
    var kit = kitArray[k].trim();
    var paket = paketArray[k] || paketArray[0] || '';
    var serial = serialArray[k] || serialArray[0] || '';

    if (kit) {
      var isMatched = false;

      // Check if this kit matches the search value (by KIT or Serial)
      if (foundIn === 'KIT Number') {
        isMatched = kit.toUpperCase() === searchValue.toUpperCase();
      } else if (foundIn === 'Serial Number') {
        isMatched = serial.trim().toUpperCase() === searchValue.toUpperCase();
      }

      allKits.push({
        kitNumber: kit,
        serialNumber: serial.trim(),
        paket: paket.trim(),
        isSelected: isMatched
      });

      Logger.log(`   - KIT ${k + 1}: ${kit} | SN: ${serial.trim()} | Paket: ${paket.trim()} | Matched: ${isMatched}`);
    }
  }

  return {
    status: 'found',
    clientStatus: sheetName,
    isActive: sheetName === "Client Aktif",
    foundBy: foundIn, // NEW: indicate what field matched
    data: {
      nama: nama.toString().trim(),
      rowNumber: rowIndex + 1,
      sheetName: sheetName,
      allKits: allKits,
      totalKits: allKits.length
    },
    warning: sheetName !== "Client Aktif" ?
      "⚠️ PERHATIAN: Client tidak berada di 'Client Aktif'. Segera update data!" : null
  };
}

/**
 * 🔄 UPDATED: Get All Client Names
 */
function getAllClientNames() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Client Aktif");
    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet "Client Aktif" tidak ditemukan'
      };
    }

    var data = sheet.getDataRange().getValues();
    var clientNames = [];

    for (var i = 1; i < data.length; i++) {
      var nama = data[i][COLUMNS.NAMA]; // A: Nama
      if (nama && nama.toString().trim()) {
        clientNames.push(nama.toString().trim());
      }
    }

    var uniqueNames = [];
    var seen = {};
    for (var k = 0; k < clientNames.length; k++) {
      if (!seen[clientNames[k]]) {
        seen[clientNames[k]] = true;
        uniqueNames.push(clientNames[k]);
      }
    }
    uniqueNames.sort();

    return {
      status: 'success',
      data: uniqueNames,
      total: uniqueNames.length
    };

  } catch (error) {
    return {
      status: 'error',
      message: 'Error mengambil data client: ' + error.message
    };
  }
}

/**
 * 🔄 FIXED: Get All Clients By Status dengan field mapping yang benar
 */
function getAllClientsByStatus() {
  try {
    console.log('=== FIXED getAllClientsByStatus START ===');
    
    var allData = {
      proses: [],
      warningReminder: [],
      reminderH1: [],
      reminderH2: [],
      reminderH3: [],
      jatuhTempo: [],
      segera: [],
      observasi: []
    };
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Client Aktif");
    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet "Client Aktif" tidak ditemukan'
      };
    }

    var data = sheet.getDataRange().getValues();
    console.log(`Processing ${data.length - 1} rows...`);

    for (var i = 1; i < data.length; i++) {
      var nama = data[i][COLUMNS.NAMA] || '';
      var loginGmail = data[i][COLUMNS.LOGIN_GMAIL] || '';
      var loginStarlink = data[i][COLUMNS.LOGIN_STARLINK] || '';
      var emailClient = data[i][COLUMNS.EMAIL_CLIENT] || '';
      var nomorWA = data[i][COLUMNS.NOMOR_CS] || '';
      
      // ✅ FIXED: Proper extraction dengan debug logging
      var kitNumberRaw = data[i][COLUMNS.KIT_NUMBER]; // Column I (index 8)
      var serialNumberRaw = data[i][COLUMNS.SERIAL_NUMBER]; // Column J (index 9)
      var kodeRaw = data[i][COLUMNS.KODE]; // Column L (index 11)
      var paymentRaw = data[i][COLUMNS.PAYMENT]; // Column M (index 12)
      var tipePelangganRaw = data[i][COLUMNS.TIPE_PELANGGAN]; // Column S (index 18)

      // ✅ FIXED: Safe conversion to strings
      var kitNumber = '';
      var serialNumber = '';
      var kode = '';
      var payment = '';
      var tipePelanggan = '';

      if (kitNumberRaw !== null && kitNumberRaw !== undefined) {
        kitNumber = kitNumberRaw.toString().trim();
      }

      if (serialNumberRaw !== null && serialNumberRaw !== undefined) {
        serialNumber = serialNumberRaw.toString().trim();
      }

      if (kodeRaw !== null && kodeRaw !== undefined) {
        kode = kodeRaw.toString().trim();
      }

      if (paymentRaw !== null && paymentRaw !== undefined) {
        payment = paymentRaw.toString().trim();
      }

      if (tipePelangganRaw !== null && tipePelangganRaw !== undefined) {
        tipePelanggan = tipePelangganRaw.toString().trim();
      }

      var jatuhTempo = data[i][COLUMNS.JATUH_TEMPO] || '';
      var transactionID = data[i][COLUMNS.TRANSACTION_ID] || '';
      var status = data[i][COLUMNS.STATUS] || '';
      var paket = data[i][COLUMNS.PAKET] || '';
      
      // Status indicators
      var jatuhTempoLabel = data[i][COLUMNS.LABEL_JATUH_TEMPO] || '';
      var prosesLabel = data[i][COLUMNS.LABEL_PROSES] || '';
      var segaraLabel = data[i][COLUMNS.LABEL_SEGERA] || '';
      var observasiLabel = data[i][COLUMNS.LABEL_OBSERVASI] || '';
      
      if (!nama.toString().trim()) continue;
      
      var jatuhTempoFormatted = '';
      if (jatuhTempo instanceof Date) {
        jatuhTempoFormatted = formatTanggalIndonesia(jatuhTempo);
      } else if (jatuhTempo) {
        var tempDate = new Date(jatuhTempo);
        if (!isNaN(tempDate.getTime())) {
          jatuhTempoFormatted = formatTanggalIndonesia(tempDate);
        }
      }
      
      // ✅ FIXED: Pastikan semua field dikirim dengan benar
      var clientData = {
        nama: nama.toString(),
        loginGmail: loginGmail.toString(),
        loginStarlink: loginStarlink.toString(),
        emailClient: emailClient.toString(),
        nomorWA: nomorWA.toString(),
        kitNumber: kitNumber, // ✅ FIXED: Kirim sebagai string, bukan undefined
        serialNumber: serialNumber, // ✅ FIXED: Kirim sebagai string, bukan undefined
        jatuhTempo: jatuhTempoFormatted,
        paket: paket.toString(),
        kode: kode, // ✅ FIXED: Kirim sebagai string, bukan undefined
        payment: payment, // ✅ FIXED: Kirim sebagai string, bukan undefined
        transactionID: transactionID.toString(),
        status: status.toString(),
        tipePelanggan: tipePelanggan, // ✅ NEW: Tipe Pelanggan dari kolom S
        statusClient: "Client Aktif",
        rowNumber: i + 1
      };
      
      // Kategorikan berdasarkan label
      if (prosesLabel.toString().toUpperCase() === 'PROSES') {
        allData.proses.push(clientData);
      }
      
      if (jatuhTempoLabel.toString().toUpperCase() === 'JATUH TEMPO') {
        allData.jatuhTempo.push(clientData);
      }
      
      if (segaraLabel.toString().toUpperCase() === 'SEGERA') {
        allData.segera.push(clientData);
      }
      
      if (observasiLabel.toString().toUpperCase() === 'OBSERVASI') {
        allData.observasi.push(clientData);
      }
      
      // ✅ Cek reminder dari kolom Y
      try {
        var reminderCell = sheet.getRange(i + 1, COLUMNS.TOMBOL_WHATSAPP + 1);
        var reminderRichText = reminderCell.getRichTextValue();
        
        if (reminderRichText) {
          var reminderText = reminderRichText.getText();
          var reminderUrl = reminderRichText.getLinkUrl();
          
          if (reminderText) {
            if (reminderUrl) clientData.waLink = reminderUrl;
            
            if (reminderText.includes('💀')) {
              allData.warningReminder.push(clientData);
            } else if (reminderText.includes('🚨')) {
              allData.reminderH1.push(clientData);
            } else if (reminderText.includes('⚠️')) {
              allData.reminderH2.push(clientData);
            } else if (reminderText.includes('📅')) {
              allData.reminderH3.push(clientData);
            }
          }
        }
        
      } catch (reminderError) {
        console.log(`❌ Error reading reminder cell for row ${i+1}: ${reminderError.message}`);
      }
    }
    
    return {
      status: 'success',
      data: allData
    };
    
  } catch (error) {
    console.log(`❌ Error in FIXED getAllClientsByStatus: ${error.message}`);
    console.log(`Stack: ${error.stack}`);
    return {
      status: 'error',
      message: 'Error mengambil data status: ' + error.message
    };
  }
}

