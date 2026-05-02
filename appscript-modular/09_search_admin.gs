// ========================================
// 🔍 4. SEARCH HELPER FUNCTIONS
// ========================================

function buildMaskedData(data, rowIndex, sheetName) {
  var nama = data[rowIndex][COLUMNS.NAMA] || '-';
  var email = data[rowIndex][COLUMNS.LOGIN_STARLINK] || '-';
  var emailClient = data[rowIndex][COLUMNS.EMAIL_CLIENT] || '-'; 
  var nomorWA = data[rowIndex][COLUMNS.NOMOR_CS] || '-';
  var kitNumber = data[rowIndex][COLUMNS.KIT_NUMBER] || '-'; // ✅ NEW
  var serialNumber = data[rowIndex][COLUMNS.SERIAL_NUMBER] || '-'; // ✅ NEW
  var jatuhTempo = data[rowIndex][COLUMNS.JATUH_TEMPO] || '-'; // ✅ NEW
  var transactionID = data[rowIndex][COLUMNS.TRANSACTION_ID] || '-';
  var status = data[rowIndex][COLUMNS.STATUS] || '-';
  var paket = data[rowIndex][COLUMNS.PAKET] || '-';
  var noRegister = data[rowIndex][COLUMNS.NO_REGISTER] || '-';
  
  var jatuhTempoFormatted = '-';
  if (jatuhTempo && jatuhTempo instanceof Date) {
    jatuhTempoFormatted = formatTanggalIndonesia(jatuhTempo);
  } else if (jatuhTempo && jatuhTempo !== '-') {
    var tempDate = new Date(jatuhTempo);
    if (!isNaN(tempDate.getTime())) {
      jatuhTempoFormatted = formatTanggalIndonesia(tempDate);
    }
  }
  
  var emailMasked = '-';
  if (email && email.includes('@')) {
    var emailParts = email.split('@');
    var localPart = emailParts[0];
    var domain = emailParts[1];
    emailMasked = localPart.substring(0, 3) + '***@' + domain;
  }
  
  var nomorWAMasked = '-';
  if (nomorWA && nomorWA.toString().length > 4) {
    nomorWAMasked = nomorWA.toString().substring(0, 4) + '****';
  }
  
  return {
    nama: nama,
    email: emailMasked,
    emailClient: emailClient,
    nomorWA: nomorWAMasked,
    kitNumber: kitNumber, // ✅ NEW
    serialNumber: serialNumber, // ✅ NEW
    jatuhTempo: jatuhTempoFormatted,
    transactionID: transactionID,
    status: status,
    paket: paket,
    noRegister: noRegister,
    statusClient: sheetName,
    rowNumber: rowIndex + 1
  };
}

function buildRowData(data, rowIndex, sheetName) {
  var rowData = {};
  var columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'];
  
  for (var col = 0; col < 25; col++) {
    var cellValue = data[rowIndex][col] || '';
    
    if (cellValue instanceof Date) {
      if (col === COLUMNS.JATUH_TEMPO) { // ✅ NEW: Kolom K
        cellValue = formatTanggalIndonesia(cellValue);
      } else {
        cellValue = cellValue.getDate() + '/' + (cellValue.getMonth() + 1) + '/' + cellValue.getFullYear();
      }
    }
    
    rowData[columns[col]] = cellValue.toString();
  }
  
  // ✅ NEW: Gabungkan KIT Number & Paket
  var kitData = rowData['I'] || '';
  var paketData = rowData['P'] || '';
  var kitPaketCombined = '';
  
  if (kitData && paketData) {
    var kitLines = kitData.includes('\n') ? kitData.split('\n') : [kitData];
    var paketLines = paketData.includes('\n') ? paketData.split('\n') : [paketData];
    var combinedLines = [];
    var maxLines = Math.max(kitLines.length, paketLines.length);
    
    for (var k = 0; k < maxLines; k++) {
      var kit = kitLines[k] || kitLines[0] || '';
      var paket = paketLines[k] || paketLines[0] || '';
      combinedLines.push(kit.trim() + ' (' + paket.trim() + ')');
    }
    
    kitPaketCombined = combinedLines.join('\n');
  }
  
  rowData['I_P_COMBINED'] = kitPaketCombined; // ✅ NEW
  rowData['statusClient'] = sheetName;
  rowData['rowNumber'] = rowIndex + 1;
  
  return rowData;
}

function buildResponse(data, rowIndex, headerRow, sheetName) {
  var rowData = {};
  var columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'];
  
  for (var col = 0; col < 25; col++) {
    var cellValue = data[rowIndex][col] || '';
    
    if (cellValue instanceof Date) {
      if (col === COLUMNS.JATUH_TEMPO) { // ✅ NEW: Kolom K
        cellValue = formatTanggalIndonesia(cellValue);
      } else {
        cellValue = cellValue.getDate() + '/' + (cellValue.getMonth() + 1) + '/' + cellValue.getFullYear();
      }
    }
    
    rowData[columns[col]] = cellValue.toString();
  }
  
  // ✅ NEW: Gabungkan KIT & Paket
  var kitData = rowData['I'] || '';
  var paketData = rowData['P'] || '';
  var kitPaketCombined = '';
  
  if (kitData && paketData) {
    var kitLines = kitData.includes('\n') ? kitData.split('\n') : [kitData];
    var paketLines = paketData.includes('\n') ? paketData.split('\n') : [paketData];
    var combinedLines = [];
    var maxLines = Math.max(kitLines.length, paketLines.length);
    
    for (var k = 0; k < maxLines; k++) {
      var kit = kitLines[k] || kitLines[0] || '';
      var paket = paketLines[k] || paketLines[0] || '';
      combinedLines.push(kit.trim() + ' (' + paket.trim() + ')');
    }
    
    kitPaketCombined = combinedLines.join('\n');
  }
  
  rowData['I_P_COMBINED'] = kitPaketCombined; // ✅ NEW
  
  return {
    status: 'success',
    data: {
      ...rowData,
      statusClient: sheetName
    },
    rowNumber: rowIndex + 1,
    headers: {
      'A': headerRow[0] || 'Nama',
      'B': headerRow[1] || 'Login GMAIL',
      'C': headerRow[2] || 'Login Starlink',
      'D': headerRow[3] || 'Login Alternatif',
      'E': headerRow[4] || 'ACC No.',
      'F': headerRow[5] || 'Email Client',
      'G': headerRow[6] || 'Nomor CS',
      'H': headerRow[7] || 'Alamat',
      'I': headerRow[8] || 'KIT Number', // ✅ NEW
      'J': headerRow[9] || 'Serial Number', // ✅ NEW
      'K': headerRow[10] || 'Tanggal Jatuh Tempo', // ✅ NEW
      'L': headerRow[11] || 'Kode',
      'M': headerRow[12] || 'Payment',
      'N': headerRow[13] || 'ID Transaksi',
      'O': headerRow[14] || 'STATUS',
      'P': headerRow[15] || 'Paket',
      'Q': headerRow[16] || 'Last 4 Digit',
      'R': headerRow[17] || 'No Register',
      'S': headerRow[18] || 'Cadangan 1',
      'T': headerRow[19] || 'Cadangan 2',
      'U': headerRow[20] || 'JATUH TEMPO',
      'V': headerRow[21] || 'PROSES',
      'W': headerRow[22] || 'SEGERA',
      'X': headerRow[23] || 'OBSERVASI',
      'Y': headerRow[24] || 'Tombol WhatsApp'
    }
  };
}

function buildHeadersObject(headerRow) {
  return {
    'A': headerRow[0] || 'Nama',
    'B': headerRow[1] || 'Login GMAIL',
    'C': headerRow[2] || 'Login Starlink',
    'D': headerRow[3] || 'Login Alternatif',
    'E': headerRow[4] || 'ACC No.',
    'F': headerRow[5] || 'Email Client',
    'G': headerRow[6] || 'Nomor CS',
    'H': headerRow[7] || 'Alamat',
    'I': headerRow[8] || 'KIT Number', // ✅ NEW
    'J': headerRow[9] || 'Serial Number', // ✅ NEW
    'K': headerRow[10] || 'Tanggal Jatuh Tempo', // ✅ NEW
    'L': headerRow[11] || 'Kode',
    'M': headerRow[12] || 'Payment',
    'N': headerRow[13] || 'ID Transaksi',
    'O': headerRow[14] || 'STATUS',
    'P': headerRow[15] || 'Paket',
    'Q': headerRow[16] || 'Last 4 Digit',
    'R': headerRow[17] || 'No Register',
    'S': headerRow[18] || 'Cadangan 1',
    'T': headerRow[19] || 'Cadangan 2',
    'U': headerRow[20] || 'JATUH TEMPO',
    'V': headerRow[21] || 'PROSES',
    'W': headerRow[22] || 'SEGERA',
    'X': headerRow[23] || 'OBSERVASI',
    'Y': headerRow[24] || 'Tombol WhatsApp'
  };
}

// ========================================
// 🔍 5. SEARCH FUNCTIONS - ADMIN VERSION
// ========================================

// 5.1 KIT SEARCH FUNCTIONS
function searchByKit(kitNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
    
      for (var i = 1; i < data.length; i++) {
        var kitCell = data[i][COLUMNS.KIT_NUMBER]; // ✅ NEW: Kolom I (KIT Number)
        
        var kitString = kitCell.toString().trim();
        var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];
        
        for (var j = 0; j < kitArray.length; j++) {
          if (kitArray[j].trim().toLowerCase() === kitNumber.toLowerCase()) {
            
            var nama = data[i][COLUMNS.NAMA] || '-';
            var email = data[i][COLUMNS.LOGIN_STARLINK] || '-';
            var emailClient = data[i][COLUMNS.EMAIL_CLIENT] || '-'; 
            var nomorWA = data[i][COLUMNS.NOMOR_CS] || '-';
            var kitNumber = data[i][COLUMNS.KIT_NUMBER] || '-'; // ✅ NEW
            var serialNumber = data[i][COLUMNS.SERIAL_NUMBER] || '-'; // ✅ NEW
            var jatuhTempo = data[i][COLUMNS.JATUH_TEMPO] || '-'; // ✅ NEW
            var transactionID = data[i][COLUMNS.TRANSACTION_ID] || '-';
            var status = data[i][COLUMNS.STATUS] || '-';
            var paket = data[i][COLUMNS.PAKET] || '-';
            var noRegister = data[i][COLUMNS.NO_REGISTER] || '-';
            
            var jatuhTempoFormatted = '-';
            if (jatuhTempo && jatuhTempo instanceof Date) {
              jatuhTempoFormatted = formatTanggalIndonesia(jatuhTempo);
            } else if (jatuhTempo && jatuhTempo !== '-') {
              var tempDate = new Date(jatuhTempo);
              if (!isNaN(tempDate.getTime())) {
                jatuhTempoFormatted = formatTanggalIndonesia(tempDate);
              }
            }
            
            var emailMasked = '-';
            if (email && email.includes('@')) {
              var emailParts = email.split('@');
              var localPart = emailParts[0];
              var domain = emailParts[1];
              emailMasked = localPart.substring(0, 3) + '***@' + domain;
            }
            
            var nomorWAMasked = '-';
            if (nomorWA && nomorWA.toString().length > 4) {
              nomorWAMasked = nomorWA.toString().substring(0, 4) + '****';
            }
            
            return {
              status: 'success',
              data: {
                nama: nama,
                email: emailMasked,
                emailClient: emailClient,
                nomorWA: nomorWAMasked,
                kitNumber: kitNumber, // ✅ NEW
                serialNumber: serialNumber, // ✅ NEW
                jatuhTempo: jatuhTempoFormatted,
                transactionID: transactionID,
                status: status,
                paket: paket,
                noRegister: noRegister,
                statusClient: sheetName,
                rowNumber: i + 1
              }
            };
          }
        }
      }
    }

    return { status: 'not_found', message: 'Nomor kit tidak ditemukan dalam database' };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByKitAdmin(kitNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      for (var i = 1; i < data.length; i++) {
        var kitCell = data[i][COLUMNS.KIT_NUMBER]; // ✅ NEW: Kolom I (KIT Number)
        
        var kitString = kitCell.toString().trim();
        var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];
        
        for (var j = 0; j < kitArray.length; j++) {
          if (kitArray[j].trim().toLowerCase() === kitNumber.toLowerCase()) {
            return buildResponse(data, i, headerRow, sheetName);
          }
        }
      }
    }
    
    return { status: 'not_found', message: 'Nomor kit tidak ditemukan dalam database' };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByKitAdminMultiple(kitNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    var headers = {};
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      if (Object.keys(headers).length === 0) {
        headers = buildHeadersObject(headerRow);
      }
      
      for (var i = 1; i < data.length; i++) {
        var kitCell = data[i][COLUMNS.KIT_NUMBER]; // ✅ NEW: Kolom I (KIT Number)
        
        var kitString = kitCell.toString().trim();
        var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];
        
        for (var j = 0; j < kitArray.length; j++) {
          if (kitArray[j].trim().toLowerCase() === kitNumber.toLowerCase()) {
            var rowData = buildRowData(data, i, sheetName);
            allResults.push(rowData);
            break;
          }
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Nomor kit tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults,
      headers: headers
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByKitMultiple(kitNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var kitCell = data[i][COLUMNS.KIT_NUMBER]; // ✅ NEW: Kolom I (KIT Number)
        
        var kitString = kitCell.toString().trim();
        var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];
        
        for (var j = 0; j < kitArray.length; j++) {
          if (kitArray[j].trim().toLowerCase() === kitNumber.toLowerCase()) {
            var maskedData = buildMaskedData(data, i, sheetName);
            allResults.push(maskedData);
            break;
          }
        }
      }
    }

    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Nomor kit tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 5.2 ✅ NEW: SERIAL NUMBER SEARCH FUNCTIONS
function searchBySerialNumber(serialNumber) {
  try {
    Logger.log('🔍 Searching by Serial Number: ' + serialNumber);
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) continue;
      
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      for (var i = 1; i < data.length; i++) {
        var serialCell = data[i][COLUMNS.SERIAL_NUMBER]; // J: Serial Number
        
        if (!serialCell) continue;
        
        var serialString = serialCell.toString().trim();
        var serialArray = serialString.includes('\n') ? serialString.split('\n') : [serialString];
        
        for (var j = 0; j < serialArray.length; j++) {
          if (serialArray[j].trim().toLowerCase() === serialNumber.toLowerCase()) {
            Logger.log('✅ Serial Number found in ' + sheetName + ' at row ' + (i + 1));
            return buildResponse(data, i, headerRow, sheetName);
          }
        }
      }
    }
    
    Logger.log('❌ Serial Number not found: ' + serialNumber);
    return { status: 'not_found', message: 'Serial Number tidak ditemukan dalam database' };
    
  } catch (error) {
    Logger.log('❌ Error in searchBySerialNumber: ' + error.message);
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchBySerialNumberMultiple(serialNumber) {
  try {
    Logger.log('🔍 Searching by Serial Number (Multiple): ' + serialNumber);
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    var headers = {};
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) continue;
      
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      if (Object.keys(headers).length === 0) {
        headers = buildHeadersObject(headerRow);
      }
      
      for (var i = 1; i < data.length; i++) {
        var serialCell = data[i][COLUMNS.SERIAL_NUMBER]; // J: Serial Number
        
        if (!serialCell) continue;
        
        var serialString = serialCell.toString().trim();
        var serialArray = serialString.includes('\n') ? serialString.split('\n') : [serialString];
        
        for (var j = 0; j < serialArray.length; j++) {
          if (serialArray[j].trim().toLowerCase() === serialNumber.toLowerCase()) {
            var rowData = buildRowData(data, i, sheetName);
            allResults.push(rowData);
            Logger.log('✅ Serial Number found in ' + sheetName + ' at row ' + (i + 1));
            break;
          }
        }
      }
    }
    
    if (allResults.length === 0) {
      Logger.log('❌ Serial Number not found: ' + serialNumber);
      return { status: 'not_found', message: 'Serial Number tidak ditemukan dalam database' };
    }
    
    Logger.log('✅ Serial Number search completed: ' + allResults.length + ' results found');
    return {
      status: 'success',
      data: allResults,
      headers: headers
    };
    
  } catch (error) {
    Logger.log('❌ Error in searchBySerialNumberMultiple: ' + error.message);
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 5.3 EMAIL SEARCH FUNCTIONS
function searchByEmailClient(emailClient) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      for (var i = 1; i < data.length; i++) {
        var emailCell = data[i][COLUMNS.EMAIL_CLIENT]; // Kolom F: Email Client
        
        if (emailCell && emailCell.toString().toLowerCase().includes(emailClient.toLowerCase())) {
          return buildResponse(data, i, headerRow, sheetName);
        }
      }
    }
    
    return { status: 'not_found', message: 'Email client tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByEmailClientMultiple(emailClient) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    var headers = {};
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      if (Object.keys(headers).length === 0) {
        headers = buildHeadersObject(headerRow);
      }
      
      for (var i = 1; i < data.length; i++) {
        var emailCell = data[i][COLUMNS.EMAIL_CLIENT]; // Kolom F: Email Client
        
        if (emailCell && emailCell.toString().toLowerCase().includes(emailClient.toLowerCase())) {
          var rowData = buildRowData(data, i, sheetName);
          allResults.push(rowData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Email client tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults,
      headers: headers
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 5.4 PHONE SEARCH FUNCTIONS
function searchByNomorCS(nomorCS) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      for (var i = 1; i < data.length; i++) {
        var phoneCell = data[i][COLUMNS.NOMOR_CS]; // Kolom G: Nomor CS
        
        var phoneString = phoneCell.toString().replace(/\D/g, '');
        var searchString = nomorCS.replace(/\D/g, '');
        
        if (phoneString.includes(searchString)) {
          return buildResponse(data, i, headerRow, sheetName);
        }
      }
    }
    
    return { status: 'not_found', message: 'Nomor CS tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByNomorCSMultiple(nomorCS) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    var headers = {};
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      if (Object.keys(headers).length === 0) {
        headers = buildHeadersObject(headerRow);
      }
      
      for (var i = 1; i < data.length; i++) {
        var phoneCell = data[i][COLUMNS.NOMOR_CS]; // Kolom G: Nomor CS
        
        var phoneString = phoneCell.toString().replace(/\D/g, '');
        var searchString = nomorCS.replace(/\D/g, '');
        
        if (phoneString.includes(searchString)) {
          var rowData = buildRowData(data, i, sheetName);
          allResults.push(rowData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Nomor CS tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults,
      headers: headers
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 5.5 LOGIN STARLINK SEARCH FUNCTIONS
function searchByLoginStarlink(loginStarlink) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      for (var i = 1; i < data.length; i++) {
        var loginCell = data[i][COLUMNS.LOGIN_STARLINK]; // Kolom C: Login Starlink
        
        if (loginCell && loginCell.toString().toLowerCase().includes(loginStarlink.toLowerCase())) {
          return buildResponse(data, i, headerRow, sheetName);
        }
      }
    }
    
    return { status: 'not_found', message: 'Login Starlink tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByLoginStarlinkMultiple(loginStarlink) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    var headers = {};
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      if (Object.keys(headers).length === 0) {
        headers = buildHeadersObject(headerRow);
      }
      
      for (var i = 1; i < data.length; i++) {
        var loginCell = data[i][COLUMNS.LOGIN_STARLINK]; // Kolom C: Login Starlink
        
        if (loginCell && loginCell.toString().toLowerCase().includes(loginStarlink.toLowerCase())) {
          var rowData = buildRowData(data, i, sheetName);
          allResults.push(rowData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Login Starlink tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults,
      headers: headers
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 5.6 ACCOUNT NUMBER SEARCH FUNCTIONS
function searchByAccountNumber(accountNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      for (var i = 1; i < data.length; i++) {
        var accountCell = data[i][COLUMNS.ACC_NO]; // Kolom E: Account Number
        
        if (accountCell && accountCell.toString().toLowerCase().includes(accountNumber.toLowerCase())) {
          return buildResponse(data, i, headerRow, sheetName);
        }
      }
    }
    
    return { status: 'not_found', message: 'Account Number tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByAccountNumberMultiple(accountNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    var headers = {};
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      if (Object.keys(headers).length === 0) {
        headers = buildHeadersObject(headerRow);
      }
      
      for (var i = 1; i < data.length; i++) {
        var accountCell = data[i][COLUMNS.ACC_NO]; // Kolom E: Account Number
        
        if (accountCell && accountCell.toString().toLowerCase().includes(accountNumber.toLowerCase())) {
          var rowData = buildRowData(data, i, sheetName);
          allResults.push(rowData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Account Number tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults,
      headers: headers
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 5.7 NAME SEARCH FUNCTIONS
function searchByNama(nama) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      for (var i = 1; i < data.length; i++) {
        var namaCell = data[i][COLUMNS.NAMA]; // Kolom A: Nama
        
        if (namaCell && namaCell.toString().toLowerCase().includes(nama.toLowerCase())) {
          return buildResponse(data, i, headerRow, sheetName);
        }
      }
    }
    
    return { status: 'not_found', message: 'Nama tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByNamaMultiple(nama) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    var headers = {};
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      var headerRow = data[0];
      
      if (Object.keys(headers).length === 0) {
        headers = buildHeadersObject(headerRow);
      }
      
      for (var i = 1; i < data.length; i++) {
        var namaCell = data[i][COLUMNS.NAMA]; // Kolom A: Nama
        
        if (namaCell && namaCell.toString().toLowerCase().includes(nama.toLowerCase())) {
          var rowData = buildRowData(data, i, sheetName);
          allResults.push(rowData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Nama tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults,
      headers: headers
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

