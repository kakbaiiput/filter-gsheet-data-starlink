// ========================================
// 🔍 6. SEARCH FUNCTIONS - CUSTOMER VERSION (with data masking)
// ========================================

// 6.1 CUSTOMER EMAIL SEARCH
function searchByEmailClientCustomer(emailClient) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var emailCell = data[i][COLUMNS.EMAIL_CLIENT]; // Kolom F: Email Client
        
        if (emailCell && emailCell.toString().toLowerCase().includes(emailClient.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          return {
            status: 'success',
            data: maskedData
          };
        }
      }
    }
    
    return { status: 'not_found', message: 'Email client tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByEmailClientCustomerMultiple(emailClient) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var emailCell = data[i][COLUMNS.EMAIL_CLIENT]; // Kolom F: Email Client
        
        if (emailCell && emailCell.toString().toLowerCase().includes(emailClient.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          allResults.push(maskedData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Email client tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 6.2 ✅ NEW: CUSTOMER SERIAL NUMBER SEARCH
function searchBySerialNumberCustomer(serialNumber) {
  try {
    Logger.log('🔍 Searching by Serial Number (Customer): ' + serialNumber);
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) continue;
      
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var serialCell = data[i][COLUMNS.SERIAL_NUMBER]; // J: Serial Number
        
        if (!serialCell) continue;
        
        var serialString = serialCell.toString().trim();
        var serialArray = serialString.includes('\n') ? serialString.split('\n') : [serialString];
        
        for (var j = 0; j < serialArray.length; j++) {
          if (serialArray[j].trim().toLowerCase() === serialNumber.toLowerCase()) {
            var maskedData = buildMaskedData(data, i, sheetName);
            Logger.log('✅ Serial Number found (Customer): ' + serialNumber);
            return {
              status: 'success',
              data: maskedData
            };
          }
        }
      }
    }
    
    Logger.log('❌ Serial Number not found (Customer): ' + serialNumber);
    return { status: 'not_found', message: 'Serial Number tidak ditemukan dalam database' };
    
  } catch (error) {
    Logger.log('❌ Error in searchBySerialNumberCustomer: ' + error.message);
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchBySerialNumberCustomerMultiple(serialNumber) {
  try {
    Logger.log('🔍 Searching by Serial Number (Customer Multiple): ' + serialNumber);
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) continue;
      
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var serialCell = data[i][COLUMNS.SERIAL_NUMBER]; // J: Serial Number
        
        if (!serialCell) continue;
        
        var serialString = serialCell.toString().trim();
        var serialArray = serialString.includes('\n') ? serialString.split('\n') : [serialString];
        
        for (var j = 0; j < serialArray.length; j++) {
          if (serialArray[j].trim().toLowerCase() === serialNumber.toLowerCase()) {
            var maskedData = buildMaskedData(data, i, sheetName);
            allResults.push(maskedData);
            Logger.log('✅ Serial Number found (Customer): ' + serialNumber);
            break;
          }
        }
      }
    }
    
    if (allResults.length === 0) {
      Logger.log('❌ Serial Number not found (Customer): ' + serialNumber);
      return { status: 'not_found', message: 'Serial Number tidak ditemukan dalam database' };
    }
    
    Logger.log('✅ Serial Number search completed (Customer): ' + allResults.length + ' results found');
    return {
      status: 'success',
      data: allResults
    };
    
  } catch (error) {
    Logger.log('❌ Error in searchBySerialNumberCustomerMultiple: ' + error.message);
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 6.3 CUSTOMER PHONE SEARCH
function searchByNomorCSCustomer(nomorCS) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var phoneCell = data[i][COLUMNS.NOMOR_CS]; // Kolom G: Nomor CS
        
        var phoneString = phoneCell.toString().replace(/\D/g, '');
        var searchString = nomorCS.replace(/\D/g, '');
        
        if (phoneString.includes(searchString)) {
          var maskedData = buildMaskedData(data, i, sheetName);
          return {
            status: 'success',
            data: maskedData
          };
        }
      }
    }
    
    return { status: 'not_found', message: 'Nomor CS tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByNomorCSCustomerMultiple(nomorCS) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var phoneCell = data[i][COLUMNS.NOMOR_CS]; // Kolom G: Nomor CS
        
        var phoneString = phoneCell.toString().replace(/\D/g, '');
        var searchString = nomorCS.replace(/\D/g, '');
        
        if (phoneString.includes(searchString)) {
          var maskedData = buildMaskedData(data, i, sheetName);
          allResults.push(maskedData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Nomor CS tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 6.4 CUSTOMER LOGIN SEARCH
function searchByLoginStarlinkCustomer(loginStarlink) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var loginCell = data[i][COLUMNS.LOGIN_STARLINK]; // Kolom C: Login Starlink
        
        if (loginCell && loginCell.toString().toLowerCase().includes(loginStarlink.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          return {
            status: 'success',
            data: maskedData
          };
        }
      }
    }
    
    return { status: 'not_found', message: 'Login Starlink tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByLoginStarlinkCustomerMultiple(loginStarlink) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var loginCell = data[i][COLUMNS.LOGIN_STARLINK]; // Kolom C: Login Starlink
        
        if (loginCell && loginCell.toString().toLowerCase().includes(loginStarlink.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          allResults.push(maskedData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Login Starlink tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 6.5 CUSTOMER ACCOUNT SEARCH
function searchByAccountNumberCustomer(accountNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var accountCell = data[i][COLUMNS.ACC_NO]; // Kolom E: Account Number
        
        if (accountCell && accountCell.toString().toLowerCase().includes(accountNumber.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          return {
            status: 'success',
            data: maskedData
          };
        }
      }
    }
    
    return { status: 'not_found', message: 'Account Number tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByAccountNumberCustomerMultiple(accountNumber) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var accountCell = data[i][COLUMNS.ACC_NO]; // Kolom E: Account Number
        
        if (accountCell && accountCell.toString().toLowerCase().includes(accountNumber.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          allResults.push(maskedData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Account Number tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// 6.6 CUSTOMER NAME SEARCH
function searchByNamaCustomer(nama) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var namaCell = data[i][COLUMNS.NAMA]; // Kolom A: Nama
        
        if (namaCell && namaCell.toString().toLowerCase().includes(nama.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          return {
            status: 'success',
            data: maskedData
          };
        }
      }
    }
    
    return { status: 'not_found', message: 'Nama tidak ditemukan dalam database' };
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

function searchByNamaCustomerMultiple(nama) {
  try {
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];    
    var allResults = [];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var namaCell = data[i][COLUMNS.NAMA]; // Kolom A: Nama
        
        if (namaCell && namaCell.toString().toLowerCase().includes(nama.toLowerCase())) {
          var maskedData = buildMaskedData(data, i, sheetName);
          allResults.push(maskedData);
        }
      }
    }
    
    if (allResults.length === 0) {
      return { status: 'not_found', message: 'Nama tidak ditemukan dalam database' };
    }
    
    return {
      status: 'success',
      data: allResults
    };
    
  } catch (error) {
    return { status: 'error', message: 'Error dalam pencarian: ' + error.message };
  }
}

// ========================================
// 🔍 7. UNIVERSAL SEARCH FUNCTION
// ========================================

function universalSearch(searchTerm) {
  try {
    Logger.log('Universal search for: ' + searchTerm);
    
    var allResults = [];
    var searchSources = [];
    
    // Search by Kit
    try {
      var kitResult = searchByKitAdminMultiple(searchTerm);
      if (kitResult.status === 'success') {
        allResults = allResults.concat(kitResult.data.map(item => ({
          ...item,
          searchSource: 'KIT Number' // ✅ NEW
        })));
        searchSources.push('KIT Number');
      }
    } catch (e) {
      Logger.log('Kit search error: ' + e.message);
    }
    
    // ✅ NEW: Search by Serial Number
    try {
      var serialResult = searchBySerialNumberMultiple(searchTerm);
      if (serialResult.status === 'success') {
        allResults = allResults.concat(serialResult.data.map(item => ({
          ...item,
          searchSource: 'Serial Number'
        })));
        searchSources.push('Serial Number');
      }
    } catch (e) {
      Logger.log('Serial search error: ' + e.message);
    }
    
    // Search by Login
    try {
      var loginResult = searchByLoginStarlinkMultiple(searchTerm);
      if (loginResult.status === 'success') {
        allResults = allResults.concat(loginResult.data.map(item => ({
          ...item,
          searchSource: 'Login Starlink'
        })));
        searchSources.push('Login Starlink');
      }
    } catch (e) {
      Logger.log('Login search error: ' + e.message);
    }
    
    // Search by Email
    try {
      var emailResult = searchByEmailClientMultiple(searchTerm);
      if (emailResult.status === 'success') {
        allResults = allResults.concat(emailResult.data.map(item => ({
          ...item,
          searchSource: 'Email Client'
        })));
        searchSources.push('Email Client');
      }
    } catch (e) {
      Logger.log('Email search error: ' + e.message);
    }
    
    // Search by Phone
    try {
      var phoneResult = searchByNomorCSMultiple(searchTerm);
      if (phoneResult.status === 'success') {
        allResults = allResults.concat(phoneResult.data.map(item => ({
          ...item,
          searchSource: 'Nomor CS'
        })));
        searchSources.push('Nomor CS');
      }
    } catch (e) {
      Logger.log('Phone search error: ' + e.message);
    }
    
    // Search by Account
    try {
      var accountResult = searchByAccountNumberMultiple(searchTerm);
      if (accountResult.status === 'success') {
        allResults = allResults.concat(accountResult.data.map(item => ({
          ...item,
          searchSource: 'Account Number'
        })));
        searchSources.push('Account Number');
      }
    } catch (e) {
      Logger.log('Account search error: ' + e.message);
    }
    
    // Search by Name
    try {
      var nameResult = searchByNamaMultiple(searchTerm);
      if (nameResult.status === 'success') {
        allResults = allResults.concat(nameResult.data.map(item => ({
          ...item,
          searchSource: 'Nama Client'
        })));
        searchSources.push('Nama Client');
      }
    } catch (e) {
      Logger.log('Name search error: ' + e.message);
    }
    
    // Remove duplicates based on row number and sheet
    var uniqueResults = [];
    var seen = new Set();
    
    allResults.forEach(item => {
      var key = item.statusClient + '_' + item.rowNumber;
      if (!seen.has(key)) {
        seen.add(key);
        uniqueResults.push(item);
      }
    });
    
    if (uniqueResults.length === 0) {
      return {
        status: 'not_found',
        message: 'Tidak ada data yang ditemukan untuk: ' + searchTerm,
        searchedIn: ['KIT Number', 'Serial Number', 'Login Starlink', 'Email Client', 'Nomor CS', 'Account Number', 'Nama Client'] // ✅ UPDATED
      };
    }
    
    return {
      status: 'success',
      data: uniqueResults,
      totalResults: uniqueResults.length,
      searchSources: searchSources,
      searchTerm: searchTerm
    };
    
  } catch (error) {
    return {
      status: 'error',
      message: 'Error dalam universal search: ' + error.message
    };
  }
}

