// ========================================
// ✅ 11. FORM VALIDATION FUNCTIONS
// ========================================

/**
 * 🔄 UPDATED: Handle Form Validation Multi Kit dengan kolom mapping baru
 */
function handleFormValidationMultiKit(kitNumber, selectedMonth, selectedYear) {
  var validation = validateKitForFormMultiKit(kitNumber);
  
  var duplicate = { status: 'no_duplicate', hasDuplicate: false, duplicateData: [], duplicateKits: [] };
  
  if (validation.status === 'found' && validation.data && validation.data.allKits) {
    var allKitNumbers = validation.data.allKits.map(function(kit) { 
      return kit.kitNumber; 
    });
    
    var month = selectedMonth;
    var year = selectedYear;
    
    if (!month || !year) {
      var currentDate = new Date();
      month = currentDate.getMonth() + 1;
      year = currentDate.getFullYear();
      Logger.log('⚠️ No month/year parameter, using current: ' + month + '/' + year);
    } else {
      Logger.log('✅ Using selected month/year: ' + month + '/' + year);
    }
    
    duplicate = checkDuplicateInReportMultiKit(allKitNumbers, month, year);
  }
  
  return {
    validation: validation,
    duplicate: duplicate,
    selectedMonth: month, 
    selectedYear: year,   
    timestamp: new Date().toISOString(),
    targetFile: REPORT_SPREADSHEET_NAME
  };
}

/**
 * 🔄 UPDATED: Validate Client By Name dengan kolom mapping baru
 */
function validateClientByName(clientName, selectedMonth, selectedYear) {
  try {
    Logger.log('=== validateClientByName START (UPDATED MAPPING) ===');
    Logger.log('Client name: ' + clientName);
    
    const sheetNames = ["Client Aktif", "Client Non Aktif", "Client Lepas"];
    
    for (let sheetName of sheetNames) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) continue;
      
      var data = sheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var namaCell = data[i][COLUMNS.NAMA]; // A: Nama
        
        if (namaCell && namaCell.toString().trim().toLowerCase() === clientName.toLowerCase()) {
          Logger.log('✅ Client found: ' + namaCell);
          
          var nama = data[i][COLUMNS.NAMA] || '';
          var kitCell = data[i][COLUMNS.KIT_NUMBER] || ''; // I: KIT Number ✅ NEW
          var serialCell = data[i][COLUMNS.SERIAL_NUMBER] || ''; // J: Serial Number ✅ NEW
          var paketCell = data[i][COLUMNS.PAKET] || ''; // P: Paket ✅ MOVED
          
          var kitString = kitCell.toString().trim();
          var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];
          
          var serialString = serialCell.toString().trim();
          var serialArray = serialString.includes('\n') ? serialString.split('\n') : [serialString];
          
          var paketString = paketCell.toString().trim();
          var paketArray = paketString.includes('\n') ? paketString.split('\n') : [paketString];
          
          var allKits = [];
          for (var k = 0; k < kitArray.length; k++) {
            var kit = kitArray[k].trim();
            var serial = serialArray[k] || serialArray[0] || '';
            var paket = paketArray[k] || paketArray[0] || '';
            
            if (kit) {
              allKits.push({
                kitNumber: kit,
                serialNumber: serial.trim(), // ✅ NEW
                paket: paket.trim(),
                isSelected: false
              });
            }
          }
          
          Logger.log('Found ' + allKits.length + ' KITs for client');
          
          var duplicate = { 
            status: 'no_duplicate', 
            hasDuplicate: false, 
            duplicateData: [], 
            duplicateKits: [] 
          };
          
          if (allKits.length > 0) {
            var allKitNumbers = allKits.map(function(kit) { 
              return kit.kitNumber; 
            });
            
            var month = selectedMonth;
            var year = selectedYear;
            
            if (!month || !year || isNaN(month) || isNaN(year)) {
              var currentDate = new Date();
              month = currentDate.getMonth() + 1;
              year = currentDate.getFullYear();
              Logger.log('⚠️ Invalid month/year, using current: ' + month + '/' + year);
            } else {
              Logger.log('✅ Using provided month/year: ' + month + '/' + year);
            }
            
            duplicate = checkDuplicateInReportMultiKit(allKitNumbers, month, year);
          }
          
          var result = {
            validation: {
              status: 'found',
              data: {
                nama: nama.toString().trim(),
                allKits: allKits,
                totalKits: allKits.length,
                rowNumber: i + 1,
                sheetName: sheetName
              }
            },
            duplicate: duplicate,
            selectedMonth: month,
            selectedYear: year,
            timestamp: new Date().toISOString(),
            targetFile: REPORT_SPREADSHEET_NAME,
            warning: sheetName !== "Client Aktif" ? 
              "⚠️ PERHATIAN: Client tidak berada di 'Client Aktif'." : null
          };
          
          return result;
        }
      }
    }
    
    Logger.log('❌ Client not found: ' + clientName);
    return {
      validation: {
        status: 'not_found',
        data: null
      },
      duplicate: { 
        status: 'no_duplicate', 
        hasDuplicate: false, 
        duplicateData: [], 
        duplicateKits: [] 
      },
      selectedMonth: selectedMonth,
      selectedYear: selectedYear,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('❌ Error in validateClientByName: ' + error.message);
    
    return {
      validation: {
        status: 'error',
        data: null,
        warning: "❌ Error dalam validasi: " + error.message
      },
      duplicate: { 
        status: 'no_duplicate', 
        hasDuplicate: false, 
        duplicateData: [], 
        duplicateKits: [] 
      },
      selectedMonth: selectedMonth,
      selectedYear: selectedYear,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🔧 Check Duplicates in Month
 */
function checkDuplicatesInMonth(month, year) {
  try {
    Logger.log(`Checking duplicates for ${month}/${year} with NEW 11-column structure...`);
    
    var sheet = getReportSheet(month, year, false);
    
    if (!sheet) {
      return {
        status: 'success',
        hasDuplicates: false,
        message: 'Aman, Tidak ada duplikasi',
        details: `Sheet ${month.toString().padStart(2, '0')}_${year} tidak ditemukan atau kosong.`
      };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        status: 'success',
        hasDuplicates: false,
        message: 'Aman, Tidak ada duplikasi',
        details: `Sheet ${sheet.getName()} kosong atau hanya ada header.`
      };
    }
    
    var kitCounts = {};
    var kitDetails = {};
    
    for (var i = 1; i < data.length; i++) {
      var reportKitCell = data[i][4]; // ✅ UNCHANGED: Kolom E: Serial Number (KIT) masih sama
      
      if (reportKitCell) {
        var kitString = reportKitCell.toString().trim();
        var kitArray = kitString.includes('\n') ? kitString.split('\n') : [kitString];
        
        for (var j = 0; j < kitArray.length; j++) {
          var kit = kitArray[j].trim().toUpperCase();
          
          if (kit) {
            kitCounts[kit] = (kitCounts[kit] || 0) + 1;
            
            if (!kitDetails[kit]) {
              kitDetails[kit] = [];
            }
            
            var tanggalRaw = data[i][1] || '';
            var tanggalFormatted = '';
            
            if (tanggalRaw instanceof Date) {
              var day = tanggalRaw.getDate().toString().padStart(2, '0');
              var monthNum = (tanggalRaw.getMonth() + 1).toString().padStart(2, '0');
              var yearNum = tanggalRaw.getFullYear();
              tanggalFormatted = `${day}/${monthNum}/${yearNum}`;
            } else if (tanggalRaw) {
              var tempDate = new Date(tanggalRaw);
              if (!isNaN(tempDate.getTime())) {
                var day = tempDate.getDate().toString().padStart(2, '0');
                var monthNum = (tempDate.getMonth() + 1).toString().padStart(2, '0');
                var yearNum = tempDate.getFullYear();
                tanggalFormatted = `${day}/${monthNum}/${yearNum}`;
              } else {
                tanggalFormatted = tanggalRaw.toString();
              }
            }
            
            // ✅ MINOR UPDATE: Struktur data dengan kolom baru
            kitDetails[kit].push({
              rowNumber: i + 1,
              transactionID: data[i][2] || 'No ID',        // C: Transaction ID
              namaClient: data[i][3] || 'No Name',         // D: Nama Client  
              tanggal: tanggalFormatted,                   
              kitNumber: kit,                              // E: KIT Number
              serialNumber: data[i][5] || '',              // F: Serial Number
              paket: data[i][6] || '',                     // G: Paket
              tipe: data[i][7] || '',                      // H: Tipe Pembayaran
              periode: data[i][8] || '',                   // I: Periode
              nominal: data[i][9] || ''                    // J: Nominal
            });
          }
        }
      }
    }
    
    var duplicates = [];
    
    for (var kit in kitCounts) {
      if (kitCounts[kit] > 1) {
        duplicates.push({
          kitNumber: kit,
          count: kitCounts[kit],
          occurrences: kitDetails[kit]
        });
      }
    }
    
    if (duplicates.length === 0) {
      return {
        status: 'success',
        hasDuplicates: false,
        message: 'Aman, Tidak ada duplikasi',
        sheetName: sheet.getName(),
        totalRows: data.length - 1,
        checkedKits: Object.keys(kitCounts).length,
        reportStructure: 'Updated to 12 columns (A-L)'
      };
    }
    
    return {
      status: 'success',
      hasDuplicates: true,
      message: `Ditemukan ${duplicates.length} KIT yang duplikasi`,
      sheetName: sheet.getName(),
      totalRows: data.length - 1,
      checkedKits: Object.keys(kitCounts).length,
      duplicates: duplicates,
      reportStructure: 'Updated to 12 columns (A-L)'
    };
    
  } catch (error) {
    Logger.log('Error in checkDuplicatesInMonth: ' + error.message);
    return {
      status: 'error',
      hasDuplicates: false,
      message: 'Error saat mengecek duplikasi: ' + error.message
    };
  }
}

/**
 * 🔧 Check Duplicates by Kit Numbers
 */
function checkDuplicatesByKitNumbers(kitNumbers, month, year) {
  try {
    Logger.log('Checking duplicates by KIT numbers:', kitNumbers);
    
    if (typeof kitNumbers === 'string') {
      kitNumbers = kitNumbers.split(',').map(kit => kit.trim());
    }
    
    if (!Array.isArray(kitNumbers) || kitNumbers.length === 0) {
      return {
        status: 'error',
        message: 'KIT numbers tidak valid',
        hasDuplicate: false,
        duplicateData: [],
        duplicateKits: []
      };
    }
    
    return checkDuplicateInReportMultiKit(kitNumbers, month, year);
    
  } catch (error) {
    Logger.log('Error in checkDuplicatesByKitNumbers: ' + error.message);
    return {
      status: 'error',
      message: 'Error checking duplicates: ' + error.message,
      hasDuplicate: false,
      duplicateData: [],
      duplicateKits: []
    };
  }
}

/**
 * Setup time-based trigger to cleanup expired requestId cache
 * Run this once manually to setup the trigger
 */
function setupCleanupTrigger() {
  // Delete existing triggers first
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cleanupExpiredRequestIds') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger - runs every hour
  ScriptApp.newTrigger('cleanupExpiredRequestIds')
    .timeBased()
    .everyHours(1)
    .create();
  
  Logger.log('✅ Cleanup trigger created - will run every hour');
}
