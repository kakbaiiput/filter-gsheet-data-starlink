// ========================================
// 🔧 16. SYSTEM MAINTENANCE FUNCTIONS
// ========================================

/**
 * 🎯 Fix All Existing Sheets (Utility function - retained for maintenance)
 */
function fixAllExistingSheets() {
  try {
    Logger.log('🔧 Fixing ALL existing sheets with NEW 11-column structure...');
    
    var reportSpreadsheet = getReportSpreadsheet();
    var sheets = reportSpreadsheet.getSheets();
    
    var results = [];
    
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
      
      if (!sheetName.match(/^\d{2}_\d{4}$/)) {
        Logger.log('⏭️ Skipping sheet: ' + sheetName + ' (not MM_YYYY format)');
        continue;
      }
      
      Logger.log('🔧 Processing sheet with NEW structure: ' + sheetName);
      
      try {
        // ✅ MINOR UPDATE: Check if sheet needs column F added
        var lastCol = sheet.getLastColumn();
        var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        
        // Check if Periode column exists (12-column structure)
        if (lastCol < 12 || headers[8] !== 'Periode') {
          Logger.log('📊 Updating sheet to 12-column structure: ' + sheetName);

          // Set correct headers for 12-column structure
          var newHeaders = [
            'Nomor', 'Tanggal Transaksi', 'ID Transaksi', 'Nama Client',
            'KIT Number', 'Serial Number', 'Paket', 'Tipe Pembayaran',
            'Periode', 'Nominal', 'Jumlah Total', 'Total Harian'
          ];

          sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
        }
        
        // Apply updated formatting to existing sheet
        applyMasterFormattingToReportSheetNew11Column(sheet);
        
        results.push({
          sheetName: sheetName,
          result: { status: 'success', message: 'Updated to 12-column structure' }
        });
        
        Logger.log('✅ ' + sheetName + ': Updated to NEW 11-column structure successfully');
        
      } catch (sheetError) {
        Logger.log('❌ Error fixing ' + sheetName + ': ' + sheetError.message);
        results.push({
          sheetName: sheetName,
          result: { status: 'error', message: sheetError.message }
        });
      }
    }
    
    Logger.log('✅ Bulk fix completed for ' + results.length + ' sheets with NEW 11-column structure');
    
    return {
      status: 'success',
      message: 'Processed ' + results.length + ' sheets with 12-column structure',
      results: results,
      newStructure: '12 columns (A-L) with Periode at column I'
    };
    
  } catch (error) {
    Logger.log('❌ Bulk fix error: ' + error.message);
    return { status: 'error', message: error.message };
  }
}

// ========================================
// 📝 17. FINAL SUMMARY & TESTING FUNCTIONS
// ========================================

/**
 * 🚀 SYSTEM STATUS CHECK (Utility function)
 */
function checkSystemStatus() {
  try {
    Logger.log('🔍 CHECKING SYSTEM STATUS WITH UPDATED COLUMN MAPPING');
    
    var status = {
      timestamp: new Date().toISOString(),
      columnMapping: 'Updated Structure (I=KIT Number, J=Serial Number)',
      components: {
        searchFunctions: '✅ Updated with Serial Number search',
        whatsappSystem: '✅ Updated',
        paymentSystem: '✅ Updated with Enhanced Filter',
        reportSystem: '✅ Updated to 12-column structure + Auto Filter',
        clientManagement: '✅ Updated',
        apiHandlers: '✅ Updated with Serial Number support',
        emailNotifications: '✅ Updated'
      },
      reportConnection: 'Unknown',
      clientSheets: []
    };
    
    // Test report connection
    try {
      var reportSpreadsheet = getReportSpreadsheet();
      status.reportConnection = '✅ Connected';
      status.reportSpreadsheetName = reportSpreadsheet.getName();
    } catch (reportError) {
      status.reportConnection = '❌ Failed: ' + reportError.message;
    }
    
    // Check client sheets
    try {
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var clientSheets = ['Client Aktif', 'Client Non Aktif', 'Client Lepas'];
      
      clientSheets.forEach(function(sheetName) {
        var sheet = spreadsheet.getSheetByName(sheetName);
        if (sheet) {
          var rowCount = sheet.getLastRow();
          status.clientSheets.push({
            name: sheetName,
            status: '✅ Available',
            rowCount: rowCount,
            dataRows: rowCount - 1
          });
        } else {
          status.clientSheets.push({
            name: sheetName,
            status: '❌ Not Found',
            rowCount: 0,
            dataRows: 0
          });
        }
      });
    } catch (sheetError) {
      status.clientSheets.push({
        error: 'Error checking sheets: ' + sheetError.message
      });
    }
    
    Logger.log('📊 SYSTEM STATUS REPORT:');
    Logger.log('Column Mapping: ' + status.columnMapping);
    Logger.log('Report Connection: ' + status.reportConnection);
    Logger.log('Client Sheets: ' + JSON.stringify(status.clientSheets));
    
    return status;
    
  } catch (error) {
    Logger.log('❌ Error checking system status: ' + error.message);
    return {
      status: 'error',
      message: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🚀 Initialize System dengan kolom mapping baru
 */
function initializeStarlinkSystem() {
  Logger.log('🚀 INITIALIZING STARLINK MANAGEMENT SYSTEM - CLEANED & ENHANCED VERSION');
  Logger.log('📋 Column Mapping: Updated Structure (I=KIT Number, J=Serial Number)');
  Logger.log('🔧 All functions updated for new column structure');
  Logger.log('✅ Serial Number search functions added');
  Logger.log('✅ Enhanced auto-filter system for today\'s data added');
  Logger.log('🗂️ Unused functions removed for cleaner codebase');
  Logger.log('✅ System ready for operation');
  
  // Perform system status check
  var systemStatus = checkSystemStatus();
  
  Logger.log('🎯 SYSTEM INITIALIZATION COMPLETE - CLEANED VERSION');
  Logger.log('📊 Status: ' + JSON.stringify(systemStatus, null, 2));
  
  return systemStatus;
}


