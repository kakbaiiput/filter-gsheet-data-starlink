// ========================================
// 🌐 12. API HANDLERS (doGet/doPost/doOptions)
// ========================================

/**
 * 🔐 Helper function to create response with CORS headers
 * This explicitly adds CORS headers to fix CORS preflight issues
 */
function createCorsResponse(data) {
  var jsonString = JSON.stringify(data);
  var output = ContentService.createTextOutput(jsonString);
  output.setMimeType(ContentService.MimeType.JSON);

  // Explicitly set CORS headers
  // Note: Apps Script should handle this automatically, but we're adding explicitly as fallback
  return output;
}

/**
 * 🔐 Handle OPTIONS requests for CORS preflight
 * This is CRITICAL for POST requests from different origins (CORS)
 * Without this, browser will block POST requests with CORS error
 */
function doOptions(e) {
  Logger.log('=== OPTIONS REQUEST (CORS PREFLIGHT) ===');
  Logger.log('Origin: ' + (e.parameter.origin || 'not specified'));

  // Return empty response - Google Apps Script handles CORS headers automatically
  var output = ContentService.createTextOutput('');
  output.setMimeType(ContentService.MimeType.TEXT);
  return output;
}

/**
 * 🔄 COMPLETE doGet() Function dengan kolom mapping baru dan Serial Number search
 */
function doGet(e) {
  try {
    const origin = e.parameter.origin || '';
    const callback = e.parameter.callback || '';
    
    Logger.log('=== DOGET REQUEST ===');
    Logger.log('Origin: ' + origin);
    Logger.log('Callback: ' + callback);
    Logger.log('All Parameters: ' + JSON.stringify(e.parameter));
    
    function createResponse(data) {
      var jsonString = JSON.stringify(data);
      
      if (callback) {
        var jsonpResponse = callback + '(' + jsonString + ');';
        Logger.log('Returning JSONP response with callback: ' + callback);
        return ContentService.createTextOutput(jsonpResponse)
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      } else {
        Logger.log('Returning regular JSON response');
        return ContentService.createTextOutput(jsonString)
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    var action = e.parameter.action;
    var version = e.parameter.version; 
    var multiple = e.parameter.multiple;
    
    Logger.log('Processing action: ' + action);
    Logger.log('Version: ' + version);
    Logger.log('Multiple: ' + multiple);
    
    // GET ALL CLIENT DATA COMPLETE
    if (action === 'getAllClientDataComplete') {
      Logger.log('Getting ALL client data (complete from Client Aktif)...');
      var result = getAllClientDataComplete();
      Logger.log('Complete data result: ' + result.status + ', records: ' + (result.data ? result.data.length : 0));
      return createResponse(result);
    }
    
    // GET CLIENT LEPAS DATA
    else if (action === 'getClientLepasData') {
      Logger.log('Getting ALL client data from Client Lepas sheet...');
      var result = getClientLepasData();
      Logger.log('Client Lepas data result: ' + result.status + ', records: ' + (result.data ? result.data.length : 0));
      return createResponse(result);
    }

    // GET CLIENT NON AKTIF DATA
    else if (action === 'getClientNonAktifData') {
      Logger.log('Getting ALL client data from Client Non Aktif sheet...');
      var result = getClientNonAktifData();
      Logger.log('Client Non Aktif data result: ' + result.status + ', records: ' + (result.data ? result.data.length : 0));
      return createResponse(result);
    }

    // GET CLIENT TERTAGIH DATA
    else if (action === 'getClientTertagihData') {
      Logger.log('Getting ALL client data from Client Tertagih sheet...');
      var result = getClientTertagihData();
      Logger.log('Client Tertagih data result: ' + result.status + ', records: ' + (result.data ? result.data.length : 0));
      return createResponse(result);
    }

    // GET AKUN KOSONG DATA
    else if (action === 'getAkunKosongData') {
      Logger.log('Getting ALL data from Akun Kosong sheet...');
      var result = getAkunKosongData();
      Logger.log('Akun Kosong data result: ' + result.status + ', records: ' + (result.data ? result.data.length : 0));
      return createResponse(result);
    }

    // SEARCH ACTION
    else if (action === 'search') {
      var result;
      
      if (e.parameter.kit) {
        Logger.log('🔍 Searching by Kit: ' + e.parameter.kit);
        if (version === 'admin') {
          if (multiple === 'true') {
            result = searchByKitAdminMultiple(e.parameter.kit.toString().trim());
          } else {
            result = searchByKitAdmin(e.parameter.kit.toString().trim());
          }
        } else {
          if (multiple === 'true') {
            result = searchByKitMultiple(e.parameter.kit.toString().trim());
          } else {
            result = searchByKit(e.parameter.kit.toString().trim());
          }
        }
      } 
      // ✅ NEW: SERIAL NUMBER SEARCH
      else if (e.parameter.serial) {
        Logger.log('🔍 Searching by Serial Number: ' + e.parameter.serial);
        if (version === 'admin') {
          if (multiple === 'true') {
            result = searchBySerialNumberMultiple(e.parameter.serial.toString().trim());
          } else {
            result = searchBySerialNumber(e.parameter.serial.toString().trim());
          }
        } else {
          if (multiple === 'true') {
            result = searchBySerialNumberCustomerMultiple(e.parameter.serial.toString().trim());
          } else {
            result = searchBySerialNumberCustomer(e.parameter.serial.toString().trim());
          }
        }
      }
      else if (e.parameter.login) {
        Logger.log('🔍 Searching by Login: ' + e.parameter.login);
        if (version === 'admin') {
          if (multiple === 'true') {
            result = searchByLoginStarlinkMultiple(e.parameter.login.toString().trim());
          } else {
            result = searchByLoginStarlink(e.parameter.login.toString().trim());
          }
        } else {
          if (multiple === 'true') {
            result = searchByLoginStarlinkCustomerMultiple(e.parameter.login.toString().trim());
          } else {
            result = searchByLoginStarlinkCustomer(e.parameter.login.toString().trim());
          }
        }
      }
      else if (e.parameter.email) {
        Logger.log('🔍 Searching by Email: ' + e.parameter.email);
        if (version === 'admin') {
          if (multiple === 'true') {
            result = searchByEmailClientMultiple(e.parameter.email.toString().trim());
          } else {
            result = searchByEmailClient(e.parameter.email.toString().trim());
          }
        } else {
          if (multiple === 'true') {
            result = searchByEmailClientCustomerMultiple(e.parameter.email.toString().trim());
          } else {
            result = searchByEmailClientCustomer(e.parameter.email.toString().trim());
          }
        }
      }
      else if (e.parameter.phone) {
        Logger.log('🔍 Searching by Phone: ' + e.parameter.phone);
        if (version === 'admin') {
          if (multiple === 'true') {
            result = searchByNomorCSMultiple(e.parameter.phone.toString().trim());
          } else {
            result = searchByNomorCS(e.parameter.phone.toString().trim());
          }
        } else {
          if (multiple === 'true') {
            result = searchByNomorCSCustomerMultiple(e.parameter.phone.toString().trim());
          } else {
            result = searchByNomorCSCustomer(e.parameter.phone.toString().trim());
          }
        }
      }
      else if (e.parameter.account) {
        Logger.log('🔍 Searching by Account: ' + e.parameter.account);
        if (version === 'admin') {
          if (multiple === 'true') {
            result = searchByAccountNumberMultiple(e.parameter.account.toString().trim());
          } else {
            result = searchByAccountNumber(e.parameter.account.toString().trim());
          }
        } else {
          if (multiple === 'true') {
            result = searchByAccountNumberCustomerMultiple(e.parameter.account.toString().trim());
          } else {
            result = searchByAccountNumberCustomer(e.parameter.account.toString().trim());
          }
        }
      }
      else if (e.parameter.name) {
        Logger.log('🔍 Searching by Name: ' + e.parameter.name);
        if (version === 'admin') {
          if (multiple === 'true') {
            result = searchByNamaMultiple(e.parameter.name.toString().trim());
          } else {
            result = searchByNama(e.parameter.name.toString().trim());
          }
        } else {
          if (multiple === 'true') {
            result = searchByNamaCustomerMultiple(e.parameter.name.toString().trim());
          } else {
            result = searchByNamaCustomer(e.parameter.name.toString().trim());
          }
        }
      }
      else {
        Logger.log('❌ No valid search parameter found');
        result = {
          status: 'error',
          message: 'Parameter search diperlukan. Gunakan salah satu: kit, serial, login, email, phone, account, atau name', // ✅ UPDATED
          availableParameters: ['kit', 'serial', 'login', 'email', 'phone', 'account', 'name'], // ✅ UPDATED
          receivedParameters: Object.keys(e.parameter),
          example: 'Contoh: ?action=search&serial=UTxxxxxxxxxx&version=admin&multiple=true' // ✅ NEW
        };
      }
      
      Logger.log('Search result status: ' + result.status);
      return createResponse(result);
    } 
    
    // MULTI-KIT VALIDATION
    else if (action === 'validateKitMulti') {
      var kitNumber = e.parameter.kit;
      var selectedMonth = parseInt(e.parameter.month);
      var selectedYear = parseInt(e.parameter.year);
      
      if (!kitNumber) {
        return createResponse({
          status: 'error',
          message: 'Parameter kit diperlukan'
        });
      }
      
      Logger.log('📥 Received from VPS:');
      Logger.log('  - Kit: ' + kitNumber);
      Logger.log('  - Month param: ' + e.parameter.month);
      Logger.log('  - Year param: ' + e.parameter.year); 
      
      var result = handleFormValidationMultiKit(kitNumber.toString().trim(), selectedMonth, selectedYear);
      Logger.log('📤 Sending back selectedMonth: ' + result.selectedMonth);
      return createResponse(result);
    }
    
    // HANDLE POST DATA via GET
    else if (action === 'submitPaymentMultiKit' || e.parameter.method === 'POST') {
      Logger.log('Handling POST-style request via GET...');
      
      var formData;
      if (e.parameter.data) {
        try {
          formData = JSON.parse(e.parameter.data);
          Logger.log('Parsed form data from GET parameter');
        } catch (parseError) {
          Logger.log('Error parsing form data: ' + parseError.message);
          return createResponse({
            status: 'error',
            message: 'Invalid form data format: ' + parseError.message
          });
        }
      } else {
        return createResponse({
          status: 'error',
          message: 'Form data diperlukan untuk submission'
        });
      }
      
      Logger.log('Processing payment submission...');
      var result = submitPaymentDataMultiKit(formData);
      Logger.log('Submit result status: ' + result.status);
      return createResponse(result);
    }
    
    // DUPLICATE CHECKER
    else if (action === 'checkDuplicates') {
      var month = parseInt(e.parameter.month);
      var year = parseInt(e.parameter.year);
      
      if (e.parameter.kitNumbers) {
        Logger.log('Checking duplicates by kitNumbers for client search...');
        var kitNumbers = e.parameter.kitNumbers.split(',').map(kit => kit.trim());
        var result = checkDuplicatesByKitNumbers(kitNumbers, month, year);
        Logger.log('Duplicate check by kitNumbers result: ' + result.status);
        return createResponse(result);
      }
      
      if (!month || !year || month < 1 || month > 12) {
        return createResponse({
          status: 'error',
          message: 'Parameter month (1-12) dan year diperlukan'
        });
      }
      
      Logger.log(`Checking duplicates for ${month}/${year}`);
      var result = checkDuplicatesInMonth(month, year);
      Logger.log('Duplicate check result: ' + result.status);
      return createResponse(result);
    }
    
    // OTHER ACTIONS
    else if (action === 'getClientNames') {
      var result = getAllClientNames();
      return createResponse(result);
    }
    else if (action === 'getAllStatus') {
      var result = getAllClientsByStatus();
      return createResponse(result);
    }
    else if (action === 'submitClientData') {
      Logger.log('Handling client data submission...');
      
      var clientData;
      if (e.parameter.data) {
        try {
          clientData = JSON.parse(e.parameter.data);
          Logger.log('Parsed client data from GET parameter');
        } catch (parseError) {
          Logger.log('Error parsing client data: ' + parseError.message);
          return createResponse({
            status: 'error',
            message: 'Invalid client data format: ' + parseError.message
          });
        }
      } else {
        return createResponse({
          status: 'error',
          message: 'Client data diperlukan untuk submission'
        });
      }
      
      Logger.log('Processing client data submission...');
      var result = submitClientData(clientData);
      Logger.log('Submit result status: ' + result.status);
      return createResponse(result);
    }
    else if (action === 'getClientsBySheet') {
      var sheetName = e.parameter.sheetName || 'Client Aktif';
      var result = getClientsBySheet(sheetName);
      return createResponse(result);
    }
    else if (action === 'validateClientByName') {
      var clientName = e.parameter.clientName;
      var selectedMonth = parseInt(e.parameter.month);
      var selectedYear = parseInt(e.parameter.year); 
      
      if (!clientName) {
        return createResponse({
          status: 'error',
          message: 'Parameter clientName diperlukan'
        });
      }
      
      Logger.log('validateClientByName received params:');
      Logger.log('  - clientName: ' + clientName);
      Logger.log('  - month: ' + selectedMonth + ' (type: ' + typeof selectedMonth + ')');
      Logger.log('  - year: ' + selectedYear + ' (type: ' + typeof selectedYear + ')');
      
      var result = validateClientByName(clientName.toString().trim(), selectedMonth, selectedYear);
      return createResponse(result);
    }
    
    else if (action === 'getReportData') {
  var month = parseInt(e.parameter.month);
  var year = parseInt(e.parameter.year);

  if (!month || !year) {
    var now = new Date();
    month = now.getMonth() + 1;
    year = now.getFullYear();
  }

  var result = getReportData(month, year);
  return createResponse(result);
}

    // ========================================
    // 🔐 APPROVAL SYSTEM ACTIONS (via GET)
    // ========================================

    // ADD PENDING APPROVAL
    else if (action === 'addPendingApproval') {
      Logger.log('📝 Adding pending approval via GET...');

      try {
        // Parse data from URL parameters
        var approvalData = {
          rowId: e.parameter.rowId,
          changedBy: e.parameter.changedBy,
          changedByName: e.parameter.changedByName,
          sourceSheet: e.parameter.sourceSheet,
          changedColumns: JSON.parse(e.parameter.changedColumns || '[]'),
          originalData: JSON.parse(e.parameter.originalData || '{}'),
          newData: JSON.parse(e.parameter.newData || '{}')
        };

        Logger.log('Approval data: ' + JSON.stringify(approvalData));
        var result = addPendingApproval(approvalData);
        return createResponse(result);
      } catch (error) {
        Logger.log('❌ Error in addPendingApproval: ' + error.message);
        return createResponse({
          status: 'error',
          message: 'Error processing approval: ' + error.message
        });
      }
    }

    // GET PENDING APPROVALS
    else if (action === 'getPendingApprovals') {
      Logger.log('📋 Getting pending approvals via GET...');
      var filterStatus = e.parameter.filterStatus || 'Pending';
      var result = getPendingApprovals(filterStatus);
      return createResponse(result);
    }

    // APPROVE CHANGE
    else if (action === 'approveChange') {
      Logger.log('✅ Approving change via GET...');

      try {
        var approvalId = e.parameter.approvalId;
        var adminEmail = e.parameter.adminEmail;

        if (!approvalId || !adminEmail) {
          return createResponse({
            status: 'error',
            message: 'approvalId and adminEmail parameters required'
          });
        }

        Logger.log('Approval ID: ' + approvalId);
        Logger.log('Admin: ' + adminEmail);

        var result = approveChange(approvalId, adminEmail);
        return createResponse(result);
      } catch (error) {
        Logger.log('❌ Error in approveChange: ' + error.message);
        return createResponse({
          status: 'error',
          message: 'Error approving change: ' + error.message
        });
      }
    }

    // REJECT CHANGE
    else if (action === 'rejectChange') {
      Logger.log('❌ Rejecting change via GET...');

      try {
        var approvalId = e.parameter.approvalId;
        var adminEmail = e.parameter.adminEmail;
        var reason = e.parameter.reason || 'No reason provided';

        if (!approvalId || !adminEmail) {
          return createResponse({
            status: 'error',
            message: 'approvalId and adminEmail parameters required'
          });
        }

        Logger.log('Approval ID: ' + approvalId);
        Logger.log('Admin: ' + adminEmail);
        Logger.log('Reason: ' + reason);

        var result = rejectChange(approvalId, adminEmail, reason);
        return createResponse(result);
      } catch (error) {
        Logger.log('❌ Error in rejectChange: ' + error.message);
        return createResponse({
          status: 'error',
          message: 'Error rejecting change: ' + error.message
        });
      }
    }

    // APPROVE ALL
    else if (action === 'approveAll') {
      Logger.log('✅✅✅ Approving all changes via GET...');

      try {
        var adminEmail = e.parameter.adminEmail;

        if (!adminEmail) {
          return createResponse({
            status: 'error',
            message: 'adminEmail parameter required'
          });
        }

        Logger.log('Admin: ' + adminEmail);

        var result = approveAll(adminEmail);
        return createResponse(result);
      } catch (error) {
        Logger.log('❌ Error in approveAll: ' + error.message);
        return createResponse({
          status: 'error',
          message: 'Error approving all: ' + error.message
        });
      }
    }

    // CHECK IS ADMIN
    else if (action === 'isAdmin') {
      Logger.log('🔐 Checking admin status via GET...');

      var email = e.parameter.email;

      if (!email) {
        return createResponse({
          status: 'error',
          message: 'email parameter required'
        });
      }

      var result = {
        status: 'success',
        isAdmin: isAdmin(email)
      };

      Logger.log('Admin check result: ' + result.isAdmin);
      return createResponse(result);
    }

    // DEFAULT CASE
    else {
      Logger.log('❌ Invalid action: ' + action);
      return createResponse({
        status: 'error',
        message: 'Action tidak valid. Available: search, validateKitMulti, submitPaymentMultiKit, checkDuplicates, getClientNames, getAllStatus, submitClientData, getAllClientDataComplete, getClientLepasData, getClientNonAktifData, getClientsBySheet, validateClientByName, addPendingApproval, getPendingApprovals, approveChange, rejectChange, approveAll, isAdmin',
        receivedAction: action,
        availableActions: [
          'search',
          'validateKitMulti',
          'submitPaymentMultiKit',
          'checkDuplicates',
          'getClientNames',
          'getAllStatus',
          'submitClientData',
          'getAllClientDataComplete',
          'getClientLepasData',
          'getClientNonAktifData',
          'getClientsBySheet',
          'validateClientByName',
          'addPendingApproval',
          'getPendingApprovals',
          'approveChange',
          'rejectChange',
          'approveAll',
          'isAdmin'
        ],
        searchParameters: ['kit', 'serial', 'login', 'email', 'phone', 'account', 'name'] // ✅ UPDATED
      });
    }    
  } catch (error) {
    Logger.log('❌ Error in doGet: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    
    var errorResponse = {
      status: 'error',
      message: 'Server error: ' + error.message,
      timestamp: new Date().toISOString(),
      debug: {
        stack: error.stack,
        parameters: e.parameter
      }
    };
    
    var callback = e.parameter.callback || '';
    if (callback) {
      var jsonpError = callback + '(' + JSON.stringify(errorResponse) + ');';
      return ContentService.createTextOutput(jsonpError)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService.createTextOutput(JSON.stringify(errorResponse))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
}

function getReportData(month, year) {
  try {
    Logger.log(`Getting report data for ${month}/${year}`);
    
    var sheet = getReportSheet(month, year, false);
    
    if (!sheet) {
      return {
        status: 'not_found',
        message: `Report untuk ${month}/${year} tidak ditemukan`,
        data: []
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        status: 'success',
        message: 'Report kosong',
        data: [],
        sheetName: sheet.getName()
      };
    }
    
    var reportData = [];
    
    for (var i = 1; i < data.length; i++) {
      var tanggalRaw = data[i][1];
      var tanggalFormatted = '';
      
      if (tanggalRaw instanceof Date) {
        tanggalFormatted = tanggalRaw.getDate().toString().padStart(2, '0') + '/' + 
                          (tanggalRaw.getMonth() + 1).toString().padStart(2, '0') + '/' + 
                          tanggalRaw.getFullYear();
      } else if (tanggalRaw) {
        tanggalFormatted = tanggalRaw.toString();
      }
      
      reportData.push({
        tanggal: tanggalFormatted,
        transactionID: data[i][2] || '',
        namaClient: data[i][3] || '',
        kitNumber: data[i][4] || '',
        serialNumber: data[i][5] || '',
        paket: data[i][6] || '',
        tipePembayaran: data[i][7] || ''
      });
    }
    
    return {
      status: 'success',
      data: reportData,
      totalRecords: reportData.length,
      sheetName: sheet.getName()
    };
    
  } catch (error) {
    return {
      status: 'error',
      message: error.message,
      data: []
    };
  }
}

/**
 * 📝 Handle POST requests
 */
function doPost(e) {
  try {
    Logger.log('=== DOPOST REQUEST ===');
    Logger.log('e.parameter: ' + JSON.stringify(e.parameter));
    Logger.log('e.postData exists: ' + (e.postData ? 'yes' : 'no'));

    var postData;
    var action;

    // ✅ PRIORITY 1: Check e.parameter first (most reliable for x-www-form-urlencoded)
    if (e.parameter && e.parameter.action) {
      Logger.log('✅ Using e.parameter (URL-encoded format auto-parsed by Apps Script)');
      action = e.parameter.action;

      // For submitClientData, parse the clientData JSON string
      if (action === 'submitClientData' && e.parameter.clientData) {
        try {
          postData = {
            action: action,
            clientData: JSON.parse(e.parameter.clientData)
          };
          Logger.log('✅ Parsed clientData from e.parameter');
        } catch (parseError) {
          Logger.log('❌ Failed to parse clientData: ' + parseError.message);
          return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: 'Invalid clientData JSON: ' + parseError.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      } else {
        // For other actions, use e.parameter directly
        postData = e.parameter;
      }
    }
    // ✅ PRIORITY 2: Fallback to e.postData.contents if e.parameter not available
    else if (e.postData && e.postData.contents) {
      Logger.log('⚠️ e.parameter not available, using e.postData.contents');
      Logger.log('POST data: ' + e.postData.contents);
      Logger.log('Content type: ' + e.postData.type);

      var contents = e.postData.contents.trim();

      // Detect format by first character
      if (contents.charAt(0) === '{') {
        // JSON format
        Logger.log('Detected JSON format');
        try {
          postData = JSON.parse(contents);
          action = postData.action;
          Logger.log('✅ Parsed as JSON');
        } catch (jsonError) {
          Logger.log('❌ JSON parse error: ' + jsonError.message);
          return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: 'Invalid JSON: ' + jsonError.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      } else {
        // URL-encoded format
        Logger.log('Detected URL-encoded format');
        try {
          var params = parseFormData(contents);
          action = params.action;

          if (params.clientData) {
            postData = {
              action: action,
              clientData: JSON.parse(params.clientData)
            };
            Logger.log('✅ Parsed URL-encoded data with clientData');
          } else {
            postData = params;
            Logger.log('✅ Parsed URL-encoded data');
          }
        } catch (urlError) {
          Logger.log('❌ URL-encoded parse error: ' + urlError.message);
          return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: 'Parse error: ' + urlError.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    // ❌ No data found
    else {
      Logger.log('❌ No POST data found in e.parameter or e.postData');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'No POST data received'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    Logger.log('POST Action: ' + action);
    
    if (action === 'submitPaymentMultiKit') {
      Logger.log('Processing multi-kit payment submission...');
      
      if (!postData.formData) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'error',
          message: 'formData diperlukan untuk submission'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const result = submitPaymentDataMultiKit(postData.formData);
      Logger.log('Submit result status: ' + result.status);
      
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    else if (action === 'submitClientData') {
      Logger.log('Processing client data submission via POST...');
      
      if (!postData.clientData) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'error',
          message: 'clientData diperlukan untuk submission'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const result = submitClientData(postData.clientData);
      Logger.log('Submit result status: ' + result.status);
      
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 🔐 APPROVAL SYSTEM ACTIONS
    else if (action === 'addPendingApproval') {
      const result = addPendingApproval(postData.approvalData);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === 'getPendingApprovals') {
      const result = getPendingApprovals(postData.filterStatus);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === 'approveChange') {
      const result = approveChange(postData.approvalId, postData.adminEmail);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === 'rejectChange') {
      const result = rejectChange(postData.approvalId, postData.adminEmail, postData.reason);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === 'approveAll') {
      const result = approveAll(postData.adminEmail);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === 'isAdmin') {
      const result = { status: 'success', isAdmin: isAdmin(postData.email) };
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    else {
      Logger.log('❌ Invalid POST action: ' + action);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'POST Action tidak valid. Available: submitPaymentMultiKit, submitClientData, addPendingApproval, getPendingApprovals, approveChange, rejectChange, approveAll, isAdmin',
        receivedAction: action,
        availableActions: ['submitPaymentMultiKit', 'submitClientData', 'addPendingApproval', 'getPendingApprovals', 'approveChange', 'rejectChange', 'approveAll', 'isAdmin']
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    Logger.log('❌ Error in doPost: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'POST Error: ' + error.message,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

