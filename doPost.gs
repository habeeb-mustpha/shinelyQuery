function doPost(e) {
  try {
    // Handle case when e is undefined
    if (!e || !e.parameter) {
      return ContentService
        .createTextOutput(JSON.stringify({
          'result': 'error',
          'error': 'No form data received'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get the spreadsheet (bound to this script)
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Sheet1');
    }
    
    // Get form data
    var fullName = e.parameter.fullName || '';
    var email = e.parameter.email || '';
    var phone = e.parameter.phone || '';
    var address = e.parameter.address || '';
    var serviceType = e.parameter.serviceType || '';
    var propertySize = e.parameter.propertySize || '';
    var message = e.parameter.message || '';
    
    // Log for debugging
    Logger.log('=== NEW SUBMISSION ===');
    Logger.log('Name: ' + fullName);
    Logger.log('Email: ' + email);
    
    // Create new row
    var newRow = [
      new Date(),        // Column A: Timestamp
      fullName,          // Column B
      email,             // Column C
      phone,             // Column D
      address,           // Column E
      serviceType,       // Column F
      propertySize,      // Column G
      message            // Column H
    ];
    
    // Add row to sheet
    sheet.appendRow(newRow);
    
    Logger.log('Row added successfully!');
    
    // Return success with proper MIME type for CORS
    return ContentService
      .createTextOutput(JSON.stringify({
        'result': 'success',
        'message': 'Data saved successfully'
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    
    return ContentService
      .createTextOutput(JSON.stringify({
        'result': 'error',
        'error': error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
