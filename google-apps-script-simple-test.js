/**
 * Simple Test Script - No External Access Required
 */

function simpleTest() {
  try {
    // Test 1: Basic functionality
    console.log('✅ Test 1: Script is running');
    
    // Test 2: Get current spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✅ Test 2: Got active spreadsheet: ' + spreadsheet.getName());
    
    // Test 3: Get active sheet
    var sheet = spreadsheet.getActiveSheet();
    console.log('✅ Test 3: Got active sheet: ' + sheet.getName());
    
    // Test 4: Read data
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    console.log('✅ Test 4: Read data - Rows: ' + values.length + ', Columns: ' + (values[0] ? values[0].length : 0));
    
    // Show success
    SpreadsheetApp.getUi().alert(
      'Test Successful!',
      'All tests passed:\n' +
      '- Script is running\n' +
      '- Can access current spreadsheet\n' +
      '- Can read data\n' +
      '\nRows: ' + values.length + '\n' +
      'Columns: ' + (values[0] ? values[0].length : 0),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ Error: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Error: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Create menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧪 Simple Test')
    .addItem('Run Test', 'simpleTest')
    .addToUi();
}