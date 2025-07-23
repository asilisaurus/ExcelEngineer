/**
 * Basic Test Script for Google Apps Script
 */

// Test function to verify basic functionality
function testBasicFunction() {
  try {
    // Get active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    
    console.log('✅ Spreadsheet accessed successfully');
    console.log(`📊 Spreadsheet name: ${spreadsheet.getName()}`);
    console.log(`📄 Active sheet name: ${sheet.getName()}`);
    
    // Get data range
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    console.log(`📊 Total rows: ${values.length}`);
    console.log(`📊 Total columns: ${values[0] ? values[0].length : 0}`);
    
    // Show first 5 rows
    console.log('\n📋 First 5 rows:');
    for (let i = 0; i < Math.min(5, values.length); i++) {
      console.log(`Row ${i + 1}: ${values[i].slice(0, 3).join(' | ')}...`);
    }
    
    SpreadsheetApp.getUi().alert(
      'Test Successful',
      `Sheet: ${sheet.getName()}\nRows: ${values.length}\nColumns: ${values[0] ? values[0].length : 0}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ Error:', error);
    SpreadsheetApp.getUi().alert('Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Create menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧪 Test')
    .addItem('Run Basic Test', 'testBasicFunction')
    .addToUi();
}