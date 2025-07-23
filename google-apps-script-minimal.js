/**
 * Minimal Test - Absolute Basics
 */

function testMinimal() {
  // Super simple test
  Browser.msgBox('Hello! Script is working!');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Test')
    .addItem('Click Me', 'testMinimal')
    .addToUi();
}