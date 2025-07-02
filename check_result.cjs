const XLSX = require('xlsx');
const fs = require('fs');

console.log('=== CHECKING RESULT FILE ===');

// Read the result file
const buffer = fs.readFileSync('new_result.xlsx');
const workbook = XLSX.read(buffer, { type: 'buffer' });

console.log('Available sheets:', workbook.SheetNames);

// Process the main sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log(`\nSheet: ${sheetName}`);
console.log(`Total rows in result file: ${jsonData.length}`);

let dataRows = 0;
let reviewRows = 0;
let commentRows = 0;
let activeDiscussionRows = 0;

console.log('\n=== ANALYZING CONTENT ===');

for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row || !Array.isArray(row)) continue;
  
  const firstCol = String(row[0] || '').trim();
  const lastCol = String(row[7] || '').trim(); // Type column
  
  if (lastCol === 'Отзывы') {
    reviewRows++;
    dataRows++;
  } else if (lastCol === 'Комментарии Топ-20 выдачи') {
    commentRows++;
    dataRows++;
  } else if (lastCol === 'Активные обсуждения (мониторинг)') {
    activeDiscussionRows++;
    dataRows++;
  }
  
  // Show first few data rows
  if (dataRows <= 5) {
    console.log(`Data row ${dataRows}: Platform="${row[0]}", Type="${row[7]}"`);
  }
  
  // Show progress for large sections
  if (dataRows > 0 && dataRows % 100 === 0) {
    console.log(`Progress: ${dataRows} data rows found...`);
  }
}

console.log('\n=== SUMMARY ===');
console.log(`Total data rows: ${dataRows}`);
console.log(`Reviews: ${reviewRows}`);
console.log(`Top-20 Comments: ${commentRows}`);
console.log(`Active Discussions: ${activeDiscussionRows}`);
console.log(`Total actual data: ${reviewRows + commentRows + activeDiscussionRows}`);

// Show last few rows to understand structure
console.log('\n=== LAST 10 ROWS ===');
for (let i = Math.max(0, jsonData.length - 10); i < jsonData.length; i++) {
  const row = jsonData[i];
  if (row && Array.isArray(row) && row.length > 0) {
    console.log(`Row ${i + 1}: "${row[0]}" | Type: "${row[7] || 'N/A'}"`);
  }
}

console.log('\n=== CHECK COMPLETE ===');