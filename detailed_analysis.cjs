const XLSX = require('xlsx');
const fs = require('fs');

// Read the source file
const buffer = fs.readFileSync('uploads/Fortedetrim ORM report source.xlsx');
const workbook = XLSX.read(buffer, { type: 'buffer' });

const sheetName = workbook.SheetNames.find(name => 
  name.includes('Мар25') || name.includes('Mar25') || name.includes('Март25')
);

const worksheet = workbook.Sheets[sheetName];
const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log(`\n=== DETAILED FILE ANALYSIS ===`);
console.log(`Total rows in file: ${jsonData.length}`);

let reviewCount = 0;
let commentCount = 0;
let otherCount = 0;
let emptyCount = 0;
let sectionMarkers = [];

console.log(`\n=== SECTION ANALYSIS ===`);

for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row || !Array.isArray(row)) {
    emptyCount++;
    continue;
  }
  
  const firstCell = String(row[0] || '').trim();
  
  if (firstCell.includes('Отзывы (отзовики)') || firstCell.includes('Отзывы (аптеки)')) {
    reviewCount++;
    if (reviewCount <= 5) {
      console.log(`Review ${reviewCount} at row ${i + 1}: ${firstCell}`);
    }
  } else if (firstCell.includes('Комментарии в обсуждениях')) {
    commentCount++;
    if (commentCount <= 5 || commentCount % 100 === 0) {
      console.log(`Comment ${commentCount} at row ${i + 1}: Platform=${row[1] || 'N/A'}`);
    }
  } else if (firstCell.includes('ТОП-20') || firstCell.includes('Тип размещения') || firstCell.includes('Активные')) {
    sectionMarkers.push({ row: i + 1, text: firstCell });
  } else if (firstCell) {
    otherCount++;
    if (otherCount <= 10) {
      console.log(`Other ${otherCount} at row ${i + 1}: "${firstCell}"`);
    }
  } else {
    emptyCount++;
  }
}

console.log(`\n=== SUMMARY ===`);
console.log(`Reviews found: ${reviewCount}`);
console.log(`Comments found: ${commentCount}`);
console.log(`Other data rows: ${otherCount}`);
console.log(`Empty rows: ${emptyCount}`);
console.log(`Total processed: ${reviewCount + commentCount + otherCount + emptyCount}`);

console.log(`\n=== SECTION MARKERS ===`);
sectionMarkers.forEach(marker => {
  console.log(`Row ${marker.row}: "${marker.text}"`);
});

// Check for data in specific ranges
console.log(`\n=== RANGE ANALYSIS ===`);
const ranges = [
  { name: 'Rows 1-50', start: 0, end: 49 },
  { name: 'Rows 51-100', start: 50, end: 99 },
  { name: 'Rows 101-200', start: 100, end: 199 },
  { name: 'Rows 201-400', start: 200, end: 399 },
  { name: 'Rows 401-600', start: 400, end: 599 },
  { name: 'Rows 601-800', start: 600, end: 799 },
  { name: 'Rows 801-1000', start: 800, end: 999 },
];

ranges.forEach(range => {
  let dataRowsInRange = 0;
  for (let i = range.start; i <= Math.min(range.end, jsonData.length - 1); i++) {
    const row = jsonData[i];
    if (row && Array.isArray(row) && row[0]) {
      const firstCell = String(row[0]).trim();
      if (firstCell.includes('Отзывы') || firstCell.includes('Комментарии')) {
        dataRowsInRange++;
      }
    }
  }
  console.log(`${range.name}: ${dataRowsInRange} data rows`);
});

// Sample a few rows from different parts of the file
console.log(`\n=== SAMPLE ROWS FROM DIFFERENT SECTIONS ===`);
const sampleRows = [100, 200, 300, 400, 500, 600, 700, 800];
sampleRows.forEach(rowNum => {
  if (rowNum < jsonData.length && jsonData[rowNum]) {
    const row = jsonData[rowNum];
    const firstCell = String(row[0] || '').trim();
    console.log(`Row ${rowNum + 1}: "${firstCell}" | Columns: ${row.length} | Platform: ${row[1] || 'N/A'}`);
  }
});

console.log(`\n=== FILE ANALYSIS COMPLETE ===`);