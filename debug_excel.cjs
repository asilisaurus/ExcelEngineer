const XLSX = require('xlsx');
const fs = require('fs');

// Read the source file
const buffer = fs.readFileSync('uploads/Fortedetrim ORM report source.xlsx');
const workbook = XLSX.read(buffer, { type: 'buffer' });

console.log('Available sheets:', workbook.SheetNames);

// Find Mar25 sheet
const sheetName = workbook.SheetNames.find(name => 
  name.includes('Мар25') || name.includes('Mar25') || name.includes('Март25')
);

if (!sheetName) {
  console.log('March sheet not found');
  process.exit(1);
}

console.log(`Processing sheet: ${sheetName}`);

const worksheet = workbook.Sheets[sheetName];
const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log(`Total rows: ${jsonData.length}`);
console.log('\nFirst 50 rows analysis:');

for (let i = 0; i < Math.min(50, jsonData.length); i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  const firstCell = String(row[0] || '').trim();
  if (firstCell) {
    console.log(`Row ${i + 1}: "${firstCell}" | Total cells: ${row.length}`);
    
    // Look for section markers
    if (firstCell.includes('Отзывы') || 
        firstCell.includes('ТОП-20') || 
        firstCell.includes('Комментарии') ||
        firstCell.includes('Тип размещения')) {
      console.log(`  *** SECTION MARKER ***`);
    }
  }
}

console.log('\nLooking for data patterns in rows 80-130:');
for (let i = 79; i < Math.min(130, jsonData.length); i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  const firstCell = String(row[0] || '').trim();
  if (firstCell) {
    console.log(`Row ${i + 1}: "${firstCell}"`);
  }
}

console.log('\nSample data row analysis (row 6):');
if (jsonData[5]) {
  const row = jsonData[5];
  console.log('Row 6 contents:');
  for (let j = 0; j < row.length; j++) {
    if (row[j] !== undefined && row[j] !== null && row[j] !== '') {
      console.log(`  Column ${String.fromCharCode(65 + j)} (${j}): "${row[j]}"`);
    }
  }
}

console.log('\nSample data row analysis (row 32):');
if (jsonData[31]) {
  const row = jsonData[31];
  console.log('Row 32 contents:');
  for (let j = 0; j < row.length; j++) {
    if (row[j] !== undefined && row[j] !== null && row[j] !== '') {
      console.log(`  Column ${String.fromCharCode(65 + j)} (${j}): "${row[j]}"`);
    }
  }
}