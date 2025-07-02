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

console.log('Analyzing views data in source file...\n');

// Function to check if a value looks like views count
function looksLikeViews(value) {
  if (!value) return false;
  
  const str = String(value).trim();
  if (str === '' || str === '-' || str.toLowerCase() === 'нет данных') return false;
  
  // Remove spaces and convert to number
  const cleaned = str.replace(/\s+/g, '').replace(/['"]/g, '').replace(',', '.');
  const num = parseFloat(cleaned);
  
  return !isNaN(num) && num > 0;
}

// Analyze several rows for views patterns
console.log('Views analysis for review rows (6-15):');
for (let i = 5; i < 15; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  console.log(`\nRow ${i + 1}:`);
  console.log(`  Platform: ${row[1]}`);
  console.log(`  Views candidates:`);
  
  for (let j = 8; j < 20; j++) {
    if (row[j] !== undefined && row[j] !== null && row[j] !== '') {
      const isViewsLike = looksLikeViews(row[j]);
      console.log(`    Column ${String.fromCharCode(65 + j)} (${j}): "${row[j]}" ${isViewsLike ? '*** LOOKS LIKE VIEWS ***' : ''}`);
    }
  }
}

console.log('\n\nViews analysis for comment rows (30-40):');
for (let i = 29; i < 40; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  console.log(`\nRow ${i + 1}:`);
  console.log(`  Platform: ${row[1]}`);
  console.log(`  Views candidates:`);
  
  for (let j = 8; j < 20; j++) {
    if (row[j] !== undefined && row[j] !== null && row[j] !== '') {
      const isViewsLike = looksLikeViews(row[j]);
      console.log(`    Column ${String.fromCharCode(65 + j)} (${j}): "${row[j]}" ${isViewsLike ? '*** LOOKS LIKE VIEWS ***' : ''}`);
    }
  }
}

// Sum all potential views
let totalViews = 0;
let viewsCount = 0;

console.log('\n\nSumming all potential views data:');
for (let i = 5; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  const firstCell = String(row[0] || '').trim();
  if (firstCell.includes('Отзывы') || firstCell.includes('Комментарии')) {
    // Look for views in columns I-P (8-15)
    for (let j = 8; j < 16; j++) {
      if (looksLikeViews(row[j])) {
        const str = String(row[j]).trim().replace(/\s+/g, '').replace(/['"]/g, '').replace(',', '.');
        const num = parseFloat(str);
        if (!isNaN(num) && num > 0) {
          totalViews += num;
          viewsCount++;
          console.log(`Row ${i + 1}, Col ${String.fromCharCode(65 + j)}: ${num}`);
          break; // Take only first valid views value per row
        }
      }
    }
  }
  
  if (firstCell.includes('Тип размещения') && i > 30) {
    break;
  }
}

console.log(`\nTotal views found: ${totalViews}`);
console.log(`Number of entries with views: ${viewsCount}`);

// Look for specific large numbers
console.log('\n\nLooking for large numbers (potential total views):');
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  for (let j = 0; j < row.length; j++) {
    if (row[j] && typeof row[j] === 'number' && row[j] > 1000000) {
      console.log(`Row ${i + 1}, Col ${String.fromCharCode(65 + j)}: ${row[j]} *** LARGE NUMBER ***`);
    }
  }
}