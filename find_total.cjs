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

console.log('Searching for total views value 3398560...\n');

// Search for exact value
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  for (let j = 0; j < row.length; j++) {
    if (row[j] && (String(row[j]).includes('3398560') || row[j] === 3398560)) {
      console.log(`Found 3398560 at Row ${i + 1}, Column ${String.fromCharCode(65 + j)} (${j})`);
      console.log(`Context: ${JSON.stringify(row.slice(0, 10))}`);
    }
  }
}

// Search for any large numbers 
console.log('\nLooking for numbers > 1,000,000:');
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  for (let j = 0; j < row.length; j++) {
    if (row[j] && typeof row[j] === 'number' && row[j] > 1000000) {
      console.log(`Large number ${row[j]} at Row ${i + 1}, Column ${String.fromCharCode(65 + j)}`);
      console.log(`Row context: ${JSON.stringify(row.slice(Math.max(0, j-2), j+3))}`);
    }
  }
}

// Search in text for "просмотр" keyword
console.log('\nSearching for "просмотр" keyword:');
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  for (let j = 0; j < row.length; j++) {
    if (row[j] && String(row[j]).toLowerCase().includes('просмотр')) {
      console.log(`Found "просмотр" at Row ${i + 1}, Column ${String.fromCharCode(65 + j)}`);
      console.log(`Value: "${row[j]}"`);
      console.log(`Row context: ${JSON.stringify(row.slice(Math.max(0, j-2), j+3))}`);
    }
  }
}

// Look for yellow highlighted cell content (summary statistics)
console.log('\nSearching for summary/statistics section:');
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  const firstCell = String(row[0] || '').toLowerCase();
  if (firstCell.includes('суммарное') || firstCell.includes('просмотр') || firstCell.includes('статистик')) {
    console.log(`Summary row ${i + 1}: ${JSON.stringify(row.slice(0, 10))}`);
  }
}

// Search entire row for potential summary data
console.log('\nSearching rows that might contain summary data (looking for patterns):');
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (!row) continue;
  
  // Look for rows with multiple large numbers
  let largeNumbers = [];
  for (let j = 0; j < row.length; j++) {
    if (row[j] && typeof row[j] === 'number' && row[j] > 10000) {
      largeNumbers.push({col: j, val: row[j]});
    }
  }
  
  if (largeNumbers.length >= 2) {
    console.log(`Row ${i + 1} has multiple large numbers:`, largeNumbers);
    console.log(`Full row: ${JSON.stringify(row.slice(0, 15))}`);
  }
}