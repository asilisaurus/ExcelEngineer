const ExcelJS = require('exceljs');

async function checkLastFile() {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('uploads/file-1751884808676-267432594_результат_20250707.xlsx');
    
    const sheet = workbook.getWorksheet(1);
    
    console.log('=== СТРУКТУРА ПОСЛЕДНЕГО ФАЙЛА ===');
    console.log('Всего строк:', sheet.rowCount);
    console.log('Всего столбцов:', sheet.columnCount);
    console.log('');
    
    // Проверим заголовки
    console.log('=== ЗАГОЛОВКИ ===');
    for (let col = 1; col <= 6; col++) {
      const cell = sheet.getCell(1, col);
      console.log(`Столбец ${col}: ${cell.value}`);
    }
    console.log('');
    
    // Проверим первые 3 строки данных
    console.log('=== ПЕРВЫЕ 3 СТРОКИ ДАННЫХ ===');
    for (let row = 2; row <= 4; row++) {
      let rowData = [];
      for (let col = 1; col <= 6; col++) {
        const cell = sheet.getCell(row, col);
        rowData.push(cell.value);
      }
      console.log(`Строка ${row}: ${JSON.stringify(rowData)}`);
    }
    console.log('');
    
    // Проверим последние 5 строк (где может быть итого)
    console.log('=== ПОСЛЕДНИЕ 5 СТРОК ===');
    for (let row = sheet.rowCount - 4; row <= sheet.rowCount; row++) {
      let rowData = [];
      for (let col = 1; col <= 6; col++) {
        const cell = sheet.getCell(row, col);
        rowData.push(cell.value);
      }
      console.log(`Строка ${row}: ${JSON.stringify(rowData)}`);
    }
    
    // Ищем строку с итого
    console.log('');
    console.log('=== ПОИСК ИТОГО ===');
    for (let row = 1; row <= sheet.rowCount; row++) {
      const cellA = sheet.getCell(row, 1);
      if (cellA.value && cellA.value.toString().toLowerCase().includes('итого')) {
        let rowData = [];
        for (let col = 1; col <= 6; col++) {
          const cell = sheet.getCell(row, col);
          rowData.push(cell.value);
        }
        console.log(`Строка итого ${row}: ${JSON.stringify(rowData)}`);
      }
    }
    
  } catch (error) {
    console.error('Ошибка:', error.message);
  }
}

checkLastFile(); 