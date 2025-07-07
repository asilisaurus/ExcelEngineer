const ExcelJS = require('exceljs');

async function checkReferenceFile() {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('attached_assets/Фортедетрим_ORM_отчет_Март_2025_результат_1751040742705.xlsx');
    
    const sheet = workbook.getWorksheet(1);
    
    console.log('=== СТРУКТУРА ЭТАЛОННОГО ФАЙЛА ===');
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
    
    // Ищем строку с итого
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
        console.log('');
        
        // Проверим следующие 3 строки после итого
        console.log('=== СТРОКИ ПОСЛЕ ИТОГО ===');
        for (let nextRow = row + 1; nextRow <= Math.min(row + 3, sheet.rowCount); nextRow++) {
          let nextRowData = [];
          for (let col = 1; col <= 6; col++) {
            const cell = sheet.getCell(nextRow, col);
            nextRowData.push(cell.value);
          }
          console.log(`Строка ${nextRow}: ${JSON.stringify(nextRowData)}`);
        }
        break;
      }
    }
    
  } catch (error) {
    console.error('Ошибка:', error.message);
  }
}

checkReferenceFile(); 