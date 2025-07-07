const ExcelJS = require('exceljs');

async function analyzeSourceColumns() {
  console.log('📊 АНАЛИЗ ИСХОДНЫХ ДАННЫХ');
  
  try {
    // Попробуем найти исходный файл
    const possibleFiles = [
      'uploads/Fortedetrim ORM report source.xlsx',
      'temp_google_download.xlsx'
    ];
    
    let sourceFile = null;
    for (const file of possibleFiles) {
      try {
        const fs = require('fs');
        if (fs.existsSync(file)) {
          sourceFile = file;
          break;
        }
      } catch (e) {}
    }
    
    if (!sourceFile) {
      console.log('❌ Исходный файл не найден');
      return;
    }
    
    console.log(`📁 Анализируем файл: ${sourceFile}`);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourceFile);
    
    const sheet = workbook.getWorksheet(1);
    
    console.log('\n🔍 ПЕРВЫЕ 10 СТРОК ИСХОДНЫХ ДАННЫХ:');
    for (let i = 1; i <= 10; i++) {
      const row = sheet.getRow(i);
      const values = [];
      for (let j = 1; j <= 10; j++) {
        const cell = row.getCell(j);
        let value = cell.text || cell.value;
        if (value === null || value === undefined) value = '';
        values.push(value.toString().substring(0, 15)); // Ограничиваем длину
      }
      console.log(`Строка ${i}: ${values.join(' | ')}`);
    }
    
    console.log('\n🎯 ПОИСК ДАТ В ИСХОДНЫХ ДАННЫХ:');
    
    // Ищем где могут быть даты в исходных данных
    for (let i = 2; i <= 20; i++) {
      const row = sheet.getRow(i);
      const rowData = [];
      let hasData = false;
      
      for (let j = 1; j <= 15; j++) {
        const cell = row.getCell(j);
        let value = cell.text || cell.value;
        if (value) hasData = true;
        
        // Проверяем на дату
        const isDate = value && (
          /\d{4}-\d{2}-\d{2}/.test(value) ||
          /\d{2}\.\d{2}\.\d{4}/.test(value) ||
          /\d{1,2}\/\d{1,2}\/\d{4}/.test(value) ||
          (typeof value === 'number' && value > 40000 && value < 50000) // Excel date number
        );
        
        if (isDate) {
          console.log(`🗓️ ДАТА НАЙДЕНА в строке ${i}, колонке ${j}: "${value}"`);
        }
        
        rowData.push(value);
      }
      
      if (!hasData) break; // Если строка пустая, прекращаем поиск
    }
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
  }
}

analyzeSourceColumns(); 