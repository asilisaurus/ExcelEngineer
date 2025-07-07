const ExcelJS = require('exceljs');

async function analyzeResultColumns() {
  console.log('📊 АНАЛИЗ КОЛОНОК В РЕЗУЛЬТИРУЮЩЕМ ФАЙЛЕ');
  
  try {
    const file = 'uploads/Fortedetrim_ORM_report_March_2025_результат_20250707.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file);
    
    const sheet = workbook.getWorksheet(1);
    
    console.log('\n🔍 ПЕРВЫЕ 15 СТРОК:');
    for (let i = 1; i <= 15; i++) {
      const row = sheet.getRow(i);
      const values = [];
      for (let j = 1; j <= 10; j++) {
        const cell = row.getCell(j);
        let value = cell.text || cell.value;
        if (value === null || value === undefined) value = '';
        values.push(value.toString().substring(0, 20)); // Ограничиваем длину
      }
      console.log(`Строка ${i}: ${values.join(' | ')}`);
    }
    
    console.log('\n🎯 АНАЛИЗ КОНКРЕТНЫХ КОЛОНОК:');
    
    // Ищем строки с данными (не заголовки)
    let dataStartRow = 5; // Предполагаем что данные начинаются с 5 строки
    for (let i = dataStartRow; i <= dataStartRow + 10; i++) {
      const row = sheet.getRow(i);
      console.log(`\nСтрока ${i}:`);
      console.log(`  Колонка D (Дата): "${row.getCell(4).text || row.getCell(4).value}"`);
      console.log(`  Колонка E (Ник): "${row.getCell(5).text || row.getCell(5).value}"`);
      console.log(`  Колонка F: "${row.getCell(6).text || row.getCell(6).value}"`);
      console.log(`  Колонка G: "${row.getCell(7).text || row.getCell(7).value}"`);
      console.log(`  Колонка H: "${row.getCell(8).text || row.getCell(8).value}"`);
    }
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
  }
}

analyzeResultColumns(); 