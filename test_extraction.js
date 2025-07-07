const XLSX = require('xlsx');
const fs = require('fs');

// Прямой импорт процессора
const { ExcelProcessor } = require('./server/services/excel-processor-improved');

async function testExtraction() {
  try {
    console.log('🔍 ТЕСТИРОВАНИЕ ИЗВЛЕЧЕНИЯ ДАННЫХ');
    console.log('================================');
    
    // Читаем исходный файл
    const buffer = fs.readFileSync('uploads/Fortedetrim ORM report source.xlsx');
    
    // Создаем процессор
    const processor = new ExcelProcessor();
    
    // Обрабатываем файл
    const result = await processor.processExcelFile(buffer, 'test.xlsx');
    
    console.log('📊 РЕЗУЛЬТАТЫ ОБРАБОТКИ:');
    console.log('Статистика:', result.statistics);
    
    // Сохраняем результат
    const outputBuffer = await result.workbook.xlsx.writeBuffer();
    fs.writeFileSync('test_output.xlsx', outputBuffer);
    
    console.log('✅ Файл сохранен как test_output.xlsx');
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
    console.error(error.stack);
  }
}

testExtraction(); 