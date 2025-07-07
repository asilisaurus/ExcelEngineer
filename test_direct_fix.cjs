const { simpleProcessor } = require('./dist/server/services/excel-processor-simple.js');

async function testDirectFix() {
  try {
    console.log('🔧 Тестирование исправленной логики...');
    
    // Тестируем с локальным файлом
    const testFile = './test_download.xlsx';
    
    console.log('📁 Тестируем файл:', testFile);
    
    const result = await simpleProcessor.processExcelFile(testFile);
    
    console.log('✅ Результат обработки:', result);
    console.log('🎉 Обработка завершена успешно!');
    console.log('📁 Файл создан:', result);
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
    console.error('Stack:', error.stack);
  }
}

testDirectFix(); 