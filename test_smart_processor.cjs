const { smartProcessor } = require('./server/services/excel-processor-smart.ts');
const fs = require('fs');
const path = require('path');

async function testSmartProcessor() {
  console.log('🧠 ТЕСТ УМНОГО ПРОЦЕССОРА');
  console.log('==========================');
  
  // Используем существующий файл
  const testFile = 'uploads/Fortedetrim ORM report source.xlsx';
  
  try {
    if (!fs.existsSync(testFile)) {
      console.error('❌ Тестовый файл не найден:', testFile);
      return;
    }
    
    console.log('📊 Обрабатываем файл:', testFile);
    
    // Запускаем обработку
    const startTime = Date.now();
    const resultPath = await smartProcessor.processExcelFile(testFile);
    const endTime = Date.now();
    
    console.log(`✅ Обработка завершена за ${endTime - startTime}ms`);
    console.log(`📁 Результат сохранен в: ${resultPath}`);
    
    // Проверяем, что файл создан
    if (fs.existsSync(resultPath)) {
      const stats = fs.statSync(resultPath);
      console.log(`📊 Размер файла: ${stats.size} байт`);
      console.log(`📅 Создан: ${stats.birthtime}`);
    } else {
      console.error('❌ Выходной файл не создан');
    }
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    if (error.stack) {
      console.error(error.stack);
    }
  }
}

testSmartProcessor().catch(console.error); 