const { simpleProcessor } = require('./server/services/excel-processor-simple.ts');
const fs = require('fs');
const path = require('path');

async function testSimpleProcessor() {
  console.log('⚡ ТЕСТ УПРОЩЕННОГО ПРОЦЕССОРА');
  console.log('=============================');
  
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
    const resultPath = await simpleProcessor.processExcelFile(testFile);
    const endTime = Date.now();
    
    console.log(`✅ Обработка завершена за ${endTime - startTime}ms`);
    console.log(`📁 Результат сохранен в: ${resultPath}`);
    
    // Проверяем, что файл создан
    if (fs.existsSync(resultPath)) {
      const stats = fs.statSync(resultPath);
      console.log(`📊 Размер файла: ${stats.size} байт`);
      console.log(`📅 Создан: ${stats.birthtime}`);
      console.log('🎉 УСПЕХ! Файл создан без ошибок!');
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

testSimpleProcessor().catch(console.error); 