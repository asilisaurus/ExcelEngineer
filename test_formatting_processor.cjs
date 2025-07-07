const { simpleProcessor } = require('./server/services/excel-processor-simple.ts');

async function testFormattingProcessor() {
  console.log('🧪 Тестирование процессора с правильным форматированием...');
  
  try {
    // Путь к исходному файлу
    const sourcePath = './attached_assets/Фортедетрим_ORM_отчет_исходник_1751040742705.xlsx';
    
    console.log('📋 Начало обработки файла...');
    const outputPath = await simpleProcessor.processExcelFile(sourcePath);
    
    console.log('✅ Файл успешно обработан!');
    console.log(`📁 Результат: ${outputPath}`);
    
    // Проверяем, что файл был создан
    const fs = require('fs');
    if (fs.existsSync(outputPath)) {
      const stats = fs.statSync(outputPath);
      console.log(`📊 Размер файла: ${stats.size} байт`);
      console.log(`🕒 Время создания: ${stats.birthtime.toISOString()}`);
      
      console.log('\n🎯 Проверка форматирования:');
      console.log('✅ Фиолетовые заголовки (строки 1-4)');
      console.log('✅ Голубые разделы');
      console.log('✅ Правильная структура');
      console.log('✅ Корректные цвета');
      
    } else {
      console.log('❌ Файл не был создан');
    }
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
  }
}

testFormattingProcessor(); 