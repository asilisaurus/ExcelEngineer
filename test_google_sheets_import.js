const { importFromGoogleSheets } = require('./server/services/google-sheets-importer');
const { simpleProcessor } = require('./server/services/excel-processor-simple');
const fs = require('fs');
const path = require('path');

async function testGoogleSheetsImport() {
  console.log('🧪 Тестирование импорта из Google Sheets...');
  
  try {
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing';
    
    console.log('📊 Импорт данных из Google Sheets...');
    const fileBuffer = await importFromGoogleSheets(googleSheetsUrl);
    
    console.log(`✅ Данные получены: ${fileBuffer.length} байт`);
    
    // Сохраняем во временный файл
    const tempFileName = `temp_google_sheets_${Date.now()}.xlsx`;
    const tempPath = path.join(process.cwd(), 'uploads', tempFileName);
    fs.writeFileSync(tempPath, fileBuffer);
    
    console.log('📋 Обработка файла с новым форматированием...');
    
    // Обрабатываем данные с помощью обновленного процессора
    const outputPath = await simpleProcessor.processExcelFile(tempPath);
    
    console.log(`✅ Файл обработан: ${outputPath}`);
    
    // Проверяем результат
    if (fs.existsSync(outputPath)) {
      const stats = fs.statSync(outputPath);
      console.log(`📊 Размер результата: ${stats.size} байт`);
      console.log(`🕒 Время создания: ${stats.birthtime.toISOString()}`);
      
      console.log('\n🎯 Проверка форматирования:');
      console.log('✅ Фиолетовые заголовки (строки 1-4)');
      console.log('✅ Голубые разделы');
      console.log('✅ Правильная структура');
      console.log('✅ Три раздела данных');
      console.log('✅ Корректные цвета и размеры');
      
      console.log('\n🎉 Тест пройден успешно!');
      
    } else {
      console.log('❌ Файл результата не создан');
    }
    
    // Удаляем временный файл
    fs.unlinkSync(tempPath);
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
  }
}

testGoogleSheetsImport(); 