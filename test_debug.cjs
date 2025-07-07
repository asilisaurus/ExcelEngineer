const fetch = require('node-fetch');

async function testGoogleSheets() {
  try {
    console.log('🚀 Тестирование Google Sheets импорта с отладкой...');
    
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing';
    
    console.log('📋 Отправляем запрос на импорт...');
    const importResponse = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ url: googleSheetsUrl }),
    });
    
    if (!importResponse.ok) {
      const errorText = await importResponse.text();
      console.error('❌ Ошибка импорта:', errorText);
      return;
    }
    
    const result = await importResponse.json();
    console.log('✅ Импорт запущен:', result.message);
    console.log('🆔 ID файла:', result.fileId);
    
    // Отслеживаем прогресс каждые 0.5 секунды
    const maxAttempts = 60; // 30 секунд
    let attempts = 0;
    
    while (attempts < maxAttempts) {
      await new Promise(resolve => setTimeout(resolve, 500));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      
      if (!statusResponse.ok) {
        console.error('❌ Ошибка получения статуса');
        break;
      }
      
      const fileInfo = await statusResponse.json();
      
      console.log(`[${attempts + 1}/${maxAttempts}] 📊 Статус: ${fileInfo.status}, Строк: ${fileInfo.rowsProcessed}`);
      
      if (fileInfo.statistics) {
        try {
          const stats = JSON.parse(fileInfo.statistics);
          console.log(`  📈 Этап: ${stats.stage}, Сообщение: ${stats.message}`);
        } catch (e) {
          console.log('  📊 Статистика:', fileInfo.statistics);
        }
      }
      
      if (fileInfo.status === 'completed') {
        console.log('\n🎉 Обработка завершена успешно!');
        console.log(`📁 Результат: ${fileInfo.processedName}`);
        
        // Проверяем файл на диске
        const fs = require('fs');
        const path = require('path');
        const filePath = path.join('uploads', fileInfo.processedName);
        if (fs.existsSync(filePath)) {
          console.log(`✅ Файл существует на диске: ${filePath}`);
          console.log(`📊 Размер файла: ${fs.statSync(filePath).size} байт`);
        } else {
          console.log(`❌ Файл НЕ найден на диске: ${filePath}`);
        }
        
        // Тестируем скачивание
        console.log('\n🔄 Тестируем скачивание...');
        const downloadResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}/download`);
        if (downloadResponse.ok) {
          console.log('✅ Скачивание работает!');
        } else {
          console.log('❌ Ошибка скачивания:', await downloadResponse.text());
        }
        
        break;
      }
      
      if (fileInfo.status === 'error') {
        console.log('\n❌ Ошибка при обработке файла:');
        console.log(`🚨 Сообщение: ${fileInfo.errorMessage}`);
        break;
      }
      
      attempts++;
    }
    
    if (attempts >= maxAttempts) {
      console.log('\n⏱️ Превышено время ожидания');
    }
    
  } catch (error) {
    console.error('❌ Ошибка тестирования:', error);
  }
}

// Запускаем тест
testGoogleSheets().catch(console.error); 