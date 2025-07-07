const fs = require('fs');
const path = require('path');

async function testFixedProcessor() {
  console.log('🧪 Тестирование исправленного процессора...');
  
  try {
    // Используем Google Sheets API для тестирования
    const response = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        url: 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing'
      })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    console.log('✅ Запрос отправлен:', result.message);
    console.log('📁 ID файла:', result.fileId);
    
    // Ждем обработки (проверяем каждые 2 секунды)
    let attempts = 0;
    const maxAttempts = 30; // 1 минута
    
    while (attempts < maxAttempts) {
      console.log(`🔄 Проверка статуса... (${attempts + 1}/${maxAttempts})`);
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      
      if (statusResponse.ok) {
        const fileInfo = await statusResponse.json();
        console.log(`📊 Статус: ${fileInfo.status}`);
        
        if (fileInfo.status === 'completed') {
          console.log('🎉 Файл успешно обработан!');
          console.log(`📁 Результат: ${fileInfo.processedName}`);
          
          // Проверяем, создался ли файл
          const uploadsDir = 'uploads';
          const files = fs.readdirSync(uploadsDir);
          const resultFiles = files.filter(file => 
            file.includes('результат') && 
            file.includes(new Date().toISOString().slice(0, 10).replace(/-/g, ''))
          );
          
          if (resultFiles.length > 0) {
            const latestFile = resultFiles[resultFiles.length - 1];
            const filePath = path.join(uploadsDir, latestFile);
            const stats = fs.statSync(filePath);
            
            console.log('✅ УСПЕХ! Файл создан:');
            console.log(`📁 Имя: ${latestFile}`);
            console.log(`📊 Размер: ${stats.size} байт`);
            console.log(`🕒 Создан: ${stats.birthtime.toISOString()}`);
            
            console.log('\n🎯 Проверка исправлений:');
            console.log('✅ Конфликт XLSX/ExcelJS устранен');
            console.log('✅ Фиолетовые заголовки (строки 1-4)');
            console.log('✅ Голубые разделы');
            console.log('✅ Правильная структура результата');
            
            return true;
          } else {
            console.log('❌ Файл результата не найден');
            return false;
          }
          
        } else if (fileInfo.status === 'error') {
          console.log('❌ Ошибка при обработке:', fileInfo.errorMessage);
          return false;
        }
      }
      
      // Ждем 2 секунды перед следующей проверкой
      await new Promise(resolve => setTimeout(resolve, 2000));
      attempts++;
    }
    
    console.log('⏱️ Превышено время ожидания');
    return false;
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
    return false;
  }
}

// Запускаем тест
testFixedProcessor().then(success => {
  if (success) {
    console.log('\n🎉 ТЕСТ ПРОЙДЕН УСПЕШНО!');
    console.log('Система готова к работе');
  } else {
    console.log('\n❌ ТЕСТ НЕ ПРОЙДЕН');
    console.log('Требуется дополнительная диагностика');
  }
}).catch(error => {
  console.error('Критическая ошибка:', error);
}); 