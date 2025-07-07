const fetch = require('node-fetch');

async function testProgressFix() {
  try {
    console.log(' Тестирование исправлений прогресса...');
    
    // Тест 1: Импорт из Google Sheets
    console.log('\n Тест 1: Импорт из Google Sheets');
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing';
    
    const importResponse = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ url: googleSheetsUrl }),
    });
    
    if (!importResponse.ok) {
      const errorText = await importResponse.text();
      console.error(' Ошибка импорта:', errorText);
      return;
    }
    
    const result = await importResponse.json();
    console.log(' Импорт запущен:', result.message);
    console.log(' ID файла:', result.fileId);
    
    // Отслеживаем прогресс
    const maxAttempts = 30; // 30 секунд
    let attempts = 0;
    
    while (attempts < maxAttempts) {
      console.log(\n Проверка прогресса (попытка {attempts + 1}/{maxAttempts})...);
      
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      const statusResponse = await fetch(http://localhost:5000/api/files/{result.fileId});
      
      if (!statusResponse.ok) {
        console.error(' Ошибка получения статуса');
        break;
      }
      
      const fileInfo = await statusResponse.json();
      
      console.log( Статус: {fileInfo.status});
      console.log( Строк обработано: {fileInfo.rowsProcessed});
      
      if (fileInfo.statistics) {
        try {
          const stats = JSON.parse(fileInfo.statistics);
          console.log( Этап: {stats.stage || 'неизвестно'});
          console.log( Сообщение: {stats.message || 'нет сообщения'});
        } catch (e) {
          console.log(' Статистика:', fileInfo.statistics);
        }
      }
      
      if (fileInfo.status === 'completed') {
        console.log('\n Обработка файла завершена успешно!');
        console.log( Результат: {fileInfo.processedName});
        break;
      }
      
      if (fileInfo.status === 'error') {
        console.log('\n Ошибка при обработке файла:');
        console.log( Сообщение: {fileInfo.errorMessage});
        break;
      }
      
      attempts++;
    }
    
    if (attempts >= maxAttempts) {
      console.log('\n Превышено время ожидания');
    }
    
  } catch (error) {
    console.error(' Ошибка тестирования:', error);
  }
}

// Запускаем тест
testProgressFix().catch(console.error);
