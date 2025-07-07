const fetch = require('node-fetch');
const FormData = require('form-data');
const fs = require('fs');

async function testFixedLogic() {
  try {
    console.log('🔧 Тестирование исправленной логики...');
    
    // Тестируем с Google Sheets URL
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing';
    
    const response = await fetch('http://localhost:3000/api/process-google-sheets', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        url: googleSheetsUrl,
        processorType: 'simple'
      })
    });
    
    const result = await response.json();
    console.log('✅ Результат обработки:', result);
    
    if (result.success) {
      console.log('🎉 Обработка завершена успешно!');
      console.log('📁 Файл создан:', result.filename);
    } else {
      console.log('❌ Ошибка:', result.error);
    }
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
  }
}

testFixedLogic(); 