async function checkLogs() {
  console.log('🔍 ПРОВЕРКА ЛОГОВ СЕРВЕРА');
  
  try {
    // Быстрый тест
    const response = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        url: 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing'
      })
    });
    
    const result = await response.json();
    console.log('✅ Запрос отправлен, ID:', result.fileId);
    console.log('🔍 Проверьте логи сервера на наличие сообщений:');
    console.log('  - "🔥🔥🔥 ИСПОЛЬЗУЕТСЯ НОВЫЙ ПРОЦЕССОР ExcelProcessorSimple!"');
    console.log('  - "🔥 НОВЫЙ ПРОЦЕССОР: Исходно найдено: X записей"');
    console.log('  - "🎯 Активные обсуждения: X записей"');
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
  }
}

checkLogs(); 