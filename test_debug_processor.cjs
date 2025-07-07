async function testDebugProcessor() {
  console.log('🔍 ТЕСТ С ДЕТАЛЬНЫМ АНАЛИЗОМ ЛОГОВ ПРОЦЕССОРА');
  console.log('='.repeat(60));
  
  try {
    console.log('📤 Отправляем запрос на обработку...');
    const response = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        url: 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing'
      })
    });
    
    const result = await response.json();
    const fileId = result.fileId;
    console.log(`✅ Запрос отправлен, ID файла: ${fileId}`);
    
    console.log('\n📋 ВАЖНЫЕ СООБЩЕНИЯ ДЛЯ ПОИСКА В ЛОГАХ:');
    console.log('  🔥 "ИСПОЛЬЗУЕТСЯ НОВЫЙ ПРОЦЕССОР ExcelProcessorSimple!"');
    console.log('  🔍 "DEBUG DATE: Checking column..."');
    console.log('  📅 "Found DATE (ISO format)..."');
    console.log('  📅 "Found DATE (day format)..."');
    console.log('  📅 "Found DATE (Date object)..."');
    console.log('  👤 "Found AUTHOR..."');
    console.log('');
    
    // Быстрое отслеживание
    console.log('⏱️  Отслеживаем прогресс...');
    
    let attempts = 0;
    while (attempts < 10) {
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${fileId}`);
      const fileInfo = await statusResponse.json();
      
      console.log(`[${attempts + 1}/10] 📊 Статус: ${fileInfo.status}, Строк: ${fileInfo.rowsProcessed || 0}`);
      
      if (fileInfo.status === 'completed') {
        console.log('✅ Обработка завершена!');
        
        // Анализируем результат
        const ExcelJS = require('exceljs');
        const workbook = new ExcelJS.Workbook();
        
        await workbook.xlsx.readFile(`uploads/${fileInfo.processedName}`);
        const sheet = workbook.getWorksheet(1);
        
        console.log('\n📋 ДЕТАЛЬНЫЙ АНАЛИЗ РЕЗУЛЬТАТА:');
        console.log('Первые 15 строк с данными:');
        
        let realDatesFound = 0;
        let authorsFound = 0;
        
        for (let i = 6; i <= 20; i++) {
          const row = sheet.getRow(i);
          const platform = row.getCell(1).text || row.getCell(1).value || '';
          const date = row.getCell(4).text || row.getCell(4).value || '';
          const author = row.getCell(5).text || row.getCell(5).value || '';
          
          // Проверяем реальные даты (не служебные)
          if (date && date !== '' && date !== 'Отзывы' && date !== 'Комментарии' && 
              date !== 'Активные обсуждения' && !date.includes('Топ-20') && 
              date.match(/\d{2}\.\d{2}\.\d{4}/)) {
            realDatesFound++;
          }
          
          if (author && author !== '' && author !== 'Отзывы' && author !== 'Комментарии' && 
              author !== 'Активные обсуждения' && !author.includes('Топ-20')) {
            authorsFound++;
          }
          
          if (platform && platform !== '' && !platform.includes('Отзывы') && 
              !platform.includes('Комментарии') && !platform.includes('обсуждения')) {
            console.log(`Строка ${i}: Платформа="${platform}" | Дата="${date}" | Автор="${author}"`);
          }
        }
        
        console.log('\n📊 ФИНАЛЬНАЯ СТАТИСТИКА:');
        console.log(`  📅 Реальных дат найдено: ${realDatesFound}`);
        console.log(`  👤 Авторов найдено: ${authorsFound}`);
        
        if (realDatesFound > 0) {
          console.log('\n🎉 УСПЕХ! Исправление работает!');
        } else {
          console.log('\n❌ ПРОБЛЕМА: Исправление не работает');
          console.log('   Проверьте логи сервера на наличие сообщений "📅 Found DATE"');
          console.log('   Возможно проблема в другом месте');
        }
        
        break;
      }
      
      if (fileInfo.status === 'error') {
        console.log('❌ Ошибка:', fileInfo.errorMessage);
        break;
      }
      
      attempts++;
    }
    
  } catch (error) {
    console.error('❌ Ошибка теста:', error.message);
  }
}

testDebugProcessor(); 