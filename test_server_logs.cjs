async function testServerLogs() {
  console.log('📋 ТЕСТ ЛОГОВ СЕРВЕРА ВО ВРЕМЯ ОБРАБОТКИ');
  console.log('='.repeat(50));
  
  try {
    console.log('📤 Отправляем запрос для обработки...');
    console.log('⚠️  ВАЖНО: Следите за логами сервера в другом окне терминала!');
    console.log('');
    
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
    console.log('');
    console.log('🔍 В логах сервера должны появиться сообщения:');
    console.log('  - "🔥🔥🔥 ИСПОЛЬЗУЕТСЯ НОВЫЙ ПРОЦЕССОР ExcelProcessorSimple!"');
    console.log('  - "🔍 DEBUG DATE: Checking column..."');
    console.log('  - "📅 Found DATE (...)"');
    console.log('  - "👤 Found AUTHOR (...)"');
    console.log('');
    
    // Отслеживаем прогресс
    let attempts = 0;
    const maxAttempts = 15;
    
    while (attempts < maxAttempts) {
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${fileId}`);
      const fileInfo = await statusResponse.json();
      
      console.log(`[${attempts + 1}/${maxAttempts}] 📊 Статус: ${fileInfo.status}, Строк: ${fileInfo.rowsProcessed || 0}`);
      
      if (fileInfo.status === 'completed') {
        console.log('✅ Обработка завершена!');
        console.log(`📁 Файл: ${fileInfo.processedName}`);
        
        // Проверяем финальный результат
        console.log('\n🔍 АНАЛИЗ СОЗДАННОГО ФАЙЛА:');
        const ExcelJS = require('exceljs');
        const workbook = new ExcelJS.Workbook();
        
        try {
          await workbook.xlsx.readFile(`uploads/${fileInfo.processedName}`);
          const sheet = workbook.getWorksheet(1);
          
          console.log('📋 Первые 10 строк результата:');
          for (let i = 5; i <= 10; i++) {
            const row = sheet.getRow(i);
            const platform = row.getCell(1).text || row.getCell(1).value || '';
            const text = (row.getCell(3).text || row.getCell(3).value || '').substring(0, 30);
            const date = row.getCell(4).text || row.getCell(4).value || '';
            const author = row.getCell(5).text || row.getCell(5).value || '';
            
            console.log(`Строка ${i}: Платформа="${platform}" | Дата="${date}" | Автор="${author}" | Текст="${text}..."`);
          }
          
          // Подсчитываем статистику
          let datesCount = 0;
          let authorsCount = 0;
          
          for (let i = 5; i <= 25; i++) {
            const row = sheet.getRow(i);
            const date = row.getCell(4).text || row.getCell(4).value;
            const author = row.getCell(5).text || row.getCell(5).value;
            
            if (date && date !== '' && date !== 'Отзывы' && date !== 'Комментарии') {
              datesCount++;
            }
            if (author && author !== '' && author !== 'Отзывы' && author !== 'Комментарии') {
              authorsCount++;
            }
          }
          
          console.log(`\n📊 СТАТИСТИКА (первые 20 строк данных):`);
          console.log(`  📅 Дат найдено: ${datesCount}`);
          console.log(`  👤 Авторов найдено: ${authorsCount}`);
          
          if (datesCount === 0) {
            console.log('\n❌ ПРОБЛЕМА: Даты не переносятся в результирующий файл!');
            console.log('   Проверьте логи сервера на наличие сообщений "📅 Found DATE"');
          } else {
            console.log('\n✅ УСПЕХ: Даты успешно переносятся в результат!');
          }
          
        } catch (fileError) {
          console.error('❌ Ошибка чтения результирующего файла:', fileError.message);
        }
        
        break;
      }
      
      if (fileInfo.status === 'error') {
        console.log('❌ Ошибка обработки:', fileInfo.errorMessage);
        break;
      }
      
      attempts++;
    }
    
    if (attempts >= maxAttempts) {
      console.log('⏰ Превышено время ожидания');
    }
    
  } catch (error) {
    console.error('❌ Ошибка теста:', error.message);
  }
}

testServerLogs(); 