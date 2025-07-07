async function testDatesFix() {
  console.log('📅 ТЕСТ ИСПРАВЛЕНИЯ ДАТ');
  console.log('='.repeat(40));
  
  try {
    // Отправляем запрос на обработку
    console.log('📤 Отправляем запрос Google Sheets...');
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
    
    // Отслеживаем прогресс
    let attempts = 0;
    const maxAttempts = 30;
    
    while (attempts < maxAttempts) {
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${fileId}`);
      const fileInfo = await statusResponse.json();
      
      console.log(`[${attempts + 1}/${maxAttempts}] 📊 Статус: ${fileInfo.status}, Строк: ${fileInfo.rowsProcessed || 0}`);
      
      if (fileInfo.status === 'completed') {
        console.log('✅ Обработка завершена!');
        
        // Анализируем результат
        const resultFile = fileInfo.processedName;
        console.log(`📁 Файл создан: ${resultFile}`);
        
        // Проверяем даты в созданном файле
        const ExcelJS = require('exceljs');
        const workbook = new ExcelJS.Workbook();
        const filePath = `uploads/${resultFile}`;
        
        try {
          await workbook.xlsx.readFile(filePath);
          const sheet = workbook.getWorksheet(1);
          
          console.log('\n📋 ПРОВЕРКА ДАТ В РЕЗУЛЬТАТЕ:');
          let datesFound = 0;
          
          for (let i = 5; i <= 15; i++) {
            const row = sheet.getRow(i);
            const dateCell = row.getCell(4); // Колонка D
            const nickCell = row.getCell(5); // Колонка E
            
            if (dateCell.value) {
              datesFound++;
              console.log(`📅 Строка ${i}: Дата="${dateCell.text || dateCell.value}", Ник="${nickCell.text || nickCell.value}"`);
            }
          }
          
          if (datesFound > 0) {
            console.log(`🎉 УСПЕХ! Найдено ${datesFound} дат в результате`);
          } else {
            console.log('❌ ПРОБЛЕМА! Даты не найдены в результирующем файле');
          }
          
        } catch (fileError) {
          console.error('❌ Ошибка чтения файла:', fileError.message);
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

testDatesFix(); 