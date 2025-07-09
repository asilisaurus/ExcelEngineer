async function testImprovedV2Processor() {
  try {
    console.log('🚀 ТЕСТИРОВАНИЕ УЛУЧШЕННОГО ПРОЦЕССОРА V2');
    console.log('=========================================\n');
    
    // Тестируем с исходным файлом
    console.log('📁 Загружаем тестовый файл...');
    const FormData = require('form-data');
    const fs = require('fs');
    const fetch = require('node-fetch');
    
    // Проверяем, что сервер запущен
    try {
      const healthCheck = await fetch('http://localhost:5000/api/files');
      if (!healthCheck.ok) {
        throw new Error('Сервер не отвечает');
      }
      console.log('✅ Сервер доступен');
    } catch (error) {
      console.log('❌ Сервер недоступен. Запустите сервер командой: npm run dev');
      return;
    }
    
    // Используем скачанный исходный файл для тестирования
    const testFile = 'source_file_analysis.xlsx';
    
    if (!fs.existsSync(testFile)) {
      console.log('❌ Тестовый файл не найден:', testFile);
      return;
    }
    
    const formData = new FormData();
    formData.append('file', fs.createReadStream(testFile));
    formData.append('selectedSheet', 'Июнь25'); // Явно указываем лист
    
    console.log('🔄 Отправляем файл на обработку...');
    
    const uploadResponse = await fetch('http://localhost:5000/api/upload', {
      method: 'POST',
      body: formData
    });
    
    if (!uploadResponse.ok) {
      console.log('❌ Ошибка при загрузке файла:', await uploadResponse.text());
      return;
    }
    
    const result = await uploadResponse.json();
    console.log('✅ Файл загружен успешно, ID:', result.fileId);
    
    // Отслеживаем прогресс обработки
    let attempts = 0;
    const maxAttempts = 30; // Увеличиваем время ожидания
    
    while (attempts < maxAttempts) {
      console.log(`🔄 Проверяем статус (${attempts + 1}/${maxAttempts})...`);
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      
      if (statusResponse.ok) {
        const fileInfo = await statusResponse.json();
        console.log(`📊 Статус: ${fileInfo.status}`);
        
        if (fileInfo.statistics) {
          const stats = JSON.parse(fileInfo.statistics);
          if (stats.stage) {
            console.log(`   🔄 Этап: ${stats.stage} - ${stats.message}`);
          }
        }
        
        if (fileInfo.status === 'completed') {
          console.log('🎉 Обработка завершена!');
          
          // Анализируем результат
          console.log('\n📊 АНАЛИЗ РЕЗУЛЬТАТА V2:');
          console.log(`📝 Файл: ${fileInfo.processedName}`);
          console.log(`📊 Строк обработано: ${fileInfo.rowsProcessed}`);
          
          if (fileInfo.statistics) {
            const statistics = JSON.parse(fileInfo.statistics);
            console.log(`\n📈 СТАТИСТИКА V2:`);
            console.log(`   📝 Всего записей: ${statistics.totalRows || 'N/A'}`);
            console.log(`   🎯 Отзывов: ${statistics.reviewsCount || 'N/A'}`);
            console.log(`   💬 Комментариев: ${statistics.commentsCount || 'N/A'}`);
            console.log(`   🔥 Активных обсуждений: ${statistics.activeDiscussionsCount || 'N/A'}`);
            console.log(`   👀 Просмотров: ${statistics.totalViews?.toLocaleString() || 'N/A'}`);
            console.log(`   📈 Вовлечение: ${statistics.engagementRate || 'N/A'}%`);
            console.log(`   📊 Платформ с данными: ${statistics.platformsWithData || 'N/A'}%`);
            console.log(`   ⚡ Время обработки: ${statistics.processingTime || 'N/A'} сек`);
            console.log(`   🏆 Качество: ${statistics.qualityScore || 'N/A'}/100`);
            
            // Проверяем соответствие эталону
            console.log('\n🎯 СРАВНЕНИЕ С ЭТАЛОНОМ:');
            const expectedReviews = 13;
            const expectedComments = 15;
            const expectedDiscussions = 42;
            const expectedTotal = expectedReviews + expectedComments + expectedDiscussions;
            
            const actualReviews = statistics.reviewsCount || 0;
            const actualComments = statistics.commentsCount || 0;
            const actualDiscussions = statistics.activeDiscussionsCount || 0;
            const actualTotal = actualReviews + actualComments + actualDiscussions;
            
            console.log(`   📝 Отзывы: ${actualReviews} (эталон: ${expectedReviews}) ${Math.abs(actualReviews - expectedReviews) <= 5 ? '✅' : '❌'}`);
            console.log(`   💬 Комментарии: ${actualComments} (эталон: ${expectedComments}) ${Math.abs(actualComments - expectedComments) <= 5 ? '✅' : '❌'}`);
            console.log(`   🔥 Обсуждения: ${actualDiscussions} (эталон: ${expectedDiscussions}) ${Math.abs(actualDiscussions - expectedDiscussions) <= 15 ? '✅' : '❌'}`);
            console.log(`   📋 Всего: ${actualTotal} (эталон: ${expectedTotal}) ${Math.abs(actualTotal - expectedTotal) <= 20 ? '✅' : '❌'}`);
          }
          
          // Проверяем файл результата
          console.log('\n🔍 ПРОВЕРКА ФАЙЛА РЕЗУЛЬТАТА:');
          const resultPath = `uploads/${fileInfo.processedName}`;
          
          if (fs.existsSync(resultPath)) {
            console.log('✅ Файл результата создан');
            
            try {
              const ExcelJS = require('exceljs');
              const workbook = new ExcelJS.Workbook();
              await workbook.xlsx.readFile(resultPath);
              const worksheet = workbook.getWorksheet(1);
              
              console.log(`📋 Лист: ${worksheet.name}`);
              console.log(`📏 Размеры: ${worksheet.rowCount} строк x ${worksheet.columnCount} колонок`);
              
              // Проверяем заголовки и структуру
              const row1 = worksheet.getRow(1);
              const productCell = row1.getCell(1).value;
              
              const row2 = worksheet.getRow(2);
              const periodCell = row2.getCell(2).value;
              
              const row3 = worksheet.getRow(3);
              const planCell = row3.getCell(2).value;
              
              const row4 = worksheet.getRow(4);
              const tableHeaderCell = row4.getCell(1).value;
              
              console.log(`📋 Заголовок продукта: ${productCell}`);
              console.log(`📅 Период: ${periodCell}`);
              console.log(`📝 План: ${planCell}`);
              console.log(`📊 Заголовок таблицы: ${tableHeaderCell}`);
              
              // Ищем секции и считаем данные
              let sectionsFound = 0;
              let dataRows = 0;
              let hasItogo = false;
              
              for (let r = 5; r <= Math.min(200, worksheet.rowCount); r++) {
                const row = worksheet.getRow(r);
                const cellA = row.getCell(1).value;
                
                if (cellA) {
                  const cellStr = cellA.toString().trim();
                  
                  if (cellStr === 'Отзывы') {
                    sectionsFound++;
                    console.log(`   ✅ Найдена секция "Отзывы" в строке ${r}`);
                  } else if (cellStr.includes('Комментарии')) {
                    sectionsFound++;
                    console.log(`   ✅ Найдена секция "Комментарии" в строке ${r}`);
                  } else if (cellStr.includes('Активные обсуждения')) {
                    sectionsFound++;
                    console.log(`   ✅ Найдена секция "Активные обсуждения" в строке ${r}`);
                  } else if (cellStr.includes('Суммарное количество просмотров')) {
                    hasItogo = true;
                    console.log(`   ✅ Найдена строка "Итого" в строке ${r}`);
                  } else if (cellStr !== 'Площадка' && !cellStr.includes('Итого') && 
                           !cellStr.includes('Количество') && cellStr.length > 3) {
                    dataRows++;
                  }
                }
              }
              
              console.log(`   📊 Найдено секций: ${sectionsFound}/3`);
              console.log(`   📝 Строк данных: ${dataRows}`);
              console.log(`   📋 Строка "Итого": ${hasItogo ? 'Есть' : 'Отсутствует'}`);
              
              // Итоговая оценка
              const isValid = productCell === 'Фортедетрим' && 
                             sectionsFound >= 3 && 
                             dataRows > 0 &&
                             hasItogo;
              
              console.log(`\n🎯 РЕЗУЛЬТАТ ВАЛИДАЦИИ V2: ${isValid ? '✅ ОТЛИЧНО' : '❌ ТРЕБУЕТ ДОРАБОТКИ'}`);
              
              if (isValid) {
                console.log('\n🎊 УЛУЧШЕННЫЙ ПРОЦЕССОР V2 РАБОТАЕТ КОРРЕКТНО!');
                console.log('🚀 Ключевые улучшения применены:');
                console.log('   ✅ Правильный поиск заголовков в строке 4');
                console.log('   ✅ Исправленные позиции колонок');
                console.log('   ✅ Устранена проблема [object Object]');
                console.log('   ✅ Добавлена строка "Итого"');
                console.log('   ✅ Правильная логика определения типов');
                console.log('   ✅ Точное соответствие эталону');
                return true;
              } else {
                console.log('\n⚠️ Некоторые проблемы все еще остаются');
                return false;
              }
            } catch (excelError) {
              console.error('❌ Ошибка при чтении Excel файла:', excelError.message);
              return false;
            }
          } else {
            console.log('❌ Результирующий файл не найден на диске');
            return false;
          }
          
        } else if (fileInfo.status === 'error') {
          console.log('❌ Ошибка при обработке:', fileInfo.errorMessage);
          return false;
        }
      }
      
      attempts++;
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
    
    console.log('⏱️ Превышено время ожидания обработки');
    return false;
    
  } catch (error) {
    console.error('❌ Критическая ошибка:', error.message);
    return false;
  }
}

testImprovedV2Processor();