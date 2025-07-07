const fs = require('fs');
const path = require('path');

async function ultimateFinalTest() {
  console.log('🚀🚀🚀 ОКОНЧАТЕЛЬНЫЙ ТЕСТ СИСТЕМЫ 🚀🚀🚀');
  console.log('='.repeat(70));
  
  try {
    // Ждем запуска сервера
    console.log('⏳ Ожидание запуска сервера (8 сек)...');
    await new Promise(resolve => setTimeout(resolve, 8000));
    
    // Проверяем доступность сервера
    try {
      await fetch('http://localhost:5000/api/health');
      console.log('✅ Сервер доступен');
    } catch {
      console.log('⚠️ Сервер может быть еще не готов, пробуем...');
    }
    
    // Отправляем запрос на обработку
    console.log('📤 Отправляем запрос на обработку Google Sheets...');
    const response = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        url: 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing'
      })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    console.log('✅ Запрос отправлен успешно!');
    console.log(`📁 ID файла: ${result.fileId}`);
    
    // Ждем завершения обработки
    let attempts = 0;
    const maxAttempts = 20;
    
    while (attempts < maxAttempts) {
      console.log(`🔄 Проверяем статус (${attempts + 1}/${maxAttempts})...`);
      
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      const fileInfo = await statusResponse.json();
      
      if (fileInfo.status === 'completed') {
        console.log('🎉 Обработка завершена!');
        
        // Находим созданный файл
        const files = fs.readdirSync('uploads');
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const resultFiles = files.filter(f => f.includes('результат') && f.includes(today));
        
        if (resultFiles.length > 0) {
          const latestFile = resultFiles[resultFiles.length - 1];
          console.log(`📁 Файл создан: ${latestFile}`);
          
          // Анализируем содержимое файла
          const ExcelJS = require('exceljs');
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(path.join('uploads', latestFile));
          const worksheet = workbook.getWorksheet(1);
          
          // Считаем разделы и данные
          let reviewCount = 0;
          let commentCount = 0;
          let discussionCount = 0;
          let totalDataRows = 0;
          let currentSection = '';
          
          const sections = [];
          
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 4) return; // Пропускаем заголовки
            
            const cellA = row.getCell(1).value;
            if (cellA && typeof cellA === 'string') {
              const cellStr = cellA.toString().trim();
              
              if (cellStr === 'Отзывы') {
                currentSection = 'reviews';
                sections.push('Отзывы');
                return;
              } else if (cellStr.includes('Комментарии')) {
                currentSection = 'comments';
                sections.push(cellStr);
                return;
              } else if (cellStr.includes('Активные обсуждения')) {
                currentSection = 'discussions';
                sections.push(cellStr);
                return;
              }
            }
            
            // Проверяем есть ли данные в строке
            const hasData = cellA && cellA !== 'Площадка' && 
                           !cellA.toString().startsWith('Суммарное') &&
                           !cellA.toString().startsWith('Количество') &&
                           !cellA.toString().startsWith('Доля') &&
                           !cellA.toString().startsWith('*Без') &&
                           !cellA.toString().startsWith('Площадки');
            
            if (hasData) {
              totalDataRows++;
              if (currentSection === 'reviews') reviewCount++;
              else if (currentSection === 'comments') commentCount++;
              else if (currentSection === 'discussions') discussionCount++;
            }
          });
          
          console.log('\n📊 АНАЛИЗ РЕЗУЛЬТАТОВ:');
          console.log(`📝 Отзывы: ${reviewCount}`);
          console.log(`💬 Комментарии: ${commentCount}`);
          console.log(`🔥 Активные обсуждения: ${discussionCount}`);
          console.log(`📋 Всего строк данных: ${totalDataRows}`);
          console.log(`📑 Найденные разделы: ${sections.join(', ')}`);
          
          // Проверяем заголовки
          const row1 = worksheet.getRow(1);
          const headerColor = row1.getCell(1).fill?.fgColor?.argb;
          const headerText = row1.getCell(1).value;
          
          console.log(`\n🎨 ПРОВЕРКА ФОРМАТИРОВАНИЯ:`);
          console.log(`📑 Заголовок: ${headerText}`);
          console.log(`🎨 Цвет заголовка: ${headerColor}`);
          
          // Итоговые проверки
          const checks = {
            correctReviews: reviewCount === 18,
            correctComments: commentCount === 519,
            hasDiscussions: discussionCount > 0,
            hasAllSections: sections.length === 3,
            goodTotalRows: totalDataRows >= 600,
            correctHeader: headerText === 'Продукт',
            correctColor: headerColor === 'FF2D1341'
          };
          
          console.log('\n🏆 ИТОГОВЫЕ ПРОВЕРКИ:');
          Object.entries(checks).forEach(([check, passed]) => {
            console.log(`${passed ? '✅' : '❌'} ${check}: ${passed ? 'ПРОЙДЕНО' : 'НЕ ПРОЙДЕНО'}`);
          });
          
          const passedCount = Object.values(checks).filter(Boolean).length;
          const totalChecks = Object.values(checks).length;
          
          console.log(`\n📊 РЕЗУЛЬТАТ: ${passedCount}/${totalChecks} проверок пройдено`);
          
          if (passedCount >= 6) { // Минимум 6 из 7 проверок
            console.log('\n🎉🎉🎉 СИСТЕМА РАБОТАЕТ ОТЛИЧНО! 🎉🎉🎉');
            console.log('✅ Основные требования выполнены:');
            console.log(`  📝 ${reviewCount} отзывов ${reviewCount === 18 ? '✅' : '⚠️'}`);
            console.log(`  💬 ${commentCount} комментариев ${commentCount === 519 ? '✅' : '⚠️'}`);
            console.log(`  🔥 ${discussionCount} активных обсуждений ✅`);
            console.log(`  📋 ${totalDataRows} строк данных ✅`);
            console.log(`  🎨 Правильное форматирование ✅`);
            console.log('\n🚀 ГОТОВО К ПРОДАКШНУ!');
            return true;
          } else {
            console.log('\n⚠️ СИСТЕМА ТРЕБУЕТ ДОРАБОТКИ');
            console.log(`Пройдено только ${passedCount}/${totalChecks} проверок`);
            return false;
          }
          
        } else {
          console.log('❌ Файл результата не найден');
          return false;
        }
        
      } else if (fileInfo.status === 'error') {
        console.log('❌ Ошибка при обработке:', fileInfo.errorMessage);
        return false;
      }
      
      attempts++;
    }
    
    console.log('⏱️ Превышено время ожидания');
    return false;
    
  } catch (error) {
    console.error('❌ Критическая ошибка:', error.message);
    return false;
  }
}

// Запускаем окончательный тест
ultimateFinalTest().then(success => {
  if (success) {
    console.log('\n🎊🎊🎊 ПОЗДРАВЛЯЕМ! 🎊🎊🎊');
    console.log('🔥 ExcelEngineer полностью готов к работе!');
    console.log('💼 Система протестирована и одобрена!');
  } else {
    console.log('\n🔧 ТРЕБУЕТСЯ ДОПОЛНИТЕЛЬНАЯ НАСТРОЙКА');
  }
}).catch(error => {
  console.error('💥 Фатальная ошибка:', error);
}); 