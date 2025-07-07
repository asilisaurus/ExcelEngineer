const fs = require('fs');
const path = require('path');

async function testWithDiscussions() {
  console.log('🔥 ТЕСТ С АКТИВНЫМИ ОБСУЖДЕНИЯМИ 🔥');
  console.log('='.repeat(60));
  
  try {
    // Ждем запуска сервера
    console.log('⏳ Ожидание запуска сервера (6 сек)...');
    await new Promise(resolve => setTimeout(resolve, 6000));
    
    // Отправляем запрос
    console.log('📤 Отправляем запрос на обработку...');
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
    console.log('✅ Запрос отправлен, ID:', result.fileId);
    
    // Ждем завершения
    for (let i = 0; i < 15; i++) {
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      const fileInfo = await statusResponse.json();
      console.log(`🔄 Статус: ${fileInfo.status}`);
      
      if (fileInfo.status === 'completed') {
        console.log('🎉 Файл создан!');
        
        // Находим файл
        const files = fs.readdirSync('uploads');
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const resultFiles = files.filter(f => f.includes('результат') && f.includes(today));
        
        if (resultFiles.length > 0) {
          const latestFile = resultFiles[resultFiles.length - 1];
          console.log(`📁 Файл: ${latestFile}`);
          
          // Проверяем содержимое
          const ExcelJS = require('exceljs');
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(path.join('uploads', latestFile));
          const worksheet = workbook.getWorksheet(1);
          
          // Подсчитываем данные по разделам
          let reviewCount = 0;
          let commentCount = 0;
          let discussionCount = 0;
          let totalDataRows = 0;
          
          let currentSection = '';
          
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 4) return; // Пропускаем заголовки
            
            const cellA = row.getCell(1).value;
            if (cellA && typeof cellA === 'string') {
              if (cellA.includes('Отзывы')) {
                currentSection = 'reviews';
                return;
              } else if (cellA.includes('Комментарии')) {
                currentSection = 'comments';
                return;
              } else if (cellA.includes('Активные обсуждения')) {
                currentSection = 'discussions';
                return;
              }
            }
            
            // Если это строка с данными (не заголовок раздела)
            if (cellA && cellA !== 'Площадка' && !cellA.toString().startsWith('Суммарное')) {
              totalDataRows++;
              if (currentSection === 'reviews') reviewCount++;
              if (currentSection === 'comments') commentCount++;
              if (currentSection === 'discussions') discussionCount++;
            }
          });
          
          console.log('\n📊 РЕЗУЛЬТАТЫ АНАЛИЗА:');
          console.log(`📝 Отзывы: ${reviewCount}`);
          console.log(`💬 Комментарии Топ-20: ${commentCount}`);
          console.log(`🔥 Активные обсуждения: ${discussionCount}`);
          console.log(`📋 Общее количество строк данных: ${totalDataRows}`);
          
          // Проверяем целевые значения
          const checks = {
            reviewsCorrect: reviewCount === 18,
            commentsCorrect: commentCount === 519,
            hasDiscussions: discussionCount > 0,
            totalRows: totalDataRows > 600, // Должно быть больше 600 строк
            hasAllSections: discussionCount > 0 && reviewCount > 0 && commentCount > 0
          };
          
          console.log('\n🏆 ПРОВЕРКИ:');
          Object.entries(checks).forEach(([check, passed]) => {
            console.log(`${passed ? '✅' : '❌'} ${check}: ${passed ? 'ПРОЙДЕНО' : 'НЕ ПРОЙДЕНО'}`);
          });
          
          const allPassed = Object.values(checks).every(Boolean);
          
          if (allPassed) {
            console.log('\n🎉🎉🎉 ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ! 🎉🎉🎉');
            console.log('✅ Система корректно распределяет данные по разделам:');
            console.log(`  📝 Отзывы: ${reviewCount}`);
            console.log(`  💬 Комментарии Топ-20: ${commentCount}`);
            console.log(`  🔥 Активные обсуждения: ${discussionCount}`);
            console.log(`  📋 Общее количество: ${totalDataRows} строк`);
            console.log('🚀 ГОТОВО К ПРОДАКШНУ!');
            return true;
          } else {
            console.log('\n❌ Некоторые проверки не пройдены');
            console.log('Ожидаемо:');
            console.log('  📝 Отзывы: 18');
            console.log('  💬 Комментарии: 519');
            console.log('  🔥 Активные обсуждения: > 0');
            console.log('  📋 Общее: > 600 строк');
            return false;
          }
        }
        break;
      } else if (fileInfo.status === 'error') {
        console.log('❌ Ошибка:', fileInfo.errorMessage);
        return false;
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
    return false;
  }
}

testWithDiscussions().then(success => {
  if (success) {
    console.log('\n🎊 ТЕСТ УСПЕШНО ПРОЙДЕН!');
    console.log('🎯 Все разделы заполнены корректно!');
  } else {
    console.log('\n⚠️ ТЕСТ НЕ ПРОЙДЕН');
  }
}); 