const fs = require('fs');
const path = require('path');

async function superFinalTest() {
  console.log('🎊 СУПЕР-ФИНАЛЬНЫЙ ТЕСТ 🎊');
  console.log('🚀 Проверка всех исправлений...');
  
  try {
    // Ждем 4 секунды для запуска сервера
    await new Promise(resolve => setTimeout(resolve, 4000));
    
    // Отправляем запрос
    const response = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        url: 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing'
      })
    });
    
    const result = await response.json();
    console.log('✅ Запрос отправлен, ID:', result.fileId);
    
    // Ждем завершения
    for (let i = 0; i < 10; i++) {
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      const fileInfo = await statusResponse.json();
      
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
          
          const row1 = worksheet.getRow(1);
          const headerColor = row1.getCell(1).fill?.fgColor?.argb;
          
          // Подсчитываем данные
          let reviewCount = 0, commentCount = 0;
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 5) {
              const postType = row.getCell(8).value;
              if (postType === 'ОС') reviewCount++;
              if (postType === 'ЦС') commentCount++;
            }
          });
          
          console.log('📊 РЕЗУЛЬТАТЫ:');
          console.log(`📝 Отзывов: ${reviewCount}`);
          console.log(`💬 Комментариев: ${commentCount}`);
          console.log(`🎨 Цвет заголовка: ${headerColor}`);
          console.log(`📑 Заголовок: ${row1.getCell(1).value}`);
          
          // Итоговая оценка
          const allGood = 
            reviewCount === 18 && 
            commentCount === 519 && 
            headerColor === 'FF2D1341' && 
            row1.getCell(1).value === 'Продукт';
          
          if (allGood) {
            console.log('\n🎉🎉🎉 ВСЕ ИДЕАЛЬНО! 🎉🎉🎉');
            console.log('✅ СИСТЕМА ПОЛНОСТЬЮ ГОТОВА К РАБОТЕ!');
            console.log('🚀 18 отзывов, 519 комментариев');
            console.log('🎨 Правильные цвета и форматирование');
            console.log('📊 Корректная структура данных');
            console.log('💼 ГОТОВО К ПРОДАКШНУ!');
            return true;
          } else {
            console.log('❌ Некоторые проверки не пройдены');
            return false;
          }
        }
        break;
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
    return false;
  }
}

superFinalTest().then(success => {
  if (success) {
    console.log('\n🎊 ПОЗДРАВЛЯЕМ! 🎊');
    console.log('🔥 ExcelEngineer работает идеально!');
  } else {
    console.log('\n⚠️ Нужны дополнительные исправления');
  }
}); 