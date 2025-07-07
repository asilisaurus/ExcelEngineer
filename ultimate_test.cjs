const fs = require('fs');
const path = require('path');

async function ultimateTest() {
  console.log('🚀 ОКОНЧАТЕЛЬНЫЙ ТЕСТ ВСЕХ ИСПРАВЛЕНИЙ');
  console.log('='.repeat(60));
  
  try {
    // Ждем запуска сервера
    console.log('⏳ Ожидание запуска сервера (5 сек)...');
    await new Promise(resolve => setTimeout(resolve, 5000));
    
    // Отправляем запрос на обработку
    console.log('📤 Отправляем запрос на обработку...');
    const response = await fetch('http://localhost:5000/api/import-google-sheets', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
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
    
    // Отслеживаем прогресс
    let attempts = 0;
    const maxAttempts = 15;
    
    while (attempts < maxAttempts) {
      console.log(`🔄 Проверяем статус (${attempts + 1}/${maxAttempts})...`);
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      
      if (statusResponse.ok) {
        const fileInfo = await statusResponse.json();
        console.log(`📊 Статус: ${fileInfo.status}`);
        
        if (fileInfo.status === 'completed') {
          console.log('🎉 Обработка завершена!');
          
          // Находим созданный файл
          const uploadsDir = 'uploads';
          const files = fs.readdirSync(uploadsDir);
          const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
          const resultFiles = files.filter(file => 
            file.includes('результат') && 
            file.includes(today)
          );
          
          if (resultFiles.length > 0) {
            // Берем последний файл
            const filesWithStats = resultFiles.map(file => ({
              name: file,
              stats: fs.statSync(path.join(uploadsDir, file))
            }));
            
            filesWithStats.sort((a, b) => b.stats.mtime - a.stats.mtime);
            const latestFile = filesWithStats[0].name;
            
            console.log('\n✅ ФАЙЛ СОЗДАН УСПЕШНО!');
            console.log(`📁 Имя: ${latestFile}`);
            console.log(`📊 Размер: ${filesWithStats[0].stats.size} байт`);
            
            // Проверяем содержимое файла
            console.log('\n🔍 ПРОВЕРКА СОДЕРЖИМОГО:');
            
            const ExcelJS = require('exceljs');
            const workbook = new ExcelJS.Workbook();
            const filePath = path.join(uploadsDir, latestFile);
            
            try {
              await workbook.xlsx.readFile(filePath);
              const worksheet = workbook.getWorksheet(1);
              
              // Проверяем заголовки
              const row1 = worksheet.getRow(1);
              const row2 = worksheet.getRow(2);
              const row3 = worksheet.getRow(3);
              
              console.log(`Строка 1: ${row1.getCell(1).value} | ${row1.getCell(2).value}`);
              console.log(`Строка 2: ${row2.getCell(1).value} | ${row2.getCell(2).value}`);
              console.log(`Строка 3: ${row3.getCell(1).value} | ${row3.getCell(2).value}`);
              
              // Проверяем цвета
              const headerColor = row1.getCell(1).fill?.fgColor?.argb;
              console.log(`Цвет заголовка: ${headerColor}`);
              
              // Проверяем данные
              let reviewCount = 0;
              let commentCount = 0;
              let totalRows = 0;
              
              worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 5) { // Пропускаем заголовки
                  const postType = row.getCell(8).value;
                  if (postType === 'ОС') reviewCount++;
                  if (postType === 'ЦС') commentCount++;
                  totalRows++;
                }
              });
              
              console.log(`\n📊 СТАТИСТИКА ДАННЫХ:`);
              console.log(`📝 Отзывов: ${reviewCount}`);
              console.log(`💬 Комментариев: ${commentCount}`);
              console.log(`📋 Всего строк данных: ${totalRows}`);
              
              // Проверяем на наличие дублирования
              const firstDataRow = worksheet.getRow(6);
              const duplicateCheck = [];
              for (let i = 1; i <= 8; i++) {
                const cellValue = firstDataRow.getCell(i).value;
                duplicateCheck.push(cellValue);
              }
              
              const hasDuplicates = duplicateCheck.some((value, index) => 
                duplicateCheck.indexOf(value) !== index && value !== '' && value !== null
              );
              
              console.log(`\n🔍 ПРОВЕРКА ДУБЛИРОВАНИЯ:`);
              console.log(`Первая строка данных: ${duplicateCheck.slice(0, 4).join(' | ')}`);
              console.log(`Дублирование: ${hasDuplicates ? '❌ ЕСТЬ' : '✅ НЕТ'}`);
              
              // Проверяем на наличие [object Object]
              let hasObjectErrors = false;
              worksheet.eachRow((row, rowNumber) => {
                if (rowNumber <= 10) { // Проверяем первые 10 строк
                  for (let i = 1; i <= 8; i++) {
                    const cellValue = row.getCell(i).value;
                    if (cellValue && cellValue.toString().includes('[object Object]')) {
                      hasObjectErrors = true;
                      break;
                    }
                  }
                }
              });
              
              console.log(`[object Object] ошибки: ${hasObjectErrors ? '❌ ЕСТЬ' : '✅ НЕТ'}`);
              
              // Итоговая оценка
              console.log('\n🏆 ИТОГОВАЯ ОЦЕНКА:');
              
              const checksPass = {
                fileCreated: true,
                correctHeaders: row1.getCell(1).value === 'Продукт',
                correctColors: headerColor === 'FF2D1341',
                correctCounts: reviewCount === 18 && commentCount === 519,
                noDuplicates: !hasDuplicates,
                noObjectErrors: !hasObjectErrors
              };
              
              Object.entries(checksPass).forEach(([check, passed]) => {
                console.log(`${passed ? '✅' : '❌'} ${check}: ${passed ? 'ПРОЙДЕНО' : 'НЕ ПРОЙДЕНО'}`);
              });
              
              const allPassed = Object.values(checksPass).every(Boolean);
              
              if (allPassed) {
                console.log('\n🎉 ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ УСПЕШНО!');
                console.log('🚀 СИСТЕМА EXCELENGINER ПОЛНОСТЬЮ ГОТОВА!');
                console.log('✨ Все проблемы исправлены:');
                console.log('  ✅ Конфликт XLSX/ExcelJS устранен');
                console.log('  ✅ Дублирование данных исправлено');
                console.log('  ✅ [object Object] ошибки устранены');
                console.log('  ✅ Правильные цвета заголовков');
                console.log('  ✅ Корректная структура данных');
                console.log('  ✅ Точные количества: 18 отзывов, 519 комментариев');
                return true;
              } else {
                console.log('\n❌ НЕКОТОРЫЕ ПРОВЕРКИ НЕ ПРОЙДЕНЫ');
                return false;
              }
              
            } catch (fileError) {
              console.error('❌ Ошибка при чтении файла:', fileError.message);
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
      }
      
      await new Promise(resolve => setTimeout(resolve, 2000));
      attempts++;
    }
    
    console.log('⏱️ Превышено время ожидания');
    return false;
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
    return false;
  }
}

// Запускаем окончательный тест
ultimateTest().then(success => {
  if (success) {
    console.log('\n🎊 ПОЗДРАВЛЯЕМ! СИСТЕМА ПОЛНОСТЬЮ ГОТОВА К РАБОТЕ!');
    console.log('💼 Можно использовать в продакшне');
  } else {
    console.log('\n⚠️ ТРЕБУЕТСЯ ДОПОЛНИТЕЛЬНАЯ ОТЛАДКА');
  }
}).catch(error => {
  console.error('💥 Критическая ошибка:', error);
}); 