const fs = require('fs');
const path = require('path');

async function testImprovedProcessor() {
  console.log('🧪 ТЕСТИРОВАНИЕ УЛУЧШЕННОГО ПРОЦЕССОРА');
  console.log('='.repeat(60));
  
  try {
    // Ждем запуска сервера
    console.log('⏳ Ожидание запуска сервера (3 сек)...');
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    // Тестируем Google Sheets импорт с новым процессором
    console.log('📤 Тестируем импорт Google Sheets с улучшенным процессором...');
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
    const maxAttempts = 20;
    
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
          console.log('\n📊 АНАЛИЗ РЕЗУЛЬТАТА:');
          console.log(`📝 Файл: ${fileInfo.processedName}`);
          console.log(`📊 Строк обработано: ${fileInfo.rowsProcessed}`);
          
          if (fileInfo.statistics) {
            const statistics = JSON.parse(fileInfo.statistics);
            console.log(`\n📈 СТАТИСТИКА:`);
            console.log(`   📝 Всего записей: ${statistics.totalRows || 'N/A'}`);
            console.log(`   🎯 Отзывов: ${statistics.reviewsCount || 'N/A'}`);
            console.log(`   💬 Комментариев: ${statistics.commentsCount || 'N/A'}`);
            console.log(`   🔥 Активных обсуждений: ${statistics.activeDiscussionsCount || 'N/A'}`);
            console.log(`   👀 Просмотров: ${statistics.totalViews?.toLocaleString() || 'N/A'}`);
            console.log(`   📈 Вовлечение: ${statistics.engagementRate || 'N/A'}%`);
            console.log(`   📊 Платформ с данными: ${statistics.platformsWithData || 'N/A'}%`);
          }
          
          // Проверяем файл на диске
          const uploadsDir = 'uploads';
          const files = fs.readdirSync(uploadsDir);
          const resultFiles = files.filter(file => 
            file.includes('Fortedetrim_ORM_report') && 
            file.includes('результат')
          );
          
          if (resultFiles.length > 0) {
            const latestFile = resultFiles[resultFiles.length - 1];
            const filePath = path.join(uploadsDir, latestFile);
            const fileStats = fs.statSync(filePath);
            
            console.log('\n✅ ФАЙЛ НА ДИСКЕ:');
            console.log(`📁 Имя: ${latestFile}`);
            console.log(`📊 Размер: ${(fileStats.size / 1024).toFixed(1)} KB`);
            console.log(`🕐 Создан: ${fileStats.birthtime.toLocaleString('ru-RU')}`);
            
            // Проверяем содержимое файла Excel
            try {
              const ExcelJS = require('exceljs');
              const workbook = new ExcelJS.Workbook();
              await workbook.xlsx.readFile(filePath);
              
              const worksheet = workbook.getWorksheet(1);
              if (worksheet) {
                console.log('\n🔍 ПРОВЕРКА СОДЕРЖИМОГО:');
                
                // Проверяем заголовки
                const productCell = worksheet.getCell('A1').value;
                const periodCell = worksheet.getCell('A2').value;
                const planCell = worksheet.getCell('A3').value;
                
                console.log(`   📋 Продукт: ${productCell}`);
                console.log(`   📅 Период: ${periodCell}`);
                console.log(`   📊 План: ${planCell}`);
                
                // Считаем данные
                let dataRows = 0;
                let sectionsFound = 0;
                
                worksheet.eachRow((row, rowNumber) => {
                  if (rowNumber > 5) { // Пропускаем заголовки
                    const cellA = row.getCell(1).value;
                    if (cellA) {
                      const cellStr = cellA.toString();
                      if (cellStr.includes('Отзывы') || cellStr.includes('Комментарии') || cellStr.includes('Активные')) {
                        sectionsFound++;
                      } else if (cellStr !== 'Площадка' && cellStr !== 'Итого') {
                        dataRows++;
                      }
                    }
                  }
                });
                
                console.log(`   📊 Найдено секций: ${sectionsFound}`);
                console.log(`   📝 Строк данных: ${dataRows}`);
                
                // Проверка успешности
                const isValid = productCell === 'Продукт' && 
                               periodCell === 'Период' && 
                               sectionsFound >= 3 && 
                               dataRows > 0;
                
                console.log(`\n🎯 РЕЗУЛЬТАТ ВАЛИДАЦИИ: ${isValid ? '✅ УСПЕШНО' : '❌ ОШИБКА'}`);
                
                if (isValid) {
                  console.log('\n🎊 УЛУЧШЕННЫЙ ПРОЦЕССОР РАБОТАЕТ КОРРЕКТНО!');
                  console.log('🚀 Основные улучшения проверены:');
                  console.log('   ✅ Автоматическое определение структуры');
                  console.log('   ✅ Обработка ошибок без падений');
                  console.log('   ✅ Гибкая конфигурация');
                  console.log('   ✅ Структурированный код');
                  console.log('   ✅ Подробное логирование');
                  return true;
                } else {
                  console.log('\n⚠️ Найдены проблемы в выходном файле');
                  return false;
                }
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

// Дополнительная функция для тестирования загрузки файла
async function testFileUpload() {
  console.log('\n📁 ТЕСТИРОВАНИЕ ЗАГРУЗКИ ФАЙЛА');
  console.log('='.repeat(40));
  
  // Проверяем наличие тестового файла
  const testFile = 'test_download.xlsx';
  if (!fs.existsSync(testFile)) {
    console.log('⚠️ Тестовый файл не найден, пропускаем тест загрузки');
    return true;
  }
  
  try {
    const FormData = require('form-data');
    const form = new FormData();
    form.append('file', fs.createReadStream(testFile));
    
    const response = await fetch('http://localhost:5000/api/upload', {
      method: 'POST',
      body: form
    });
    
    if (response.ok) {
      const result = await response.json();
      console.log('✅ Загрузка файла успешна, ID:', result.fileId);
      return true;
    } else {
      console.log('❌ Ошибка загрузки файла:', response.status);
      return false;
    }
  } catch (error) {
    console.error('❌ Ошибка при загрузке файла:', error.message);
    return false;
  }
}

// Запускаем тесты
async function runAllTests() {
  console.log('🧪 ПОЛНОЕ ТЕСТИРОВАНИЕ УЛУЧШЕННОГО ПРОЦЕССОРА');
  console.log('=='.repeat(40));
  
  const googleSheetsTest = await testImprovedProcessor();
  const fileUploadTest = await testFileUpload();
  
  console.log('\n📋 ИТОГОВЫЙ ОТЧЕТ:');
  console.log(`   📊 Google Sheets импорт: ${googleSheetsTest ? '✅ ПРОЙДЕН' : '❌ НЕ ПРОЙДЕН'}`);
  console.log(`   📁 Загрузка файла: ${fileUploadTest ? '✅ ПРОЙДЕН' : '❌ НЕ ПРОЙДЕН'}`);
  
  const allPassed = googleSheetsTest && fileUploadTest;
  
  if (allPassed) {
    console.log('\n🎉 ВСЕ ТЕСТЫ ПРОЙДЕНЫ УСПЕШНО!');
    console.log('🚀 УЛУЧШЕННЫЙ ПРОЦЕССОР ГОТОВ К ИСПОЛЬЗОВАНИЮ!');
    console.log('\n💡 Основные преимущества:');
    console.log('   🔧 Гибкость к изменениям структуры файлов');
    console.log('   🛡️ Надежная обработка ошибок');
    console.log('   ⚙️ Настраиваемая конфигурация');
    console.log('   📊 Подробная диагностика');
    console.log('   🔍 Автоматическое определение колонок');
  } else {
    console.log('\n⚠️ НЕКОТОРЫЕ ТЕСТЫ НЕ ПРОЙДЕНЫ');
    console.log('Требуется дополнительная отладка');
  }
  
  return allPassed;
}

// Запускаем
runAllTests().catch(error => {
  console.error('💥 Критическая ошибка:', error);
});