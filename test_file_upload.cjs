const fs = require('fs');
const path = require('path');
const FormData = require('form-data');

async function testFileUpload() {
  console.log('📤 ТЕСТ ЗАГРУЗКИ ФАЙЛА');
  console.log('='.repeat(50));
  
  try {
    // Находим исходный файл
    const sourceFile = path.join('uploads', 'Fortedetrim ORM report source.xlsx');
    
    if (!fs.existsSync(sourceFile)) {
      console.log('❌ Исходный файл не найден:', sourceFile);
      return;
    }
    
    console.log('📂 Найден исходный файл:', sourceFile);
    
    // Создаем FormData для загрузки файла
    const formData = new FormData();
    formData.append('file', fs.createReadStream(sourceFile), {
      filename: 'test_upload.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    console.log('📤 Загружаем файл через API...');
    
    // Отправляем файл
    const response = await fetch('http://localhost:5000/api/upload', {
      method: 'POST',
      body: formData
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    console.log('✅ Файл загружен успешно!');
    console.log(`📁 ID файла: ${result.fileId}`);
    
    // Ждем обработки
    let attempts = 0;
    const maxAttempts = 20;
    
    while (attempts < maxAttempts) {
      console.log(`🔄 Проверяем статус (${attempts + 1}/${maxAttempts})...`);
      
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      const statusResponse = await fetch(`http://localhost:5000/api/files/${result.fileId}`);
      const fileInfo = await statusResponse.json();
      
      if (fileInfo.status === 'completed') {
        console.log('🎉 Обработка файла завершена!');
        
        // Анализируем результат
        const files = fs.readdirSync('uploads');
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const resultFiles = files.filter(f => f.includes('результат') && f.includes(today));
        
        if (resultFiles.length > 0) {
          const latestFile = resultFiles[resultFiles.length - 1];
          console.log(`📁 Файл создан: ${latestFile}`);
          
          // Анализируем содержимое
          const ExcelJS = require('exceljs');
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(path.join('uploads', latestFile));
          const worksheet = workbook.getWorksheet(1);
          
          let reviewCount = 0;
          let commentCount = 0;
          let discussionCount = 0;
          let currentSection = '';
          
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 4) return;
            
            const cellA = row.getCell(1).value;
            if (cellA && typeof cellA === 'string') {
              const cellStr = cellA.toString().trim();
              
              if (cellStr === 'Отзывы') {
                currentSection = 'reviews';
                return;
              } else if (cellStr.includes('Комментарии')) {
                currentSection = 'comments';
                return;
              } else if (cellStr.includes('Активные обсуждения')) {
                currentSection = 'discussions';
                return;
              }
            }
            
            const hasData = cellA && cellA !== 'Площадка' && 
                           !cellA.toString().startsWith('Суммарное');
            
            if (hasData) {
              if (currentSection === 'reviews') reviewCount++;
              else if (currentSection === 'comments') commentCount++;
              else if (currentSection === 'discussions') discussionCount++;
            }
          });
          
          console.log('\n📊 РЕЗУЛЬТАТ ТЕСТИРОВАНИЯ ЗАГРУЗКИ:');
          console.log(`📝 Отзывы: ${reviewCount}`);
          console.log(`💬 Комментарии: ${commentCount}`);
          console.log(`🔥 Активные обсуждения: ${discussionCount}`);
          console.log(`📋 Всего: ${reviewCount + commentCount + discussionCount}`);
          
          if (reviewCount === 18 && commentCount === 519 && discussionCount > 0) {
            console.log('\n🎉🎉🎉 ЗАГРУЗКА ФАЙЛА РАБОТАЕТ ОТЛИЧНО! 🎉🎉🎉');
            console.log('✅ Сервер правильно обрабатывает загруженные файлы');
            console.log('✅ Все разделы заполнены корректно');
            console.log('🚀 СИСТЕМА ГОТОВА К РАБОТЕ!');
            return true;
          } else {
            console.log('\n❌ ЗАГРУЗКА ФАЙЛА НЕ РАБОТАЕТ');
            console.log('⚠️ Сервер дает неправильные результаты');
            return false;
          }
        }
        break;
      } else if (fileInfo.status === 'error') {
        console.log('❌ Ошибка при обработке:', fileInfo.errorMessage);
        return false;
      }
      
      attempts++;
    }
    
    console.log('⏱️ Превышено время ожидания');
    return false;
    
  } catch (error) {
    console.error('❌ Ошибка при загрузке файла:', error.message);
    return false;
  }
}

testFileUpload().then(success => {
  if (success) {
    console.log('\n🎊 ТЕСТ ЗАГРУЗКИ ПРОЙДЕН!');
    console.log('🎯 Система готова к продакшну!');
  } else {
    console.log('\n🔧 ТРЕБУЕТСЯ НАСТРОЙКА ЗАГРУЗКИ ФАЙЛОВ');
  }
}); 