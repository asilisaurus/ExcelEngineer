const axios = require('axios');

// Тестирование исправленного процессора
async function testFixedProcessor() {
  try {
    console.log('🧪 Тестирование исправленного процессора...');
    
    // URL Google Sheets для тестирования
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing';
    
    // Отправляем запрос на импорт
    const response = await axios.post('http://localhost:5000/api/import-google-sheets', {
      url: googleSheetsUrl
    });
    
    console.log('✅ Запрос отправлен успешно:', response.data);
    
    if (response.data.fileId) {
      console.log('📊 Отслеживаем прогресс обработки...');
      
      // Отслеживаем статус обработки
      let attempts = 0;
      const maxAttempts = 30; // 30 секунд
      
      while (attempts < maxAttempts) {
        await new Promise(resolve => setTimeout(resolve, 1000)); // Ждем 1 секунду
        
        try {
          const statusResponse = await axios.get(`http://localhost:5000/api/files/${response.data.fileId}`);
          const file = statusResponse.data;
          
          console.log(`📈 Статус: ${file.status}, Обработано строк: ${file.rowsProcessed || 0}`);
          
          if (file.status === 'completed') {
            console.log('🎉 Обработка завершена успешно!');
            console.log('📊 Статистика:', JSON.parse(file.statistics || '{}'));
            console.log('📁 Файл готов для скачивания:', file.processedName);
            
            // Пытаемся скачать файл
            try {
              const downloadResponse = await axios.get(
                `http://localhost:5000/api/files/${file.id}/download`,
                { responseType: 'arraybuffer' }
              );
              console.log('💾 Файл скачан, размер:', downloadResponse.data.length, 'байт');
              
              // Проверяем результат
              if (downloadResponse.data.length > 0) {
                console.log('✅ Исправленный процессор работает корректно!');
                console.log('📄 Проверьте файл в папке uploads/');
              } else {
                console.warn('⚠️ Файл пустой');
              }
            } catch (downloadError) {
              console.warn('⚠️ Ошибка скачивания:', downloadError.message);
            }
            
            break;
          } else if (file.status === 'error') {
            console.error('❌ Ошибка обработки:', file.errorMessage);
            break;
          }
          
          attempts++;
        } catch (statusError) {
          console.warn(`⚠️ Ошибка получения статуса (попытка ${attempts + 1}):`, statusError.message);
          attempts++;
        }
      }
      
      if (attempts >= maxAttempts) {
        console.warn('⏰ Превышено время ожидания обработки');
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка тестирования:', error.message);
    if (error.response) {
      console.error('📄 Ответ сервера:', error.response.data);
    }
  }
}

// Запускаем тест
testFixedProcessor();