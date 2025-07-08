const fs = require('fs');

async function testWebInterface() {
  console.log('🌐 ТЕСТИРОВАНИЕ ВЕБ-ИНТЕРФЕЙСА');
  console.log('='.repeat(50));
  
  try {
    // Проверяем доступность сервера
    console.log('🔍 Проверяем доступность сервера...');
    const response = await fetch('http://localhost:5000');
    
    if (response.ok) {
      console.log('✅ Сервер доступен');
      console.log(`📊 Статус: ${response.status}`);
      console.log(`📝 Content-Type: ${response.headers.get('content-type')}`);
      
      // Проверяем API
      console.log('\n🔍 Проверяем API...');
      const apiResponse = await fetch('http://localhost:5000/api/files');
      
      if (apiResponse.ok) {
        const files = await apiResponse.json();
        console.log('✅ API работает');
        console.log(`📁 Найдено файлов: ${files.length}`);
      } else {
        console.log('❌ API недоступен');
      }
      
      // Получаем информацию о сети
      console.log('\n🌐 ИНФОРМАЦИЯ ДЛЯ ДОСТУПА:');
      console.log('📍 Локальный доступ: http://localhost:5000');
      console.log('📍 Сеть (найдите ваш IP): http://[IP_ADDRESS]:5000');
      
      console.log('\n💡 КАК НАЙТИ IP:');
      console.log('Windows: ipconfig');
      console.log('Mac/Linux: ifconfig | grep "inet "');
      
      console.log('\n🎯 ВОЗМОЖНОСТИ ВЕБ-ИНТЕРФЕЙСА:');
      console.log('✅ Загрузка Excel файлов (Drag & Drop)');
      console.log('✅ Импорт из Google Sheets');
      console.log('✅ Отслеживание прогресса в реальном времени');
      console.log('✅ Просмотр статистики обработки');
      console.log('✅ Скачивание результатов');
      console.log('✅ История обработок');
      console.log('✅ Обработка ошибок с подробными сообщениями');
      
      console.log('\n🚀 ИНТЕРФЕЙС ГОТОВ К ИСПОЛЬЗОВАНИЮ!');
      return true;
      
    } else {
      console.log(`❌ Сервер недоступен. Статус: ${response.status}`);
      console.log('💡 Убедитесь что запущен: npm run dev');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Ошибка подключения:', error.message);
    console.log('\n💡 ВОЗМОЖНЫЕ ПРИЧИНЫ:');
    console.log('• Сервер не запущен (npm run dev)');
    console.log('• Порт 5000 занят другим приложением');
    console.log('• Проблемы с сетью');
    
    console.log('\n🔧 РЕШЕНИЕ:');
    console.log('1. cd ExcelEngineer');
    console.log('2. npm run dev');
    console.log('3. Ждите "serving on port 5000"');
    console.log('4. Откройте http://localhost:5000');
    
    return false;
  }
}

// Дополнительная функция для проверки производительности
async function performanceTest() {
  console.log('\n⚡ ТЕСТ ПРОИЗВОДИТЕЛЬНОСТИ');
  console.log('='.repeat(30));
  
  try {
    const start = Date.now();
    const response = await fetch('http://localhost:5000/api/files');
    const duration = Date.now() - start;
    
    if (response.ok) {
      console.log(`✅ API отвечает за ${duration}мс`);
      if (duration < 100) {
        console.log('🚀 Отличная скорость!');
      } else if (duration < 500) {
        console.log('✅ Хорошая скорость');
      } else {
        console.log('⚠️ Медленный отклик');
      }
    }
  } catch (error) {
    console.log('❌ Тест производительности не удался');
  }
}

// Запускаем тесты
testWebInterface()
  .then(success => {
    if (success) {
      return performanceTest();
    }
  })
  .catch(error => {
    console.error('💥 Критическая ошибка:', error);
  });