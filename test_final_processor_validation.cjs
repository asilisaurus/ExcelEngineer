const fs = require('fs');
const path = require('path');

// Импортируем тестировщик и процессор
const { FinalGoogleAppsScriptTester } = require('./google-apps-script-testing-final.js');

async function runFinalValidation() {
  console.log('🔍 ФИНАЛЬНАЯ ПРОВЕРКА ИСПРАВЛЕННОГО ПРОЦЕССОРА');
  console.log('=' .repeat(60));
  
  try {
    // Создаем экземпляр тестировщика
    const tester = new FinalGoogleAppsScriptTester();
    
    console.log('📋 Запуск полного тестирования...');
    
    // Запускаем тестирование
    const results = await tester.runFinalTesting();
    
    console.log('\n📊 РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ:');
    console.log('=' .repeat(40));
    
    if (results && results.length > 0) {
      results.forEach((result, index) => {
        console.log(`\n📅 Тест ${index + 1}:`);
        console.log(`   Месяц: ${result.month || 'Не указан'}`);
        console.log(`   Статус: ${result.success ? '✅ УСПЕШНО' : '❌ ОШИБКА'}`);
        
        if (result.stats) {
          console.log('   Статистика:');
          console.log(`   - Отзывы: ${result.stats.reviews || 0}`);
          console.log(`   - Комментарии: ${result.stats.comments || 0}`);
          console.log(`   - Обсуждения: ${result.stats.discussions || 0}`);
          console.log(`   - Всего: ${result.stats.total || 0}`);
        }
        
        if (result.accuracy) {
          console.log(`   Точность: ${result.accuracy.toFixed(2)}%`);
        }
        
        if (result.errors && result.errors.length > 0) {
          console.log('   Ошибки:');
          result.errors.forEach(error => {
            console.log(`   - ${error}`);
          });
        }
      });
      
      // Общая статистика
      const successCount = results.filter(r => r.success).length;
      const totalCount = results.length;
      const overallSuccess = (successCount / totalCount) * 100;
      
      console.log('\n📈 ОБЩАЯ СТАТИСТИКА:');
      console.log(`   Успешных тестов: ${successCount}/${totalCount}`);
      console.log(`   Общая успешность: ${overallSuccess.toFixed(2)}%`);
      console.log(`   Статус: ${overallSuccess >= 95 ? '✅ ОТЛИЧНО' : '⚠️ ТРЕБУЕТ ДОРАБОТКИ'}`);
      
    } else {
      console.log('⚠️ Результаты тестирования не получены');
    }
    
  } catch (error) {
    console.error('❌ Ошибка при выполнении тестирования:', error);
    
    // Попробуем альтернативный подход - прямое тестирование
    console.log('\n🔄 Попытка альтернативного тестирования...');
    await alternativeTest();
  }
}

async function alternativeTest() {
  console.log('🧪 Альтернативное тестирование исправленного процессора');
  
  try {
    // Загружаем исправленный процессор
    const processorPath = path.join(__dirname, 'google-apps-script-processor-fixed-boundaries.js');
    if (fs.existsSync(processorPath)) {
      console.log('✅ Исправленный процессор найден');
      
      // Проверяем, что исправление применено
      const processorContent = fs.readFileSync(processorPath, 'utf8');
      const hasFix = processorContent.includes('sectionStart = i + 1');
      
      console.log(`${hasFix ? '✅' : '❌'} Критическое исправление ${hasFix ? 'применено' : 'НЕ ПРИМЕНЕНО'}`);
      
      if (hasFix) {
        console.log('🎯 Исправление в строке findSectionBoundaries(): sectionStart = i + 1');
        console.log('📋 Это должно исправить проблему с включением заголовков в данные секций');
      }
      
    } else {
      console.log('❌ Исправленный процессор не найден');
    }
    
    // Проверяем наличие эталонных данных
    const referenceFiles = [
      'attached_assets/Фортедетрим_ORM_отчет_исходник_1751040742705.xlsx',
      'attached_assets/Фортедетрим_ORM_отчет_Март_2025_результат_1751040742705.xlsx'
    ];
    
    console.log('\n📁 Проверка эталонных файлов:');
    referenceFiles.forEach(file => {
      const exists = fs.existsSync(path.join(__dirname, file));
      console.log(`   ${exists ? '✅' : '❌'} ${file}`);
    });
    
  } catch (error) {
    console.error('❌ Ошибка альтернативного тестирования:', error);
  }
}

// Запуск
runFinalValidation().catch(console.error); 