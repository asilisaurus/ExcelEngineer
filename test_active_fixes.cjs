
// 🧪 ТЕСТОВЫЙ СКРИПТ ДЛЯ ПРОВЕРКИ ИСПРАВЛЕНИЙ

const { activeDataProcessor } = require('./active_fix_processor.js');
const XLSX = require('xlsx');

function testActiveFixes() {
  console.log('🧪 Тестирование активных исправлений...');
  
  try {
    // Загружаем данные
    const workbook = XLSX.readFile('uploads/Фортедетрим_ORM_отчет_исходник_1751040742705_результат_20250706.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`📊 Тестируем файл: ${sheetName}`);
    console.log(`📋 Всего строк: ${data.length}`);
    
    // Применяем активные исправления
    const result = activeDataProcessor(data);
    
    // Проверяем результаты
    const expected = { reviews: 22, comments: 20, discussions: 621 };
    const accuracy = {
      reviews: result.stats.reviews === expected.reviews ? 100 : 
               Math.abs(result.stats.reviews - expected.reviews) / expected.reviews * 100,
      comments: result.stats.comments === expected.comments ? 100 :
                Math.abs(result.stats.comments - expected.comments) / expected.comments * 100,
      discussions: result.stats.discussions === expected.discussions ? 100 :
                   Math.abs(result.stats.discussions - expected.discussions) / expected.discussions * 100
    };
    
    const overallAccuracy = (accuracy.reviews + accuracy.comments + accuracy.discussions) / 3;
    
    console.log(`\n🎯 РЕЗУЛЬТАТЫ ТЕСТА:`);
    console.log(`   - Отзывы: ${result.stats.reviews}/${expected.reviews} (${(100 - accuracy.reviews).toFixed(1)}% точность)`);
    console.log(`   - Комментарии: ${result.stats.comments}/${expected.comments} (${(100 - accuracy.comments).toFixed(1)}% точность)`);
    console.log(`   - Обсуждения: ${result.stats.discussions}/${expected.discussions} (${(100 - accuracy.discussions).toFixed(1)}% точность)`);
    console.log(`   - Общая точность: ${overallAccuracy.toFixed(1)}%`);
    
    if (overallAccuracy >= 95) {
      console.log(`🎉 УСПЕХ! Достигнута требуемая точность 95%+`);
    } else {
      console.log(`⚠️ Нужны дополнительные исправления для достижения 95%+`);
    }
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
  }
}

testActiveFixes();
