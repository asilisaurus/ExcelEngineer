/**
 * 🧪 ТЕСТ УЛУЧШЕНИЙ GOOGLE APPS SCRIPT
 * Проверка исправлений на основе реальных данных
 */

function testImprovedProcessor() {
  console.log('🧪 ТЕСТ УЛУЧШЕНИЙ GOOGLE APPS SCRIPT');
  console.log('=' .repeat(60));
  
  try {
    // Создаем экземпляр улучшенного процессора
    const processor = new FinalMonthlyReportProcessor();
    
    // Тестовые данные на основе реальной структуры
    const testData = [
      // Строки 1-3: Мета-информация
      ['Мета-информация 1'],
      ['Мета-информация 2'],
      ['Мета-информация 3'],
      
      // Строка 4: Заголовки
      ['Площадка', 'Ссылка', 'Текст', 'Дата', 'Автор', 'F', 'G', 'H', 'I', 'J', 'K', 'Просмотры', 'Тип'],
      
      // Строка 5: Заголовок секции
      ['Отзывы'],
      
      // Строки 6+: Реальные данные
      ['otzovik.com', 'https://otzovik.com/reviews/fortedetrim/', 'Отличный витамин D! Принимаю уже месяц, чувствую себя намного лучше.', '07.03.2025', 'Лилушан', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 100, 'ОС'],
      ['irecommend.ru', 'https://irecommend.ru/content/fortedetrim', 'Витамин Д организму идет на пользу. У меня лишний вес, сахар на границе нормы.', '10.03.2025', 'Marinasmk', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 150, 'ОС'],
      ['www.otzyvru.com', 'https://www.otzyvru.com/fortedetrim', 'Достойный витамин Д. Рабочий. Принимал Фортедетрим 4000 по назначению эндокринолога.', '06.03.2025', 'Антон Д.', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 200, 'ОС'],
      
      // Заголовок секции комментариев
      ['Комментарии'],
      
      // Данные комментариев
      ['otzovik.com', 'https://otzovik.com/comments/fortedetrim/', 'Интересный отзыв, спасибо за информацию!', '08.03.2025', 'Комментатор1', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 50, 'ЦС'],
      ['irecommend.ru', 'https://irecommend.ru/comments/fortedetrim', 'Согласен с автором, препарат действительно эффективный.', '09.03.2025', 'Комментатор2', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 75, 'ЦС'],
      
      // Заголовок секции обсуждений
      ['Обсуждения'],
      
      // Данные обсуждений
      ['forum.ru', 'https://forum.ru/discussion/fortedetrim', 'Обсуждение эффективности витамина D при различных заболеваниях.', '11.03.2025', 'Участник1', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 300, 'ЦС'],
      ['vk.com', 'https://vk.com/discussion/fortedetrim', 'Группа для обсуждения витамина D и его влияния на здоровье.', '12.03.2025', 'Участник2', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 'Нет данных', 250, 'ЦС']
    ];
    
    console.log('📊 Тестовые данные загружены');
    console.log(`📋 Всего строк: ${testData.length}`);
    
    // Анализируем структуру данных
    processor.analyzeDataStructure(testData);
    
    // Обрабатываем данные
    const processedData = processor.processData(testData);
    
    console.log('\n📊 РЕЗУЛЬТАТЫ ОБРАБОТКИ:');
    console.log('-' .repeat(40));
    console.log(`📈 Отзывы: ${processedData.statistics.totalReviews}`);
    console.log(`💬 Целевые: ${processedData.statistics.totalTargeted}`);
    console.log(`🔥 Социальные: ${processedData.statistics.totalSocial}`);
    console.log(`👀 Всего просмотров: ${processedData.statistics.totalViews}`);
    
    // Проверяем детали
    console.log('\n📋 ДЕТАЛИ ОБРАБОТКИ:');
    console.log('Отзывы:');
    processedData.reviews.forEach((review, i) => {
      console.log(`  ${i + 1}. ${review.platform} - ${review.author} (${review.date})`);
    });
    
    console.log('\nЦелевые:');
    processedData.targeted.forEach((targeted, i) => {
      console.log(`  ${i + 1}. ${targeted.platform} - ${targeted.author} (${targeted.date})`);
    });
    
    console.log('\nСоциальные:');
    processedData.social.forEach((social, i) => {
      console.log(`  ${i + 1}. ${social.platform} - ${social.author} (${social.date})`);
    });
    
    // Валидация результатов
    const expectedReviews = 3;
    const expectedTargeted = 2;
    const expectedSocial = 2;
    
    const isValid = (
      processedData.statistics.totalReviews === expectedReviews &&
      processedData.statistics.totalTargeted === expectedTargeted &&
      processedData.statistics.totalSocial === expectedSocial
    );
    
    console.log('\n✅ ВАЛИДАЦИЯ:');
    console.log(`Ожидалось: ${expectedReviews}/${expectedTargeted}/${expectedSocial}`);
    console.log(`Получено: ${processedData.statistics.totalReviews}/${processedData.statistics.totalTargeted}/${processedData.statistics.totalSocial}`);
    console.log(`Результат: ${isValid ? '✅ ТЕСТ ПРОЙДЕН' : '❌ ТЕСТ НЕ ПРОЙДЕН'}`);
    
    return {
      success: isValid,
      data: processedData,
      expected: { reviews: expectedReviews, targeted: expectedTargeted, social: expectedSocial }
    };
    
  } catch (error) {
    console.error('❌ Ошибка тестирования:', error);
    return { success: false, error: error.toString() };
  }
}

// Запуск теста
function runTest() {
  const result = testImprovedProcessor();
  
  console.log('\n🏁 ТЕСТИРОВАНИЕ ЗАВЕРШЕНО');
  console.log('=' .repeat(60));
  
  if (result.success) {
    console.log('🎉 УЛУЧШЕНИЯ РАБОТАЮТ КОРРЕКТНО!');
    console.log('🚀 Google Apps Script готов к использованию!');
  } else {
    console.log('❌ Требуются дополнительные исправления');
    if (result.error) {
      console.log('Ошибка:', result.error);
    }
  }
  
  return result;
} 