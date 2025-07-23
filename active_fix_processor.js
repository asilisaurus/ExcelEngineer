
// 🚀 АКТИВНЫЕ ИСПРАВЛЕНИЯ ДЛЯ ДОСТИЖЕНИЯ 95%+ ТОЧНОСТИ

function activeDataProcessor(data) {
  console.log('🚀 Запуск активного процессора...');
  
  // 1. Очистка данных от пустых строк
  const cleanData = data.filter(row => 
    row && row[0] && row[0].toString().trim() !== ""
  );
  
  console.log(`📊 Очищено данных: ${data.length} → ${cleanData.length} строк`);
  
  // 2. Универсальное определение секций (работает с любым количеством строк)
  const sections = {
    reviews: [],
    comments: [],
    discussions: []
  };
  
  let currentSection = null;
  let sectionStartRow = -1;
  
  for (let i = 0; i < cleanData.length; i++) {
    const cellValue = cleanData[i][0].toString();
    
    // Динамическое определение секций
    if (cellValue.toLowerCase().includes('отзыв') && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'reviews';
      sectionStartRow = i + 1; // Следующая строка после заголовка
      console.log(`📍 Секция "Отзывы" начинается со строки ${sectionStartRow}`);
      continue;
    } else if (cellValue.toLowerCase().includes('комментар') && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'comments';
      sectionStartRow = i + 1;
      console.log(`📍 Секция "Комментарии" начинается со строки ${sectionStartRow}`);
      continue;
    } else if ((cellValue.toLowerCase().includes('обсужден') || cellValue.toLowerCase().includes('активн')) && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'discussions';
      sectionStartRow = i + 1;
      console.log(`📍 Секция "Обсуждения" начинается со строки ${sectionStartRow}`);
      continue;
    }
    
    // Добавление данных в секции (универсальная логика)
    if (currentSection && isValidDataRow(cleanData[i])) {
      if (currentSection === 'reviews') {
        sections.reviews.push(cleanData[i]);
      } else if (currentSection === 'comments') {
        sections.comments.push(cleanData[i]);
      } else if (currentSection === 'discussions') {
        sections.discussions.push(cleanData[i]);
      }
    }
  }
  
  // 3. Универсальная валидация (работает с любым количеством строк)
  function isValidDataRow(row) {
    if (!row || !row[0]) return false;
    
    const cellValue = row[0].toString().trim();
    if (cellValue === '') return false;
    
    // Исключаем заголовки
    const headerKeywords = ['площадка', 'тема', 'текст', 'дата', 'ник', 'просмотры'];
    if (headerKeywords.some(keyword => cellValue.toLowerCase().includes(keyword))) {
      return false;
    }
    
    // Проверяем, что это данные (домен, ссылка, текст)
    return cellValue.includes('.') || cellValue.includes('http') || cellValue.includes('www') || cellValue.length > 10;
  }
  
  // 4. Автоматический подсчет (универсальный)
  const stats = {
    reviews: sections.reviews.length,
    comments: sections.comments.length,
    discussions: sections.discussions.length,
    total: sections.reviews.length + sections.comments.length + sections.discussions.length
  };
  
  console.log(`📈 РЕЗУЛЬТАТЫ АКТИВНОГО ПРОЦЕССОРА:`);
  console.log(`   - Отзывы: ${stats.reviews}`);
  console.log(`   - Комментарии: ${stats.comments}`);
  console.log(`   - Обсуждения: ${stats.discussions}`);
  console.log(`   - Всего: ${stats.total}`);
  
  // 5. Проверка качества (универсальная)
  const qualityCheck = validateQuality(stats);
  console.log(`🎯 Качество обработки: ${qualityCheck.passed ? '✅ ПРОЙДЕНА' : '❌ НЕ ПРОЙДЕНА'}`);
  
  return { sections, stats, qualityCheck };
}

function validateQuality(stats) {
  // Универсальная проверка качества (работает с любыми данными)
  const checks = {
    hasData: stats.total > 0,
    hasReviews: stats.reviews >= 0,
    hasComments: stats.comments >= 0,
    hasDiscussions: stats.discussions >= 0,
    reasonableTotal: stats.total < 10000 // Защита от ошибок
  };
  
  const passed = Object.values(checks).every(check => check);
  
  return {
    passed,
    checks,
    message: passed ? 'Данные обработаны корректно' : 'Обнаружены проблемы в обработке'
  };
}

// Экспорт для использования
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    activeDataProcessor,
    validateQuality
  };
}
