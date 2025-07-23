
// 🌍 ФИНАЛЬНЫЙ УНИВЕРСАЛЬНЫЙ ПРОЦЕССОР (работает с любым количеством строк)

function finalUniversalProcessor(data) {
  console.log('🌍 Запуск финального универсального процессора...');
  
  // 1. Очистка данных (универсальная)
  const cleanData = data.filter(row => 
    row && row[0] && row[0].toString().trim() !== ""
  );
  
  console.log(`📊 Очищено данных: ${data.length} → ${cleanData.length} строк`);
  
  // 2. Универсальное определение секций
  const sections = {
    reviews: [],
    comments: [],
    discussions: []
  };
  
  let currentSection = null;
  
  for (let i = 0; i < cleanData.length; i++) {
    const cellValue = cleanData[i][0].toString();
    
    // Динамическое определение секций (универсальное)
    if (cellValue.toLowerCase().includes('отзыв') && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'reviews';
      console.log(`📍 Секция "Отзывы" найдена в строке ${i + 1}`);
      continue;
    } else if (cellValue.toLowerCase().includes('комментар') && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'comments';
      console.log(`📍 Секция "Комментарии" найдена в строке ${i + 1}`);
      continue;
    } else if ((cellValue.toLowerCase().includes('обсужден') || cellValue.toLowerCase().includes('активн')) && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'discussions';
      console.log(`📍 Секция "Обсуждения" найдена в строке ${i + 1}`);
      continue;
    }
    
    // Универсальная валидация и добавление данных
    if (currentSection && isValidDataRow(cleanData[i], currentSection)) {
      sections[currentSection].push(cleanData[i]);
    }
  }
  
  // Универсальная валидация (работает с любым количеством строк)
  function isValidDataRow(row, sectionType) {
    if (!row || !row[0]) return false;
    
    const cellValue = row[0].toString().trim();
    if (cellValue === '') return false;
    
    // Исключаем заголовки
    const headerKeywords = ['площадка', 'тема', 'текст', 'дата', 'ник', 'просмотры'];
    if (headerKeywords.some(keyword => cellValue.toLowerCase().includes(keyword))) {
      return false;
    }
    
    // Разные критерии для разных секций
    if (sectionType === 'reviews') {
      return cellValue.includes('.') || cellValue.includes('http') || cellValue.includes('www');
    } else if (sectionType === 'comments') {
      // Более мягкие критерии для комментариев
      return cellValue.includes('.') || cellValue.includes('http') || cellValue.length > 5;
    } else if (sectionType === 'discussions') {
      // Включаем все валидные данные для обсуждений
      return cellValue.includes('.') || cellValue.includes('http') || cellValue.includes('www') || cellValue.length > 10;
    }
    
    return false;
  }
  
  // 3. Автоматический подсчет (универсальный)
  const stats = {
    reviews: sections.reviews.length,
    comments: sections.comments.length,
    discussions: sections.discussions.length,
    total: sections.reviews.length + sections.comments.length + sections.discussions.length
  };
  
  console.log(`📈 ФИНАЛЬНЫЕ РЕЗУЛЬТАТЫ:`);
  console.log(`   - Отзывы: ${stats.reviews}`);
  console.log(`   - Комментарии: ${stats.comments}`);
  console.log(`   - Обсуждения: ${stats.discussions}`);
  console.log(`   - Всего: ${stats.total}`);
  
  return { sections, stats };
}

// Экспорт для использования
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { finalUniversalProcessor };
}
