
// ИСПРАВЛЕННЫЙ ПРОЦЕССОР - УЛУЧШЕННАЯ ВЕРСИЯ

function processDataWithFixes(data) {
  // 1. Удаляем пустые строки
  const cleanData = data.filter(row => 
    row && row[0] && row[0].toString().trim() !== ""
  );
  
  // 2. Правильное определение секций
  let reviews = [];
  let comments = [];
  let discussions = [];
  
  let currentSection = null;
  
  for (let i = 0; i < cleanData.length; i++) {
    const cellValue = cleanData[i][0].toString();
    
    // Определяем секции (исключаем заголовки)
    if (cellValue.includes('Отзывы') && !cellValue.includes('Количество')) {
      currentSection = 'reviews';
      continue; // Пропускаем заголовок
    } else if (cellValue.includes('Комментарии') && !cellValue.includes('Количество')) {
      currentSection = 'comments';
      continue; // Пропускаем заголовок
    } else if (cellValue.includes('Активные обсуждения') || cellValue.includes('Обсуждения')) {
      currentSection = 'discussions';
      continue; // Пропускаем заголовок
    }
    
    // Добавляем данные в соответствующую секцию
    if (currentSection === 'reviews' && cellValue.includes('.')) {
      reviews.push(cleanData[i]);
    } else if (currentSection === 'comments' && cellValue.includes('.')) {
      comments.push(cleanData[i]);
    } else if (currentSection === 'discussions' && cellValue.includes('.')) {
      discussions.push(cleanData[i]);
    }
  }
  
  // 3. Проверка полноты данных
  const total = reviews.length + comments.length + discussions.length;
  console.log(`📊 Результаты: Отзывы: ${reviews.length}, Комментарии: ${comments.length}, Обсуждения: ${discussions.length}`);
  console.log(`📋 Всего: ${total}`);
  
  return { reviews, comments, discussions, total };
}
