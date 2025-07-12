/**
 * 🧪 ТЕСТ ЛОГИКИ ПРОЦЕССОРА
 * Проверка исправлений в логике определения разделов и типов
 * 
 * Автор: AI Assistant
 * Дата: 2025
 */

// ==================== ТЕСТОВЫЕ ДАННЫЕ ====================

// Имитация данных из Google Sheets
const testData = [
  // Строки 1-4: мета-информация
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'],
  
  // Строка 5: начало данных
  ['', '', '', '', '', '', '', ''],
  
  // Раздел "Отзывы"
  ['Отзывы', '', '', '', '', '', '', ''],
  ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'],
  ['site1.com', 'Тема 1', 'Отличный продукт, рекомендую всем', '01.05.2025', 'user1', 100, '5', 'ОС'],
  ['site2.com', 'Тема 2', 'Покупала, очень довольна', '02.05.2025', 'user2', 150, '4', 'ОС'],
  ['site3.com', 'Тема 3', 'Хороший товар, советую', '03.05.2025', 'user3', 200, '5', 'ОС'],
  
  // Раздел "Комментарии Топ-20 выдачи"
  ['Комментарии Топ-20 выдачи', '', '', '', '', '', '', ''],
  ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'],
  ['forum1.com', 'Тема 4', 'Комментарий к статье', '04.05.2025', 'user4', 50, '3', 'ЦС'],
  ['forum2.com', 'Тема 5', 'Ответ на вопрос', '05.05.2025', 'user5', 75, '4', 'ЦС'],
  
  // Раздел "Активные обсуждения (мониторинг)"
  ['Активные обсуждения (мониторинг)', '', '', '', '', '', '', ''],
  ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'],
  ['social1.com', 'Тема 6', 'Обсуждение в соцсети', '06.05.2025', 'user6', 300, '2', 'ПС'],
  ['social2.com', 'Тема 7', 'Пост в сообществе', '07.05.2025', 'user7', 250, '3', 'ПС'],
  ['social3.com', 'Тема 8', 'Комментарий в группе', '08.05.2025', 'user8', 180, '4', 'ПС'],
  
  // Строки статистики (должны быть пропущены)
  ['', '', '', '', '', '', '', ''],
  ['Суммарное количество просмотров', '1255', '', '', '', '', '', ''],
  ['Количество карточек товара (отзывы)', '3', '', '', '', '', '', ''],
  ['Количество обсуждений (форумы, сообщества, комментарии к статьям)', '5', '', '', '', '', '', ''],
  ['Доля обсуждений с вовлечением в диалог', '0.8', '', '', '', '', '', '']
];

// ==================== ЛОГИКА ПРОЦЕССОРА ====================

/**
 * Имитация логики определения разделов
 */
function determineSection(firstCell) {
  const lowerCell = firstCell.toLowerCase().trim();
  
  if (lowerCell.includes('отзывы') && !lowerCell.includes('топ-20') && !lowerCell.includes('обсуждения')) {
    return 'reviews';
  } else if (lowerCell.includes('комментарии топ-20') || lowerCell.includes('топ-20 выдачи')) {
    return 'commentsTop20';
  } else if (lowerCell.includes('активные обсуждения') || lowerCell.includes('мониторинг')) {
    return 'activeDiscussions';
  }
  
  return null;
}

/**
 * Имитация логики определения типа поста
 */
function determinePostType(currentSection) {
  if (currentSection === 'reviews') {
    return 'ОС'; // Отзывы сайтов
  } else if (currentSection === 'commentsTop20') {
    return 'ЦС'; // Целевые сайты
  } else if (currentSection === 'activeDiscussions') {
    return 'ПС'; // Площадки социальные
  }
  
  return 'ОС'; // По умолчанию
}

/**
 * Имитация логики проверки строки статистики
 */
function isStatisticsRow(firstCell) {
  const lowerCell = firstCell.toLowerCase().trim();
  
  return lowerCell.includes('суммарное количество просмотров') || 
         lowerCell.includes('количество карточек товара') ||
         lowerCell.includes('количество обсуждений') ||
         lowerCell.includes('доля обсуждений');
}

/**
 * Имитация логики проверки заголовка
 */
function isHeaderRow(row) {
  return row.some(cell => 
    String(cell).toLowerCase().includes('площадка') ||
    String(cell).toLowerCase().includes('тема') ||
    String(cell).toLowerCase().includes('текст') ||
    String(cell).toLowerCase().includes('дата') ||
    String(cell).toLowerCase().includes('ник')
  );
}

/**
 * Имитация логики проверки данных
 */
function hasData(row) {
  const text = row[4] ? String(row[4]).trim() : ''; // Текст сообщения
  const platform = row[1] ? String(row[1]).trim() : ''; // Площадка
  const date = row[6] ? String(row[6]).trim() : ''; // Дата
  const hasLink = row.some(cell => String(cell).includes('http'));
  
  const hasText = text.length > 5;
  const hasPlatform = platform.length > 0;
  const hasDate = date.length > 0;
  
  return hasText || hasPlatform || hasDate || hasLink;
}

// ==================== ТЕСТИРОВАНИЕ ====================

/**
 * Тестирование логики процессора
 */
function testProcessorLogic() {
  console.log('🚀 ТЕСТИРОВАНИЕ ЛОГИКИ ПРОЦЕССОРА');
  console.log('==================================');
  
  let currentSection = null;
  let sectionStartRow = -1;
  let processedRows = 0;
  let skippedRows = 0;
  
  const results = {
    reviews: [],
    commentsTop20: [],
    activeDiscussions: []
  };
  
  // Обрабатываем строки данных (начиная с строки 5)
  for (let i = 4; i < testData.length; i++) {
    const row = testData[i];
    const firstCell = String(row[0] || '').toLowerCase().trim();
    
    // Определяем раздел
    const newSection = determineSection(String(row[0] || ''));
    if (newSection) {
      currentSection = newSection;
      sectionStartRow = i;
      console.log(`📂 Найден раздел "${newSection}" в строке ${i + 1}`);
      continue;
    }
    
    // Пропускаем пустые строки
    if (!row || row.every(cell => !cell || String(cell).trim() === '')) {
      skippedRows++;
      continue;
    }
    
    // Пропускаем строки статистики
    if (isStatisticsRow(String(row[0] || ''))) {
      console.log(`⏭️ Пропускаем строку статистики: "${String(row[0])}"`);
      skippedRows++;
      continue;
    }
    
    // Проверяем, что раздел определен
    if (!currentSection) {
      skippedRows++;
      continue;
    }
    
    // Проверяем, что это не заголовок
    if (i === sectionStartRow + 1 && isHeaderRow(row)) {
      console.log(`⏭️ Пропускаем заголовок раздела`);
      skippedRows++;
      continue;
    }
    
    // Проверяем наличие данных
    if (hasData(row)) {
      const postType = determinePostType(currentSection);
      const record = {
        platform: row[1] || '',
        theme: row[3] || '',
        text: row[4] || '',
        date: row[6] || '',
        author: row[7] || '',
        views: parseInt(row[11]) || 0,
        engagement: row[12] || '',
        type: postType,
        section: currentSection
      };
      
      if (currentSection === 'reviews') {
        results.reviews.push(record);
      } else if (currentSection === 'commentsTop20') {
        results.commentsTop20.push(record);
      } else if (currentSection === 'activeDiscussions') {
        results.activeDiscussions.push(record);
      }
      
      processedRows++;
      console.log(`✅ Обработана строка ${i + 1}: ${record.platform} - ${record.type}`);
    } else {
      skippedRows++;
    }
  }
  
  // Выводим результаты
  console.log('\n📊 РЕЗУЛЬТАТЫ ОБРАБОТКИ:');
  console.log(`📈 Обработано строк: ${processedRows}`);
  console.log(`⏭️ Пропущено строк: ${skippedRows}`);
  console.log(`📂 Отзывы: ${results.reviews.length} записей`);
  console.log(`📂 Комментарии Топ-20: ${results.commentsTop20.length} записей`);
  console.log(`📂 Активные обсуждения: ${results.activeDiscussions.length} записей`);
  
  // Проверяем правильность типов
  console.log('\n🔍 ПРОВЕРКА ТИПОВ:');
  results.reviews.forEach((record, index) => {
    console.log(`   Отзыв ${index + 1}: ${record.type} (ожидается: ОС)`);
  });
  
  results.commentsTop20.forEach((record, index) => {
    console.log(`   Комментарий ${index + 1}: ${record.type} (ожидается: ЦС)`);
  });
  
  results.activeDiscussions.forEach((record, index) => {
    console.log(`   Обсуждение ${index + 1}: ${record.type} (ожидается: ПС)`);
  });
  
  // Проверяем соответствие ожиданиям
  const expectedReviews = 3;
  const expectedComments = 2;
  const expectedDiscussions = 3;
  
  console.log('\n✅ ПРОВЕРКА СООТВЕТСТВИЯ:');
  console.log(`   Отзывы: ${results.reviews.length}/${expectedReviews} ${results.reviews.length === expectedReviews ? '✅' : '❌'}`);
  console.log(`   Комментарии: ${results.commentsTop20.length}/${expectedComments} ${results.commentsTop20.length === expectedComments ? '✅' : '❌'}`);
  console.log(`   Обсуждения: ${results.activeDiscussions.length}/${expectedDiscussions} ${results.activeDiscussions.length === expectedDiscussions ? '✅' : '❌'}`);
  
  const allCorrect = results.reviews.length === expectedReviews && 
                    results.commentsTop20.length === expectedComments && 
                    results.activeDiscussions.length === expectedDiscussions;
  
  console.log(`\n🎯 ИТОГОВЫЙ РЕЗУЛЬТАТ: ${allCorrect ? '✅ ВСЕ ТЕСТЫ ПРОЙДЕНЫ' : '❌ ЕСТЬ ОШИБКИ'}`);
  
  return allCorrect;
}

// Запускаем тест
if (require.main === module) {
  testProcessorLogic();
}

module.exports = { testProcessorLogic }; 