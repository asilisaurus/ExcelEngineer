/**
 * 🔍 ДИАГНОСТИКА ГРАНИЦ РАЗДЕЛОВ
 * Быстрый тест для понимания проблемы с определением границ разделов
 */

console.log('🔍 ДИАГНОСТИКА ГРАНИЦ РАЗДЕЛОВ');
console.log('===============================');

// Симулируем данные из лога
const mockData = [
  // Строки 1-5: инфо и заголовки
  ['', '', '', ''],
  ['', '', '', ''],
  ['', '', '', ''],
  ['Площадка', 'Ссылка', 'Тема', 'Текст'],
  ['', '', '', ''],
  
  // Строка 6: заголовок "Отзывы"
  ['отзывы', '', '', ''],
  
  // Строки 7-27: данные отзывов (21 строка)
  ['платформа1', 'ссылка1', 'тема1', 'текст отзыва 1'],
  ['платформа2', 'ссылка2', 'тема2', 'текст отзыва 2'],
  ['платформа3', 'ссылка3', 'тема3', 'текст отзыва 3'],
  // ... еще 18 строк отзывов
  ...new Array(18).fill(['платформа', 'ссылка', 'тема', 'текст отзыва']),
  
  // Строка 28: заголовок "Комментарии Топ-20"
  ['комментарии топ-20 выдачи', '', '', ''],
  
  // Строки 29-48: данные комментариев (20 строк)
  ['платформа1', 'ссылка1', 'тема1', 'текст комментария 1'],
  ['платформа2', 'ссылка2', 'тема2', 'текст комментария 2'],
  // ... еще 18 строк комментариев
  ...new Array(18).fill(['платформа', 'ссылка', 'тема', 'текст комментария']),
  
  // Строка 49: заголовок "Активные обсуждения"
  ['активные обсуждения (мониторинг)', '', '', ''],
  
  // Строки 50-680: данные обсуждений (631 строка)
  ['платформа1', 'ссылка1', 'тема1', 'текст обсуждения 1'],
  ['платформа2', 'ссылка2', 'тема2', 'текст обсуждения 2'],
  // ... еще 629 строк обсуждений
  ...new Array(629).fill(['платформа', 'ссылка', 'тема', 'текст обсуждения']),
  
  // Строки 681-690: статистика
  ['', '', '', ''],
  ['', '', '', ''],
  ['суммарное количество просмотров* 3333564', '', '', ''],
  ['количество карточек товара (отзывы) 22', '', '', ''],
  ['количество обсуждений (форумы, сообщества, комментарии к статьям) 631', '', '', ''],
  ['доля обсуждений с вовлечением в диалог 0.20', '', '', ''],
  ['', '', '', ''],
  ['*без учета площадок с закрытой статистикой прочтений', '', '', ''],
  ['площадки со статистикой просмотров 0.70', '', '', ''],
  ['количество прочтений увеличивается в среднем на 30%', '', '', ''],
];

console.log(`📊 Тестовые данные: ${mockData.length} строк`);

// Симулируем логику findSectionBoundaries (УЛУЧШЕННАЯ ВЕРСИЯ)
function simulateFindSectionBoundaries(data) {
  console.log('\n🔍 СИМУЛЯЦИЯ findSectionBoundaries() - УЛУЧШЕННАЯ ВЕРСИЯ');
  
  const sections = [];
  let currentSection = null;
  let sectionStart = -1;
  
  const dataStartRow = 5; // CONFIG.STRUCTURE.dataStartRow - 1
  
  for (let i = dataStartRow - 1; i < data.length; i++) {
    const row = data[i];
    const firstCell = String(row[0] || '').toLowerCase().trim();
    
    // ✅ ИСПРАВЛЕНИЕ: Пропускаем строки статистики
    if (isStatisticsRow(row)) {
      continue;
    }
    
    // Определяем тип раздела
    let sectionType = null;
    let sectionName = '';
    
    // ✅ ИСПРАВЛЕНИЕ: Более строгие условия для определения заголовков разделов
    if (firstCell === 'отзывы' || (firstCell.includes('отзывы') && !firstCell.includes('топ-20') && !firstCell.includes('обсуждения') && !firstCell.includes('количество'))) {
      sectionType = 'reviews';
      sectionName = 'Отзывы';
    } else if (firstCell.includes('комментарии топ-20') || firstCell.includes('топ-20 выдачи')) {
      sectionType = 'commentsTop20';
      sectionName = 'Комментарии Топ-20';
    } else if (firstCell.includes('активные обсуждения') || firstCell.includes('мониторинг')) {
      sectionType = 'activeDiscussions';
      sectionName = 'Активные обсуждения';
    }
    
    // Если найден новый раздел
    if (sectionType && sectionType !== currentSection) {
      // Закрываем предыдущий раздел
      if (currentSection && sectionStart !== -1) {
        // ✅ ИСПРАВЛЕНИЕ: Улучшенное определение конца раздела
        let endRow = i - 1;
        
        // Ищем последнюю строку данных (исключаем статистику и пустые строки)
        for (let j = i - 1; j >= sectionStart; j--) {
          const checkRow = data[j];
          if (!isStatisticsRow(checkRow) && !isEmptyRow(checkRow)) {
            endRow = j;
            break;
          }
        }
        
        sections.push({
          type: currentSection,
          name: getSectionName(currentSection),
          startRow: sectionStart,
          endRow: endRow,
          dataRows: endRow - sectionStart + 1
        });
      }
      
      // Начинаем новый раздел
      currentSection = sectionType;
      sectionStart = i + 1; // ✅ ИСПРАВЛЕНО: данные начинаются со следующей строки
      console.log(`📂 Найден раздел "${sectionName}" в строке ${i + 1}, данные с строки ${sectionStart + 1}`);
    }
  }
  
  // Закрываем последний раздел
  if (currentSection && sectionStart !== -1) {
    // ✅ ИСПРАВЛЕНИЕ: Улучшенное определение конца последнего раздела
    let endRow = data.length - 1;
    
    // Ищем последнюю строку данных (исключаем статистику)
    for (let j = data.length - 1; j >= sectionStart; j--) {
      const checkRow = data[j];
      if (!isStatisticsRow(checkRow) && !isEmptyRow(checkRow)) {
        endRow = j;
        break;
      }
    }
    
    sections.push({
      type: currentSection,
      name: getSectionName(currentSection),
      startRow: sectionStart,
      endRow: endRow,
      dataRows: endRow - sectionStart + 1
    });
  }
  
  return sections;
}

// Вспомогательные функции для проверки
function isStatisticsRow(row) {
  if (!row || row.length === 0) return false;
  
  const firstCell = String(row[0] || '').toLowerCase().trim();
  return firstCell.includes('суммарное количество просмотров') || 
         firstCell.includes('количество карточек товара') ||
         firstCell.includes('количество обсуждений') ||
         firstCell.includes('доля обсуждений') ||
         firstCell.includes('площадки со статистикой') ||
         firstCell.includes('количество прочтений увеличивается');
}

function isEmptyRow(row) {
  return !row || row.every(cell => !cell || String(cell).trim() === '');
}

function getSectionName(sectionType) {
  const names = {
    'reviews': 'Отзывы',
    'commentsTop20': 'Комментарии Топ-20',
    'activeDiscussions': 'Активные обсуждения'
  };
  return names[sectionType] || sectionType;
}

// Запускаем тест
const sections = simulateFindSectionBoundaries(mockData);

console.log('\n📋 РЕЗУЛЬТАТЫ ОПРЕДЕЛЕНИЯ РАЗДЕЛОВ:');
sections.forEach((section, index) => {
  console.log(`${index + 1}. ${section.name}:`);
  console.log(`   Строки: ${section.startRow + 1}-${section.endRow + 1}`);
  console.log(`   Данные: ${section.dataRows} строк`);
  console.log(`   Тип: ${section.type}`);
});

console.log('\n📊 СРАВНЕНИЕ С ЭТАЛОНОМ:');
console.log('✅ Ожидается:');
console.log('   - Отзывы: 22 строки');
console.log('   - Комментарии: 20 строк');
console.log('   - Обсуждения: 631 строка');

console.log('\n🔍 Получается:');
sections.forEach(section => {
  console.log(`   - ${section.name}: ${section.dataRows} строк`);
});

// Проверяем правильность
const expectedCounts = { reviews: 22, commentsTop20: 20, activeDiscussions: 631 };
let allCorrect = true;

sections.forEach(section => {
  const expected = expectedCounts[section.type];
  const actual = section.dataRows;
  
  if (expected !== actual) {
    console.log(`❌ ${section.name}: ожидалось ${expected}, получено ${actual}`);
    allCorrect = false;
  } else {
    console.log(`✅ ${section.name}: ${actual} строк - ПРАВИЛЬНО`);
  }
});

console.log('\n🎯 ИТОГОВЫЙ РЕЗУЛЬТАТ:');
if (allCorrect) {
  console.log('✅ ВСЕ РАЗДЕЛЫ ОПРЕДЕЛЕНЫ ПРАВИЛЬНО!');
} else {
  console.log('❌ ЕСТЬ ПРОБЛЕМЫ С ОПРЕДЕЛЕНИЕМ РАЗДЕЛОВ');
} 