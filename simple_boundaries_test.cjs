/**
 * 🔍 ПРОСТОЙ ТЕСТ ГРАНИЦ РАЗДЕЛОВ
 */

console.log('🔍 ПРОСТОЙ ТЕСТ ГРАНИЦ РАЗДЕЛОВ');
console.log('==============================');

// Простые тестовые данные
const testData = [
  ['', '', '', ''],                          // 1
  ['', '', '', ''],                          // 2
  ['', '', '', ''],                          // 3
  ['Площадка', 'Ссылка', 'Тема', 'Текст'],  // 4 - заголовки
  ['', '', '', ''],                          // 5
  ['отзывы', '', '', ''],                    // 6 - заголовок отзывов
  ['платформа1', 'ссылка1', 'тема1', 'текст1'], // 7 - данные
  ['платформа2', 'ссылка2', 'тема2', 'текст2'], // 8 - данные
  ['комментарии топ-20 выдачи', '', '', ''],    // 9 - заголовок комментариев
  ['платформа3', 'ссылка3', 'тема3', 'текст3'], // 10 - данные
  ['активные обсуждения (мониторинг)', '', '', ''], // 11 - заголовок обсуждений
  ['платформа4', 'ссылка4', 'тема4', 'текст4'], // 12 - данные
  ['', '', '', ''],                          // 13
  ['суммарное количество просмотров* 1000', '', '', ''], // 14 - статистика
  ['количество карточек товара (отзывы) 2', '', '', ''], // 15 - статистика
];

console.log(`📊 Тестовые данные: ${testData.length} строк`);

// Функции проверки
function isStatisticsRow(row) {
  if (!row || row.length === 0) return false;
  
  const firstCell = String(row[0] || '').toLowerCase().trim();
  return firstCell.includes('суммарное количество') || 
         firstCell.includes('количество карточек') ||
         firstCell.includes('количество обсуждений') ||
         firstCell.includes('доля обсуждений');
}

function isEmptyRow(row) {
  return !row || row.every(cell => !cell || String(cell).trim() === '');
}

// Основная функция
function findSectionBoundaries(data) {
  console.log('\n🔍 ПОИСК ГРАНИЦ РАЗДЕЛОВ');
  
  const sections = [];
  let currentSection = null;
  let sectionStart = -1;
  
  const dataStartRow = 5; // CONFIG.STRUCTURE.dataStartRow - 1
  
  for (let i = dataStartRow - 1; i < data.length; i++) {
    const row = data[i];
    const firstCell = String(row[0] || '').toLowerCase().trim();
    
    console.log(`Строка ${i + 1}: "${firstCell}" (статистика: ${isStatisticsRow(row)})`);
    
    // Пропускаем строки статистики
    if (isStatisticsRow(row)) {
      console.log(`  ⏭️ Пропускаем статистику`);
      continue;
    }
    
    // Определяем тип раздела
    let sectionType = null;
    let sectionName = '';
    
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
      console.log(`  📂 Найден новый раздел: "${sectionName}"`);
      
      // Закрываем предыдущий раздел
      if (currentSection && sectionStart !== -1) {
        let endRow = i - 1;
        
        // Ищем последнюю строку данных
        for (let j = i - 1; j >= sectionStart; j--) {
          const checkRow = data[j];
          if (!isStatisticsRow(checkRow) && !isEmptyRow(checkRow)) {
            endRow = j;
            break;
          }
        }
        
        const dataRows = endRow - sectionStart + 1;
        console.log(`  ✅ Закрываем раздел "${getSectionName(currentSection)}" (строки ${sectionStart + 1}-${endRow + 1}, данных: ${dataRows})`);
        
        sections.push({
          type: currentSection,
          name: getSectionName(currentSection),
          startRow: sectionStart,
          endRow: endRow,
          dataRows: dataRows
        });
      }
      
      // Начинаем новый раздел
      currentSection = sectionType;
      sectionStart = i + 1; // Данные начинаются со следующей строки
      console.log(`  🚀 Начинаем новый раздел "${sectionName}", данные с строки ${sectionStart + 1}`);
    }
  }
  
  // Закрываем последний раздел
  if (currentSection && sectionStart !== -1) {
    let endRow = data.length - 1;
    
    // Ищем последнюю строку данных
    for (let j = data.length - 1; j >= sectionStart; j--) {
      const checkRow = data[j];
      if (!isStatisticsRow(checkRow) && !isEmptyRow(checkRow)) {
        endRow = j;
        break;
      }
    }
    
    const dataRows = endRow - sectionStart + 1;
    console.log(`  ✅ Закрываем последний раздел "${getSectionName(currentSection)}" (строки ${sectionStart + 1}-${endRow + 1}, данных: ${dataRows})`);
    
    sections.push({
      type: currentSection,
      name: getSectionName(currentSection),
      startRow: sectionStart,
      endRow: endRow,
      dataRows: dataRows
    });
  }
  
  return sections;
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
const sections = findSectionBoundaries(testData);

console.log('\n📋 ИТОГОВЫЕ РЕЗУЛЬТАТЫ:');
sections.forEach((section, index) => {
  console.log(`${index + 1}. ${section.name}:`);
  console.log(`   Строки: ${section.startRow + 1}-${section.endRow + 1}`);
  console.log(`   Данные: ${section.dataRows} строк`);
});

console.log('\n🎯 ОЖИДАЕМЫЕ РЕЗУЛЬТАТЫ:');
console.log('   - Отзывы: 2 строки данных (строки 7-8)');
console.log('   - Комментарии: 1 строка данных (строка 10)');
console.log('   - Обсуждения: 1 строка данных (строка 12)');

console.log('\n🔍 АНАЛИЗ:');
if (sections.length === 3) {
  console.log('✅ Найдено 3 раздела (правильно)');
} else {
  console.log(`❌ Найдено ${sections.length} разделов (ожидалось 3)`);
}

// Проверка отсутствия дубликатов
const sectionTypes = sections.map(s => s.type);
const uniqueTypes = [...new Set(sectionTypes)];
if (sectionTypes.length === uniqueTypes.length) {
  console.log('✅ Нет дубликатов разделов');
} else {
  console.log('❌ Есть дубликаты разделов');
} 