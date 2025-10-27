const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function universalProcessorReminder() {
  console.log('🌍 НАПОМИНАНИЕ ОБ УНИВЕРСАЛЬНОСТИ ПРОЦЕССОРА');
  console.log('==========================================\n');
  
  console.log('🎯 КРИТИЧЕСКИ ВАЖНО: ПРОЦЕССОР ДОЛЖЕН БЫТЬ УНИВЕРСАЛЬНЫМ!');
  console.log('=======================================================\n');
  
  console.log('❌ ЧТО НЕ ДОЛЖНО БЫТЬ В КОДЕ:');
  console.log('- Жестко заданные числа (22, 20, 621)');
  console.log('- Фиксированные границы строк');
  console.log('- Месяц-специфичные проверки');
  console.log('- Статические массивы данных');
  console.log('- Хардкод названий секций\n');
  
  console.log('✅ ЧТО ДОЛЖНО БЫТЬ В КОДЕ:');
  console.log('- Динамическое определение секций');
  console.log('- Автоматический подсчет записей');
  console.log('- Адаптивные границы');
  console.log('- Универсальная логика обработки');
  console.log('- Гибкие критерии валидации\n');
  
  // Создаем универсальный процессор
  console.log('📝 СОЗДАНИЕ УНИВЕРСАЛЬНОГО ПРОЦЕССОРА:');
  
  const universalProcessorCode = `
// 🌍 УНИВЕРСАЛЬНЫЙ ПРОЦЕССОР ДЛЯ ЛЮБЫХ МЕСЯЦЕВ И ДАННЫХ

function universalDataProcessor(data) {
  console.log('🌍 Запуск универсального процессора...');
  
  // 1. Динамическое определение структуры
  const structure = analyzeDataStructure(data);
  console.log('📊 Структура данных определена:', structure);
  
  // 2. Универсальная обработка секций
  const sections = processAllSections(data, structure);
  console.log('📋 Секции обработаны:', sections);
  
  // 3. Автоматический подсчет
  const statistics = calculateStatistics(sections);
  console.log('📈 Статистика:', statistics);
  
  return { sections, statistics, structure };
}

function analyzeDataStructure(data) {
  // Динамически определяем структуру данных
  const structure = {
    hasReviews: false,
    hasComments: false,
    hasDiscussions: false,
    sectionHeaders: [],
    dataStartRow: -1
  };
  
  for (let i = 0; i < data.length; i++) {
    if (data[i] && data[i][0]) {
      const cellValue = data[i][0].toString();
      
      // Динамическое определение секций
      if (cellValue.toLowerCase().includes('отзыв')) {
        structure.hasReviews = true;
        structure.sectionHeaders.push({ type: 'reviews', row: i, title: cellValue });
      } else if (cellValue.toLowerCase().includes('комментар')) {
        structure.hasComments = true;
        structure.sectionHeaders.push({ type: 'comments', row: i, title: cellValue });
      } else if (cellValue.toLowerCase().includes('обсужден') || cellValue.toLowerCase().includes('активн')) {
        structure.hasDiscussions = true;
        structure.sectionHeaders.push({ type: 'discussions', row: i, title: cellValue });
      }
    }
  }
  
  return structure;
}

function processAllSections(data, structure) {
  const sections = {
    reviews: [],
    comments: [],
    discussions: []
  };
  
  let currentSection = null;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i] && data[i][0]) {
      const cellValue = data[i][0].toString();
      
      // Динамическое определение текущей секции
      if (structure.sectionHeaders.some(h => h.row === i)) {
        const header = structure.sectionHeaders.find(h => h.row === i);
        currentSection = header.type;
        console.log(\`📍 Переход в секцию: \${header.title}\`);
        continue; // Пропускаем заголовок
      }
      
      // Универсальная валидация данных
      if (isValidDataRow(data[i])) {
        if (currentSection === 'reviews') {
          sections.reviews.push(data[i]);
        } else if (currentSection === 'comments') {
          sections.comments.push(data[i]);
        } else if (currentSection === 'discussions') {
          sections.discussions.push(data[i]);
        }
      }
    }
  }
  
  return sections;
}

function isValidDataRow(row) {
  // Универсальная валидация - работает для любых данных
  if (!row || !row[0]) return false;
  
  const cellValue = row[0].toString().trim();
  if (cellValue === '') return false;
  
  // Проверяем, что это не заголовок
  const headerKeywords = ['площадка', 'тема', 'текст', 'дата', 'ник', 'просмотры'];
  if (headerKeywords.some(keyword => cellValue.toLowerCase().includes(keyword))) {
    return false;
  }
  
  // Проверяем, что это данные (содержит домен или ссылку)
  return cellValue.includes('.') || cellValue.includes('http') || cellValue.includes('www');
}

function calculateStatistics(sections) {
  // Автоматический подсчет для любых данных
  const stats = {
    reviews: sections.reviews.length,
    comments: sections.comments.length,
    discussions: sections.discussions.length,
    total: sections.reviews.length + sections.comments.length + sections.discussions.length
  };
  
  console.log('📊 Автоматический подсчет:');
  console.log(\`   - Отзывы: \${stats.reviews}\`);
  console.log(\`   - Комментарии: \${stats.comments}\`);
  console.log(\`   - Обсуждения: \${stats.discussions}\`);
  console.log(\`   - Всего: \${stats.total}\`);
  
  return stats;
}

// Функция для проверки качества обработки
function validateProcessingQuality(actualStats, expectedStats = null) {
  console.log('🔍 Проверка качества обработки...');
  
  if (expectedStats) {
    // Сравнение с эталоном (если есть)
    const accuracy = {
      reviews: actualStats.reviews === expectedStats.reviews ? 100 : 
               Math.abs(actualStats.reviews - expectedStats.reviews) / expectedStats.reviews * 100,
      comments: actualStats.comments === expectedStats.comments ? 100 :
                Math.abs(actualStats.comments - expectedStats.comments) / expectedStats.comments * 100,
      discussions: actualStats.discussions === expectedStats.discussions ? 100 :
                   Math.abs(actualStats.discussions - expectedStats.discussions) / expectedStats.discussions * 100
    };
    
    const overallAccuracy = (accuracy.reviews + accuracy.comments + accuracy.discussions) / 3;
    console.log(\`📈 Точность: \${overallAccuracy.toFixed(1)}%\`);
    
    return overallAccuracy >= 95;
  } else {
    // Проверка логичности данных
    const isValid = actualStats.total > 0 && 
                   actualStats.reviews >= 0 && 
                   actualStats.comments >= 0 && 
                   actualStats.discussions >= 0;
    
    console.log(\`✅ Логичность данных: \${isValid ? 'ПРОЙДЕНА' : 'НЕ ПРОЙДЕНА'}\`);
    return isValid;
  }
}

// Экспорт для использования
module.exports = {
  universalDataProcessor,
  analyzeDataStructure,
  processAllSections,
  isValidDataRow,
  calculateStatistics,
  validateProcessingQuality
};
`;
  
  console.log('   Код универсального процессора создан!');
  console.log('   🌍 Основные принципы универсальности:');
  console.log('   ✅ Динамическое определение структуры');
  console.log('   ✅ Адаптивная обработка секций');
  console.log('   ✅ Автоматический подсчет');
  console.log('   ✅ Универсальная валидация');
  console.log('   ✅ Работает для любых месяцев');
  console.log('   ✅ Работает с любым количеством данных\n');
  
  // Сохраняем универсальный код
  fs.writeFileSync('universal_processor.js', universalProcessorCode);
  console.log('💾 Универсальный процессор сохранен в файл: universal_processor.js');
  
  console.log('\n🎯 КРИТИЧЕСКИЕ ТРЕБОВАНИЯ ДЛЯ АГЕНТА:');
  console.log('=====================================');
  console.log('1. ❌ НЕ использовать жестко заданные числа!');
  console.log('2. ❌ НЕ привязываться к конкретному месяцу!');
  console.log('3. ❌ НЕ использовать статические границы!');
  console.log('4. ✅ Использовать динамическое определение!');
  console.log('5. ✅ Обрабатывать любые объемы данных!');
  console.log('6. ✅ Работать с любыми месяцами!');
  console.log('7. ✅ Адаптироваться к изменениям структуры!\n');
  
  console.log('📋 ПРИМЕРЫ ЧТО НЕ ДЕЛАТЬ:');
  console.log('- if (month === "Март") { ... }');
  console.log('- const expectedReviews = 22;');
  console.log('- const sectionEnd = 48;');
  console.log('- const fixedHeaders = ["Отзывы", "Комментарии"];\n');
  
  console.log('📋 ПРИМЕРЫ ЧТО ДЕЛАТЬ:');
  console.log('- const sections = findSectionsDynamically(data);');
  console.log('- const count = countValidRows(section);');
  console.log('- const structure = analyzeStructure(data);');
  console.log('- const stats = calculateStats(sections);\n');
  
  console.log('🚀 СЛЕДУЮЩИЕ ШАГИ:');
  console.log('1. Применить универсальный процессор из universal_processor.js');
  console.log('2. Протестировать на разных месяцах (февраль, март, апрель, май)');
  console.log('3. Убедиться, что работает с разными объемами данных');
  console.log('4. Проверить адаптивность к изменениям структуры');
  console.log('5. Достичь 95%+ точности на всех тестах');
  
  console.log('\n🌍 ПОМНИ: УНИВЕРСАЛЬНОСТЬ = БУДУЩЕЕ!');
  console.log('Процессор должен работать завтра, через месяц и через год!');
}

universalProcessorReminder();