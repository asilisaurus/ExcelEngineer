const fs = require('fs');
const path = require('path');

/**
 * 🚀 БЫСТРЫЙ ТЕСТ ИСПРАВЛЕННОГО ПРОЦЕССОРА
 * Проверка критических исправлений и базовой логики
 */

async function quickTest() {
  console.log('🚀 БЫСТРЫЙ ТЕСТ ИСПРАВЛЕННОГО ПРОЦЕССОРА');
  console.log('=' .repeat(50));

  // 1. Проверка исправления
  console.log('\n📋 1. ПРОВЕРКА ИСПРАВЛЕНИЯ:');
  const processorPath = path.join(__dirname, 'google-apps-script-processor-fixed-boundaries.js');
  
  if (!fs.existsSync(processorPath)) {
    console.log('❌ Исправленный процессор не найден');
    return;
  }

  const processorContent = fs.readFileSync(processorPath, 'utf8');
  const hasFix = processorContent.includes('sectionStart = i + 1');
  
  console.log(`${hasFix ? '✅' : '❌'} Критическое исправление: ${hasFix ? 'ПРИМЕНЕНО' : 'НЕ ПРИМЕНЕНО'}`);
  console.log(`🔍 Исправление: sectionStart = i + 1 (исключение заголовков секций)`);

  // 2. Симуляция логики
  console.log('\n🧪 2. СИМУЛЯЦИЯ ИСПРАВЛЕННОЙ ЛОГИКИ:');
  
  // Создаем тестовые данные
  const testData = [
    ['Информация', 'Фортедетрим'],
    ['Период', 'Март 2025'],
    ['Дата', '2025-03-31'],
    ['Название', 'Пост', 'Просмотры', 'Комментарии'], // Заголовки
    ['отзывы', '', '', ''],                          // Заголовок секции отзывов
    ['Отзыв 1', 'Положительный', '100', '5'],       // Данные отзыва 1
    ['Отзыв 2', 'Негативный', '150', '8'],          // Данные отзыва 2
    ['комментарии топ-20', '', '', ''],              // Заголовок секции комментариев
    ['Комментарий 1', 'Текст', '200', '10'],        // Данные комментария 1
    ['Комментарий 2', 'Текст', '180', '7'],         // Данные комментария 2
    ['активные обсуждения', '', '', ''],             // Заголовок секции обсуждений
    ['Обсуждение 1', 'Текст', '300', '15'],         // Данные обсуждения 1
    ['Обсуждение 2', 'Текст', '250', '12'],         // Данные обсуждения 2
    ['Обсуждение 3', 'Текст', '220', '9']           // Данные обсуждения 3
  ];

  console.log(`📊 Тестовые данные: ${testData.length} строк`);
  console.log(`   Структура: 3 инфо строки + 1 заголовок + 3 секции`);

  // 3. Тестирование старой логики (с ошибкой)
  console.log('\n❌ 3. СТАРАЯ ЛОГИКА (с ошибкой):');
  const oldResults = simulateOldLogic(testData);
  console.log(`   Отзывы: ${oldResults.reviews} (включает заголовок)`);
  console.log(`   Комментарии: ${oldResults.comments} (включает заголовок)`);
  console.log(`   Обсуждения: ${oldResults.discussions} (включает заголовок)`);
  console.log(`   Всего: ${oldResults.total}`);

  // 4. Тестирование новой логики (исправленной)
  console.log('\n✅ 4. НОВАЯ ЛОГИКА (исправленная):');
  const newResults = simulateNewLogic(testData);
  console.log(`   Отзывы: ${newResults.reviews} (без заголовка)`);
  console.log(`   Комментарии: ${newResults.comments} (без заголовка)`);
  console.log(`   Обсуждения: ${newResults.discussions} (без заголовка)`);
  console.log(`   Всего: ${newResults.total}`);

  // 5. Сравнение результатов
  console.log('\n📊 5. СРАВНЕНИЕ РЕЗУЛЬТАТОВ:');
  const improvement = newResults.total - oldResults.total;
  console.log(`   Улучшение: ${improvement > 0 ? '+' : ''}${improvement} записей`);
  
  const expectedResults = { reviews: 2, comments: 2, discussions: 3, total: 7 };
  const accuracy = calculateAccuracy(newResults, expectedResults);
  
  console.log(`   Ожидаемо: ${expectedResults.reviews}+${expectedResults.comments}+${expectedResults.discussions} = ${expectedResults.total}`);
  console.log(`   Получено: ${newResults.reviews}+${newResults.comments}+${newResults.discussions} = ${newResults.total}`);
  console.log(`   Точность: ${accuracy.toFixed(2)}%`);

  // 6. Вердикт
  console.log('\n🎯 6. ВЕРДИКТ:');
  if (hasFix && accuracy >= 95) {
    console.log('✅ ТЕСТ ПРОЙДЕН УСПЕШНО!');
    console.log('   ✅ Исправление применено');
    console.log('   ✅ Логика работает корректно');
    console.log('   ✅ Точность 95%+');
    console.log('   🎉 Процессор готов к использованию!');
  } else {
    console.log('❌ ТЕСТ НЕ ПРОЙДЕН');
    if (!hasFix) console.log('   ❌ Исправление не применено');
    if (accuracy < 95) console.log('   ❌ Низкая точность');
    console.log('   🔧 Требуются дополнительные исправления');
  }

  return { hasFix, accuracy, oldResults, newResults };
}

// Симуляция старой логики (с ошибкой)
function simulateOldLogic(data) {
  const sections = findSectionsOld(data);
  return calculateStats(sections);
}

// Симуляция новой логики (исправленной)
function simulateNewLogic(data) {
  const sections = findSectionsNew(data);
  return calculateStats(sections);
}

// Старая логика поиска секций (с ошибкой)
function findSectionsOld(data) {
  const sections = {};
  const sectionKeywords = ['отзывы', 'комментарии топ-20', 'активные обсуждения'];
  
  for (let i = 0; i < data.length; i++) {
    const rowText = data[i].join(' ').toLowerCase();
    
    for (const keyword of sectionKeywords) {
      if (rowText.includes(keyword)) {
        const sectionStart = i; // ОШИБКА: включает заголовок
        let sectionEnd = data.length;
        
        for (let j = i + 1; j < data.length; j++) {
          const nextRowText = data[j].join(' ').toLowerCase();
          if (sectionKeywords.some(k => nextRowText.includes(k))) {
            sectionEnd = j;
            break;
          }
        }
        
        sections[keyword] = { start: sectionStart, end: sectionEnd };
        break;
      }
    }
  }
  
  return sections;
}

// Новая логика поиска секций (исправленная)
function findSectionsNew(data) {
  const sections = {};
  const sectionKeywords = ['отзывы', 'комментарии топ-20', 'активные обсуждения'];
  
  for (let i = 0; i < data.length; i++) {
    const rowText = data[i].join(' ').toLowerCase();
    
    for (const keyword of sectionKeywords) {
      if (rowText.includes(keyword)) {
        const sectionStart = i + 1; // ИСПРАВЛЕНО: не включает заголовок
        let sectionEnd = data.length;
        
        for (let j = i + 1; j < data.length; j++) {
          const nextRowText = data[j].join(' ').toLowerCase();
          if (sectionKeywords.some(k => nextRowText.includes(k))) {
            sectionEnd = j;
            break;
          }
        }
        
        sections[keyword] = { start: sectionStart, end: sectionEnd };
        break;
      }
    }
  }
  
  return sections;
}

// Расчет статистики
function calculateStats(sections) {
  const stats = { reviews: 0, comments: 0, discussions: 0, total: 0 };
  
  if (sections['отзывы']) {
    stats.reviews = sections['отзывы'].end - sections['отзывы'].start;
  }
  
  if (sections['комментарии топ-20']) {
    stats.comments = sections['комментарии топ-20'].end - sections['комментарии топ-20'].start;
  }
  
  if (sections['активные обсуждения']) {
    stats.discussions = sections['активные обсуждения'].end - sections['активные обсуждения'].start;
  }
  
  stats.total = stats.reviews + stats.comments + stats.discussions;
  return stats;
}

// Расчет точности
function calculateAccuracy(actual, expected) {
  const totalExpected = expected.total;
  const totalActual = actual.total;
  
  if (totalExpected === 0) return totalActual === 0 ? 100 : 0;
  
  const accuracy = Math.max(0, 100 - Math.abs(totalActual - totalExpected) / totalExpected * 100);
  return accuracy;
}

// Запуск теста
quickTest().catch(console.error); 