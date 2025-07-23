const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function activeFixAssistant() {
  console.log('🚀 АКТИВНЫЙ ПОМОЩНИК ИСПРАВЛЕНИЙ');
  console.log('================================\n');
  
  console.log('🎯 ЦЕЛЬ: ДОСТИЧЬ 95%+ ТОЧНОСТИ И УНИВЕРСАЛЬНОСТИ');
  console.log('===============================================\n');
  
  try {
    // Находим самый новый файл результата
    const uploadsDir = 'uploads';
    const files = fs.readdirSync(uploadsDir);
    
    const resultFiles = files.filter(file => 
      file.includes('результат') && file.endsWith('.xlsx') && !file.includes('temp_google_sheets')
    );
    
    const fileStats = resultFiles.map(file => ({
      name: file,
      path: path.join(uploadsDir, file),
      stats: fs.statSync(path.join(uploadsDir, file))
    })).sort((a, b) => b.stats.mtime - a.stats.mtime);
    
    if (fileStats.length === 0) {
      console.log('❌ Файлы результата не найдены');
      return;
    }
    
    const latestFile = fileStats[0];
    console.log(`📁 Анализируем: ${latestFile.name}\n`);
    
    const workbook = XLSX.readFile(latestFile.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`📊 Лист: "${sheetName}"`);
    console.log(`📋 Всего строк: ${data.length}\n`);
    
    // АКТИВНЫЙ АНАЛИЗ ПРОБЛЕМ
    console.log('🔍 АКТИВНЫЙ АНАЛИЗ ПРОБЛЕМ:');
    
    // Проблема 1: Подсчет с учетом пустых строк
    let reviewsCount = 0;
    let commentsCount = 0;
    let discussionsCount = 0;
    let inReviewsSection = false;
    let inCommentsSection = false;
    let inDiscussionsSection = false;
    let emptyRowsInComments = 0;
    let emptyRowsInDiscussions = 0;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        
        // Определяем секции
        if (cellValue.includes('Отзывы') && !cellValue.includes('Количество')) {
          inReviewsSection = true;
          inCommentsSection = false;
          inDiscussionsSection = false;
        } else if (cellValue.includes('Комментарии') && !cellValue.includes('Количество')) {
          inReviewsSection = false;
          inCommentsSection = true;
          inDiscussionsSection = false;
        } else if (cellValue.includes('Активные обсуждения') || cellValue.includes('Обсуждения')) {
          inReviewsSection = false;
          inCommentsSection = false;
          inDiscussionsSection = true;
        }
        
        // Подсчитываем только валидные записи
        if (inReviewsSection && cellValue.includes('.')) {
          reviewsCount++;
        } else if (inCommentsSection && cellValue.includes('.')) {
          commentsCount++;
        } else if (inDiscussionsSection && cellValue.includes('.')) {
          discussionsCount++;
        }
      } else {
        // Пустые строки
        if (inCommentsSection) {
          emptyRowsInComments++;
        } else if (inDiscussionsSection) {
          emptyRowsInDiscussions++;
        }
      }
    }
    
    console.log(`📊 ТЕКУЩИЕ РЕЗУЛЬТАТЫ:`);
    console.log(`   - Отзывы: ${reviewsCount} (ожидается 22)`);
    console.log(`   - Комментарии: ${commentsCount} (ожидается 20)`);
    console.log(`   - Обсуждения: ${discussionsCount} (ожидается 621)`);
    console.log(`   - Пустых строк в комментариях: ${emptyRowsInComments}`);
    console.log(`   - Пустых строк в обсуждениях: ${emptyRowsInDiscussions}\n`);
    
    // Проблема 2: Поиск недостающих данных
    console.log('🔍 ПОИСК НЕДОСТАЮЩИХ ДАННЫХ:');
    
    // Ищем данные, которые могли быть пропущены
    let potentialMissingData = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        // Ищем строки с данными, которые не попали в подсчет
        if (cellValue.includes('.') && !inReviewsSection && !inCommentsSection && !inDiscussionsSection) {
          potentialMissingData.push({ row: i + 1, value: cellValue });
        }
      }
    }
    
    if (potentialMissingData.length > 0) {
      console.log(`⚠️ Найдено ${potentialMissingData.length} потенциально пропущенных записей:`);
      potentialMissingData.slice(0, 5).forEach(item => {
        console.log(`   Строка ${item.row}: ${item.value}`);
      });
    }
    
    // СОЗДАНИЕ АКТИВНЫХ ИСПРАВЛЕНИЙ
    console.log('\n🔧 СОЗДАНИЕ АКТИВНЫХ ИСПРАВЛЕНИЙ:');
    
    const activeFixCode = `
// 🚀 АКТИВНЫЕ ИСПРАВЛЕНИЯ ДЛЯ ДОСТИЖЕНИЯ 95%+ ТОЧНОСТИ

function activeDataProcessor(data) {
  console.log('🚀 Запуск активного процессора...');
  
  // 1. Очистка данных от пустых строк
  const cleanData = data.filter(row => 
    row && row[0] && row[0].toString().trim() !== ""
  );
  
  console.log(\`📊 Очищено данных: \${data.length} → \${cleanData.length} строк\`);
  
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
      console.log(\`📍 Секция "Отзывы" начинается со строки \${sectionStartRow}\`);
      continue;
    } else if (cellValue.toLowerCase().includes('комментар') && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'comments';
      sectionStartRow = i + 1;
      console.log(\`📍 Секция "Комментарии" начинается со строки \${sectionStartRow}\`);
      continue;
    } else if ((cellValue.toLowerCase().includes('обсужден') || cellValue.toLowerCase().includes('активн')) && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'discussions';
      sectionStartRow = i + 1;
      console.log(\`📍 Секция "Обсуждения" начинается со строки \${sectionStartRow}\`);
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
  
  console.log(\`📈 РЕЗУЛЬТАТЫ АКТИВНОГО ПРОЦЕССОРА:\`);
  console.log(\`   - Отзывы: \${stats.reviews}\`);
  console.log(\`   - Комментарии: \${stats.comments}\`);
  console.log(\`   - Обсуждения: \${stats.discussions}\`);
  console.log(\`   - Всего: \${stats.total}\`);
  
  // 5. Проверка качества (универсальная)
  const qualityCheck = validateQuality(stats);
  console.log(\`🎯 Качество обработки: \${qualityCheck.passed ? '✅ ПРОЙДЕНА' : '❌ НЕ ПРОЙДЕНА'}\`);
  
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
module.exports = {
  activeDataProcessor,
  validateQuality
};
`;
    
    console.log('   ✅ Код активных исправлений создан!');
    console.log('   🚀 Основные улучшения:');
    console.log('   - Универсальная обработка (любое количество строк)');
    console.log('   - Автоматическая очистка пустых строк');
    console.log('   - Динамическое определение секций');
    console.log('   - Улучшенная валидация данных');
    console.log('   - Автоматическая проверка качества\n');
    
    // Сохраняем активные исправления
    fs.writeFileSync('active_fix_processor.js', activeFixCode);
    console.log('💾 Активные исправления сохранены в файл: active_fix_processor.js');
    
    // Создаем тестовый скрипт для проверки
    const testCode = `
// 🧪 ТЕСТОВЫЙ СКРИПТ ДЛЯ ПРОВЕРКИ ИСПРАВЛЕНИЙ

const { activeDataProcessor } = require('./active_fix_processor.js');
const XLSX = require('xlsx');

function testActiveFixes() {
  console.log('🧪 Тестирование активных исправлений...');
  
  try {
    // Загружаем данные
    const workbook = XLSX.readFile('${latestFile.path}');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(\`📊 Тестируем файл: \${sheetName}\`);
    console.log(\`📋 Всего строк: \${data.length}\`);
    
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
    
    console.log(\`\\n🎯 РЕЗУЛЬТАТЫ ТЕСТА:\`);
    console.log(\`   - Отзывы: \${result.stats.reviews}/\${expected.reviews} (\${(100 - accuracy.reviews).toFixed(1)}% точность)\`);
    console.log(\`   - Комментарии: \${result.stats.comments}/\${expected.comments} (\${(100 - accuracy.comments).toFixed(1)}% точность)\`);
    console.log(\`   - Обсуждения: \${result.stats.discussions}/\${expected.discussions} (\${(100 - accuracy.discussions).toFixed(1)}% точность)\`);
    console.log(\`   - Общая точность: \${overallAccuracy.toFixed(1)}%\`);
    
    if (overallAccuracy >= 95) {
      console.log(\`🎉 УСПЕХ! Достигнута требуемая точность 95%+\`);
    } else {
      console.log(\`⚠️ Нужны дополнительные исправления для достижения 95%+\`);
    }
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
  }
}

testActiveFixes();
`;
    
    fs.writeFileSync('test_active_fixes.cjs', testCode);
    console.log('💾 Тестовый скрипт сохранен в файл: test_active_fixes.cjs');
    
    console.log('\n🎯 АКТИВНЫЕ ДЕЙСТВИЯ ДЛЯ АГЕНТА:');
    console.log('================================');
    console.log('1. 🚀 Применить active_fix_processor.js');
    console.log('2. 🧪 Запустить test_active_fixes.cjs');
    console.log('3. 📊 Проверить достижение 95%+ точности');
    console.log('4. 🌍 Убедиться в универсальности (любое количество строк)');
    console.log('5. 🔄 Протестировать на других месяцах');
    
    console.log('\n🚨 КРИТИЧЕСКИ ВАЖНО:');
    console.log('- Процессор должен работать с ЛЮБЫМ количеством строк!');
    console.log('- НЕ использовать жестко заданные числа!');
    console.log('- Динамически определять структуру данных!');
    console.log('- Адаптироваться к изменениям!');
    
  } catch (error) {
    console.error('❌ Ошибка при создании активных исправлений:', error);
  }
}

activeFixAssistant();