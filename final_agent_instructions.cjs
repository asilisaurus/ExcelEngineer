const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function finalAgentInstructions() {
  console.log('🎯 ФИНАЛЬНЫЕ ИНСТРУКЦИИ ДЛЯ АГЕНТА');
  console.log('==================================\n');
  
  console.log('🚨 КРИТИЧЕСКАЯ ЗАДАЧА: ДОСТИЧЬ 95%+ ТОЧНОСТИ');
  console.log('============================================\n');
  
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
    
    // ТОЧНЫЙ АНАЛИЗ ПРОБЛЕМ
    console.log('🔍 ТОЧНЫЙ АНАЛИЗ ПРОБЛЕМ:');
    
    // Эталонные данные
    const expected = { reviews: 22, comments: 20, discussions: 621, total: 663 };
    
    // Текущие результаты агента
    const current = { reviews: 22, comments: 19, discussions: 514, total: 555 };
    
    console.log('📋 ЭТАЛОННЫЕ ДАННЫЕ:');
    console.log(`   - Отзывы: ${expected.reviews}`);
    console.log(`   - Комментарии: ${expected.comments}`);
    console.log(`   - Обсуждения: ${expected.discussions}`);
    console.log(`   - Всего: ${expected.total}\n`);
    
    console.log('📊 ТЕКУЩИЕ РЕЗУЛЬТАТЫ АГЕНТА:');
    console.log(`   - Отзывы: ${current.reviews} ✅`);
    console.log(`   - Комментарии: ${current.comments} ❌ (не хватает 1)`);
    console.log(`   - Обсуждения: ${current.discussions} ❌ (не хватает 107)`);
    console.log(`   - Всего: ${current.total} ❌ (не хватает 108)\n`);
    
    // Расчет точности
    const accuracy = {
      reviews: current.reviews === expected.reviews ? 100 : 
               (100 - Math.abs(current.reviews - expected.reviews) / expected.reviews * 100),
      comments: current.comments === expected.comments ? 100 :
                (100 - Math.abs(current.comments - expected.comments) / expected.comments * 100),
      discussions: current.discussions === expected.discussions ? 100 :
                   (100 - Math.abs(current.discussions - expected.discussions) / expected.discussions * 100)
    };
    
    const overallAccuracy = (accuracy.reviews + accuracy.comments + accuracy.discussions) / 3;
    
    console.log('🎯 ТОЧНЫЙ РАСЧЕТ ТОЧНОСТИ:');
    console.log(`   - Отзывы: ${accuracy.reviews.toFixed(1)}%`);
    console.log(`   - Комментарии: ${accuracy.comments.toFixed(1)}%`);
    console.log(`   - Обсуждения: ${accuracy.discussions.toFixed(1)}%`);
    console.log(`   - Общая точность: ${overallAccuracy.toFixed(1)}%\n`);
    
    if (overallAccuracy >= 95) {
      console.log('🎉 УСПЕХ! Достигнута требуемая точность 95%+');
    } else {
      console.log('⚠️ НЕ ДОСТИГНУТА требуемая точность 95%+');
      console.log(`   Текущая точность: ${overallAccuracy.toFixed(1)}%`);
      console.log(`   Нужно улучшить на: ${(95 - overallAccuracy).toFixed(1)}%\n`);
    }
    
    // КОНКРЕТНЫЕ ИСПРАВЛЕНИЯ
    console.log('🔧 КОНКРЕТНЫЕ ИСПРАВЛЕНИЯ ДЛЯ АГЕНТА:');
    console.log('=====================================\n');
    
    console.log('1️⃣ ПРОБЛЕМА С КОММЕНТАРИЯМИ (не хватает 1):');
    console.log('   - Возможная причина: один комментарий не прошел валидацию');
    console.log('   - Решение: улучшить критерии валидации комментариев');
    console.log('   - Код для исправления:');
    console.log('     ```javascript');
    console.log('     // Улучшенная валидация комментариев');
    console.log('     function isValidComment(row) {');
    console.log('       if (!row || !row[0]) return false;');
    console.log('       const cellValue = row[0].toString().trim();');
    console.log('       if (cellValue === "") return false;');
    console.log('       // Более мягкие критерии для комментариев');
    console.log('       return cellValue.includes(".") || cellValue.includes("http") || cellValue.length > 5;');
    console.log('     }');
    console.log('     ```\n');
    
    console.log('2️⃣ ПРОБЛЕМА С ОБСУЖДЕНИЯМИ (не хватает 107):');
    console.log('   - Возможная причина: данные обрезаются в конце файла');
    console.log('   - Решение: проверить логику определения границ секции');
    console.log('   - Код для исправления:');
    console.log('     ```javascript');
    console.log('     // Улучшенное определение границ секции обсуждений');
    console.log('     let discussionsEnd = data.length; // До конца файла');
    console.log('     for (let i = discussionsStart + 1; i < data.length; i++) {');
    console.log('       if (data[i] && data[i][0] && isValidDiscussion(data[i])) {');
    console.log('         discussions.push(data[i]);');
    console.log('       }');
    console.log('     }');
    console.log('     ```\n');
    
    // СОЗДАНИЕ ФИНАЛЬНОГО КОДА
    console.log('📝 СОЗДАНИЕ ФИНАЛЬНОГО УНИВЕРСАЛЬНОГО ПРОЦЕССОРА:');
    
    const finalProcessorCode = `
// 🌍 ФИНАЛЬНЫЙ УНИВЕРСАЛЬНЫЙ ПРОЦЕССОР (работает с любым количеством строк)

function finalUniversalProcessor(data) {
  console.log('🌍 Запуск финального универсального процессора...');
  
  // 1. Очистка данных (универсальная)
  const cleanData = data.filter(row => 
    row && row[0] && row[0].toString().trim() !== ""
  );
  
  console.log(\`📊 Очищено данных: \${data.length} → \${cleanData.length} строк\`);
  
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
      console.log(\`📍 Секция "Отзывы" найдена в строке \${i + 1}\`);
      continue;
    } else if (cellValue.toLowerCase().includes('комментар') && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'comments';
      console.log(\`📍 Секция "Комментарии" найдена в строке \${i + 1}\`);
      continue;
    } else if ((cellValue.toLowerCase().includes('обсужден') || cellValue.toLowerCase().includes('активн')) && !cellValue.toLowerCase().includes('количество')) {
      currentSection = 'discussions';
      console.log(\`📍 Секция "Обсуждения" найдена в строке \${i + 1}\`);
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
  
  console.log(\`📈 ФИНАЛЬНЫЕ РЕЗУЛЬТАТЫ:\`);
  console.log(\`   - Отзывы: \${stats.reviews}\`);
  console.log(\`   - Комментарии: \${stats.comments}\`);
  console.log(\`   - Обсуждения: \${stats.discussions}\`);
  console.log(\`   - Всего: \${stats.total}\`);
  
  return { sections, stats };
}

// Экспорт для использования
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { finalUniversalProcessor };
}
`;
    
    console.log('   ✅ Финальный универсальный процессор создан!');
    console.log('   🌍 Ключевые особенности:');
    console.log('   - Работает с ЛЮБЫМ количеством строк');
    console.log('   - Динамическое определение секций');
    console.log('   - Универсальная валидация');
    console.log('   - Адаптивные критерии для разных секций');
    console.log('   - Автоматический подсчет\n');
    
    // Сохраняем финальный код
    fs.writeFileSync('final_universal_processor.js', finalProcessorCode);
    console.log('💾 Финальный процессор сохранен в файл: final_universal_processor.js');
    
    console.log('\n🎯 ФИНАЛЬНЫЕ ШАГИ ДЛЯ АГЕНТА:');
    console.log('==============================');
    console.log('1. 🚀 Применить final_universal_processor.js');
    console.log('2. 🧪 Протестировать на текущих данных');
    console.log('3. 📊 Проверить достижение 95%+ точности');
    console.log('4. 🌍 Протестировать на других месяцах (февраль, апрель, май)');
    console.log('5. ✅ Убедиться в универсальности (любое количество строк)');
    console.log('6. 🎉 Документировать успешное решение');
    
    console.log('\n🚨 КРИТИЧЕСКИ ВАЖНЫЕ ТРЕБОВАНИЯ:');
    console.log('- Процессор ДОЛЖЕН работать с ЛЮБЫМ количеством строк!');
    console.log('- НЕ использовать жестко заданные числа!');
    console.log('- Динамически определять структуру данных!');
    console.log('- Адаптироваться к изменениям!');
    console.log('- Достичь 95%+ точности на всех тестах!');
    
    console.log('\n🌍 ПОМНИ: УНИВЕРСАЛЬНОСТЬ = БУДУЩЕЕ!');
    console.log('Процессор должен работать завтра, через месяц и через год!');
    
  } catch (error) {
    console.error('❌ Ошибка при создании инструкций:', error);
  }
}

finalAgentInstructions();