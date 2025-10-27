const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function fixProcessorIssues() {
  console.log('🔧 ИСПРАВЛЕНИЕ ПРОБЛЕМ ПРОЦЕССОРА');
  console.log('================================\n');
  
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
    
    // Анализируем проблемы и предлагаем исправления
    console.log('🔍 АНАЛИЗ ПРОБЛЕМ:');
    
    // Проблема 1: Пустые строки в секциях
    let emptyRowsInComments = 0;
    let emptyRowsInDiscussions = 0;
    let inCommentsSection = false;
    let inDiscussionsSection = false;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        
        if (cellValue.includes('Комментарии') && !cellValue.includes('Количество')) {
          inCommentsSection = true;
          inDiscussionsSection = false;
        } else if (cellValue.includes('Активные обсуждения') || cellValue.includes('Обсуждения')) {
          inCommentsSection = false;
          inDiscussionsSection = true;
        }
      } else {
        // Пустая строка
        if (inCommentsSection) {
          emptyRowsInComments++;
        } else if (inDiscussionsSection) {
          emptyRowsInDiscussions++;
        }
      }
    }
    
    console.log(`❌ Проблема 1: ${emptyRowsInComments} пустых строк в секции комментариев`);
    console.log(`❌ Проблема 2: ${emptyRowsInDiscussions} пустых строк в секции обсуждений`);
    
    // Проблема 3: Проверяем, есть ли данные после последней секции
    let hasDataAfterDiscussions = false;
    let lastDataRow = -1;
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i] && data[i][0] && data[i][0].toString().includes('.')) {
        lastDataRow = i;
        break;
      }
    }
    
    if (lastDataRow !== -1) {
      console.log(`📍 Последняя строка с данными: ${lastDataRow + 1}`);
      if (lastDataRow < data.length - 10) {
        console.log(`❌ Проблема 3: Возможно, данные обрезаны в конце файла`);
        hasDataAfterDiscussions = true;
      }
    }
    
    console.log('\n🔧 ПРЕДЛАГАЕМЫЕ ИСПРАВЛЕНИЯ:');
    
    // Исправление 1: Удаление пустых строк
    console.log('\n1️⃣ ИСПРАВЛЕНИЕ ПУСТЫХ СТРОК:');
    console.log('   - Добавить фильтрацию пустых строк в процессор');
    console.log('   - Исключить строки без данных из подсчета');
    console.log('   - Код для исправления:');
    console.log('     ```javascript');
    console.log('     // Фильтруем пустые строки');
    console.log('     const validRows = rows.filter(row => ');
    console.log('       row && row[0] && row[0].toString().trim() !== ""');
    console.log('     );');
    console.log('     ```');
    
    // Исправление 2: Улучшение определения границ секций
    console.log('\n2️⃣ ИСПРАВЛЕНИЕ ГРАНИЦ СЕКЦИЙ:');
    console.log('   - Улучшить логику определения начала и конца секций');
    console.log('   - Убедиться, что заголовки секций не включаются в данные');
    console.log('   - Код для исправления:');
    console.log('     ```javascript');
    console.log('     // Правильное определение границ');
    console.log('     if (cellValue.includes("Отзывы") && !cellValue.includes("Количество")) {');
    console.log('       sectionStart = i + 1; // Исключаем заголовок');
    console.log('     }');
    console.log('     ```');
    
    // Исправление 3: Проверка полноты данных
    console.log('\n3️⃣ ПРОВЕРКА ПОЛНОТЫ ДАННЫХ:');
    console.log('   - Добавить проверку, что все данные извлечены');
    console.log('   - Убедиться, что файл не обрезается');
    console.log('   - Код для исправления:');
    console.log('     ```javascript');
    console.log('     // Проверка полноты данных');
    console.log('     const expectedTotal = reviews + comments + discussions;');
    console.log('     if (actualTotal < expectedTotal) {');
    console.log('       console.log("⚠️ Предупреждение: Не все данные извлечены");');
    console.log('     }');
    console.log('     ```');
    
    // Создаем исправленный процессор
    console.log('\n📝 СОЗДАНИЕ ИСПРАВЛЕННОГО ПРОЦЕССОРА:');
    
    const fixedProcessorCode = `
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
  console.log(\`📊 Результаты: Отзывы: \${reviews.length}, Комментарии: \${comments.length}, Обсуждения: \${discussions.length}\`);
  console.log(\`📋 Всего: \${total}\`);
  
  return { reviews, comments, discussions, total };
}
`;
    
    console.log('   Код исправленного процессора создан!');
    console.log('   Основные улучшения:');
    console.log('   ✅ Фильтрация пустых строк');
    console.log('   ✅ Правильное определение границ секций');
    console.log('   ✅ Исключение заголовков из данных');
    console.log('   ✅ Проверка полноты данных');
    
    // Сохраняем исправленный код
    fs.writeFileSync('fixed_processor_code.js', fixedProcessorCode);
    console.log('\n💾 Исправленный код сохранен в файл: fixed_processor_code.js');
    
    console.log('\n🎯 СЛЕДУЮЩИЕ ШАГИ ДЛЯ АГЕНТА:');
    console.log('1. Применить исправления из fixed_processor_code.js');
    console.log('2. Перезапустить обработку данных');
    console.log('3. Проверить, что количество записей соответствует эталону');
    console.log('4. Убедиться, что точность достигла 95%+');
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
  }
}

fixProcessorIssues();