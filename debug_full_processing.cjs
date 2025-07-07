const ExcelJS = require('exceljs');

async function debugFullProcessing() {
  console.log('🔍 ДЕТАЛЬНАЯ ОТЛАДКА ПОЛНОЙ ОБРАБОТКИ ДАННЫХ');
  console.log('='.repeat(60));
  
  try {
    // Шаг 1: Читаем исходные данные
    console.log('\n📂 ШАГ 1: ЧТЕНИЕ ИСХОДНЫХ ДАННЫХ');
    const sourceFile = 'uploads/Fortedetrim ORM report source.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourceFile);
    
    const worksheet = workbook.getWorksheet(1);
    const data = [];
    
    // Читаем данные так же, как в процессоре
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Пропускаем первую строку
        const rowData = [];
        for (let i = 1; i <= 15; i++) {
          const cell = row.getCell(i);
          rowData.push(cell.value);
        }
        data.push(rowData);
      }
    });
    
    console.log(`📊 Прочитано ${data.length} строк из исходного файла`);
    
    // Шаг 2: Анализируем типы строк
    console.log('\n🔍 ШАГ 2: АНАЛИЗ ТИПОВ СТРОК');
    const processedRows = [];
    
    for (let i = 0; i < Math.min(data.length, 20); i++) {
      const row = data[i];
      const type = analyzeRowType(row);
      
      console.log(`\nСтрока ${i + 2}: Тип = "${type}"`);
      console.log(`  Колонка A: "${row[0]}"`);
      console.log(`  Колонка B: "${row[1]}"`);
      console.log(`  Колонка E: "${row[4]}"`);
      console.log(`  Колонка G: "${row[6]}"`);
      console.log(`  Колонка H: "${row[7]}"`);
      
      if (type === 'review_otzovik' || type === 'review_pharmacy' || type === 'comment') {
        console.log(`  ✅ Строка подходит для обработки`);
        
        // Имитируем логику поиска дат и авторов
        let dateValue = '';
        let authorValue = '';
        
        // Ищем дату
        const potentialDateColumns = [6, 3, 5]; // G, D, F
        for (const colIndex of potentialDateColumns) {
          const cellValue = row[colIndex];
          if (cellValue) {
            const cellStr = cellValue.toString();
            console.log(`  🔍 Проверяем дату в колонке ${String.fromCharCode(65 + colIndex)}: "${cellStr}"`);
            
            if (typeof cellValue === 'number' && cellValue > 40000 && cellValue < 50000) {
              dateValue = convertExcelDateToString(cellValue);
              console.log(`  📅 ДАТА НАЙДЕНА (Excel число): ${dateValue}`);
              break;
            } else if (cellStr.match(/\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4}/)) {
              dateValue = convertExcelDateToString(cellValue);
              console.log(`  📅 ДАТА НАЙДЕНА (формат слеш): ${dateValue}`);
              break;
            } else if (cellStr.match(/(Mon|Tue|Wed|Thu|Fri|Sat|Sun).+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec).+\d{4}/)) {
              dateValue = convertExcelDateToString(cellValue);
              console.log(`  📅 ДАТА НАЙДЕНА (формат день): ${dateValue}`);
              break;
            } else if (cellValue instanceof Date || (typeof cellValue === 'object' && cellValue.toString().includes('GMT'))) {
              dateValue = convertExcelDateToString(cellValue);
              console.log(`  📅 ДАТА НАЙДЕНА (Date объект): ${dateValue}`);
              break;
            }
          }
        }
        
        // Ищем автора
        const potentialAuthorColumns = [7, 4, 8]; // H, E, I
        for (const colIndex of potentialAuthorColumns) {
          const cellValue = row[colIndex];
          if (cellValue && typeof cellValue === 'string') {
            const cellStr = cellValue.toString().trim();
            const isValidAuthor = cellStr.length > 2 && cellStr.length < 50 && 
                                 !cellStr.includes('http') && 
                                 !cellStr.includes('.com') &&
                                 !cellStr.match(/\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4}/) &&
                                 !cellStr.match(/^\d+$/);
            
            if (isValidAuthor) {
              authorValue = cellStr;
              console.log(`  👤 АВТОР НАЙДЕН в колонке ${String.fromCharCode(65 + colIndex)}: "${authorValue}"`);
              break;
            }
          }
        }
        
        // Создаем запись
        const processedRow = {
          type,
          text: (row[4] || '').toString(),
          url: (row[1] || '').toString() || (row[3] || '').toString(),
          date: dateValue,
          author: authorValue,
          postType: (row[13] || '').toString(),
          originalRow: row
        };
        
        processedRows.push(processedRow);
        
        console.log(`  ➡️ РЕЗУЛЬТАТ: date="${dateValue}", author="${authorValue}"`);
        
        if (!dateValue) {
          console.log(`  ⚠️ ВНИМАНИЕ: Дата не найдена для этой строки!`);
        }
      }
    }
    
    console.log(`\n📊 Обработано ${processedRows.length} строк с данными`);
    
    // Шаг 3: Проверяем результат
    console.log('\n📋 ШАГ 3: АНАЛИЗ РЕЗУЛЬТАТА');
    const rowsWithDates = processedRows.filter(row => row.date);
    const rowsWithoutDates = processedRows.filter(row => !row.date);
    
    console.log(`✅ Строк с датами: ${rowsWithDates.length}`);
    console.log(`❌ Строк без дат: ${rowsWithoutDates.length}`);
    
    if (rowsWithDates.length > 0) {
      console.log('\n🎯 СТРОКИ С ДАТАМИ:');
      rowsWithDates.forEach((row, index) => {
        console.log(`${index + 1}. Дата: "${row.date}", Автор: "${row.author}", Тип: "${row.type}"`);
      });
    }
    
    if (rowsWithoutDates.length > 0) {
      console.log('\n❌ СТРОКИ БЕЗ ДАТ:');
      rowsWithoutDates.forEach((row, index) => {
        console.log(`${index + 1}. Автор: "${row.author}", Тип: "${row.type}"`);
        console.log(`    Исходная строка G: "${row.originalRow[6]}"`);
        console.log(`    Исходная строка D: "${row.originalRow[3]}"`);
        console.log(`    Исходная строка F: "${row.originalRow[5]}"`);
      });
    }
    
  } catch (error) {
    console.error('❌ Ошибка отладки:', error.message);
    console.error(error.stack);
  }
}

// Копируем функции из процессора для тестирования
function analyzeRowType(row) {
  if (!row || row.length === 0) return 'empty';
  
  const colA = (row[0] || '').toString().toLowerCase();
  const colB = (row[1] || '').toString().toLowerCase();
  const colD = (row[3] || '').toString().toLowerCase();
  const colE = (row[4] || '').toString().toLowerCase();
  
  // Заголовки и служебные строки
  if (colA.includes('тип размещения') || colA.includes('площадка') || 
      colB.includes('площадка') || colE.includes('текст сообщения') ||
      colA.includes('план') || colA.includes('итого')) {
    return 'header';
  }
  
  // Google Sheets специфичная логика
  if (colA.includes('отзывы (отзовики)')) {
    return 'review_otzovik';
  }
  
  if (colA.includes('отзывы (аптеки)')) {
    return 'review_pharmacy';
  }
  
  if (colA.includes('комментарии в обсуждениях')) {
    return 'comment';
  }
  
  // Анализ по URL
  const urlText = colB + ' ' + colD;
  
  // Проверяем платформы
  const reviewPlatforms = ['otzovik', 'irecommend', 'otzyvru', 'pravogolosa', 'medum', 'vseotzyvy', 'otzyvy.pro'];
  const pharmacyPlatforms = ['megapteka', 'apteka', 'pharmacy', 'piluli', 'zdravcity'];
  
  const isReviewPlatform = reviewPlatforms.some(platform => urlText.toLowerCase().includes(platform));
  const isPharmacyPlatform = pharmacyPlatforms.some(platform => urlText.toLowerCase().includes(platform));
  
  if (isReviewPlatform && (colB || colD || colE)) {
    return 'review_otzovik';
  }
  
  if (isPharmacyPlatform && (colB || colD || colE)) {
    return 'review_pharmacy';
  }
  
  // Если есть контент, но тип неясен
  if (colB || colD || colE) {
    return 'content';
  }
  
  return 'empty';
}

function convertExcelDateToString(dateValue) {
  if (!dateValue) return '';
  
  try {
    let jsDate;
    
    if (typeof dateValue === 'string') {
      if (dateValue.includes('/')) {
        const parts = dateValue.split('/');
        if (parts.length === 3) {
          const month = parseInt(parts[0]);
          const day = parseInt(parts[1]);
          const year = parseInt(parts[2]);
          jsDate = new Date(year, month - 1, day);
        } else {
          jsDate = new Date(dateValue);
        }
      } else if (dateValue.includes('.')) {
        const parts = dateValue.split('.');
        if (parts.length === 3) {
          const day = parseInt(parts[0]);
          const month = parseInt(parts[1]);
          const year = parseInt(parts[2]);
          jsDate = new Date(year, month - 1, day);
        } else {
          jsDate = new Date(dateValue);
        }
      } else if (dateValue.match(/(Mon|Tue|Wed|Thu|Fri|Sat|Sun).+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec).+\d{4}/)) {
        jsDate = new Date(dateValue);
      } else {
        jsDate = new Date(dateValue);
      }
    } else if (typeof dateValue === 'number' && dateValue > 1) {
      jsDate = new Date((dateValue - 25569) * 86400 * 1000);
    } else if (dateValue instanceof Date) {
      jsDate = dateValue;
    } else {
      return '';
    }
    
    if (isNaN(jsDate.getTime())) {
      return '';
    }
    
    const day = jsDate.getDate().toString().padStart(2, '0');
    const month = (jsDate.getMonth() + 1).toString().padStart(2, '0');
    const year = jsDate.getFullYear().toString();
    
    return `${day}.${month}.${year}`;
  } catch (error) {
    return '';
  }
}

debugFullProcessing(); 