const ExcelJS = require('exceljs');

async function debugDateDetection() {
  console.log('🔍 ОТЛАДКА ОПРЕДЕЛЕНИЯ ДАТ');
  console.log('='.repeat(50));
  
  try {
    const sourceFile = 'uploads/Fortedetrim ORM report source.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourceFile);
    
    const sheet = workbook.getWorksheet(1);
    
    console.log('\n📊 АНАЛИЗ ИСХОДНЫХ ДАННЫХ С ДАТАМИ:');
    
    // Проверяем строки с данными (начиная с 6-й строки где есть данные)
    for (let i = 6; i <= 15; i++) {
      const row = sheet.getRow(i);
      
      console.log(`\n--- СТРОКА ${i} ---`);
      
      // Проверяем колонки где могут быть даты
      const potentialDateColumns = [6, 3, 5]; // G, D, F
      
      for (const colIndex of potentialDateColumns) {
        const cellValue = row.getCell(colIndex + 1).value; // ExcelJS использует 1-based индексы
        const cellText = row.getCell(colIndex + 1).text;
        
        console.log(`Колонка ${String.fromCharCode(65 + colIndex)}(${colIndex}): value="${cellValue}", text="${cellText}"`);
        
        if (cellValue) {
          const cellStr = cellValue.toString();
          
          // Тестируем наши регулярные выражения
          const isExcelNumber = typeof cellValue === 'number' && cellValue > 40000 && cellValue < 50000;
          const isDateSlash = cellStr.match(/\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4}/);
          const isDateFormat = cellStr.match(/(Mon|Tue|Wed|Thu|Fri|Sat|Sun).+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec).+\d{4}/);
          
          console.log(`  ✓ Excel number (${isExcelNumber}): ${cellValue}`);
          console.log(`  ✓ Date slash (${!!isDateSlash}): ${cellStr}`);
          console.log(`  ✓ Date format (${!!isDateFormat}): ${cellStr}`);
          
          if (isExcelNumber || isDateSlash || isDateFormat) {
            console.log(`  🎯 ДАТА НАЙДЕНА в колонке ${String.fromCharCode(65 + colIndex)}!`);
            
            // Тестируем конвертацию
            try {
              const convertedDate = convertExcelDateToString(cellValue);
              console.log(`  📅 Конвертированная дата: "${convertedDate}"`);
            } catch (e) {
              console.log(`  ❌ Ошибка конвертации: ${e.message}`);
            }
          }
        }
      }
      
      // Проверяем ники
      const potentialAuthorColumns = [7, 4, 8]; // H, E, I
      
      for (const colIndex of potentialAuthorColumns) {
        const cellValue = row.getCell(colIndex + 1).value;
        
        if (cellValue && typeof cellValue === 'string') {
          const cellStr = cellValue.toString().trim();
          const isValidAuthor = cellStr.length > 2 && cellStr.length < 50 && 
                               !cellStr.includes('http') && 
                               !cellStr.includes('.com') &&
                               !cellStr.match(/\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4}/) &&
                               !cellStr.match(/^\d+$/);
          
          if (isValidAuthor) {
            console.log(`  👤 НИК НАЙДЕН в колонке ${String.fromCharCode(65 + colIndex)}: "${cellStr}"`);
          }
        }
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
  }
}

// Копируем функцию конвертации дат из процессора
function convertExcelDateToString(dateValue) {
  if (!dateValue) {
    return '';
  }
  
  try {
    let jsDate;
    
    // Если это уже строка даты
    if (typeof dateValue === 'string') {
      // Обрабатываем формат "3/4/2025" или "03.04.2025"
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
        // Обрабатываем формат "Fri Mar 07 2025"
        jsDate = new Date(dateValue);
      } else {
        jsDate = new Date(dateValue);
      }
    }
    // Если это число Excel
    else if (typeof dateValue === 'number' && dateValue > 1) {
      jsDate = new Date((dateValue - 25569) * 86400 * 1000);
    }
    // Если это уже объект Date
    else if (dateValue instanceof Date) {
      jsDate = dateValue;
    }
    else {
      return '';
    }
    
    // Проверяем валидность даты
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

debugDateDetection(); 