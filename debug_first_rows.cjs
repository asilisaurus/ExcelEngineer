const ExcelJS = require('exceljs');

async function debugFirstRows() {
  try {
    console.log('=== АНАЛИЗ ПЕРВЫХ СТРОК ФАЙЛА ===');
    
    const filePath = 'uploads/temp_google_sheets_1751807589999.xlsx';
    
    // Читаем файл
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];
    
    // Показываем первые 30 строк
    for (let i = 1; i <= 30; i++) {
      const row = worksheet.getRow(i);
      const values = [];
      
      // Получаем первые 8 колонок
      for (let j = 1; j <= 8; j++) {
        const cell = row.getCell(j);
        let value = cell.value;
        if (value && typeof value === 'string') {
          value = value.length > 50 ? value.substring(0, 50) + '...' : value;
        }
        values.push(value || '');
      }
      
      console.log(`${i.toString().padStart(2)}: [${values.join(' | ')}]`);
    }
    
    console.log('\\n=== ПОИСК ОТЗЫВОВ ===');
    
    // Поиск строк с otzovik, irecommend и т.д.
    let foundReviews = 0;
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const colB = (row.getCell(2).value || '').toString().toLowerCase();
        const colD = (row.getCell(4).value || '').toString().toLowerCase();
        const urlText = colB + ' ' + colD;
        
        if (/otzovik|irecommend|otzyvru|pravogolosa|medum|vseotzyvy|otzyvy\\.pro|market\\.yandex|dialog\\.ru|goodapteka|megapteka|uteka|nfapteka|piluli\\.ru|eapteka\\.ru|pharmspravka\\.ru|gde\\.ru|ozon\\.ru/i.test(urlText)) {
          foundReviews++;
          if (foundReviews <= 5) {
            console.log(`Строка ${rowNumber}: [${colB}] [${colD}]`);
          }
        }
      }
    });
    
    console.log(`\\n🎯 Найдено отзывов: ${foundReviews}`);
    
  } catch (error) {
    console.error('❌ Ошибка:', error);
  }
}

debugFirstRows(); 