const ExcelJS = require('exceljs');

async function analyzeLatestFile() {
  try {
    const filePath = './uploads/temp_google_sheets_1751818962604_Март_2025_результат_20250706.xlsx';
    
    console.log('🔍 Анализируем файл:', filePath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    
    let reviewsCount = 0;
    let commentsCount = 0;
    let discussionsCount = 0;
    
    let currentSection = '';
    let rowIndex = 0;
    
    worksheet.eachRow((row, rowNumber) => {
      const firstCell = row.getCell(1).value;
      if (firstCell) {
        const cellValue = firstCell.toString();
        
        if (cellValue.includes('Отзывы') && !cellValue.includes('Комментарии')) {
          currentSection = 'reviews';
          console.log(`📝 Найдена секция "Отзывы" в строке ${rowNumber}`);
        } else if (cellValue.includes('Комментарии Топ-20')) {
          currentSection = 'comments';
          console.log(`💬 Найдена секция "Комментарии Топ-20" в строке ${rowNumber}`);
        } else if (cellValue.includes('Активные обсуждения')) {
          currentSection = 'discussions';
          console.log(`🔥 Найдена секция "Активные обсуждения" в строке ${rowNumber}`);
        } else if (currentSection && cellValue && 
                  !cellValue.includes('Продукт') && 
                  !cellValue.includes('Период') && 
                  !cellValue.includes('План') &&
                  !cellValue.includes('Суммарное') &&
                  !cellValue.includes('Количество')) {
          // Это строка с данными
          if (currentSection === 'reviews') {
            reviewsCount++;
            if (reviewsCount <= 3) {
              console.log(`  📝 Отзыв ${reviewsCount}: ${cellValue.substring(0, 50)}...`);
            }
          } else if (currentSection === 'comments') {
            commentsCount++;
            if (commentsCount <= 3) {
              console.log(`  💬 Комментарий ${commentsCount}: ${cellValue.substring(0, 50)}...`);
            }
          } else if (currentSection === 'discussions') {
            discussionsCount++;
            if (discussionsCount <= 3) {
              console.log(`  🔥 Обсуждение ${discussionsCount}: ${cellValue.substring(0, 50)}...`);
            }
          }
        }
      }
    });
    
    console.log('\n📊 РЕЗУЛЬТАТЫ АНАЛИЗА:');
    console.log(`📝 Отзывы: ${reviewsCount}`);
    console.log(`💬 Комментарии Топ-20: ${commentsCount}`);
    console.log(`🔥 Активные обсуждения: ${discussionsCount}`);
    console.log(`📋 Общее количество записей: ${reviewsCount + commentsCount + discussionsCount}`);
    
    if (discussionsCount > 0) {
      console.log('\n✅ ПРОБЛЕМА ИСПРАВЛЕНА! Активные обсуждения присутствуют в файле!');
    } else {
      console.log('\n❌ Проблема не исправлена - активные обсуждения по-прежнему отсутствуют');
    }
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error.message);
  }
}

analyzeLatestFile(); 