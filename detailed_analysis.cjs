const ExcelJS = require('exceljs');
async function detailedAnalysis() {
  try {
    console.log(' АНАЛИЗ ФАЙЛА...');
    const filePath = './uploads/temp_google_sheets_1751817154291_Март_2025_результат_20250706.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);
    
    let reviewsCount = 0;
    let topCommentsCount = 0;
    let discussionsCount = 0;
    let currentSection = '';
    
    worksheet.eachRow((row, rowNumber) => {
      const firstCell = row.getCell(1).value;
      if (firstCell) {
        const cellValue = firstCell.toString();
        
        if (cellValue.includes('Отзывы') && !cellValue.includes('Комментарии')) {
          currentSection = 'reviews';
          console.log(` Секция "Отзывы" найдена в строке ${rowNumber}`);
        } else if (cellValue.includes('Комментарии Топ-20')) {
          currentSection = 'topComments';
          console.log(` Секция "Комментарии Топ-20" найдена в строке ${rowNumber}`);
        } else if (cellValue.includes('Активные обсуждения')) {
          currentSection = 'discussions';
          console.log(` Секция "Активные обсуждения" найдена в строке ${rowNumber}`);
        } else if (currentSection && cellValue && !cellValue.includes('Продукт') && !cellValue.includes('Период') && !cellValue.includes('План')) {
          if (currentSection === 'reviews') {
            reviewsCount++;
          } else if (currentSection === 'topComments') {
            topCommentsCount++;
          } else if (currentSection === 'discussions') {
            discussionsCount++;
          }
        }
      }
    });
    
    console.log(` Отзывы: ${reviewsCount}`);
    console.log(` Комментарии Топ-20: ${topCommentsCount}`);
    console.log(` Активные обсуждения: ${discussionsCount}`);
    console.log(` Общее количество: ${reviewsCount + topCommentsCount + discussionsCount}`);
    
  } catch (error) {
    console.error(' Ошибка:', error.message);
  }
}
detailedAnalysis();
