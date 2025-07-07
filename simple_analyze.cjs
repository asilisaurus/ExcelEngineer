const ExcelJS = require('exceljs');
async function analyzeUploadedFile() {
  try {
    console.log(' Анализ выгруженного файла...');
    const filePath = './test_download.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);
    let totalRows = 0;
    let discussionsCount = 0;
    worksheet.eachRow((row, rowNumber) => {
      totalRows++;
      const firstCell = row.getCell(1).value;
      if (firstCell && firstCell.toString().includes('Комментарии в обсуждениях')) {
        discussionsCount++;
        if (discussionsCount <= 10) {
          console.log( Обсуждение  (строка ): ...);
        }
      }
    });
    console.log( Всего строк в файле: );
    console.log( Найдено обсуждений: );
  } catch (error) {
    console.error(' Ошибка при анализе:', error.message);
  }
}
analyzeUploadedFile();
