const { GoogleSheetsService } = require('./dist/server/services/google-sheets.js');
const ExcelJS = require('exceljs');

async function testBothSources() {
  try {
    console.log('🔍 Тестирование обеих версий источников данных...\n');

    // 1. Тестируем загрузку из Google Sheets
    console.log('=== ТЕСТ 1: ЗАГРУЗКА ИЗ GOOGLE SHEETS ===');
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing';
    
    const googleSheetsService = new GoogleSheetsService();
    const googleFilePath = await googleSheetsService.downloadAndSaveSheet(googleSheetsUrl);
    
    console.log('📁 Google Sheets файл скачан:', googleFilePath);
    
    // Анализируем скачанный файл
    const workbook1 = new ExcelJS.Workbook();
    await workbook1.xlsx.readFile(googleFilePath);
    
    const worksheet1 = workbook1.getWorksheet(1);
    let googleRows = 0;
    let googleDataRows = [];
    
    worksheet1.eachRow((row, rowNumber) => {
      googleRows++;
      const rowData = [];
      row.eachCell((cell, colNumber) => {
        rowData[colNumber - 1] = cell.value;
      });
      googleDataRows.push(rowData);
    });
    
    console.log(`📊 Google Sheets - всего строк: ${googleRows}`);
    console.log(`📊 Google Sheets - первые 5 строк:`)
    googleDataRows.slice(0, 5).forEach((row, i) => {
      console.log(`  ${i + 1}: ${row.slice(0, 3).map(c => c ? c.toString().substring(0, 50) : '').join(' | ')}`);
    });
    
    // 2. Тестируем загрузку выгруженного файла
    console.log('\n=== ТЕСТ 2: ЗАГРУЗКА ВЫГРУЖЕННОГО ФАЙЛА ===');
    const uploadedFilePath = './test_download.xlsx';
    
    console.log('📁 Выгруженный файл:', uploadedFilePath);
    
    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(uploadedFilePath);
    
    const worksheet2 = workbook2.getWorksheet(1);
    let uploadedRows = 0;
    let uploadedDataRows = [];
    
    worksheet2.eachRow((row, rowNumber) => {
      uploadedRows++;
      const rowData = [];
      row.eachCell((cell, colNumber) => {
        rowData[colNumber - 1] = cell.value;
      });
      uploadedDataRows.push(rowData);
    });
    
    console.log(`📊 Выгруженный файл - всего строк: ${uploadedRows}`);
    console.log(`📊 Выгруженный файл - первые 5 строк:`)
    uploadedDataRows.slice(0, 5).forEach((row, i) => {
      console.log(`  ${i + 1}: ${row.slice(0, 3).map(c => c ? c.toString().substring(0, 50) : '').join(' | ')}`);
    });
    
    // 3. Сравниваем данные
    console.log('\n=== СРАВНЕНИЕ ДАННЫХ ===');
    console.log(`Google Sheets строк: ${googleRows}`);
    console.log(`Выгруженный файл строк: ${uploadedRows}`);
    console.log(`Разница: ${Math.abs(googleRows - uploadedRows)}`);
    
    // Проверяем количество "Комментарии в обсуждениях"
    const googleDiscussions = googleDataRows.filter(row => 
      row[0] && row[0].toString().includes('Комментарии в обсуждениях')
    );
    const uploadedDiscussions = uploadedDataRows.filter(row => 
      row[0] && row[0].toString().includes('Комментарии в обсуждениях')
    );
    
    console.log(`\n📊 "Комментарии в обсуждениях":`);
    console.log(`Google Sheets: ${googleDiscussions.length}`);
    console.log(`Выгруженный файл: ${uploadedDiscussions.length}`);
    
    // Проверяем отзывы
    const googleReviews = googleDataRows.filter(row => 
      row[0] && (row[0].toString().includes('Отзыв') || row[0].toString().includes('отзыв'))
    );
    const uploadedReviews = uploadedDataRows.filter(row => 
      row[0] && (row[0].toString().includes('Отзыв') || row[0].toString().includes('отзыв'))
    );
    
    console.log(`\n📊 Отзывы:`);
    console.log(`Google Sheets: ${googleReviews.length}`);
    console.log(`Выгруженный файл: ${uploadedReviews.length}`);
    
    if (googleRows === uploadedRows && googleDiscussions.length === uploadedDiscussions.length) {
      console.log('\n✅ Данные идентичны!');
    } else {
      console.log('\n⚠️ Данные различаются - требуется дальнейшее исследование');
    }
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
    console.error('Stack:', error.stack);
  }
}

testBothSources(); 