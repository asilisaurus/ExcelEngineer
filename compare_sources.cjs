const ExcelJS = require('exceljs');
const https = require('https');
const fs = require('fs');

async function downloadGoogleSheet(url) {
  // Конвертируем Google Sheets URL в Excel export URL
  const sheetId = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)?.[1];
  if (!sheetId) {
    throw new Error('Не удалось извлечь ID таблицы из URL');
  }
  
  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  const tempFilePath = './temp_google_download.xlsx';
  
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(tempFilePath);
    
    https.get(exportUrl, (response) => {
      if (response.statusCode !== 200) {
        reject(new Error(`HTTP ${response.statusCode}: ${response.statusMessage}`));
        return;
      }
      
      response.pipe(file);
      
      file.on('finish', () => {
        file.close();
        console.log('📁 Google Sheets файл скачан:', tempFilePath);
        resolve(tempFilePath);
      });
      
      file.on('error', (err) => {
        fs.unlink(tempFilePath, () => {});
        reject(err);
      });
    }).on('error', (err) => {
      reject(err);
    });
  });
}

async function analyzeSources() {
  try {
    console.log('🔍 Сравнение источников данных...\n');

    // 1. Скачиваем из Google Sheets
    console.log('=== ЗАГРУЗКА ИЗ GOOGLE SHEETS ===');
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing';
    const googleFilePath = await downloadGoogleSheet(googleSheetsUrl);
    
    // 2. Анализируем Google Sheets файл
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
    
    // 3. Анализируем выгруженный файл
    console.log('\n=== АНАЛИЗ ВЫГРУЖЕННОГО ФАЙЛА ===');
    const uploadedFilePath = './test_download.xlsx';
    
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
    
    // 4. Сравниваем данные
    console.log('\n=== СРАВНЕНИЕ ДАННЫХ ===');
    console.log(`Google Sheets строк: ${googleRows}`);
    console.log(`Выгруженный файл строк: ${uploadedRows}`);
    console.log(`Разница: ${Math.abs(googleRows - uploadedRows)}`);
    
    // Ищем "Комментарии в обсуждениях"
    const googleDiscussions = googleDataRows.filter(row => 
      row[0] && row[0].toString().includes('Комментарии в обсуждениях')
    );
    const uploadedDiscussions = uploadedDataRows.filter(row => 
      row[0] && row[0].toString().includes('Комментарии в обсуждениях')
    );
    
    console.log(`\n📊 "Комментарии в обсуждениях":`);
    console.log(`Google Sheets: ${googleDiscussions.length}`);
    console.log(`Выгруженный файл: ${uploadedDiscussions.length}`);
    
    // Показываем диапазон строк с обсуждениями
    if (googleDiscussions.length > 0) {
      const firstDiscussion = googleDataRows.findIndex(row => 
        row[0] && row[0].toString().includes('Комментарии в обсуждениях')
      );
      const lastDiscussion = googleDataRows.map((row, index) => 
        row[0] && row[0].toString().includes('Комментарии в обсуждениях') ? index : -1
      ).filter(i => i !== -1).pop();
      
      console.log(`Google Sheets: обсуждения в строках ${firstDiscussion + 1} - ${lastDiscussion + 1}`);
    }
    
    if (uploadedDiscussions.length > 0) {
      const firstDiscussion = uploadedDataRows.findIndex(row => 
        row[0] && row[0].toString().includes('Комментарии в обсуждениях')
      );
      const lastDiscussion = uploadedDataRows.map((row, index) => 
        row[0] && row[0].toString().includes('Комментарии в обсуждениях') ? index : -1
      ).filter(i => i !== -1).pop();
      
      console.log(`Выгруженный файл: обсуждения в строках ${firstDiscussion + 1} - ${lastDiscussion + 1}`);
    }
    
    // Проверяем заголовок
    console.log('\n=== ЗАГОЛОВОК ===');
    console.log('Google Sheets:');
    console.log(`  Строка 1: ${googleDataRows[0]?.slice(0, 4).map(c => c ? c.toString().substring(0, 50) : '').join(' | ')}`);
    console.log('Выгруженный файл:');
    console.log(`  Строка 1: ${uploadedDataRows[0]?.slice(0, 4).map(c => c ? c.toString().substring(0, 50) : '').join(' | ')}`);
    
    // Очищаем временный файл
    fs.unlinkSync(googleFilePath);
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
  }
}

analyzeSources(); 