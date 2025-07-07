const axios = require('axios');
const fs = require('fs');
const ExcelJS = require('exceljs');

async function testGoogleSheetsExport() {
  try {
    console.log('=== ТЕСТ GOOGLE SHEETS ЭКСПОРТА ===');
    
    const url = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?usp=sharing';
    
    // Извлекаем ID таблицы
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      throw new Error('Неверный URL');
    }
    
    const spreadsheetId = match[1];
    console.log('Spreadsheet ID:', spreadsheetId);
    
    // Тестируем разные способы экспорта
    const exportFormats = [
      { format: 'xlsx', description: 'Excel (.xlsx)' },
      { format: 'csv', description: 'CSV' },
      { format: 'tsv', description: 'TSV' }
    ];
    
    for (const fmt of exportFormats) {
      console.log(`\\n📥 Тестируем экспорт в формате: ${fmt.description}`);
      
      const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=${fmt.format}`;
      console.log('Export URL:', exportUrl);
      
      try {
        const response = await axios.get(exportUrl, {
          responseType: 'arraybuffer',
          timeout: 30000,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
          }
        });
        
        console.log(`✅ Размер файла: ${response.data.byteLength} байт`);
        
        if (fmt.format === 'xlsx') {
          // Анализируем Excel файл
          const buffer = Buffer.from(response.data);
          const filename = `test_export_${Date.now()}.xlsx`;
          fs.writeFileSync(filename, buffer);
          
          console.log(`📁 Файл сохранен: ${filename}`);
          
          // Анализируем содержимое
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(filename);
          
          const worksheet = workbook.worksheets[0];
          console.log(`📊 Найдено листов: ${workbook.worksheets.length}`);
          console.log(`📋 Строк в первом листе: ${worksheet.actualRowCount}`);
          
          // Подсчитываем строки с данными
          let dataRows = 0;
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Пропускаем заголовок
              const hasData = row.values.some(cell => cell && cell.toString().trim() !== '');
              if (hasData) {
                dataRows++;
              }
            }
          });
          
          console.log(`📝 Строк с данными: ${dataRows}`);
          
          // Показываем первые и последние 5 строк
          console.log('\\n🔍 Первые 5 строк:');
          for (let i = 1; i <= Math.min(6, worksheet.actualRowCount); i++) {
            const row = worksheet.getRow(i);
            const colA = row.getCell(1).value || '';
            const colB = row.getCell(2).value || '';
            const colE = row.getCell(5).value || '';
            console.log(`  ${i}: A="${colA}" | B="${colB}" | E="${colE.toString().substring(0, 50)}..."`);
          }
          
          console.log('\\n🔍 Последние 5 строк:');
          for (let i = Math.max(1, worksheet.actualRowCount - 4); i <= worksheet.actualRowCount; i++) {
            const row = worksheet.getRow(i);
            const colA = row.getCell(1).value || '';
            const colB = row.getCell(2).value || '';
            const colE = row.getCell(5).value || '';
            console.log(`  ${i}: A="${colA}" | B="${colB}" | E="${colE.toString().substring(0, 50)}..."`);
          }
          
          // Удаляем тестовый файл
          fs.unlinkSync(filename);
        }
        
      } catch (error) {
        console.log(`❌ Ошибка: ${error.message}`);
      }
    }
    
    console.log('\\n=== ТЕСТ ЗАВЕРШЕН ===');
    
  } catch (error) {
    console.error('❌ Критическая ошибка:', error);
  }
}

testGoogleSheetsExport(); 