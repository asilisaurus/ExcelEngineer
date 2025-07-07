const axios = require('axios');
const XLSX = require('xlsx');

function extractGoogleSheetsInfo(url) {
  try {
    const patterns = [
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)\/edit/,
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)\/edit#gid=(\d+)/
    ];
    
    let spreadsheetId = null;
    let gid = null;
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) {
        spreadsheetId = match[1];
        if (match[2]) {
          gid = match[2];
        }
        break;
      }
    }
    
    if (!gid) {
      const gidMatch = url.match(/[#&]gid=(\d+)/);
      if (gidMatch) {
        gid = gidMatch[1];
      }
    }
    
    return { spreadsheetId, gid };
  } catch (error) {
    console.error('Ошибка при разборе URL:', error);
    return { spreadsheetId: null, gid: null };
  }
}

async function importFromGoogleSheets(url) {
  const { spreadsheetId, gid } = extractGoogleSheetsInfo(url);
  
  if (!spreadsheetId) {
    throw new Error('Неверный URL Google Таблиц');
  }
  
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx${gid ? `&gid=${gid}` : ''}`;
  
  try {
    const response = await axios.get(exportUrl, {
      responseType: 'arraybuffer',
      timeout: 30000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    return Buffer.from(response.data);
  } catch (error) {
    throw new Error(`Ошибка загрузки: ${error.message}`);
  }
}

async function analyzeResultFile() {
  console.log('🔍 АНАЛИЗ ФАЙЛА РЕЗУЛЬТАТА');
  console.log('=========================');
  
  const resultUrl = 'https://docs.google.com/spreadsheets/d/18jkeKNIn5QJpzQrDN3RAT3mEYRXlSNKOnNvA3pxoBx8/edit?gid=535445992#gid=535445992';
  
  try {
    const resultBuffer = await importFromGoogleSheets(resultUrl);
    const resultWorkbook = XLSX.read(resultBuffer, { type: 'buffer' });
    
    const sheetName = resultWorkbook.SheetNames[0];
    const sheet = resultWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    console.log(`📊 Всего строк в результате: ${data.length}`);
    
    // Анализ секций
    let reviewSection = false;
    let commentSection = false;
    let reviewCount = 0;
    let commentCount = 0;
    
    console.log('\n📝 АНАЛИЗ ОТЗЫВОВ В РЕЗУЛЬТАТЕ:');
    console.log('==============================');
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row && row[0]) {
        const cellText = String(row[0]).toLowerCase();
        
        if (cellText === 'отзывы') {
          reviewSection = true;
          commentSection = false;
          console.log(`📋 Секция отзывов начинается с строки ${i + 1}`);
          continue;
        }
        
        if (cellText.includes('комментар')) {
          reviewSection = false;
          commentSection = true;
          console.log(`💬 Секция комментариев начинается с строки ${i + 1}`);
          console.log(`📊 Всего отзывов в результате: ${reviewCount}`);
          continue;
        }
        
        if (reviewSection && row[1]) { // Есть данные во втором столбце
          reviewCount++;
          if (reviewCount <= 5) {
            const text = (row[2] || '').toString().substring(0, 80);
            console.log(`📝 Отзыв ${reviewCount}: ${text}...`);
          }
        }
        
        if (commentSection && row[1]) { // Есть данные во втором столбце
          commentCount++;
          if (commentCount <= 5) {
            const text = (row[2] || '').toString().substring(0, 80);
            console.log(`💬 Комментарий ${commentCount}: ${text}...`);
          }
        }
      }
    }
    
    console.log(`\n📊 ИТОГОВАЯ СТАТИСТИКА РЕЗУЛЬТАТА:`);
    console.log(`📝 Отзывы: ${reviewCount}`);
    console.log(`💬 Комментарии: ${commentCount}`);
    console.log(`📊 Всего: ${reviewCount + commentCount}`);
    
    // Поиск статистики внизу файла
    console.log('\n🔍 ПОИСК СТАТИСТИКИ В КОНЦЕ ФАЙЛА:');
    console.log('==================================');
    
    for (let i = data.length - 20; i < data.length; i++) {
      if (data[i] && data[i].length > 0) {
        const row = data[i];
        const rowText = row.join(' ').toLowerCase();
        
        if (rowText.includes('просмотр') || rowText.includes('3398560') || 
            rowText.includes('519') || rowText.includes('18')) {
          console.log(`📊 Строка ${i + 1}: ${row.join(' | ')}`);
        }
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка при анализе результата:', error.message);
  }
}

analyzeResultFile().catch(console.error); 