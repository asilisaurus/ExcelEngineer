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

async function analyzeStructure() {
  console.log('🔍 ДЕТАЛЬНЫЙ АНАЛИЗ СТРУКТУРЫ ДАННЫХ');
  console.log('====================================');
  
  const sourceUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?gid=1258324199#gid=1258324199';
  
  try {
    const sourceBuffer = await importFromGoogleSheets(sourceUrl);
    const sourceWorkbook = XLSX.read(sourceBuffer, { type: 'buffer' });
    
    const sheetName = sourceWorkbook.SheetNames[0];
    const sheet = sourceWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    console.log('📊 АНАЛИЗ КЛЮЧЕВЫХ СТРОК:');
    console.log('========================');
    
    // Анализируем первые 10 строк
    for (let i = 0; i < Math.min(10, data.length); i++) {
      console.log(`\n📋 СТРОКА ${i + 1}:`);
      
      if (data[i]) {
        data[i].forEach((cell, colIndex) => {
          if (cell !== undefined && cell !== null && cell !== '') {
            console.log(`  Колонка ${colIndex} (${String.fromCharCode(65 + colIndex)}): ${cell}`);
          }
        });
      }
    }
    
    console.log('\n🔍 ПОИСК КЛЮЧЕВЫХ ЗНАЧЕНИЙ:');
    console.log('===========================');
    
    // Ищем строки с важными данными
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row && row.length > 0) {
        
        // Поиск чисел 18, 519, 3398560
        const rowText = row.join(' ').toLowerCase();
        
        if (rowText.includes('18') || rowText.includes('519') || rowText.includes('3398560')) {
          console.log(`\n📍 СТРОКА ${i + 1} (содержит ключевые значения):`);
          row.forEach((cell, colIndex) => {
            if (cell !== undefined && cell !== null && cell !== '') {
              console.log(`  ${String.fromCharCode(65 + colIndex)}${i + 1}: ${cell}`);
            }
          });
        }
        
        // Ищем строки с отзывами и комментариями
        if (rowText.includes('отзыв') || rowText.includes('комментар') || rowText.includes('топ')) {
          console.log(`\n📝 СТРОКА ${i + 1} (отзывы/комментарии):`);
          row.forEach((cell, colIndex) => {
            if (cell !== undefined && cell !== null && cell !== '') {
              console.log(`  ${String.fromCharCode(65 + colIndex)}${i + 1}: ${cell}`);
            }
          });
        }
      }
    }
    
    console.log('\n🔍 ПОИСК СТАТИСТИЧЕСКИХ ДАННЫХ:');
    console.log('==============================');
    
    // Ищем числовые значения в определенных диапазонах
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row && row.length > 0) {
        row.forEach((cell, colIndex) => {
          if (typeof cell === 'number') {
            if (cell === 18 || cell === 519 || cell === 3398560 || cell === 22 || cell === 650) {
              console.log(`📊 ЧИСЛО ${cell} найдено в ${String.fromCharCode(65 + colIndex)}${i + 1}`);
            }
          }
        });
      }
    }
    
    console.log('\n🎯 АНАЛИЗ ДИАПАЗОНОВ ПО PYTHON ЛОГИКЕ:');
    console.log('=====================================');
    
    // Проверяем диапазоны как в Python коде
    console.log('\n📊 ОТЗЫВЫ OTZ (строки 6-15):');
    let otzCount = 0;
    for (let i = 6; i < 15; i++) {
      if (data[i]) {
        // Проверяем, есть ли данные в ключевых колонках
        const hasData = data[i][1] || data[i][3] || data[i][4]; // B, D, E
        if (hasData) {
          otzCount++;
          console.log(`  Строка ${i + 1}: ${data[i][1] || ''} | ${data[i][3] || ''} | ${data[i][4] || ''}`);
        }
      }
    }
    console.log(`📝 Отзывы OTZ: ${otzCount}`);
    
    console.log('\n📊 ОТЗЫВЫ APT (строки 15-28):');
    let aptCount = 0;
    for (let i = 15; i < 28; i++) {
      if (data[i]) {
        const hasData = data[i][1] || data[i][3] || data[i][4]; // B, D, E
        if (hasData) {
          aptCount++;
          console.log(`  Строка ${i + 1}: ${data[i][1] || ''} | ${data[i][3] || ''} | ${data[i][4] || ''}`);
        }
      }
    }
    console.log(`📝 Отзывы APT: ${aptCount}`);
    
    console.log('\n📊 ТОП-20 (строки 31-51):');
    let top20Count = 0;
    for (let i = 31; i < 51; i++) {
      if (data[i]) {
        const hasData = data[i][1] || data[i][3] || data[i][4]; // B, D, E
        if (hasData) {
          top20Count++;
          console.log(`  Строка ${i + 1}: ${data[i][1] || ''} | ${data[i][3] || ''} | ${data[i][4] || ''}`);
        }
      }
    }
    console.log(`📝 Комментарии ТОП-20: ${top20Count}`);
    
    console.log('\n🎯 ИТОГОВЫЕ РЕЗУЛЬТАТЫ:');
    console.log('=======================');
    console.log(`Отзывы OTZ: ${otzCount}`);
    console.log(`Отзывы APT: ${aptCount}`);
    console.log(`Всего отзывов: ${otzCount + aptCount}`);
    console.log(`Комментарии ТОП-20: ${top20Count}`);
    console.log(`Всего комментариев: ${top20Count}` + ' (активные обсуждения = 0)');
    
    // Извлекаем просмотры из заголовка
    const viewsMatch = data[0] && data[0][5] && String(data[0][5]).match(/(\d+)/);
    const totalViews = viewsMatch ? parseInt(viewsMatch[1]) : 0;
    console.log(`Просмотры: ${totalViews.toLocaleString()}`);
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error.message);
  }
}

analyzeStructure().catch(console.error); 