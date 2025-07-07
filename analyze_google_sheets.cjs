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
  console.log('=== ИМПОРТ ИЗ GOOGLE ТАБЛИЦ ===');
  console.log('URL:', url);
  
  const { spreadsheetId, gid } = extractGoogleSheetsInfo(url);
  
  if (!spreadsheetId) {
    throw new Error('Неверный URL Google Таблиц');
  }
  
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx${gid ? `&gid=${gid}` : ''}`;
  
  console.log('Export URL:', exportUrl);
  
  try {
    const response = await axios.get(exportUrl, {
      responseType: 'arraybuffer',
      timeout: 30000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    console.log('Данные загружены, размер:', response.data.byteLength, 'байт');
    
    return Buffer.from(response.data);
    
  } catch (error) {
    console.error('Ошибка при загрузке из Google Таблиц:', error.message);
    
    if (error.response?.status === 403) {
      throw new Error('Доступ к таблице запрещен. Убедитесь, что таблица доступна для просмотра.');
    } else if (error.response?.status === 404) {
      throw new Error('Таблица не найдена. Проверьте URL.');
    } else {
      throw new Error(`Ошибка загрузки: ${error.message}`);
    }
  }
}

async function analyzeGoogleSheets() {
  const sourceUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?gid=1258324199#gid=1258324199';
  const resultUrl = 'https://docs.google.com/spreadsheets/d/18jkeKNIn5QJpzQrDN3RAT3mEYRXlSNKOnNvA3pxoBx8/edit?gid=535445992#gid=535445992';
  
  console.log('🔍 АНАЛИЗ ИСХОДНОГО ФАЙЛА');
  console.log('========================');
  
  try {
    const sourceBuffer = await importFromGoogleSheets(sourceUrl);
    const sourceWorkbook = XLSX.read(sourceBuffer, { type: 'buffer' });
    
    console.log('📊 ИСХОДНЫЙ ФАЙЛ - ЛИСТЫ:', sourceWorkbook.SheetNames);
    
    // Анализируем первый лист
    const sourceSheetName = sourceWorkbook.SheetNames[0];
    const sourceSheet = sourceWorkbook.Sheets[sourceSheetName];
    const sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1 });
    
    console.log('📋 ИСХОДНЫЙ ФАЙЛ - ЗАГОЛОВКИ:');
    console.log(sourceData[0]);
    
    console.log('📋 ИСХОДНЫЙ ФАЙЛ - ВСЕГО СТРОК:', sourceData.length);
    
    // Показываем первые 5 строк данных
    console.log('📋 ИСХОДНЫЙ ФАЙЛ - ПЕРВЫЕ 5 СТРОК:');
    for (let i = 0; i < Math.min(5, sourceData.length); i++) {
      console.log(`Строка ${i + 1}:`, sourceData[i]);
    }
    
    // Анализируем структуру данных
    console.log('\n📊 АНАЛИЗ СТРУКТУРЫ ИСХОДНЫХ ДАННЫХ:');
    const headers = sourceData[0];
    headers.forEach((header, index) => {
      console.log(`Колонка ${index} (${String.fromCharCode(65 + index)}): ${header}`);
    });
    
    console.log('\n' + '='.repeat(50));
    
  } catch (error) {
    console.error('❌ Ошибка при анализе исходного файла:', error.message);
  }
  
  console.log('\n🎯 АНАЛИЗ ФАЙЛА РЕЗУЛЬТАТА');
  console.log('=========================');
  
  try {
    const resultBuffer = await importFromGoogleSheets(resultUrl);
    const resultWorkbook = XLSX.read(resultBuffer, { type: 'buffer' });
    
    console.log('📊 ФАЙЛ РЕЗУЛЬТАТА - ЛИСТЫ:', resultWorkbook.SheetNames);
    
    // Анализируем первый лист
    const resultSheetName = resultWorkbook.SheetNames[0];
    const resultSheet = resultWorkbook.Sheets[resultSheetName];
    const resultData = XLSX.utils.sheet_to_json(resultSheet, { header: 1 });
    
    console.log('📋 ФАЙЛ РЕЗУЛЬТАТА - ЗАГОЛОВКИ:');
    console.log(resultData[0]);
    
    console.log('📋 ФАЙЛ РЕЗУЛЬТАТА - ВСЕГО СТРОК:', resultData.length);
    
    // Показываем первые 10 строк данных
    console.log('📋 ФАЙЛ РЕЗУЛЬТАТА - ПЕРВЫЕ 10 СТРОК:');
    for (let i = 0; i < Math.min(10, resultData.length); i++) {
      console.log(`Строка ${i + 1}:`, resultData[i]);
    }
    
    // Ищем статистику
    console.log('\n📊 ПОИСК СТАТИСТИКИ В ФАЙЛЕ РЕЗУЛЬТАТА:');
    
    for (let i = 0; i < resultData.length; i++) {
      const row = resultData[i];
      if (row && row.length > 0) {
        const firstCell = String(row[0]).toLowerCase();
        if (firstCell.includes('суммарное') || firstCell.includes('количество') || firstCell.includes('доля') || firstCell.includes('площадки')) {
          console.log(`Строка ${i + 1} (СТАТИСТИКА):`, row);
        }
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка при анализе файла результата:', error.message);
  }
}

// Запускаем анализ
analyzeGoogleSheets().catch(console.error); 