const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function checkLatestResult() {
  console.log('🔍 Проверка последнего результата...');
  
  try {
    // Ищем самый новый файл результата
    const uploadsDir = 'uploads';
    const files = fs.readdirSync(uploadsDir);
    
    const resultFiles = files.filter(file => 
      file.includes('результат') && file.endsWith('.xlsx') && !file.includes('temp_google_sheets')
    );
    
    // Сортируем по времени создания
    const fileStats = resultFiles.map(file => ({
      name: file,
      path: path.join(uploadsDir, file),
      stats: fs.statSync(path.join(uploadsDir, file))
    })).sort((a, b) => b.stats.mtime - a.stats.mtime);
    
    if (fileStats.length === 0) {
      console.log('❌ Файлы результата не найдены');
      return;
    }
    
    const latestFile = fileStats[0];
    console.log(`📁 Проверяем: ${latestFile.name}`);
    console.log(`📊 Размер: ${latestFile.stats.size} байт`);
    console.log(`🕒 Создан: ${latestFile.stats.mtime.toISOString()}`);
    
    const workbook = XLSX.readFile(latestFile.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`\n📊 Лист: "${sheetName}"`);
    console.log(`📋 Всего строк: ${data.length}`);
    
    // Проверяем структуру результата vs исходника
    console.log('\n🎯 Анализ структуры:');
    
    if (data[0]) {
      const row1 = data[0];
      console.log('✅ Строка 1:', row1.slice(0, 4));
      
      // Проверяем, является ли это правильным результатом
      if (row1[0] === 'Продукт' && row1[1] === 'Акрихин - Фортедетрим') {
        console.log('🎉 ПРАВИЛЬНАЯ СТРУКТУРА РЕЗУЛЬТАТА!');
      } else if (row1[0] === 'Фортедетрим' || row1.includes('Фортедетрим')) {
        console.log('❌ НЕПРАВИЛЬНАЯ СТРУКТУРА - ВЫГЛЯДИТ КАК ИСХОДНИК!');
      }
    }
    
    if (data[1]) {
      console.log('✅ Строка 2:', data[1].slice(0, 4));
    }
    if (data[2]) {
      console.log('✅ Строка 3:', data[2].slice(0, 4));
    }
    if (data[3]) {
      console.log('✅ Строка 4:', data[3].slice(0, 4));
    }
    
    // Проверяем разделы
    console.log('\n🎯 Проверка разделов:');
    for (let i = 4; i < Math.min(data.length, 20); i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        if (cellValue.includes('Отзывы') || cellValue.includes('Комментарии') || cellValue.includes('Активные')) {
          console.log(`✅ Раздел ${i}: ${cellValue}`);
        }
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка при проверке:', error);
  }
}

checkLatestResult(); 