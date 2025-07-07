const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function simpleCheck() {
  console.log('📋 Простая проверка последнего файла...');
  
  try {
    // Находим все файлы результата
    const uploadsDir = 'uploads';
    const files = fs.readdirSync(uploadsDir);
    const resultFiles = files.filter(file => file.includes('результат'));
    
    if (resultFiles.length === 0) {
      console.log('❌ Файлы результата не найдены');
      return;
    }
    
    // Сортируем по времени создания
    const filesWithStats = resultFiles.map(file => ({
      name: file,
      stats: fs.statSync(path.join(uploadsDir, file))
    }));
    
    filesWithStats.sort((a, b) => b.stats.mtime - a.stats.mtime);
    const latestFile = filesWithStats[0].name;
    
    console.log(`📁 Проверяем файл: ${latestFile}`);
    console.log(`📊 Размер: ${filesWithStats[0].stats.size} байт`);
    console.log(`🕒 Изменен: ${filesWithStats[0].stats.mtime}`);
    
    // Загружаем файл
    const filePath = path.join(uploadsDir, latestFile);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    
    console.log('\n📋 Первые 10 строк:');
    
    for (let i = 1; i <= 10; i++) {
      const row = worksheet.getRow(i);
      const values = [];
      
      for (let j = 1; j <= 8; j++) {
        const cell = row.getCell(j);
        values.push(cell.value || '');
      }
      
      console.log(`${i}: ${values.join(' | ')}`);
    }
    
    // Подсчитываем общее количество строк
    let totalRows = 0;
    worksheet.eachRow((row) => {
      totalRows++;
    });
    
    console.log(`\n📊 Общее количество строк: ${totalRows}`);
    
    // Проверяем цвета
    const row1 = worksheet.getRow(1);
    const headerBgColor = row1.getCell(1).fill?.fgColor?.argb;
    console.log(`🎨 Цвет заголовка: ${headerBgColor}`);
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
  }
}

simpleCheck(); 