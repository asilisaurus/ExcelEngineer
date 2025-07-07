const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function verifyFormatting() {
  console.log('🔍 Проверка форматирования результата...');
  
  try {
    // Проверяем самый последний файл
    const latestFile = 'uploads/file-1751801349824-177095888_Март_2025_результат_20250706.xlsx';
    
    if (!fs.existsSync(latestFile)) {
      console.log('❌ Файл не найден:', latestFile);
      return;
    }
    
    const workbook = XLSX.readFile(latestFile);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`📊 Лист: ${sheetName}`);
    console.log(`📋 Всего строк: ${data.length}`);
    
    // Проверяем структуру заголовков
    console.log('\n🎯 Проверка структуры заголовков:');
    
    if (data[0]) {
      console.log('✅ Строка 1:', data[0].slice(0, 8));
    }
    if (data[1]) {
      console.log('✅ Строка 2:', data[1].slice(0, 8));
    }
    if (data[2]) {
      console.log('✅ Строка 3:', data[2].slice(0, 8));
    }
    if (data[3]) {
      console.log('✅ Строка 4:', data[3].slice(0, 8));
    }
    
    // Проверяем разделы
    console.log('\n🎯 Проверка разделов:');
    let sectionCount = 0;
    
    for (let i = 4; i < Math.min(data.length, 50); i++) {
      if (data[i] && data[i][0] && (
        data[i][0].includes('Отзывы') || 
        data[i][0].includes('Комментарии') ||
        data[i][0].includes('Активные')
      )) {
        console.log(`✅ Раздел ${++sectionCount}: ${data[i][0]}`);
      }
    }
    
    // Проверяем итоговую статистику
    console.log('\n🎯 Проверка итоговой статистики:');
    const lastRows = data.slice(-10);
    lastRows.forEach((row, index) => {
      if (row && row[0] && (
        row[0].includes('Суммарное') || 
        row[0].includes('Количество') ||
        row[0].includes('Доля')
      )) {
        console.log(`✅ Статистика: ${row[0]} = ${row[5] || 'N/A'}`);
      }
    });
    
    console.log('\n🎉 Проверка завершена!');
    console.log(`📁 Файл: ${latestFile}`);
    
  } catch (error) {
    console.error('❌ Ошибка при проверке:', error);
  }
}

verifyFormatting(); 