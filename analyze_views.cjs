const XLSX = require('xlsx');
const fs = require('fs');

console.log('🔍 ДЕТАЛЬНЫЙ АНАЛИЗ ПРОСМОТРОВ');
console.log('==============================');

// Читаем исходный файл
const buffer = fs.readFileSync('uploads/Fortedetrim ORM report source.xlsx');
const workbook = XLSX.read(buffer, { type: 'buffer' });
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

console.log('📊 Анализ просмотров в комментариях...');

let totalViews = 0;
let viewsFound = 0;
let viewsData = [];

// Анализируем только комментарии
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (!row || !row[0]) continue;
  
  const firstCell = row[0].toString().toLowerCase();
  
  if (firstCell.includes('комментарии')) {
    // Проверяем все колонки на наличие просмотров
    let rowViews = 0;
    let viewsColumn = '';
    
    for (let col = 6; col < 20; col++) {
      if (row[col] && typeof row[col] === 'number' && row[col] > 100 && row[col] < 1000000) {
        // Это может быть количество просмотров
        if (row[col] > rowViews) {
          rowViews = row[col];
          viewsColumn = `Колонка ${String.fromCharCode(65 + col)} (${col})`;
        }
      }
    }
    
    if (rowViews > 0) {
      totalViews += rowViews;
      viewsFound++;
      viewsData.push({
        строка: i + 1,
        площадка: row[1] || 'Не указана',
        просмотры: rowViews,
        колонка: viewsColumn,
        ник: row[7] || 'Не указан'
      });
    }
  }
}

console.log(`\n📊 ИТОГИ АНАЛИЗА ПРОСМОТРОВ:`);
console.log('============================');
console.log('👁️ Общие просмотры:', totalViews);
console.log('📈 Записей с просмотрами:', viewsFound);
console.log('📉 Записей без просмотров:', 649 - viewsFound);

console.log('\n🔍 ПЕРВЫЕ 20 ЗАПИСЕЙ С ПРОСМОТРАМИ:');
console.log('===================================');
viewsData.slice(0, 20).forEach((item, index) => {
  console.log(`${index + 1}. Строка ${item.строка}: ${item.просмотры} просмотров (${item.колонка})`);
  console.log(`   Площадка: ${item.площадка.substring(0, 50)}...`);
  console.log(`   Ник: ${item.ник}`);
  console.log('');
});

console.log('\n📋 АНАЛИЗ КОЛОНОК:');
console.log('==================');
const columnStats = {};
viewsData.forEach(item => {
  if (!columnStats[item.колонка]) {
    columnStats[item.колонка] = 0;
  }
  columnStats[item.колонка]++;
});

Object.entries(columnStats).forEach(([column, count]) => {
  console.log(`${column}: ${count} записей`);
});

console.log('\n✅ Анализ завершен. Нужно использовать правильные колонки для просмотров.'); 