const XLSX = require('xlsx');
const fs = require('fs');

console.log('🔧 ТЕСТ ИСПРАВЛЕННОЙ ЛОГИКИ ПРОСМОТРОВ');
console.log('====================================');

// Читаем исходный файл
const buffer = fs.readFileSync('uploads/Fortedetrim ORM report source.xlsx');
const workbook = XLSX.read(buffer, { type: 'buffer' });
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

// Функция извлечения просмотров (исправленная)
function extractViews(row) {
  // ВАЖНО: Колонка G (6) содержит ДАТЫ в формате Excel serial number, НЕ просмотры!
  // Проверяем только правильные колонки для просмотров
  
  // Проверяем колонку K (10) - там чаще всего просмотры
  if (row[10] && typeof row[10] === 'number' && row[10] > 100 && row[10] < 1000000) {
    return Math.round(row[10]);
  }
  
  // Проверяем колонку L (11) 
  if (row[11] && typeof row[11] === 'number' && row[11] > 100 && row[11] < 1000000) {
    return Math.round(row[11]);
  }
  
  // Проверяем колонку M (12)
  if (row[12] && typeof row[12] === 'number' && row[12] > 100 && row[12] < 1000000) {
    return Math.round(row[12]);
  }
  
  // НЕ проверяем колонку G (6) - там даты!
  
  return 'Нет данных';
}

// Функция конверсии даты
function convertExcelDateToString(dateValue) {
  if (!dateValue) return '';
  
  // Если это Excel serial number (число больше 40000)
  if (typeof dateValue === 'number' && dateValue > 40000) {
    try {
      // Конвертируем Excel serial number в дату
      const date = new Date((dateValue - 25569) * 86400 * 1000);
      const day = date.getDate().toString().padStart(2, '0');
      const month = (date.getMonth() + 1).toString().padStart(2, '0');
      const year = date.getFullYear();
      return `${day}.${month}.${year}`;
    } catch (error) {
      console.warn('Ошибка конверсии даты:', dateValue, error);
      return dateValue.toString();
    }
  }
  
  return dateValue.toString();
}

let reviewCount = 0;
let commentCount = 0;
let totalViews = 0;
let commentsWithViews = 0;
let datesConverted = 0;

console.log('📊 Обработка данных с исправленной логикой...');

for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (!row || !row[0]) continue;
  
  const firstCell = row[0].toString().toLowerCase();
  
  if (firstCell.includes('отзыв')) {
    reviewCount++;
    
    // Для отзывов дата также в колонке G (6)
    if (row[6] && typeof row[6] === 'number' && row[6] > 40000) {
      datesConverted++;
      const dateStr = convertExcelDateToString(row[6]);
      if (datesConverted <= 5) {
        console.log(`📅 Отзыв - дата: ${row[6]} -> ${dateStr}`);
      }
    }
  }
  
  if (firstCell.includes('комментарии')) {
    commentCount++;
    
    // Дата в колонке G (6)
    if (row[6] && typeof row[6] === 'number' && row[6] > 40000) {
      datesConverted++;
      const dateStr = convertExcelDateToString(row[6]);
      if (datesConverted <= 10) {
        console.log(`📅 Комментарий - дата: ${row[6]} -> ${dateStr}`);
      }
    }
    
    // Просмотры НЕ в колонке G!
    const views = extractViews(row);
    if (typeof views === 'number') {
      totalViews += views;
      commentsWithViews++;
      if (commentsWithViews <= 5) {
        console.log(`👁️ Просмотры: ${views} (строка ${i + 1})`);
      }
    }
  }
}

console.log('\n📊 РЕЗУЛЬТАТЫ ИСПРАВЛЕННОЙ ОБРАБОТКИ:');
console.log('=====================================');
console.log('📝 Отзывов:', reviewCount);
console.log('💬 Комментариев:', commentCount);
console.log('📅 Дат обработано:', datesConverted);
console.log('👁️ Общие просмотры:', totalViews);
console.log('📈 Комментариев с просмотрами:', commentsWithViews);
console.log('📉 Комментариев без просмотров:', commentCount - commentsWithViews);

const engagementRate = commentCount > 0 ? Math.round((commentsWithViews / commentCount) * 100) : 0;
console.log('📊 Процент с данными о просмотрах:', engagementRate + '%');

console.log('\n✅ Тест завершен. Ожидаемые результаты:');
console.log('- Общие просмотры должны быть около 3398560');
console.log('- Количество дат должно быть равно количеству записей');
console.log('- Просмотры НЕ должны содержать даты'); 