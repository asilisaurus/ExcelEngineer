const XLSX = require('xlsx');
const fs = require('fs');

console.log('🔍 ПРЯМОЕ ТЕСТИРОВАНИЕ ИЗВЛЕЧЕНИЯ ДАННЫХ');
console.log('========================================');

try {
  // Читаем исходный файл
  const buffer = fs.readFileSync('uploads/Fortedetrim ORM report source.xlsx');
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  
  console.log('📋 Анализ данных:');
  console.log('Общее количество строк:', jsonData.length);
  
  // Тестируем новую логику извлечения
  const allData = [];
  let reviewsCount = 0;
  let commentsCount = 0;
  
  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (!row || !Array.isArray(row)) continue;

    const типРазмещения = row[0] ? row[0].toString().trim() : '';
    
    if (типРазмещения.toLowerCase().includes('отзыв')) {
      // Структура для отзывов
      const площадка = row[1] ? row[1].toString().trim() : '';
      const продукт = row[2] ? row[2].toString().trim() : '';
      const текст = row[4] ? row[4].toString().trim() : '';
      const дата = row[6] || '';
      const ник = row[7] ? row[7].toString().trim() : '';
      
      // Для отзывов просмотры в колонке 6 (дата как число Excel)
      const просмотры = typeof row[6] === 'number' ? row[6] : 'Нет данных';
      
      if (площадка || текст) {
        allData.push({
          площадка,
          тема: продукт,
          текст: текст.substring(0, 100) + '...',
          дата,
          ник,
          просмотры,
          вовлечение: 'Нет данных',
          типПоста: 'Отзывы'
        });
        reviewsCount++;
      }
    } else if (типРазмещения.toLowerCase().includes('комментарии')) {
      // Структура для комментариев
      const площадка = row[1] ? row[1].toString().trim() : '';
      const продукт = row[2] ? row[2].toString().trim() : '';
      const текст = row[4] ? row[4].toString().trim() : '';
      const дата = row[6] || '';
      const ник = row[7] ? row[7].toString().trim() : '';
      const просмотровПолучено = typeof row[11] === 'number' ? row[11] : 'Нет данных';
      const вовлечение = row[12] ? row[12].toString().trim() : 'Нет данных';
      
      if (площадка || текст) {
        allData.push({
          площадка,
          тема: продукт,
          текст: текст.substring(0, 100) + '...',
          дата,
          ник,
          просмотры: просмотровПолучено,
          вовлечение,
          типПоста: 'Комментарии Топ-20 выдачи'
        });
        commentsCount++;
      }
    }
  }
  
  console.log('📊 РЕЗУЛЬТАТЫ:');
  console.log('Всего извлечено строк:', allData.length);
  console.log('Отзывы:', reviewsCount);
  console.log('Комментарии:', commentsCount);
  
  // Показываем первые 5 записей каждого типа
  console.log('\n📝 ПРИМЕРЫ ОТЗЫВОВ:');
  const reviews = allData.filter(item => item.типПоста === 'Отзывы');
  reviews.slice(0, 5).forEach((item, index) => {
    console.log(`${index + 1}. ${item.площадка} | Просмотры: ${item.просмотры} | Ник: ${item.ник}`);
  });
  
  console.log('\n💬 ПРИМЕРЫ КОММЕНТАРИЕВ:');
  const comments = allData.filter(item => item.типПоста === 'Комментарии Топ-20 выдачи');
  comments.slice(0, 5).forEach((item, index) => {
    console.log(`${index + 1}. ${item.площадка} | Просмотры: ${item.просмотры} | Вовлечение: ${item.вовлечение}`);
  });
  
  // Анализируем вовлечение
  console.log('\n🔍 АНАЛИЗ ВОВЛЕЧЕНИЯ:');
  const commentsWithEngagement = comments.filter(item => 
    item.вовлечение && 
    item.вовлечение !== 'Нет данных' && 
    item.вовлечение.trim() !== ''
  );
  console.log('Комментарии с вовлечением:', commentsWithEngagement.length);
  
  // Показываем примеры вовлечения
  console.log('\n📈 ПРИМЕРЫ ВОВЛЕЧЕНИЯ:');
  commentsWithEngagement.slice(0, 10).forEach((item, index) => {
    console.log(`${index + 1}. Вовлечение: "${item.вовлечение}" | Площадка: ${item.площадка.substring(0, 50)}`);
  });
  
  // Считаем суммы
  const totalViews = allData.reduce((sum, item) => {
    return sum + (typeof item.просмотры === 'number' ? item.просмотры : 0);
  }, 0);
  
  console.log('\n📈 СТАТИСТИКА:');
  console.log('Общее количество просмотров:', totalViews);
  
  // Проверяем структуру одной строки комментария
  console.log('\n🔍 СТРУКТУРА СТРОКИ КОММЕНТАРИЯ:');
  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (row && row[0] && row[0].toString().toLowerCase().includes('комментарии')) {
      console.log(`Строка ${i + 1}:`);
      for (let j = 0; j < Math.min(15, row.length); j++) {
        console.log(`  Колонка ${j}: "${row[j]}"`);
      }
      break;
    }
  }
  
} catch (error) {
  console.error('❌ Ошибка:', error.message);
} 