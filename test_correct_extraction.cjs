const XLSX = require('xlsx');

async function testCorrectExtraction() {
  console.log('🔍 ТЕСТ ПРАВИЛЬНОГО ИЗВЛЕЧЕНИЯ ДАННЫХ');
  console.log('=====================================');
  
  // Используем последний загруженный файл
  const testFile = 'uploads/Фортедетрим_ORM_отчет_исходник_1751040742705.xlsx';
  
  try {
    const workbook = XLSX.readFile(testFile);
    console.log('📋 Листы в файле:', workbook.SheetNames);
    
    // Найдем лист с данными месяца
    const months = ["Янв25", "Фев25", "Мар25", "Март25", "Апр25", "Май25", "Июн25", 
                   "Июл25", "Авг25", "Сен25", "Окт25", "Ноя25", "Дек25"];
    
    const sheetName = workbook.SheetNames.find(name => 
      months.some(month => name.includes(month))
    );
    
    if (!sheetName) {
      console.error('❌ Лист с данными месяца не найден');
      return;
    }
    
    console.log('📊 Найден лист:', sheetName);
    
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    console.log(`📊 Общее количество строк в листе: ${data.length}`);
    
    // Извлекаем данные по правильным диапазонам
    // Колонки: [1, 3, 4, 6, 7, 10, 16, 13] = B, D, E, G, H, K, Q, N (индексы 1, 3, 4, 6, 7, 10, 16, 13)
    
    console.log('\n🔸 ОТЗЫВЫ OTZ (строки 6-15):');
    const reviewsOtz = [];
    for (let i = 6; i < 15; i++) {
      if (data[i]) {
        const row = [
          data[i][1], // B - Площадка
          data[i][3], // D - Тема
          data[i][4], // E - Текст
          data[i][6], // G - Дата
          data[i][7], // H - Ник
          data[i][10], // K - Просмотры
          data[i][16], // Q - Вовлечение
          data[i][13]  // N - Тип поста
        ];
        reviewsOtz.push(row);
        console.log(`Строка ${i + 1}:`, row);
      }
    }
    
    console.log('\n🔸 ОТЗЫВЫ APT (строки 15-28):');
    const reviewsApt = [];
    for (let i = 15; i < 28; i++) {
      if (data[i]) {
        const row = [
          data[i][1], // B - Площадка
          data[i][3], // D - Тема
          data[i][4], // E - Текст
          data[i][6], // G - Дата
          data[i][7], // H - Ник
          data[i][10], // K - Просмотры
          data[i][16], // Q - Вовлечение
          data[i][13]  // N - Тип поста
        ];
        reviewsApt.push(row);
        console.log(`Строка ${i + 1}:`, row);
      }
    }
    
    console.log('\n🔸 ТОП-20 (строки 31-51):');
    const top20 = [];
    for (let i = 31; i < 51; i++) {
      if (data[i]) {
        const row = [
          data[i][1], // B - Площадка
          data[i][3], // D - Тема
          data[i][4], // E - Текст
          data[i][6], // G - Дата
          data[i][7], // H - Ник
          data[i][10], // K - Просмотры
          data[i][16], // Q - Вовлечение
          data[i][13]  // N - Тип поста
        ];
        top20.push(row);
        console.log(`Строка ${i + 1}:`, row);
      }
    }
    
    console.log('\n📊 СТАТИСТИКА:');
    console.log('===============');
    
    // Подсчитываем отзывы (только непустые строки)
    const validReviewsOtz = reviewsOtz.filter(row => row[0] || row[1] || row[2]);
    const validReviewsApt = reviewsApt.filter(row => row[0] || row[1] || row[2]);
    const allReviews = validReviewsOtz.length + validReviewsApt.length;
    
    console.log(`📝 Отзывы OTZ: ${validReviewsOtz.length}`);
    console.log(`📝 Отзывы APT: ${validReviewsApt.length}`);
    console.log(`📝 Всего отзывов: ${allReviews}`);
    
    // Подсчитываем комментарии (только непустые строки)
    const validTop20 = top20.filter(row => row[0] || row[1] || row[2]);
    const activeDiscussions = 0; // По логике Python - пустой DataFrame
    const allComments = validTop20.length + activeDiscussions;
    
    console.log(`💬 Комментарии Топ-20: ${validTop20.length}`);
    console.log(`💬 Активные обсуждения: ${activeDiscussions}`);
    console.log(`💬 Всего комментариев: ${allComments}`);
    
    // Подсчитываем просмотры (только валидные числа)
    function cleanViews(value) {
      if (!value || value === 'Нет данных' || value === '') return 0;
      const cleaned = String(value).replace(/[^0-9.]/g, '');
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    }
    
    let totalViews = 0;
    const allData = [...validReviewsOtz, ...validReviewsApt, ...validTop20];
    
    console.log('\n🔍 АНАЛИЗ ПРОСМОТРОВ:');
    allData.forEach((row, idx) => {
      const views = cleanViews(row[5]); // Индекс 5 = колонка K (просмотры)
      if (views > 0) {
        totalViews += views;
        console.log(`${idx + 1}. Просмотры: ${row[5]} → ${views}`);
      }
    });
    
    console.log(`👀 Общие просмотры: ${totalViews.toLocaleString()}`);
    
    // Подсчитываем вовлечение
    const discussionsData = [...validTop20]; // Только топ-20, active пустой
    const engagementCount = discussionsData.filter(row => 
      String(row[6]).toLowerCase().includes('есть')
    ).length;
    
    const engagementPct = discussionsData.length > 0 ? 
      Math.round((engagementCount / discussionsData.length) * 100) : 0;
    
    console.log(`💬 Обсуждения с вовлечением: ${engagementCount} из ${discussionsData.length}`);
    console.log(`📊 Доля вовлечения: ${engagementPct}%`);
    
    console.log('\n🎯 ФИНАЛЬНЫЕ РЕЗУЛЬТАТЫ:');
    console.log('========================');
    console.log(`Отзывы: ${allReviews}`);
    console.log(`Комментарии: ${allComments}`);
    console.log(`Просмотры: ${totalViews.toLocaleString()}`);
    console.log(`Вовлечение: ${engagementPct}%`);
    
    console.log('\n🔍 СРАВНЕНИЕ С ОЖИДАЕМЫМИ:');
    console.log('==========================');
    console.log(`Отзывы: ${allReviews} (ожидается: 18) ${allReviews === 18 ? '✅' : '❌'}`);
    console.log(`Комментарии: ${allComments} (ожидается: 519) ${allComments === 519 ? '✅' : '❌'}`);
    console.log(`Просмотры: ${totalViews.toLocaleString()} (ожидается: 3,398,560) ${totalViews === 3398560 ? '✅' : '❌'}`);
    console.log(`Вовлечение: ${engagementPct}% (ожидается: 20%) ${engagementPct === 20 ? '✅' : '❌'}`);
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
  }
}

testCorrectExtraction().catch(console.error); 