const ExcelJS = require('exceljs');

async function debugCommentsCount() {
  try {
    console.log('=== ПОДСЧЕТ КОММЕНТАРИЕВ В ИСХОДНОМ ФАЙЛЕ ===');
    
    const filePath = 'uploads/temp_google_sheets_1751807589999.xlsx';
    
    // Читаем файл
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];
    
    // Преобразуем в массив
    const data = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Пропускаем заголовок
        const rowData = [];
        row.eachCell((cell, colNumber) => {
          rowData[colNumber - 1] = cell.value;
        });
        data.push(rowData);
      }
    });
    
    console.log(`📋 Всего строк данных: ${data.length}`);
    
    // Анализируем каждую строку
    let reviewsOTZ = 0;
    let reviewsAPT = 0;
    let comments = 0;
    let empty = 0;
    let headers = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const colA = (row[0] || '').toString().toLowerCase();
      const colB = (row[1] || '').toString().toLowerCase();
      const colD = (row[3] || '').toString().toLowerCase();
      const colE = (row[4] || '').toString().toLowerCase();
      
      // Определяем тип строки
      if (!row || row.length === 0 || (!colA && !colB && !colD && !colE)) {
        empty++;
        continue;
      }
      
      // Заголовки (только точные совпадения в колонке A)
      if (colA === 'площадка' || colA === 'план' || colA === 'итого' || 
          colA === 'отзывы' || colA === 'комментарии топ-20 выдачи' || 
          colA === 'активные обсуждения (мониторинг)' || colA === 'продукт' || 
          colA === 'период' || colE === 'текст сообщения') {
        headers++;
        continue;
      }
      
      // Проверяем URL и платформы
      const urlText = colB + ' ' + colD;
      const isReviewPlatform = /otzovik|irecommend|otzyvru|pravogolosa|medum|vseotzyvy|otzyvy\.pro/i.test(urlText);
      const isPharmacyPlatform = /market\.yandex|dialog\.ru|goodapteka|megapteka|uteka|nfapteka|piluli\.ru|eapteka\.ru|pharmspravka\.ru|gde\.ru|ozon\.ru/i.test(urlText);
      const isCommentPlatform = /dzen\.ru|woman\.ru|forum\.baby\.ru|vk\.com|t\.me|ok\.ru|otvet\.mail\.ru|babyblog\.ru|mom\.life|youtube\.com|pikabu\.ru|livejournal\.com|facebook\.com/i.test(urlText);
      
      if (isReviewPlatform) {
        reviewsOTZ++;
      } else if (isPharmacyPlatform) {
        reviewsAPT++;
      } else if (isCommentPlatform || colE.length > 10) {
        comments++;
      }
    }
    
    console.log(`📊 Результаты анализа:`);
    console.log(`   🔹 Отзывы OTZ: ${reviewsOTZ}`);
    console.log(`   🔹 Отзывы APT: ${reviewsAPT}`);
    console.log(`   🔹 Всего отзывов: ${reviewsOTZ + reviewsAPT}`);
    console.log(`   🔹 Комментарии: ${comments}`);
    console.log(`   🔹 Пустые строки: ${empty}`);
    console.log(`   🔹 Заголовки: ${headers}`);
    console.log(`   🔹 Всего: ${reviewsOTZ + reviewsAPT + comments + empty + headers}`);
    
    console.log(`\\n🎯 Ожидаемые результаты:`);
    console.log(`   🔹 Отзывы: ${reviewsOTZ + reviewsAPT} (ожидается 22)`);
    console.log(`   🔹 Комментарии ТОП-20: 20 (первые 20 из ${comments})`);
    console.log(`   🔹 Активные обсуждения: ${comments - 20} (остальные из ${comments})`);
    console.log(`   🔹 Общее количество: ${reviewsOTZ + reviewsAPT + comments}`);
    
  } catch (error) {
    console.error('❌ Ошибка:', error);
  }
}

debugCommentsCount(); 