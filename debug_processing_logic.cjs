const ExcelJS = require('exceljs');

// Копируем логику из процессора
function analyzeRowType(row) {
  if (!row || row.length === 0) return 'empty';
  
  const colA = (row[0] || '').toString().toLowerCase();
  const colB = (row[1] || '').toString().toLowerCase();
  const colD = (row[3] || '').toString().toLowerCase();
  const colE = (row[4] || '').toString().toLowerCase();
  const colN = (row[13] || '').toString().toLowerCase();
  
  // Заголовки и служебные строки
  if (colA.includes('тип размещения') || colA.includes('площадка') || 
      colB.includes('площадка') || colE.includes('текст сообщения') ||
      colA.includes('план') || colA.includes('итого') || colA.includes('ключевые сообщения') ||
      colA.includes('топ-20 выдачи') || colA === 'о т з ы в ы') {
    return 'header';
  }
  
  // Google Sheets специфичная логика: проверяем колонку A для типа
  if (colA.includes('отзывы (отзовики)')) {
    return 'review_otzovik';
  }
  
  if (colA.includes('отзывы (аптеки)')) {
    return 'review_pharmacy';
  }
  
  if (colA.includes('комментарии в обсуждениях')) {
    return 'comment';
  }
  
  // Секционные заголовки из исходника
  if (colA.includes('отзывы') || colA.includes('комментарии') || colA.includes('топ-20 выдачи')) {
    return 'section_header';
  }
  
  // Анализ по URL и платформам
  const urlText = colB + ' ' + colD;
  
  // Проверяем отзывы на платформах отзовиков
  const reviewPlatforms = [
    'otzovik', 'irecommend', 'otzyvru', 'pravogolosa', 'medum', 
    'vseotzyvy', 'otzyvy.pro'
  ];
  
  // Проверяем отзывы на аптечных платформах
  const pharmacyPlatforms = [
    'market.yandex', 'dialog.ru', 'goodapteka', 'megapteka', 
    'uteka', 'nfapteka', 'piluli.ru', 'eapteka.ru', 'pharmspravka.ru', 
    'gde.ru', 'ozon.ru'
  ];
  
  // Проверяем комментарии
  const commentPlatforms = [
    'dzen.ru', 'woman.ru', 'forum.baby.ru', 'vk.com', 't.me',
    'ok.ru', 'otvet.mail.ru', 'babyblog.ru', 'mom.life', 
    'youtube.com', 'pikabu.ru', 'livejournal.com', 'facebook.com'
  ];
  
  const isReviewPlatform = reviewPlatforms.some(platform => 
    urlText.toLowerCase().includes(platform)
  );
  
  const isPharmacyPlatform = pharmacyPlatforms.some(platform => 
    urlText.toLowerCase().includes(platform)
  );
  
  const isCommentPlatform = commentPlatforms.some(platform => 
    urlText.toLowerCase().includes(platform)
  );
  
  // Анализ типа поста в колонке N
  const postType = colN;
  
  // Определяем тип по платформе и типу поста
  if ((colB || colD || colE) && (isReviewPlatform || (postType === 'ос' && isReviewPlatform))) {
    return 'review_otzovik';
  }
  
  if ((colB || colD || colE) && (isPharmacyPlatform || (postType === 'ос' && isPharmacyPlatform))) {
    return 'review_pharmacy';
  }
  
  if ((colB || colD || colE) && (isCommentPlatform || postType === 'цс')) {
    return 'comment';
  }
  
  // Если есть контент, но тип неясен
  if (colB || colD || colE) {
    return 'content';
  }
  
  return 'empty';
}

async function debugProcessingLogic() {
  try {
    console.log('=== ДИАГНОСТИКА ЛОГИКИ ОБРАБОТКИ ===');
    
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
    const typeCount = {};
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const type = analyzeRowType(row);
      
      if (!typeCount[type]) {
        typeCount[type] = 0;
      }
      typeCount[type]++;
    }
    
    console.log('\\n📊 Результаты анализа по типам:');
    Object.entries(typeCount).forEach(([type, count]) => {
      console.log(`   🔹 ${type}: ${count}`);
    });
    
    // Фильтруем только нужные типы
    const relevantTypes = ['review_otzovik', 'review_pharmacy', 'comment'];
    const processedRows = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const type = analyzeRowType(row);
      
      if (relevantTypes.includes(type)) {
        processedRows.push({ type, row, index: i + 2 }); // +2 для учета заголовка
      }
    }
    
    console.log('\\n🎯 Записи для обработки:');
    console.log(`   📝 review_otzovik: ${processedRows.filter(r => r.type === 'review_otzovik').length}`);
    console.log(`   📝 review_pharmacy: ${processedRows.filter(r => r.type === 'review_pharmacy').length}`);
    console.log(`   📝 comment: ${processedRows.filter(r => r.type === 'comment').length}`);
    console.log(`   📝 Общее количество: ${processedRows.length}`);
    
    // Показываем первые 10 комментариев
    const comments = processedRows.filter(r => r.type === 'comment').slice(0, 10);
    console.log('\\n🔍 Первые 10 комментариев:');
    comments.forEach((item, index) => {
      const text = (item.row[4] || '').toString();
      console.log(`   ${index + 1}. Строка ${item.index}: ${text.substring(0, 100)}...`);
    });
    
  } catch (error) {
    console.error('❌ Ошибка:', error);
  }
}

debugProcessingLogic(); 