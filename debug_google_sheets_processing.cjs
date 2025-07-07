const ExcelJS = require('exceljs');
const path = require('path');

async function debugGoogleSheetsProcessing() {
  try {
    console.log('=== ДИАГНОСТИКА ОБРАБОТКИ GOOGLE SHEETS ===');
    
    const filePath = 'uploads/temp_google_sheets_1751807589999.xlsx';
    
    // Читаем файл с помощью ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    console.log(`📋 Найдено листов: ${workbook.worksheets.length}`);
    workbook.worksheets.forEach((ws, index) => {
      console.log(`  Лист ${index + 1}: "${ws.name}" (${ws.actualRowCount} строк)`);
    });
    
    const worksheet = workbook.worksheets[0];
    console.log(`\n🔍 Анализируем первый лист: "${worksheet.name}"`);
    
    // Преобразуем данные в массив
    const data = [];
    let totalRows = 0;
    let emptyRows = 0;
    let contentRows = 0;
    
    worksheet.eachRow((row, rowNumber) => {
      totalRows++;
      const rowData = [];
      let hasContent = false;
      
      row.eachCell((cell, colNumber) => {
        let value = cell.value;
        
        if (value && typeof value === 'object') {
          if (value instanceof Date) {
            value = value.toISOString().split('T')[0];
          } else if (value.text) {
            value = value.text;
          } else if (value.result) {
            value = value.result;
          } else if (value.toString) {
            value = value.toString();
          }
        }
        
        if (value && value.toString().trim()) {
          hasContent = true;
        }
        
        rowData[colNumber - 1] = value;
      });
      
      if (hasContent) {
        contentRows++;
      } else {
        emptyRows++;
      }
      
      data[rowNumber - 1] = rowData;
      
      // Выводим первые 10 строк для анализа
      if (rowNumber <= 10) {
        console.log(`Строка ${rowNumber}: ${hasContent ? '✅' : '❌'} | ${JSON.stringify(rowData.slice(0, 5))}...`);
      }
    });
    
    console.log(`\n📊 СТАТИСТИКА ЧТЕНИЯ:`);
    console.log(`Всего строк: ${totalRows}`);
    console.log(`Строк с контентом: ${contentRows}`);
    console.log(`Пустых строк: ${emptyRows}`);
    
    // Анализируем типы данных
    let reviewOtzovikCount = 0;
    let reviewPharmacyCount = 0;
    let commentCount = 0;
    let otherCount = 0;
    let headerCount = 0;
    
    console.log(`\n🔍 АНАЛИЗ ТИПОВ СТРОК:`);
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const type = analyzeRowType(row);
      
      switch (type) {
        case 'review_otzovik':
          reviewOtzovikCount++;
          break;
        case 'review_pharmacy':
          reviewPharmacyCount++;
          break;
        case 'comment':
          commentCount++;
          break;
        case 'header':
        case 'section_header':
          headerCount++;
          break;
        default:
          otherCount++;
          break;
      }
      
      // Показываем примеры первых найденных записей каждого типа
      if ((type === 'review_otzovik' && reviewOtzovikCount <= 3) ||
          (type === 'review_pharmacy' && reviewPharmacyCount <= 3) ||
          (type === 'comment' && commentCount <= 3)) {
        console.log(`Строка ${i + 1} [${type}]: ${(row[0] || '').toString().substring(0, 50)}...`);
      }
    }
    
    console.log(`\n📊 РЕЗУЛЬТАТЫ АНАЛИЗА:`);
    console.log(`Отзывы-отзовики: ${reviewOtzovikCount}`);
    console.log(`Отзывы-аптеки: ${reviewPharmacyCount}`);
    console.log(`Комментарии: ${commentCount}`);
    console.log(`Заголовки: ${headerCount}`);
    console.log(`Прочие: ${otherCount}`);
    console.log(`ИТОГО ОБРАБОТАННЫХ: ${reviewOtzovikCount + reviewPharmacyCount + commentCount}`);
    
    // Анализируем первые 60 строк более подробно
    console.log(`\n🔍 ПОДРОБНЫЙ АНАЛИЗ ПЕРВЫХ 60 СТРОК:`);
    for (let i = 0; i < Math.min(60, data.length); i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const colA = (row[0] || '').toString();
      const colB = (row[1] || '').toString();
      const colE = (row[4] || '').toString();
      const type = analyzeRowType(row);
      
      if (colA || colB || colE) {
                 console.log(`${(i + 1).toString().padStart(2)}: [${type.padEnd(15)}] A:"${colA.substring(0, 30)}" | E:"${colE.substring(0, 40)}"`);
       }
     }
     
   } catch (error) {
     console.error('❌ Ошибка:', error);
   }
}

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
      colA.includes('план') || colA.includes('итого')) {
    return 'header';
  }
  
  // Секционные заголовки
  if ((colA === 'отзывы' || colA === 'комментарии') && !colB && !colD && !colE) {
    return 'section_header';
  }
  
  // Платформы отзывов (отзовики)
  const reviewPlatforms = ['otzovik', 'irecommend', 'otzyvru', 'pravogolosa', 'medum', 'vseotzyvy', 'otzyvy.pro'];
  
  // Аптечные платформы
  const pharmacyPlatforms = ['market.yandex', 'dialog.ru', 'goodapteka', 'megapteka', 'uteka', 'nfapteka', 'piluli.ru', 'eapteka.ru', 'pharmspravka.ru', 'gde.ru', 'ozon.ru'];
  
  // Платформы комментариев
  const commentPlatforms = ['dzen.ru', 'woman.ru', 'forum.baby.ru', 'vk.com', 't.me', 'ok.ru', 'otvet.mail.ru', 'babyblog.ru', 'mom.life', 'youtube.com', 'pikabu.ru', 'livejournal.com', 'facebook.com'];
  
  // Анализ по URL и платформам
  const urlText = colB + ' ' + colD;
  const isReviewPlatform = reviewPlatforms.some(platform => urlText.toLowerCase().includes(platform));
  const isPharmacyPlatform = pharmacyPlatforms.some(platform => urlText.toLowerCase().includes(platform));
  const isCommentPlatform = commentPlatforms.some(platform => urlText.toLowerCase().includes(platform));
  
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

// Запускаем диагностику
debugGoogleSheetsProcessing(); 