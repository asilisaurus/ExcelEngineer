const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function diagnoseMissingData() {
  console.log('🔍 ДИАГНОСТИКА ОТСУТСТВУЮЩИХ ДАННЫХ');
  console.log('==================================\n');
  
  try {
    // Находим самый новый файл результата
    const uploadsDir = 'uploads';
    const files = fs.readdirSync(uploadsDir);
    
    const resultFiles = files.filter(file => 
      file.includes('результат') && file.endsWith('.xlsx') && !file.includes('temp_google_sheets')
    );
    
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
    console.log(`📁 Анализируем: ${latestFile.name}\n`);
    
    const workbook = XLSX.readFile(latestFile.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`📊 Лист: "${sheetName}"`);
    console.log(`📋 Всего строк: ${data.length}\n`);
    
    // Анализируем каждую секцию детально
    let inReviewsSection = false;
    let inCommentsSection = false;
    let inDiscussionsSection = false;
    let reviewsStart = -1;
    let commentsStart = -1;
    let discussionsStart = -1;
    let reviewsEnd = -1;
    let commentsEnd = -1;
    let discussionsEnd = -1;
    
    // Находим границы секций
    for (let i = 0; i < data.length; i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        
        if (cellValue.includes('Отзывы') && !cellValue.includes('Количество')) {
          reviewsStart = i;
          inReviewsSection = true;
          inCommentsSection = false;
          inDiscussionsSection = false;
          console.log(`📍 Секция "Отзывы" начинается в строке ${i + 1}`);
        } else if (cellValue.includes('Комментарии') && !cellValue.includes('Количество')) {
          if (inReviewsSection) {
            reviewsEnd = i - 1;
            console.log(`📍 Секция "Отзывы" заканчивается в строке ${i}`);
          }
          commentsStart = i;
          inReviewsSection = false;
          inCommentsSection = true;
          inDiscussionsSection = false;
          console.log(`📍 Секция "Комментарии" начинается в строке ${i + 1}`);
        } else if (cellValue.includes('Активные обсуждения') || cellValue.includes('Обсуждения')) {
          if (inCommentsSection) {
            commentsEnd = i - 1;
            console.log(`📍 Секция "Комментарии" заканчивается в строке ${i}`);
          }
          discussionsStart = i;
          inReviewsSection = false;
          inCommentsSection = false;
          inDiscussionsSection = true;
          console.log(`📍 Секция "Активные обсуждения" начинается в строке ${i + 1}`);
        }
      }
    }
    
    // Устанавливаем конец последней секции
    if (inDiscussionsSection) {
      discussionsEnd = data.length - 1;
      console.log(`📍 Секция "Активные обсуждения" заканчивается в строке ${data.length}`);
    }
    
    console.log('\n📊 АНАЛИЗ СЕКЦИЙ:');
    
    // Анализируем секцию отзывов
    if (reviewsStart !== -1 && reviewsEnd !== -1) {
      const reviewsSection = data.slice(reviewsStart + 1, reviewsEnd + 1);
      const validReviews = reviewsSection.filter(row => 
        row && row[0] && row[0].toString().includes('.')
      );
      console.log(`📝 Отзывы: ${validReviews.length} записей (строки ${reviewsStart + 2}-${reviewsEnd + 1})`);
      
      // Показываем первые 3 отзыва
      console.log('   Первые 3 отзыва:');
      for (let i = 0; i < Math.min(3, validReviews.length); i++) {
        console.log(`   ${i + 1}. ${validReviews[i][0]}`);
      }
    }
    
    // Анализируем секцию комментариев
    if (commentsStart !== -1 && commentsEnd !== -1) {
      const commentsSection = data.slice(commentsStart + 1, commentsEnd + 1);
      const validComments = commentsSection.filter(row => 
        row && row[0] && row[0].toString().includes('.')
      );
      console.log(`💬 Комментарии: ${validComments.length} записей (строки ${commentsStart + 2}-${commentsEnd + 1})`);
      
      // Показываем первые 3 комментария
      console.log('   Первые 3 комментария:');
      for (let i = 0; i < Math.min(3, validComments.length); i++) {
        console.log(`   ${i + 1}. ${validComments[i][0]}`);
      }
      
      // Проверяем, есть ли пустые строки или заголовки в секции
      const emptyRows = commentsSection.filter(row => 
        !row || !row[0] || row[0].toString().trim() === ''
      );
      if (emptyRows.length > 0) {
        console.log(`   ⚠️ Найдено ${emptyRows.length} пустых строк в секции комментариев`);
      }
    }
    
    // Анализируем секцию обсуждений
    if (discussionsStart !== -1 && discussionsEnd !== -1) {
      const discussionsSection = data.slice(discussionsStart + 1, discussionsEnd + 1);
      const validDiscussions = discussionsSection.filter(row => 
        row && row[0] && row[0].toString().includes('.')
      );
      console.log(`🔥 Обсуждения: ${validDiscussions.length} записей (строки ${discussionsStart + 2}-${discussionsEnd + 1})`);
      
      // Показываем первые 3 обсуждения
      console.log('   Первые 3 обсуждения:');
      for (let i = 0; i < Math.min(3, validDiscussions.length); i++) {
        console.log(`   ${i + 1}. ${validDiscussions[i][0]}`);
      }
      
      // Проверяем, есть ли пустые строки или заголовки в секции
      const emptyRows = discussionsSection.filter(row => 
        !row || !row[0] || row[0].toString().trim() === ''
      );
      if (emptyRows.length > 0) {
        console.log(`   ⚠️ Найдено ${emptyRows.length} пустых строк в секции обсуждений`);
      }
      
      // Проверяем, есть ли строки с заголовками в секции
      const headerRows = discussionsSection.filter(row => 
        row && row[0] && (
          row[0].toString().includes('Площадка') ||
          row[0].toString().includes('Тема') ||
          row[0].toString().includes('Текст') ||
          row[0].toString().includes('Дата')
        )
      );
      if (headerRows.length > 0) {
        console.log(`   ⚠️ Найдено ${headerRows.length} строк с заголовками в секции обсуждений`);
      }
    }
    
    console.log('\n🔧 РЕКОМЕНДАЦИИ ДЛЯ АГЕНТА:');
    console.log('1. Проверить логику определения границ секций');
    console.log('2. Убедиться, что заголовки секций не включаются в подсчет данных');
    console.log('3. Проверить фильтрацию пустых строк');
    console.log('4. Убедиться, что все строки с данными содержат точки (домены)');
    console.log('5. Проверить, не обрезаются ли данные в конце файла');
    
    // Специфические рекомендации
    if (commentsStart !== -1 && commentsEnd !== -1) {
      const commentsSection = data.slice(commentsStart + 1, commentsEnd + 1);
      const validComments = commentsSection.filter(row => 
        row && row[0] && row[0].toString().includes('.')
      );
      if (validComments.length === 19) {
        console.log('\n💡 ДЛЯ КОММЕНТАРИЕВ:');
        console.log('- Найдено 19 комментариев, ожидается 20');
        console.log('- Возможно, один комментарий не прошел фильтрацию');
        console.log('- Проверить критерии валидации комментариев');
      }
    }
    
    if (discussionsStart !== -1 && discussionsEnd !== -1) {
      const discussionsSection = data.slice(discussionsStart + 1, discussionsEnd + 1);
      const validDiscussions = discussionsSection.filter(row => 
        row && row[0] && row[0].toString().includes('.')
      );
      if (validDiscussions.length === 514) {
        console.log('\n💡 ДЛЯ ОБСУЖДЕНИЙ:');
        console.log('- Найдено 514 обсуждений, ожидается 621');
        console.log('- Не хватает 107 обсуждений');
        console.log('- Возможно, проблема с определением границ секции');
        console.log('- Проверить, не обрезаются ли данные в конце файла');
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
  }
}

diagnoseMissingData();