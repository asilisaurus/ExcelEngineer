const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function compareWithReference() {
  console.log('🔍 СРАВНЕНИЕ С ЭТАЛОННЫМИ ДАННЫМИ');
  console.log('================================\n');
  
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
    
    // Эталонные данные для Марта 2025
    const referenceData = {
      reviews: 22,
      comments: 20,
      discussions: 621,
      total: 663
    };
    
    console.log('📋 ЭТАЛОННЫЕ ДАННЫЕ (Март 2025):');
    console.log(`- Отзывы: ${referenceData.reviews}`);
    console.log(`- Комментарии: ${referenceData.comments}`);
    console.log(`- Активные обсуждения: ${referenceData.discussions}`);
    console.log(`- Всего: ${referenceData.total}\n`);
    
    // Подсчитываем фактические данные
    let reviewsCount = 0;
    let commentsCount = 0;
    let discussionsCount = 0;
    let inReviewsSection = false;
    let inCommentsSection = false;
    let inDiscussionsSection = false;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        
        // Определяем секции
        if (cellValue.includes('Отзывы') && !cellValue.includes('Количество')) {
          inReviewsSection = true;
          inCommentsSection = false;
          inDiscussionsSection = false;
          console.log(`📍 Найдена секция "Отзывы" в строке ${i + 1}`);
        } else if (cellValue.includes('Комментарии') && !cellValue.includes('Количество')) {
          inReviewsSection = false;
          inCommentsSection = true;
          inDiscussionsSection = false;
          console.log(`📍 Найдена секция "Комментарии" в строке ${i + 1}`);
        } else if (cellValue.includes('Активные обсуждения') || cellValue.includes('Обсуждения')) {
          inReviewsSection = false;
          inCommentsSection = false;
          inDiscussionsSection = true;
          console.log(`📍 Найдена секция "Активные обсуждения" в строке ${i + 1}`);
        }
        
        // Подсчитываем записи в каждой секции
        if (inReviewsSection && data[i][0] && data[i][0].toString().includes('.')) {
          reviewsCount++;
        } else if (inCommentsSection && data[i][0] && data[i][0].toString().includes('.')) {
          commentsCount++;
        } else if (inDiscussionsSection && data[i][0] && data[i][0].toString().includes('.')) {
          discussionsCount++;
        }
      }
    }
    
    const totalCount = reviewsCount + commentsCount + discussionsCount;
    
    console.log('\n📊 ФАКТИЧЕСКИЕ ДАННЫЕ:');
    console.log(`- Отзывы: ${reviewsCount}`);
    console.log(`- Комментарии: ${commentsCount}`);
    console.log(`- Активные обсуждения: ${discussionsCount}`);
    console.log(`- Всего: ${totalCount}\n`);
    
    // Вычисляем точность
    const reviewsAccuracy = Math.abs(reviewsCount - referenceData.reviews) / referenceData.reviews * 100;
    const commentsAccuracy = Math.abs(commentsCount - referenceData.comments) / referenceData.comments * 100;
    const discussionsAccuracy = Math.abs(discussionsCount - referenceData.discussions) / referenceData.discussions * 100;
    const totalAccuracy = Math.abs(totalCount - referenceData.total) / referenceData.total * 100;
    
    console.log('🎯 АНАЛИЗ ТОЧНОСТИ:');
    console.log(`- Отзывы: ${(100 - reviewsAccuracy).toFixed(1)}% точность`);
    console.log(`- Комментарии: ${(100 - commentsAccuracy).toFixed(1)}% точность`);
    console.log(`- Обсуждения: ${(100 - discussionsAccuracy).toFixed(1)}% точность`);
    console.log(`- Общая: ${(100 - totalAccuracy).toFixed(1)}% точность\n`);
    
    // Проверяем соответствие требованиям 95%+
    const overallAccuracy = (100 - reviewsAccuracy + 100 - commentsAccuracy + 100 - discussionsAccuracy + 100 - totalAccuracy) / 4;
    
    console.log('🎯 ИТОГОВАЯ ОЦЕНКА:');
    console.log(`📊 Общая точность: ${overallAccuracy.toFixed(1)}%`);
    
    if (overallAccuracy >= 95) {
      console.log('🎉 ОТЛИЧНО! Требование 95%+ выполнено!');
      console.log('✅ Агент успешно справился с задачей!');
    } else if (overallAccuracy >= 80) {
      console.log('⚠️ ХОРОШО, но нужны улучшения для достижения 95%+');
    } else {
      console.log('❌ КРИТИЧЕСКИЕ ПРОБЛЕМЫ! Нужны срочные исправления');
    }
    
    // Детальные рекомендации
    console.log('\n🔧 ДЕТАЛЬНЫЕ РЕКОМЕНДАЦИИ:');
    
    if (reviewsCount !== referenceData.reviews) {
      console.log(`- Отзывы: ожидалось ${referenceData.reviews}, получено ${reviewsCount}`);
      if (reviewsCount < referenceData.reviews) {
        console.log('  → Возможно, не все отзывы извлечены');
      } else {
        console.log('  → Возможно, включены лишние записи');
      }
    }
    
    if (commentsCount !== referenceData.comments) {
      console.log(`- Комментарии: ожидалось ${referenceData.comments}, получено ${commentsCount}`);
      if (commentsCount < referenceData.comments) {
        console.log('  → Возможно, не все комментарии извлечены');
      } else {
        console.log('  → Возможно, включены лишние записи');
      }
    }
    
    if (discussionsCount !== referenceData.discussions) {
      console.log(`- Обсуждения: ожидалось ${referenceData.discussions}, получено ${discussionsCount}`);
      if (discussionsCount < referenceData.discussions) {
        console.log('  → Возможно, не все обсуждения извлечены');
      } else {
        console.log('  → Возможно, включены лишние записи');
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
  }
}

compareWithReference();