const XLSX = require('xlsx');
const fs = require('fs');

async function debugFiltering() {
  console.log('🔍 ДЕБАГ ФИЛЬТРАЦИИ');
  console.log('===================');
  
  const testFile = 'uploads/Fortedetrim ORM report source.xlsx';
  
  try {
    const workbook = XLSX.readFile(testFile);
    
    // Найдем лист с данными месяца
    const months = ["Янв25", "Фев25", "Мар25", "Март25", "Апр25", "Май25", "Июн25", 
                   "Июл25", "Авг25", "Сен25", "Окт25", "Ноя25", "Дек25"];
    
    const sheetName = workbook.SheetNames.find(name => 
      months.some(month => name.includes(month))
    ) || workbook.SheetNames[0];
    
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    console.log(`📊 Лист: ${sheetName}`);
    console.log(`📋 Всего строк: ${data.length}`);
    
    // Анализ каждой строки
    let totalRows = 0;
    let withText = 0;
    let withUrls = 0;
    let withDates = 0;
    let withAuthors = 0;
    let withPostTypes = 0;
    let qualityRows = 0;
    
    const reviewPlatforms = ['otzovik', 'irecommend', 'otzyvru', 'pravogolosa', 'medum', 
                            'vseotzyvy', 'otzyvy.pro', 'market.yandex', 'dialog.ru',
                            'goodapteka', 'megapteka', 'uteka', 'nfapteka'];
    
    const commentPlatforms = ['dzen.ru', 'woman.ru', 'forum.baby.ru', 'vk.com', 't.me',
                             'ok.ru', 'otvet.mail.ru', 'babyblog.ru', 'mom.life', 
                             'youtube.com'];
    
    console.log('\n🔍 АНАЛИЗ ПЕРВЫХ 50 СТРОК:');
    console.log('==========================');
    
    for (let i = 0; i < Math.min(50, data.length); i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      totalRows++;
      
      const colA = (row[0] || '').toString().toLowerCase();
      const colB = (row[1] || '').toString();
      const colD = (row[3] || '').toString();
      const colE = (row[4] || '').toString();
      const colG = row[6];
      const colH = (row[7] || '').toString();
      const colN = (row[13] || '').toString();
      
      // Проверяем фильтры
      if (colE && colE.length >= 20) withText++;
      if (colD && colD.includes('http')) withUrls++;
      if (colG && typeof colG === 'number') withDates++;
      if (colH && colH.length >= 3) withAuthors++;
      if (colN && (colN.toLowerCase() === 'ос' || colN.toLowerCase() === 'цс')) withPostTypes++;
      
      let qualityScore = 100;
      if (!colE || colE.length < 20) qualityScore -= 30;
      if (!colD || !colD.includes('http')) qualityScore -= 25;
      if (!colG || typeof colG !== 'number') qualityScore -= 20;
      if (!colH || colH.length < 3) qualityScore -= 15;
      if (!colN || (colN.toLowerCase() !== 'ос' && colN.toLowerCase() !== 'цс')) qualityScore -= 10;
      
      if (qualityScore >= 40) qualityRows++;
      
      // Показываем примеры проблемных строк
      if (i < 10) {
        const type = colA.includes('отзыв') ? 'review' : 
                     colA.includes('комментар') ? 'comment' : 
                     (colB || colD || colE) ? 'content' : 'empty';
        
        console.log(`Строка ${i + 1} (${type}): A="${colA.substring(0, 20)}..." B="${colB.substring(0, 30)}..." D="${colD.substring(0, 30)}..." E="${colE.substring(0, 40)}..." N="${colN}" | Score: ${qualityScore}`);
      }
    }
    
    console.log('\n📊 СТАТИСТИКА ФИЛЬТРАЦИИ:');
    console.log('=========================');
    console.log(`Всего строк: ${totalRows}`);
    console.log(`С текстом (≥20 символов): ${withText} (${Math.round(withText/totalRows*100)}%)`);
    console.log(`С URL (содержит http): ${withUrls} (${Math.round(withUrls/totalRows*100)}%)`);
    console.log(`С датами (число): ${withDates} (${Math.round(withDates/totalRows*100)}%)`);
    console.log(`С авторами (≥3 символа): ${withAuthors} (${Math.round(withAuthors/totalRows*100)}%)`);
    console.log(`С типами постов (ОС/ЦС): ${withPostTypes} (${Math.round(withPostTypes/totalRows*100)}%)`);
    console.log(`Качество ≥40: ${qualityRows} (${Math.round(qualityRows/totalRows*100)}%)`);
    
    // Подсчитаем все записи с контентом
    let contentRows = 0;
    let reviewRows = 0;
    let commentRows = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const colA = (row[0] || '').toString().toLowerCase();
      const colB = (row[1] || '').toString();
      const colD = (row[3] || '').toString();
      const colE = (row[4] || '').toString();
      const colN = (row[13] || '').toString().toLowerCase();
      
      // Заголовки и служебные строки
      if (colA.includes('тип размещения') || colA.includes('площадка') || 
          colB.includes('площадка') || colE.includes('текст сообщения') ||
          colA.includes('план') || colA.includes('итого')) {
        continue;
      }
      
      // Секционные заголовки
      if ((colA === 'отзывы' || colA === 'комментарии') && !colB && !colD && !colE) {
        continue;
      }
      
      // Анализ по URL и платформам
      const urlText = colB + ' ' + colD;
      const isReviewPlatform = reviewPlatforms.some(platform => 
        urlText.toLowerCase().includes(platform)
      );
      const isCommentPlatform = commentPlatforms.some(platform => 
        urlText.toLowerCase().includes(platform)
      );
      
      // Определяем тип по платформе и типу поста
      if ((colB || colD || colE) && (isReviewPlatform || colN === 'ос')) {
        reviewRows++;
        contentRows++;
      } else if ((colB || colD || colE) && (isCommentPlatform || colN === 'цс')) {
        commentRows++;
        contentRows++;
      } else if (colB || colD || colE) {
        contentRows++;
      }
    }
    
    console.log('\n📋 КЛАССИФИКАЦИЯ КОНТЕНТА:');
    console.log('=========================');
    console.log(`Всего записей с контентом: ${contentRows}`);
    console.log(`Отзывы: ${reviewRows}`);
    console.log(`Комментарии: ${commentRows}`);
    console.log(`Неклассифицированные: ${contentRows - reviewRows - commentRows}`);
    
    // Рекомендации по улучшению фильтров
    console.log('\n💡 РЕКОМЕНДАЦИИ:');
    console.log('================');
    console.log('1. Ослабить фильтр URL - много записей без валидных URL');
    console.log('2. Уменьшить минимальную длину текста до 10 символов');
    console.log('3. Не требовать обязательные даты и авторов');
    console.log('4. Сосредоточиться на классификации по типам постов');
    console.log('5. Использовать более гибкую систему оценки качества');
    
  } catch (error) {
    console.error('❌ Ошибка при дебаге:', error);
  }
}

debugFiltering().catch(console.error); 