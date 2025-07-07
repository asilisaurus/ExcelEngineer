const axios = require('axios');
const XLSX = require('xlsx');

function extractGoogleSheetsInfo(url) {
  try {
    const patterns = [
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)\/edit/,
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)\/edit#gid=(\d+)/
    ];
    
    let spreadsheetId = null;
    let gid = null;
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) {
        spreadsheetId = match[1];
        if (match[2]) {
          gid = match[2];
        }
        break;
      }
    }
    
    if (!gid) {
      const gidMatch = url.match(/[#&]gid=(\d+)/);
      if (gidMatch) {
        gid = gidMatch[1];
      }
    }
    
    return { spreadsheetId, gid };
  } catch (error) {
    console.error('Ошибка при разборе URL:', error);
    return { spreadsheetId: null, gid: null };
  }
}

async function importFromGoogleSheets(url) {
  const { spreadsheetId, gid } = extractGoogleSheetsInfo(url);
  
  if (!spreadsheetId) {
    throw new Error('Неверный URL Google Таблиц');
  }
  
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx${gid ? `&gid=${gid}` : ''}`;
  
  try {
    const response = await axios.get(exportUrl, {
      responseType: 'arraybuffer',
      timeout: 30000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    return Buffer.from(response.data);
  } catch (error) {
    throw new Error(`Ошибка загрузки: ${error.message}`);
  }
}

function analyzeRowType(row, rowIndex) {
  if (!row || row.length === 0) return 'empty';
  
  const colA = (row[0] || '').toString().toLowerCase();
  const colB = (row[1] || '').toString().toLowerCase();
  const colD = (row[3] || '').toString().toLowerCase();
  const colE = (row[4] || '').toString().toLowerCase();
  const colN = (row[13] || '').toString().toLowerCase();
  
  // Заголовки и служебные строки
  if (colA.includes('тип размещения') || colA.includes('площадка') || 
      colB.includes('площадка') || colE.includes('текст сообщения')) {
    return 'header';
  }
  
  // Секционные заголовки
  if ((colA.includes('отзыв') && colA.length < 15) || 
      (colA.includes('комментар') && colA.length < 25)) {
    return 'section_header';
  }
  
  // Определяем тип контента
  if (colA.includes('отзыв') && (colB || colD || colE)) {
    return 'review';
  }
  
  if (colA.includes('комментар') && (colB || colD || colE)) {
    return 'comment';
  }
  
  // Анализ по URL и платформам
  const urlText = colB + ' ' + colD;
  
  // Определяем платформы отзывов
  const reviewPlatforms = ['otzovik', 'irecommend', 'otzyvru', 'pravogolosa', 'medum', 'vseotzyvy', 'otzyvy.pro', 'market.yandex', 'dialog.ru', 'goodapteka', 'megapteka', 'uteka', 'nfapteka'];
  
  // Определяем платформы комментариев  
  const commentPlatforms = ['dzen.ru', 'woman.ru', 'forum.baby.ru', 'vk.com', 't.me', 'ok.ru', 'otvet.mail.ru', 'babyblog.ru', 'mom.life', 'youtube.com'];
  
  const isReviewPlatform = reviewPlatforms.some(platform => urlText.includes(platform));
  const isCommentPlatform = commentPlatforms.some(platform => urlText.includes(platform));
  
  // Анализ типа поста в колонке N
  const postType = colN;
  
  if ((colB || colD || colE) && (isReviewPlatform || postType === 'ос')) {
    return 'review';
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

function analyzeDataQuality(row, rowIndex) {
  const issues = [];
  
  const colB = (row[1] || '').toString();
  const colD = (row[3] || '').toString();
  const colE = (row[4] || '').toString();
  const colG = row[6]; // Дата
  const colH = (row[7] || '').toString(); // Ник
  const colN = (row[13] || '').toString(); // Тип поста
  
  // Проверки качества данных
  if (!colE || colE.length < 10) {
    issues.push('no_text');
  }
  
  if (!colD || !colD.includes('http')) {
    issues.push('no_url');
  }
  
  if (!colG || typeof colG !== 'number') {
    issues.push('no_date');
  }
  
  if (!colH) {
    issues.push('no_author');
  }
  
  if (!colN || (colN !== 'ос' && colN !== 'цс')) {
    issues.push('no_post_type');
  }
  
  return {
    score: Math.max(0, 100 - (issues.length * 20)),
    issues
  };
}

async function smartAnalysis() {
  console.log('🧠 УМНЫЙ АНАЛИЗ БИЗНЕС-ЛОГИКИ');
  console.log('=============================');
  
  const sourceUrl = 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?gid=1258324199#gid=1258324199';
  const resultUrl = 'https://docs.google.com/spreadsheets/d/18jkeKNIn5QJpzQrDN3RAT3mEYRXlSNKOnNvA3pxoBx8/edit?gid=535445992#gid=535445992';
  
  try {
    // Анализ исходного файла
    console.log('📊 АНАЛИЗ ИСХОДНОГО ФАЙЛА');
    console.log('========================');
    
    const sourceBuffer = await importFromGoogleSheets(sourceUrl);
    const sourceWorkbook = XLSX.read(sourceBuffer, { type: 'buffer' });
    const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
    const sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1 });
    
    const stats = {
      total: 0,
      reviews: 0,
      comments: 0,
      headers: 0,
      empty: 0,
      lowQuality: 0
    };
    
    const platformStats = {};
    const qualityIssues = {};
    
    console.log('\n📋 КЛАССИФИКАЦИЯ СТРОК:');
    console.log('=======================');
    
    for (let i = 0; i < sourceData.length; i++) {
      const row = sourceData[i];
      const type = analyzeRowType(row, i);
      const quality = analyzeDataQuality(row, i);
      
      stats.total++;
      stats[type] = (stats[type] || 0) + 1;
      
      if (quality.score < 60) {
        stats.lowQuality++;
      }
      
      // Анализ платформ
      if (type === 'review' || type === 'comment') {
        const url = (row[1] || '') + ' ' + (row[3] || '');
        const domain = url.match(/https?:\/\/([^\/]+)/);
        if (domain) {
          const platform = domain[1];
          platformStats[platform] = platformStats[platform] || { reviews: 0, comments: 0 };
          platformStats[platform][type === 'review' ? 'reviews' : 'comments']++;
        }
      }
      
      // Сбор проблем качества
      quality.issues.forEach(issue => {
        qualityIssues[issue] = (qualityIssues[issue] || 0) + 1;
      });
      
      // Показываем первые несколько примеров каждого типа
      if (stats[type] <= 3 && (type === 'review' || type === 'comment')) {
        const text = (row[4] || '').toString().substring(0, 60);
        const postType = row[13] || 'не указан';
        console.log(`${type === 'review' ? '📝' : '💬'} Строка ${i + 1} (${type}): ${text}... [${postType}]`);
      }
    }
    
    console.log('\n📊 СТАТИСТИКА КЛАССИФИКАЦИИ:');
    console.log('============================');
    Object.entries(stats).forEach(([key, value]) => {
      console.log(`${key}: ${value}`);
    });
    
    console.log('\n🌐 СТАТИСТИКА ПО ПЛАТФОРМАМ:');
    console.log('============================');
    Object.entries(platformStats)
      .sort(([,a], [,b]) => (b.reviews + b.comments) - (a.reviews + a.comments))
      .slice(0, 10)
      .forEach(([platform, data]) => {
        console.log(`${platform}: ${data.reviews} отзывов, ${data.comments} комментариев`);
      });
    
    console.log('\n⚠️ ПРОБЛЕМЫ КАЧЕСТВА ДАННЫХ:');
    console.log('============================');
    Object.entries(qualityIssues).forEach(([issue, count]) => {
      console.log(`${issue}: ${count} записей`);
    });
    
    // Анализ файла результата для понимания правил фильтрации
    console.log('\n🎯 АНАЛИЗ ПРАВИЛ ФИЛЬТРАЦИИ');
    console.log('===========================');
    
    const resultBuffer = await importFromGoogleSheets(resultUrl);
    const resultWorkbook = XLSX.read(resultBuffer, { type: 'buffer' });
    const resultSheet = resultWorkbook.Sheets[resultWorkbook.SheetNames[0]];
    const resultData = XLSX.utils.sheet_to_json(resultSheet, { header: 1 });
    
    // Извлекаем URL из результата для сравнения
    const resultUrls = new Set();
    let resultReviews = 0;
    let resultComments = 0;
    let inReviewSection = false;
    let inCommentSection = false;
    
    for (let i = 0; i < resultData.length; i++) {
      const row = resultData[i];
      if (row && row[0]) {
        const cellText = String(row[0]).toLowerCase();
        if (cellText === 'отзывы') {
          inReviewSection = true;
          inCommentSection = false;
          continue;
        }
        if (cellText.includes('комментар')) {
          inReviewSection = false;
          inCommentSection = true;
          continue;
        }
      }
      
      if (row && row[1] && (inReviewSection || inCommentSection)) {
        const url = row[1].toString();
        resultUrls.add(url);
        
        if (inReviewSection) resultReviews++;
        if (inCommentSection) resultComments++;
      }
    }
    
    console.log(`\n📊 РЕЗУЛЬТАТ СОДЕРЖИТ:`);
    console.log(`📝 Отзывы: ${resultReviews}`);
    console.log(`💬 Комментарии: ${resultComments}`);
    console.log(`🔗 Уникальные URL: ${resultUrls.size}`);
    
    // Найдем какие записи были исключены
    console.log('\n🔍 АНАЛИЗ ИСКЛЮЧЕННЫХ ЗАПИСЕЙ:');
    console.log('==============================');
    
    let sourceReviews = 0;
    let sourceComments = 0;
    
    for (let i = 0; i < sourceData.length; i++) {
      const row = sourceData[i];
      const type = analyzeRowType(row, i);
      
      if (type === 'review') sourceReviews++;
      if (type === 'comment') sourceComments++;
    }
    
    console.log(`Исходно отзывов: ${sourceReviews}, в результате: ${resultReviews} (исключено: ${sourceReviews - resultReviews})`);
    console.log(`Исходно комментариев: ${sourceComments}, в результате: ${resultComments} (исключено: ${sourceComments - resultComments})`);
    
    // Выводим предлагаемые правила фильтрации
    console.log('\n💡 ПРЕДЛАГАЕМЫЕ ПРАВИЛА ФИЛЬТРАЦИИ:');
    console.log('===================================');
    console.log('1. Исключать записи без URL-ов или с некорректными URL-ами');
    console.log('2. Исключать записи без текста или с текстом < 10 символов');
    console.log('3. Исключать записи без корректного типа поста (ОС/ЦС)');
    console.log('4. Приоритизировать записи с датами и никами авторов');
    console.log('5. Классифицировать по платформам: отзывы vs комментарии');
    
  } catch (error) {
    console.error('❌ Ошибка при умном анализе:', error.message);
  }
}

smartAnalysis().catch(console.error); 