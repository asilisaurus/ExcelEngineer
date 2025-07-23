/**
 * 🧪 GOOGLE APPS SCRIPT TESTING V4
 * Тестовый скрипт для процессора V4
 * 
 * Автор: AI Assistant  
 * Дата: 2025
 */

// ID таблиц для тестирования
const TEST_CONFIG = {
  SOURCE_SHEET_ID: '1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI',
  REFERENCE_SHEET_ID: '1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk',
  
  MONTHS_TO_TEST: {
    'фев25': { name: 'Февраль', year: 2025 },
    'март25': { name: 'Март', year: 2025 },
    'апр25': { name: 'Апрель', year: 2025 },
    'май25': { name: 'Май', year: 2025 }
  }
};

/**
 * Основная функция тестирования
 */
function runFinalTestingV4() {
  console.log('🧪 ЗАПУСК ТЕСТИРОВАНИЯ PROCESSOR V4');
  console.log('=' .repeat(60));
  
  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    details: []
  };
  
  try {
    // Открываем исходную таблицу
    const sourceSpreadsheet = SpreadsheetApp.openById(TEST_CONFIG.SOURCE_SHEET_ID);
    const referenceSpreadsheet = SpreadsheetApp.openById(TEST_CONFIG.REFERENCE_SHEET_ID);
    
    // Тестируем каждый месяц
    for (const [sheetName, monthInfo] of Object.entries(TEST_CONFIG.MONTHS_TO_TEST)) {
      console.log(`\n📅 Тестирование месяца: ${monthInfo.name} ${monthInfo.year}`);
      console.log('-' .repeat(40));
      
      results.total++;
      
      const testResult = testMonth(sourceSpreadsheet, referenceSpreadsheet, sheetName, monthInfo);
      
      if (testResult.passed) {
        results.passed++;
        console.log(`✅ ${monthInfo.name}: ПРОЙДЕН (${testResult.accuracy}% точность)`);
      } else {
        results.failed++;
        console.log(`❌ ${monthInfo.name}: НЕ ПРОЙДЕН`);
      }
      
      results.details.push(testResult);
    }
    
  } catch (error) {
    console.error('❌ Ошибка тестирования:', error);
  }
  
  // Выводим итоги
  console.log('\n' + '=' .repeat(60));
  console.log('📊 ИТОГИ ТЕСТИРОВАНИЯ:');
  console.log(`Всего тестов: ${results.total}`);
  console.log(`✅ Пройдено: ${results.passed}`);
  console.log(`❌ Провалено: ${results.failed}`);
  console.log(`🎯 Общий результат: ${results.passed === results.total ? 'УСПЕХ!' : 'ТРЕБУЕТСЯ ДОРАБОТКА'}`);
  
  return results;
}

/**
 * Тестирование одного месяца
 */
function testMonth(sourceSpreadsheet, referenceSpreadsheet, sheetName, monthInfo) {
  const result = {
    month: monthInfo.name,
    year: monthInfo.year,
    passed: false,
    accuracy: 0,
    details: {
      reviews: { expected: 0, actual: 0, match: false },
      comments: { expected: 0, actual: 0, match: false },
      discussions: { expected: 0, actual: 0, match: false },
      totalViews: { expected: 0, actual: 0, match: false }
    },
    errors: []
  };
  
  try {
    // 1. Обрабатываем исходные данные
    const processor = new ProcessorV4();
    const processResult = processor.processReport(sourceSpreadsheet.getId(), sheetName);
    
    if (!processResult.success) {
      result.errors.push(`Ошибка обработки: ${processResult.error}`);
      return result;
    }
    
    // 2. Получаем эталонные данные
    const referenceSheet = referenceSpreadsheet.getSheetByName(`${monthInfo.name} ${monthInfo.year}`);
    if (!referenceSheet) {
      result.errors.push(`Эталонный лист "${monthInfo.name} ${monthInfo.year}" не найден`);
      return result;
    }
    
    const referenceData = analyzeReferenceSheet(referenceSheet);
    
    // 3. Сравниваем результаты
    result.details.reviews.expected = referenceData.reviews;
    result.details.reviews.actual = processResult.statistics.reviewsCount;
    result.details.reviews.match = result.details.reviews.expected === result.details.reviews.actual;
    
    result.details.comments.expected = referenceData.comments;
    result.details.comments.actual = processResult.statistics.commentsCount;
    result.details.comments.match = result.details.comments.expected === result.details.comments.actual;
    
    result.details.discussions.expected = referenceData.discussions;
    result.details.discussions.actual = processResult.statistics.discussionsCount;
    result.details.discussions.match = result.details.discussions.expected === result.details.discussions.actual;
    
    result.details.totalViews.expected = referenceData.totalViews;
    result.details.totalViews.actual = processResult.statistics.totalViews;
    result.details.totalViews.match = Math.abs(result.details.totalViews.expected - result.details.totalViews.actual) < 100;
    
    // 4. Вычисляем точность
    let matches = 0;
    if (result.details.reviews.match) matches++;
    if (result.details.comments.match) matches++;
    if (result.details.discussions.match) matches++;
    if (result.details.totalViews.match) matches++;
    
    result.accuracy = Math.round((matches / 4) * 100);
    result.passed = result.accuracy >= 95;
    
    // 5. Выводим детали
    console.log(`📊 Отзывы: ожидается ${result.details.reviews.expected}, получено ${result.details.reviews.actual} ${result.details.reviews.match ? '✅' : '❌'}`);
    console.log(`📊 Комментарии: ожидается ${result.details.comments.expected}, получено ${result.details.comments.actual} ${result.details.comments.match ? '✅' : '❌'}`);
    console.log(`📊 Обсуждения: ожидается ${result.details.discussions.expected}, получено ${result.details.discussions.actual} ${result.details.discussions.match ? '✅' : '❌'}`);
    console.log(`📊 Просмотры: ожидается ${result.details.totalViews.expected}, получено ${result.details.totalViews.actual} ${result.details.totalViews.match ? '✅' : '❌'}`);
    
  } catch (error) {
    result.errors.push(error.toString());
    console.error(`❌ Ошибка при тестировании ${monthInfo.name}:`, error);
  }
  
  return result;
}

/**
 * Анализ эталонного листа
 */
function analyzeReferenceSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  const result = {
    reviews: 0,
    comments: 0,
    discussions: 0,
    totalViews: 0
  };
  
  let inReviews = false;
  let inComments = false;
  let inDiscussions = false;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const firstCell = String(row[0] || '').toLowerCase();
    
    // Определяем разделы
    if (firstCell === 'отзывы') {
      inReviews = true;
      inComments = false;
      inDiscussions = false;
      continue;
    } else if (firstCell.includes('комментарии топ-20')) {
      inReviews = false;
      inComments = true;
      inDiscussions = false;
      continue;
    } else if (firstCell.includes('активные обсуждения')) {
      inReviews = false;
      inComments = false;
      inDiscussions = true;
      continue;
    }
    
    // Считаем записи
    if (row[0] && !firstCell.includes('суммарное') && !firstCell.includes('количество') && !firstCell.includes('доля')) {
      if (inReviews) result.reviews++;
      else if (inComments) result.comments++;
      else if (inDiscussions) result.discussions++;
    }
    
    // Ищем статистику
    if (firstCell.includes('суммарное количество просмотров')) {
      for (let j = 1; j < row.length; j++) {
        const value = parseFloat(String(row[j] || '').replace(/[^\d]/g, ''));
        if (!isNaN(value) && value > 0) {
          result.totalViews = Math.floor(value);
          break;
        }
      }
    }
  }
  
  return result;
}

/**
 * Быстрый тест текущего листа
 */
function quickTestCurrentSheetV4() {
  const processor = new ProcessorV4();
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  
  console.log(`🧪 Быстрый тест листа: ${sheetName}`);
  
  const result = processor.processReport(spreadsheetId, sheetName);
  
  if (result.success) {
    console.log('✅ Обработка успешна');
    console.log(`📊 Результаты:`);
    console.log(`- Отзывов: ${result.statistics.reviewsCount}`);
    console.log(`- Комментариев: ${result.statistics.commentsCount}`);
    console.log(`- Обсуждений: ${result.statistics.discussionsCount}`);
    console.log(`- Просмотров: ${result.statistics.totalViews}`);
    console.log(`📄 Отчет: ${result.reportUrl}`);
  } else {
    console.error('❌ Ошибка:', result.error);
  }
  
  return result;
}

/**
 * Добавление пунктов меню
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧪 Тестирование V4')
    .addItem('🚀 Полное тестирование', 'runFinalTestingV4')
    .addItem('⚡ Быстрый тест текущего листа', 'quickTestCurrentSheetV4')
    .addToUi();
}