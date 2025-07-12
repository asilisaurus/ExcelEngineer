/**
 * ============================================================================
 * GOOGLE APPS SCRIPT PROCESSOR - REAL DATA VERSION
 * ============================================================================
 * Исправленная версия для работы с реальными данными из эталонной таблицы
 * ID таблицы: 1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk
 * ============================================================================
 */

// =============================================================================
// КОНФИГУРАЦИЯ
// =============================================================================

// ID эталонной таблицы с 4 месяцами
const REFERENCE_SHEET_ID = '1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk';

// Целевые метрики для проверки
const TARGET_METRICS = {
  reviews: 13,    // ОС записей
  comments: 15,   // ЦС записей
  discussions: 42 // ПС записей (игнорируются)
};

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Главная функция обработки реальных данных
 */
function processRealGoogleSheets() {
  try {
    console.log('🚀 Начинаю обработку реальных данных из эталонной таблицы...');
    
    // 1. Получаем данные из таблицы
    const sourceData = getRealSourceData();
    console.log(`📊 Загружено ${sourceData.length} строк из источника`);
    
    // 2. Обрабатываем данные
    const processedData = processRealData(sourceData);
    console.log(`✅ Обработано ${processedData.length} записей`);
    
    // 3. Проверяем соответствие целевым метрикам
    const metricsCheck = checkTargetMetrics(processedData);
    console.log(`📊 Проверка метрик: ${metricsCheck.passed ? 'ПРОЙДЕНА' : 'НЕ ПРОЙДЕНА'}`);
    
    // 4. Создаем результирующий файл
    const resultFileId = createRealResultFile(processedData, metricsCheck);
    console.log(`💾 Создан файл результата: ${resultFileId}`);
    
    return {
      success: true,
      sourceRows: sourceData.length,
      processedRows: processedData.length,
      metricsCheck: metricsCheck,
      resultFileId: resultFileId
    };
    
  } catch (error) {
    console.error('❌ Ошибка при обработке:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Получение данных из эталонной таблицы
 */
function getRealSourceData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(REFERENCE_SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    console.log(`📋 Найдено ${sheets.length} листов в таблице`);
    
    // Автоматический поиск наиболее подходящего листа
    const targetSheet = findBestSheetInRealData(sheets);
    
    if (!targetSheet) {
      throw new Error('Не найден подходящий лист для обработки');
    }
    
    console.log(`📋 Выбран лист: ${targetSheet.getName()}`);
    
    // Получаем все данные листа
    const range = targetSheet.getDataRange();
    const values = range.getValues();
    
    console.log(`📊 Загружено ${values.length} строк из листа ${targetSheet.getName()}`);
    
    return values;
    
  } catch (error) {
    console.error('❌ Ошибка при получении данных:', error);
    throw error;
  }
}

/**
 * Поиск лучшего листа в реальных данных
 */
function findBestSheetInRealData(sheets) {
  let bestSheet = null;
  let bestScore = 0;
  
  console.log('🔍 Анализируем листы для выбора лучшего...');
  
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length === 0) continue;
    
    let score = 0;
    
    // Анализируем структуру листа
    const headerInfo = findRealHeaders(values);
    
    // Баллы за наличие ключевых заголовков
    const keyHeaders = ['тип размещения', 'площадка', 'текст сообщения', 'тип поста'];
    const foundHeaders = headerInfo.headers.map(h => h.toLowerCase());
    
    keyHeaders.forEach(key => {
      if (foundHeaders.some(h => h.includes(key))) {
        score += 20;
      }
    });
    
    // Баллы за количество данных
    const dataRows = values.slice(headerInfo.row + 1);
    score += Math.min(dataRows.length * 0.1, 50);
    
    // Анализируем типы контента
    const contentTypes = new Set();
    let sampleCount = 0;
    
    for (const row of dataRows) {
      if (sampleCount >= 50) break;
      
      const contentType = determineRealContentType(row);
      if (contentType !== 'unknown') {
        contentTypes.add(contentType);
        sampleCount++;
      }
    }
    
    score += contentTypes.size * 15;
    
    // Предпочтение листам с названием месяца
    const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 
                       'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    
    if (monthNames.some(month => sheetName.toLowerCase().includes(month))) {
      score += 10;
    }
    
    console.log(`📊 Лист "${sheetName}": ${score} баллов (${sampleCount} записей, ${contentTypes.size} типов)`);
    
    if (score > bestScore) {
      bestScore = score;
      bestSheet = sheet;
    }
  }
  
  return bestSheet;
}

/**
 * Обработка реальных данных
 */
function processRealData(rawData) {
  if (!rawData || rawData.length === 0) {
    throw new Error('Нет данных для обработки');
  }
  
  // 1. Поиск заголовков
  const headerInfo = findRealHeaders(rawData);
  console.log(`🔍 Найдена строка заголовков: ${headerInfo.row}`);
  
  // 2. Создание маппинга колонок
  const columnMapping = createRealColumnMapping(headerInfo.headers);
  console.log(`🗂️ Создан маппинг для ${Object.keys(columnMapping).length} колонок`);
  
  // 3. Извлечение данных
  const dataRows = rawData.slice(headerInfo.row + 1);
  
  // 4. Обработка строк
  const processedRows = [];
  let skippedRows = 0;
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    
    // Пропускаем пустые строки
    if (isEmptyRow(row)) {
      skippedRows++;
      continue;
    }
    
    // Пропускаем строки-заголовки
    if (isRealHeaderRow(row)) {
      console.log(`⏭️ Пропущена строка-заголовок: ${row[0]}`);
      skippedRows++;
      continue;
    }
    
    // Определяем тип контента
    const contentType = determineRealContentType(row);
    
    if (contentType !== 'unknown') {
      const processedRow = extractRealRowData(row, contentType, columnMapping);
      if (processedRow) {
        processedRows.push(processedRow);
      }
    } else {
      skippedRows++;
    }
  }
  
  console.log(`📊 Обработано строк: ${processedRows.length}, пропущено: ${skippedRows}`);
  
  // 5. Группировка по типам
  const reviews = processedRows.filter(row => row.type === 'review');
  const comments = processedRows.filter(row => row.type === 'comment');
  const discussions = processedRows.filter(row => row.type === 'discussion');
  
  console.log(`📊 Найдено: отзывов=${reviews.length}, комментариев=${comments.length}, обсуждений=${discussions.length}`);
  
  return processedRows;
}

/**
 * Поиск заголовков в реальных данных
 */
function findRealHeaders(data) {
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    
    const rowStr = row.map(cell => (cell || '').toString().toLowerCase()).join(' ');
    
    if (rowStr.includes('тип размещения') || 
        rowStr.includes('площадка') || 
        rowStr.includes('текст сообщения') ||
        rowStr.includes('тип поста')) {
      
      return {
        row: i,
        headers: row.map(cell => (cell || '').toString().trim())
      };
    }
  }
  
  // Если заголовки не найдены, используем первую строку
  return {
    row: 0,
    headers: data[0] ? data[0].map(cell => (cell || '').toString().trim()) : []
  };
}

/**
 * Создание маппинга колонок для реальных данных
 */
function createRealColumnMapping(headers) {
  const columnMapping = {};
  
  headers.forEach((header, index) => {
    const cleanHeader = header.toLowerCase().trim();
    if (cleanHeader) {
      columnMapping[cleanHeader] = index;
    }
  });
  
  // Стандартный маппинг для fallback
  const defaultMapping = {
    'тип размещения': 0,
    'площадка': 1,
    'продукт': 2,
    'ссылка на сообщение': 3,
    'текст сообщения': 4,
    'согласование/комментарии': 5,
    'дата': 6,
    'ник': 7,
    'автор': 8,
    'просмотры темы на старте': 9,
    'просмотры в конце месяца': 10,
    'просмотров получено': 11,
    'вовлечение': 12,
    'тип поста': 13
  };
  
  // Объединяем маппинги
  return { ...defaultMapping, ...columnMapping };
}

/**
 * Определение типа контента в реальных данных
 */
function determineRealContentType(row) {
  if (!row || row.length === 0) return 'unknown';
  
  // Проверяем последнюю колонку (тип поста)
  const lastColIndex = row.length - 1;
  const lastColValue = (row[lastColIndex] || '').toString().toLowerCase().trim();
  
  if (lastColValue === 'ос' || lastColValue === 'основное сообщение') {
    return 'review';
  }
  
  if (lastColValue === 'цс' || lastColValue === 'целевое сообщение') {
    return 'comment';
  }
  
  if (lastColValue === 'пс' || lastColValue === 'посты сообщения') {
    return 'discussion';
  }
  
  // Дополнительная проверка по первой колонке
  const firstColValue = (row[0] || '').toString().toLowerCase().trim();
  
  if (firstColValue.includes('отзыв') || firstColValue.includes('основное')) {
    return 'review';
  }
  
  if (firstColValue.includes('комментарий') || firstColValue.includes('целевое')) {
    return 'comment';
  }
  
  if (firstColValue.includes('обсуждение') || firstColValue.includes('пост')) {
    return 'discussion';
  }
  
  // Проверка по тексту сообщения (если есть содержательный текст)
  if (row.length > 4) {
    const textValue = (row[4] || '').toString().trim();
    if (textValue.length > 10) {
      // Если есть содержательный текст, возможно это review
      return 'review';
    }
  }
  
  return 'unknown';
}

/**
 * Проверка, является ли строка заголовком секции
 */
function isRealHeaderRow(row) {
  if (!row || row.length === 0) return false;
  
  const firstCell = (row[0] || '').toString().toLowerCase().trim();
  const headerPatterns = ['отзывы', 'комментарии', 'обсуждения', 'итого', 'всего'];
  
  return headerPatterns.some(pattern => firstCell.includes(pattern));
}

/**
 * Проверка, является ли строка пустой
 */
function isEmptyRow(row) {
  return !row || row.every(cell => !cell || cell.toString().trim() === '');
}

/**
 * Извлечение данных из реальной строки
 */
function extractRealRowData(row, type, mapping) {
  try {
    // Индексы колонок (с fallback на стандартные позиции)
    const siteIndex = mapping['площадка'] || 1;
    const linkIndex = mapping['ссылка на сообщение'] || 3;
    const textIndex = mapping['текст сообщения'] || 4;
    const dateIndex = mapping['дата'] || 6;
    const authorIndex = mapping['автор'] || 8;
    const viewsIndex = mapping['просмотров получено'] || 11;
    const engagementIndex = mapping['вовлечение'] || 12;
    
    // Извлекаем данные
    const site = cleanValue(row[siteIndex]);
    const link = cleanValue(row[linkIndex]);
    const text = cleanValue(row[textIndex]);
    const date = cleanValue(row[dateIndex]);
    const author = cleanValue(row[authorIndex]);
    const views = extractViews(row[viewsIndex]);
    const engagement = extractEngagement(row[engagementIndex]);
    
    // Валидация: нужна либо площадка, либо текст
    if (!site && !text) {
      return null;
    }
    
    // Если нет площадки, пытаемся извлечь из ссылки
    let finalSite = site;
    if (!finalSite && link) {
      finalSite = extractSiteFromLink(link);
    }
    
    return {
      type: type,
      site: finalSite || 'Неизвестная площадка',
      link: link,
      text: text || 'Нет текста',
      date: date,
      author: author,
      views: views,
      engagement: engagement
    };
    
  } catch (error) {
    console.error('❌ Ошибка при извлечении данных строки:', error);
    return null;
  }
}

/**
 * Извлечение сайта из ссылки
 */
function extractSiteFromLink(link) {
  try {
    if (!link) return '';
    
    const linkStr = link.toString().toLowerCase();
    
    if (linkStr.includes('market.yandex') || linkStr.includes('yandex.ru')) {
      return 'Яндекс.Маркет';
    }
    
    if (linkStr.includes('ozon.ru') || linkStr.includes('ozon.com')) {
      return 'Ozon';
    }
    
    if (linkStr.includes('wildberries.ru') || linkStr.includes('wb.ru')) {
      return 'Wildberries';
    }
    
    if (linkStr.includes('avito.ru')) {
      return 'Avito';
    }
    
    if (linkStr.includes('aliexpress')) {
      return 'AliExpress';
    }
    
    // Пытаемся извлечь домен
    const match = linkStr.match(/https?:\/\/([^\/]+)/);
    if (match && match[1]) {
      return match[1];
    }
    
    return '';
    
  } catch (error) {
    return '';
  }
}

/**
 * Очистка значения ячейки
 */
function cleanValue(value) {
  if (!value) return '';
  return value.toString().trim();
}

/**
 * Извлечение просмотров
 */
function extractViews(value) {
  if (!value) return 0;
  
  const str = value.toString().replace(/\D/g, '');
  const num = parseInt(str);
  
  return isNaN(num) ? 0 : num;
}

/**
 * Извлечение вовлечения
 */
function extractEngagement(value) {
  if (!value) return 0;
  
  const str = value.toString().replace(/[^\d.,]/g, '').replace(',', '.');
  const num = parseFloat(str);
  
  return isNaN(num) ? 0 : num;
}

/**
 * Проверка соответствия целевым метрикам
 */
function checkTargetMetrics(processedData) {
  const reviews = processedData.filter(row => row.type === 'review');
  const comments = processedData.filter(row => row.type === 'comment');
  const discussions = processedData.filter(row => row.type === 'discussion');
  
  const results = {
    reviews: {
      actual: reviews.length,
      target: TARGET_METRICS.reviews,
      passed: reviews.length === TARGET_METRICS.reviews
    },
    comments: {
      actual: comments.length,
      target: TARGET_METRICS.comments,
      passed: comments.length === TARGET_METRICS.comments
    },
    discussions: {
      actual: discussions.length,
      target: TARGET_METRICS.discussions,
      passed: discussions.length === TARGET_METRICS.discussions
    }
  };
  
  const totalPassed = results.reviews.passed && results.comments.passed;
  
  console.log(`📊 Метрики:`);
  console.log(`   Отзывы: ${results.reviews.actual}/${results.reviews.target} ${results.reviews.passed ? '✅' : '❌'}`);
  console.log(`   Комментарии: ${results.comments.actual}/${results.comments.target} ${results.comments.passed ? '✅' : '❌'}`);
  console.log(`   Обсуждения: ${results.discussions.actual}/${results.discussions.target} ${results.discussions.passed ? '✅' : '❌'}`);
  
  return {
    ...results,
    passed: totalPassed,
    accuracy: totalPassed ? 100 : Math.round((results.reviews.actual + results.comments.actual) / (TARGET_METRICS.reviews + TARGET_METRICS.comments) * 100)
  };
}

/**
 * Создание результирующего файла
 */
function createRealResultFile(processedData, metricsCheck) {
  try {
    // Создаем новый спредшит
    const currentDate = new Date();
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    
    const fileName = `Результат_Реальных_Данных_${monthNames[currentDate.getMonth()]}_${currentDate.getFullYear()}`;
    const spreadsheet = SpreadsheetApp.create(fileName);
    
    // Лист 1: Результаты
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName('Результаты');
    
    // Заголовки
    const headers = ['Тип', 'Площадка', 'Ссылка', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Вовлечение'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Данные
    const dataForSheet = processedData.map(row => [
      row.type === 'review' ? 'Отзыв' : row.type === 'comment' ? 'Комментарий' : 'Обсуждение',
      row.site,
      row.link,
      row.text,
      row.date,
      row.author,
      row.views,
      row.engagement
    ]);
    
    if (dataForSheet.length > 0) {
      sheet.getRange(2, 1, dataForSheet.length, headers.length).setValues(dataForSheet);
    }
    
    // Форматирование
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
    
    // Лист 2: Метрики
    const metricsSheet = spreadsheet.insertSheet('Метрики');
    
    const metricsHeaders = ['Тип', 'Цель', 'Фактически', 'Статус', 'Точность %'];
    metricsSheet.getRange(1, 1, 1, metricsHeaders.length).setValues([metricsHeaders]);
    
    const metricsData = [
      ['Отзывы', metricsCheck.reviews.target, metricsCheck.reviews.actual, 
       metricsCheck.reviews.passed ? 'ПРОЙДЕНО' : 'НЕ ПРОЙДЕНО',
       Math.round(metricsCheck.reviews.actual / metricsCheck.reviews.target * 100)],
      ['Комментарии', metricsCheck.comments.target, metricsCheck.comments.actual, 
       metricsCheck.comments.passed ? 'ПРОЙДЕНО' : 'НЕ ПРОЙДЕНО',
       Math.round(metricsCheck.comments.actual / metricsCheck.comments.target * 100)],
      ['Обсуждения', metricsCheck.discussions.target, metricsCheck.discussions.actual, 
       metricsCheck.discussions.passed ? 'ПРОЙДЕНО' : 'НЕ ПРОЙДЕНО',
       Math.round(metricsCheck.discussions.actual / metricsCheck.discussions.target * 100)]
    ];
    
    metricsSheet.getRange(2, 1, metricsData.length, metricsHeaders.length).setValues(metricsData);
    metricsSheet.getRange(1, 1, 1, metricsHeaders.length).setFontWeight('bold');
    metricsSheet.autoResizeColumns(1, metricsHeaders.length);
    
    // Добавляем общую точность
    metricsSheet.getRange(6, 1, 1, 2).setValues([['Общая точность:', `${metricsCheck.accuracy}%`]]);
    metricsSheet.getRange(6, 1, 1, 2).setFontWeight('bold');
    
    console.log(`💾 Создан файл: ${fileName}`);
    console.log(`🔗 ID файла: ${spreadsheet.getId()}`);
    
    return spreadsheet.getId();
    
  } catch (error) {
    console.error('❌ Ошибка при создании файла:', error);
    throw error;
  }
}

// =============================================================================
// ТЕСТОВЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Тестирование процессора на реальных данных
 */
function testRealProcessor() {
  console.log('🧪 Запуск тестирования на реальных данных...');
  
  try {
    const result = processRealGoogleSheets();
    
    if (result.success) {
      console.log('✅ Тестирование успешно завершено');
      console.log(`📊 Обработано строк: ${result.processedRows}`);
      console.log(`🎯 Точность: ${result.metricsCheck.accuracy}%`);
      console.log(`🔗 ID результата: ${result.resultFileId}`);
    } else {
      console.log('❌ Тестирование завершилось с ошибкой:', result.error);
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Анализ структуры реальных данных
 */
function analyzeRealDataStructure() {
  console.log('🔍 Анализ структуры реальных данных...');
  
  try {
    const sourceData = getRealSourceData();
    
    if (sourceData.length === 0) {
      console.log('❌ Нет данных для анализа');
      return;
    }
    
    // Анализируем первые 10 строк
    console.log('📋 Первые 10 строк:');
    for (let i = 0; i < Math.min(10, sourceData.length); i++) {
      const row = sourceData[i];
      console.log(`Строка ${i + 1}:`, row.slice(0, 5).map(cell => 
        (cell || '').toString().substring(0, 30)
      ));
    }
    
    // Поиск заголовков
    const headerInfo = findRealHeaders(sourceData);
    console.log('📋 Найденные заголовки:', headerInfo.headers);
    
    // Анализ типов контента
    const contentTypes = new Map();
    const dataRows = sourceData.slice(headerInfo.row + 1);
    
    for (let i = 0; i < Math.min(100, dataRows.length); i++) {
      const row = dataRows[i];
      const type = determineRealContentType(row);
      contentTypes.set(type, (contentTypes.get(type) || 0) + 1);
    }
    
    console.log('📊 Типы контента (первые 100 строк):');
    for (const [type, count] of contentTypes) {
      console.log(`   ${type}: ${count}`);
    }
    
    return {
      totalRows: sourceData.length,
      headerRow: headerInfo.row,
      headers: headerInfo.headers,
      contentTypes: Object.fromEntries(contentTypes)
    };
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
    return { error: error.message };
  }
}

// =============================================================================
// ГЛАВНЫЕ ФУНКЦИИ ДЛЯ ЗАПУСКА
// =============================================================================

/**
 * Основная функция для запуска из Google Apps Script
 */
function main() {
  return processRealGoogleSheets();
}

/**
 * Функция для тестирования на реальных данных
 */
function runRealTest() {
  return testRealProcessor();
}

/**
 * Функция для анализа реальных данных
 */
function runRealAnalysis() {
  return analyzeRealDataStructure();
}