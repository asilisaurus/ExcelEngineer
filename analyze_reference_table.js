/**
 * ============================================================================
 * АНАЛИЗ ЭТАЛОННОЙ ТАБЛИЦЫ С 4 МЕСЯЦАМИ
 * ============================================================================
 * Скрипт для тестирования Google Apps Script процессора на реальных данных
 * ID таблицы: 1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk
 * ============================================================================
 */

// =============================================================================
// КОНСТАНСТЫ
// =============================================================================

const REFERENCE_SHEET_ID = '1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk';

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ АНАЛИЗА
// =============================================================================

/**
 * Главная функция анализа эталонной таблицы
 */
function analyzeReferenceTable() {
  console.log('🔍 Анализ эталонной таблицы с 4 месяцами...');
  
  try {
    // 1. Получаем данные из таблицы
    const spreadsheet = SpreadsheetApp.openById(REFERENCE_SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    console.log(`📊 Найдено листов: ${sheets.length}`);
    
    // 2. Анализируем каждый лист
    const sheetsAnalysis = {};
    
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      console.log(`\n📋 Анализируем лист: ${sheetName}`);
      
      const analysis = analyzeSheet(sheet);
      sheetsAnalysis[sheetName] = analysis;
      
      // Показываем краткий отчет
      console.log(`   📊 Строк: ${analysis.totalRows}`);
      console.log(`   📊 Колонок: ${analysis.totalColumns}`);
      console.log(`   📊 Заголовков: ${analysis.headers.length}`);
      console.log(`   📊 Данных: ${analysis.dataRows.length}`);
      
      if (analysis.contentTypes.length > 0) {
        console.log(`   📊 Типы контента: ${analysis.contentTypes.join(', ')}`);
      }
    }
    
    // 3. Определяем лучший лист для тестирования
    const bestSheet = findBestSheetForTesting(sheetsAnalysis);
    console.log(`\n🎯 Лучший лист для тестирования: ${bestSheet}`);
    
    // 4. Тестируем процессор на выбранном листе
    const testResults = testProcessorOnSheet(spreadsheet, bestSheet);
    
    return {
      sheetsAnalysis,
      bestSheet,
      testResults
    };
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
    return { error: error.message };
  }
}

/**
 * Анализ отдельного листа
 */
function analyzeSheet(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  const analysis = {
    name: sheet.getName(),
    totalRows: values.length,
    totalColumns: values[0] ? values[0].length : 0,
    headers: [],
    dataRows: [],
    contentTypes: [],
    sampleData: []
  };
  
  if (values.length === 0) {
    return analysis;
  }
  
  // Поиск заголовков
  const headerInfo = findHeadersInSheet(values);
  analysis.headers = headerInfo.headers;
  analysis.headerRow = headerInfo.row;
  
  // Анализ данных
  const dataRows = values.slice(headerInfo.row + 1);
  analysis.dataRows = dataRows.filter(row => !isEmptyRow(row));
  
  // Определение типов контента
  const contentTypes = new Set();
  let sampleCount = 0;
  
  for (const row of analysis.dataRows) {
    if (sampleCount >= 5) break;
    
    const contentType = determineContentType(row);
    if (contentType !== 'unknown') {
      contentTypes.add(contentType);
      analysis.sampleData.push({
        row: row.slice(0, 5), // Первые 5 колонок для образца
        type: contentType
      });
      sampleCount++;
    }
  }
  
  analysis.contentTypes = Array.from(contentTypes);
  
  return analysis;
}

/**
 * Поиск заголовков в листе
 */
function findHeadersInSheet(values) {
  for (let i = 0; i < Math.min(10, values.length); i++) {
    const row = values[i];
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
  
  return {
    row: 0,
    headers: values[0] ? values[0].map(cell => (cell || '').toString().trim()) : []
  };
}

/**
 * Определение типа контента
 */
function determineContentType(row) {
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
  if (firstColValue.includes('отзыв') || firstColValue.includes('ос')) {
    return 'review';
  }
  
  if (firstColValue.includes('комментарий') || firstColValue.includes('цс')) {
    return 'comment';
  }
  
  if (firstColValue.includes('обсуждение') || firstColValue.includes('пс')) {
    return 'discussion';
  }
  
  return 'unknown';
}

/**
 * Проверка, является ли строка пустой
 */
function isEmptyRow(row) {
  return !row || row.every(cell => !cell || cell.toString().trim() === '');
}

/**
 * Поиск лучшего листа для тестирования
 */
function findBestSheetForTesting(sheetsAnalysis) {
  let bestSheet = null;
  let bestScore = 0;
  
  for (const [sheetName, analysis] of Object.entries(sheetsAnalysis)) {
    let score = 0;
    
    // Баллы за наличие данных
    score += analysis.dataRows.length * 0.1;
    
    // Баллы за типы контента
    score += analysis.contentTypes.length * 10;
    
    // Баллы за наличие ключевых заголовков
    const keyHeaders = ['площадка', 'текст сообщения', 'тип поста'];
    const hasKeyHeaders = keyHeaders.some(header => 
      analysis.headers.some(h => h.toLowerCase().includes(header))
    );
    
    if (hasKeyHeaders) {
      score += 50;
    }
    
    // Предпочтение листам с названием месяца
    const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 
                       'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    
    if (monthNames.some(month => sheetName.toLowerCase().includes(month))) {
      score += 20;
    }
    
    console.log(`📊 Лист "${sheetName}" получил ${score} баллов`);
    
    if (score > bestScore) {
      bestScore = score;
      bestSheet = sheetName;
    }
  }
  
  return bestSheet;
}

/**
 * Тестирование процессора на выбранном листе
 */
function testProcessorOnSheet(spreadsheet, sheetName) {
  console.log(`\n🧪 Тестирование процессора на листе: ${sheetName}`);
  
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    // Используем функции из нашего процессора
    const processedData = processDataFromSheet(values);
    
    const reviews = processedData.filter(row => row.type === 'review');
    const comments = processedData.filter(row => row.type === 'comment');
    const discussions = processedData.filter(row => row.type === 'discussion');
    
    const results = {
      totalProcessed: processedData.length,
      reviews: reviews.length,
      comments: comments.length,
      discussions: discussions.length,
      success: true
    };
    
    console.log(`📊 Результаты тестирования:`);
    console.log(`   ✅ Обработано записей: ${results.totalProcessed}`);
    console.log(`   📝 Отзывы: ${results.reviews}`);
    console.log(`   💬 Комментарии: ${results.comments}`);
    console.log(`   🗨️ Обсуждения: ${results.discussions}`);
    
    // Показываем примеры данных
    console.log(`\n📋 Примеры обработанных данных:`);
    processedData.slice(0, 5).forEach((row, index) => {
      console.log(`   ${index + 1}. ${row.type}: ${row.site} - ${row.text ? row.text.substring(0, 30) : 'Нет текста'}...`);
    });
    
    return results;
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Обработка данных из листа (адаптировано под Google Apps Script)
 */
function processDataFromSheet(rawData) {
  if (!rawData || rawData.length === 0) {
    throw new Error('Нет данных для обработки');
  }
  
  // 1. Поиск заголовков
  const headerInfo = findHeadersInSheet(rawData);
  
  // 2. Создание маппинга колонок
  const columnMapping = {};
  headerInfo.headers.forEach((header, index) => {
    const cleanHeader = header.toLowerCase().trim();
    if (cleanHeader) {
      columnMapping[cleanHeader] = index;
    }
  });
  
  // Добавляем стандартный маппинг для fallback
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
  const finalMapping = { ...defaultMapping, ...columnMapping };
  
  // 3. Извлечение данных
  const dataRows = rawData.slice(headerInfo.row + 1);
  const processedRows = [];
  
  for (const row of dataRows) {
    // Пропускаем пустые строки
    if (isEmptyRow(row)) continue;
    
    // Пропускаем строки-заголовки
    if (isHeaderRow(row)) continue;
    
    // Определяем тип контента
    const contentType = determineContentType(row);
    
    if (contentType !== 'unknown') {
      const processedRow = extractRowDataFromSheet(row, contentType, finalMapping);
      if (processedRow) {
        processedRows.push(processedRow);
      }
    }
  }
  
  return processedRows;
}

/**
 * Проверка, является ли строка заголовком
 */
function isHeaderRow(row) {
  const firstCell = (row[0] || '').toString().toLowerCase().trim();
  const headerPatterns = ['отзывы', 'комментарии', 'обсуждения'];
  
  return headerPatterns.includes(firstCell);
}

/**
 * Извлечение данных из строки
 */
function extractRowDataFromSheet(row, type, mapping) {
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
    
    // Валидация обязательных полей
    if (!site || !text) {
      return null;
    }
    
    return {
      type: type,
      site: site,
      link: link,
      text: text,
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
  
  const str = value.toString().replace(/[^\d.]/g, '');
  const num = parseFloat(str);
  
  return isNaN(num) ? 0 : num;
}

// =============================================================================
// ФУНКЦИИ ДЛЯ ЗАПУСКА
// =============================================================================

/**
 * Главная функция для запуска анализа
 */
function main() {
  return analyzeReferenceTable();
}

/**
 * Быстрый анализ структуры
 */
function quickAnalysis() {
  console.log('⚡ Быстрый анализ структуры таблицы...');
  
  const spreadsheet = SpreadsheetApp.openById(REFERENCE_SHEET_ID);
  const sheets = spreadsheet.getSheets();
  
  console.log(`📊 Найдено листов: ${sheets.length}`);
  
  sheets.forEach((sheet, index) => {
    const name = sheet.getName();
    const range = sheet.getDataRange();
    const rowCount = range.getNumRows();
    const colCount = range.getNumColumns();
    
    console.log(`${index + 1}. Лист "${name}": ${rowCount} строк, ${colCount} колонок`);
  });
  
  return { totalSheets: sheets.length };
}