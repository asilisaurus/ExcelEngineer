/**
 * ============================================================================
 * GOOGLE APPS SCRIPT PROCESSOR - FINAL VERSION
 * ============================================================================
 * Исправленная версия на основе реального анализа данных
 * Основные улучшения:
 * - Определение типа контента по колонке "тип поста"
 * - Фиксированный маппинг колонок для Google Sheets
 * - Фильтрация заголовков
 * - Упрощенные методы извлечения
 * ============================================================================
 */

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Главная функция обработки Google Sheets
 */
function processGoogleSheets() {
  try {
    console.log('🚀 Начинаю обработку Google Sheets...');
    
    // 1. Получаем исходные данные
    const sourceData = getSourceData();
    console.log(`📊 Загружено ${sourceData.length} строк из источника`);
    
    // 2. Обрабатываем данные
    const processedData = processData(sourceData);
    console.log(`✅ Обработано ${processedData.length} записей`);
    
    // 3. Создаем результирующий файл
    const resultFileId = createResultFile(processedData);
    console.log(`💾 Создан файл результата: ${resultFileId}`);
    
    return {
      success: true,
      sourceRows: sourceData.length,
      processedRows: processedData.length,
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
 * Получение исходных данных из Google Sheets
 */
function getSourceData() {
  // Замените на ID вашего Google Sheets
  const SHEET_ID = 'YOUR_SHEET_ID_HERE';
  
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    // Автоматический поиск листа с текущим месяцем
    const targetSheet = findCurrentMonthSheet(sheets);
    
    if (!targetSheet) {
      throw new Error('Не найден лист для текущего месяца');
    }
    
    console.log(`📋 Используется лист: ${targetSheet.getName()}`);
    
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
 * Поиск листа с текущим месяцем
 */
function findCurrentMonthSheet(sheets) {
  const currentMonth = new Date().getMonth();
  const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 
                     'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
  
  // Сначала ищем точное совпадение
  for (let sheet of sheets) {
    const sheetName = sheet.getName().toLowerCase();
    if (sheetName.includes(monthNames[currentMonth])) {
      return sheet;
    }
  }
  
  // Если не найдено, берем первый лист
  return sheets[0];
}

/**
 * Основная обработка данных
 */
function processData(rawData) {
  if (!rawData || rawData.length === 0) {
    throw new Error('Нет данных для обработки');
  }
  
  // 1. Поиск заголовков
  const headerInfo = findHeaders(rawData);
  console.log(`🔍 Найдена строка заголовков: ${headerInfo.row}`);
  
  // 2. Извлечение данных
  const dataRows = rawData.slice(headerInfo.row + 1);
  
  // 3. Обработка строк
  const processedRows = [];
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    
    // Пропускаем пустые строки
    if (isEmptyRow(row)) continue;
    
    // Пропускаем строки-заголовки
    if (isHeaderRow(row)) {
      console.log(`⏭️ Пропущена строка-заголовок: ${row[0]}`);
      continue;
    }
    
    // Определяем тип контента
    const contentType = determineContentType(row);
    
    if (contentType === 'review' || contentType === 'comment') {
      const processedRow = extractRowData(row, contentType, headerInfo.mapping);
      if (processedRow) {
        processedRows.push(processedRow);
      }
    }
  }
  
  // 4. Группировка по типам
  const reviews = processedRows.filter(row => row.type === 'review');
  const comments = processedRows.filter(row => row.type === 'comment');
  
  console.log(`📊 Найдено отзывов: ${reviews.length}, комментариев: ${comments.length}`);
  
  return processedRows;
}

/**
 * Поиск строки заголовков
 */
function findHeaders(data) {
  const columnMapping = {};
  
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    const rowStr = row.map(cell => (cell || '').toString().toLowerCase()).join(' ');
    
    if (rowStr.includes('тип размещения') || 
        rowStr.includes('площадка') || 
        rowStr.includes('текст сообщения')) {
      
      // Создаем маппинг колонок
      row.forEach((cell, index) => {
        const cleanHeader = (cell || '').toString().toLowerCase().trim();
        if (cleanHeader) {
          columnMapping[cleanHeader] = index;
        }
      });
      
      return {
        row: i,
        mapping: columnMapping
      };
    }
  }
  
  // Если заголовки не найдены, используем стандартный маппинг
  return {
    row: 0,
    mapping: getDefaultMapping()
  };
}

/**
 * Стандартный маппинг колонок для Google Sheets
 */
function getDefaultMapping() {
  return {
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
  
  // Дополнительная проверка по первой колонке
  const firstColValue = (row[0] || '').toString().toLowerCase().trim();
  if (firstColValue.includes('отзыв') || firstColValue.includes('ос')) {
    return 'review';
  }
  
  if (firstColValue.includes('комментарий') || firstColValue.includes('цс')) {
    return 'comment';
  }
  
  return 'unknown';
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
 * Проверка, является ли строка пустой
 */
function isEmptyRow(row) {
  return !row || row.every(cell => !cell || cell.toString().trim() === '');
}

/**
 * Извлечение данных из строки
 */
function extractRowData(row, type, mapping) {
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

/**
 * Создание результирующего файла
 */
function createResultFile(processedData) {
  try {
    // Создаем новый спредшит
    const currentDate = new Date();
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    
    const fileName = `Результат_${monthNames[currentDate.getMonth()]}_${currentDate.getFullYear()}`;
    const spreadsheet = SpreadsheetApp.create(fileName);
    
    // Получаем активный лист
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName('Результаты');
    
    // Заголовки
    const headers = ['Тип', 'Площадка', 'Ссылка', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Вовлечение'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Данные
    const dataForSheet = processedData.map(row => [
      row.type === 'review' ? 'Отзыв' : 'Комментарий',
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
 * Тестирование процессора
 */
function testProcessor() {
  console.log('🧪 Запуск тестирования...');
  
  try {
    const result = processGoogleSheets();
    
    if (result.success) {
      console.log('✅ Тестирование успешно завершено');
      console.log(`📊 Обработано строк: ${result.processedRows}`);
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
 * Анализ структуры данных
 */
function analyzeDataStructure() {
  console.log('🔍 Анализ структуры данных...');
  
  try {
    const sourceData = getSourceData();
    
    if (sourceData.length === 0) {
      console.log('❌ Нет данных для анализа');
      return;
    }
    
    // Анализируем первые 10 строк
    for (let i = 0; i < Math.min(10, sourceData.length); i++) {
      const row = sourceData[i];
      console.log(`Строка ${i + 1}:`, row.slice(0, 5).map(cell => 
        (cell || '').toString().substring(0, 30)
      ));
    }
    
    // Поиск заголовков
    const headerInfo = findHeaders(sourceData);
    console.log('📋 Найденные заголовки:', Object.keys(headerInfo.mapping));
    
    return {
      totalRows: sourceData.length,
      headerRow: headerInfo.row,
      columns: Object.keys(headerInfo.mapping)
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
  return processGoogleSheets();
}

/**
 * Функция для тестирования
 */
function runTest() {
  return testProcessor();
}

/**
 * Функция для анализа данных
 */
function runAnalysis() {
  return analyzeDataStructure();
}