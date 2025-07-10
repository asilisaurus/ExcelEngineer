/**
 * ============================================================================
 * ТЕСТИРОВАНИЕ СКРИПТА НА ЛИСТЕ "МАЙ"
 * ============================================================================
 * Специальный скрипт для тестирования процессора на данных из листа "Май"
 * Эталонная таблица: 1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk
 * ============================================================================
 */

// =============================================================================
// КОНФИГУРАЦИЯ
// =============================================================================

// ID эталонной таблицы с 4 месяцами
const REFERENCE_SHEET_ID = '1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk';

// Название целевого листа
const TARGET_SHEET_NAME = 'Май';

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ ТЕСТИРОВАНИЯ
// =============================================================================

/**
 * Главная функция тестирования листа "Май"
 */
function testMaySheetProcessor() {
  console.log('🧪 Запуск тестирования скрипта на листе "Май"...');
  
  try {
    // 1. Получаем данные из листа "Май"
    const mayData = getMaySheetData();
    console.log(`📊 Загружено ${mayData.length} строк из листа "Май"`);
    
    // 2. Анализируем структуру данных
    const structureAnalysis = analyzeMaySheetStructure(mayData);
    console.log('📋 Структура данных проанализирована');
    
    // 3. Обрабатываем данные
    const processedData = processMaySheetData(mayData);
    console.log(`✅ Обработано ${processedData.length} записей`);
    
    // 4. Анализируем результаты
    const results = analyzeMayResults(processedData);
    console.log('📊 Результаты проанализированы');
    
    // 5. Создаем отчет
    const testReport = createMayTestReport(mayData, processedData, results);
    
    // 6. Создаем результирующий файл
    const resultFileId = createMayResultFile(processedData, results, testReport);
    console.log(`💾 Создан файл результата: ${resultFileId}`);
    
    return {
      success: true,
      sourceRows: mayData.length,
      processedRows: processedData.length,
      results: results,
      testReport: testReport,
      resultFileId: resultFileId
    };
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Получение данных из листа "Май"
 */
function getMaySheetData() {
  try {
    console.log('🔍 Получаем данные из листа "Май"...');
    
    const spreadsheet = SpreadsheetApp.openById(REFERENCE_SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    console.log(`📋 Найдено ${sheets.length} листов в таблице`);
    
    // Ищем лист "Май"
    let maySheet = null;
    for (const sheet of sheets) {
      const sheetName = sheet.getName().toLowerCase();
      console.log(`📋 Проверяем лист: ${sheet.getName()}`);
      
      if (sheetName.includes('май') || sheetName.includes('may')) {
        maySheet = sheet;
        console.log(`✅ Найден лист "Май": ${sheet.getName()}`);
        break;
      }
    }
    
    if (!maySheet) {
      throw new Error('Лист "Май" не найден в таблице');
    }
    
    // Получаем все данные листа
    const range = maySheet.getDataRange();
    const values = range.getValues();
    
    console.log(`📊 Загружено ${values.length} строк из листа "${maySheet.getName()}"`);
    
    return values;
    
  } catch (error) {
    console.error('❌ Ошибка при получении данных:', error);
    throw error;
  }
}

/**
 * Анализ структуры данных листа "Май"
 */
function analyzeMaySheetStructure(data) {
  console.log('🔍 Анализируем структуру данных листа "Май"...');
  
  if (!data || data.length === 0) {
    throw new Error('Нет данных для анализа');
  }
  
  const analysis = {
    totalRows: data.length,
    headers: [],
    headerRow: 0,
    dataRowsCount: 0,
    emptyRowsCount: 0,
    sampleData: []
  };
  
  // Поиск заголовков
  const headerInfo = findMayHeaders(data);
  analysis.headers = headerInfo.headers;
  analysis.headerRow = headerInfo.row;
  
  console.log(`📋 Найдены заголовки в строке ${analysis.headerRow + 1}:`);
  headerInfo.headers.forEach((header, index) => {
    if (header) {
      console.log(`   ${index + 1}. ${header}`);
    }
  });
  
  // Анализ данных
  const dataRows = data.slice(headerInfo.row + 1);
  
  for (const row of dataRows) {
    if (isEmptyRow(row)) {
      analysis.emptyRowsCount++;
    } else {
      analysis.dataRowsCount++;
      
      // Добавляем образцы данных
      if (analysis.sampleData.length < 10) {
        analysis.sampleData.push({
          rowData: row.slice(0, 6).map(cell => (cell || '').toString().substring(0, 30)),
          contentType: determineContentType(row)
        });
      }
    }
  }
  
  console.log(`📊 Анализ завершен:`);
  console.log(`   - Всего строк: ${analysis.totalRows}`);
  console.log(`   - Строк с данными: ${analysis.dataRowsCount}`);
  console.log(`   - Пустых строк: ${analysis.emptyRowsCount}`);
  console.log(`   - Строка заголовков: ${analysis.headerRow + 1}`);
  
  return analysis;
}

/**
 * Поиск заголовков в данных листа "Май"
 */
function findMayHeaders(data) {
  console.log('🔍 Ищем заголовки в данных...');
  
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    
    const rowStr = row.map(cell => (cell || '').toString().toLowerCase()).join(' ');
    
    // Поиск ключевых заголовков
    if (rowStr.includes('тип размещения') || 
        rowStr.includes('площадка') || 
        rowStr.includes('текст сообщения') ||
        rowStr.includes('тип поста')) {
      
      console.log(`✅ Найдены заголовки в строке ${i + 1}`);
      return {
        row: i,
        headers: row.map(cell => (cell || '').toString().trim())
      };
    }
  }
  
  // Если заголовки не найдены, используем первую строку
  console.log('⚠️ Ключевые заголовки не найдены, используем первую строку');
  return {
    row: 0,
    headers: data[0] ? data[0].map(cell => (cell || '').toString().trim()) : []
  };
}

/**
 * Обработка данных листа "Май"
 */
function processMaySheetData(rawData) {
  console.log('⚙️ Обрабатываем данные листа "Май"...');
  
  if (!rawData || rawData.length === 0) {
    throw new Error('Нет данных для обработки');
  }
  
  // 1. Поиск заголовков
  const headerInfo = findMayHeaders(rawData);
  
  // 2. Создание маппинга колонок
  const columnMapping = createMayColumnMapping(headerInfo.headers);
  
  // 3. Обработка строк данных
  const dataRows = rawData.slice(headerInfo.row + 1);
  const processedRows = [];
  let skippedRows = 0;
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    
    // Пропускаем пустые строки
    if (isEmptyRow(row)) {
      skippedRows++;
      continue;
    }
    
    // Пропускаем строки-заголовки секций
    if (isSectionHeaderRow(row)) {
      console.log(`⏭️ Пропущена строка-заголовок: ${row[0]}`);
      skippedRows++;
      continue;
    }
    
    // Определяем тип контента
    const contentType = determineContentType(row);
    
    if (contentType !== 'unknown') {
      const processedRow = extractRowData(row, contentType, columnMapping);
      if (processedRow) {
        processedRows.push(processedRow);
      }
    } else {
      skippedRows++;
    }
  }
  
  console.log(`📊 Обработка завершена:`);
  console.log(`   - Обработано строк: ${processedRows.length}`);
  console.log(`   - Пропущено строк: ${skippedRows}`);
  
  return processedRows;
}

/**
 * Создание маппинга колонок для листа "Май"
 */
function createMayColumnMapping(headers) {
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
  
  return { ...defaultMapping, ...columnMapping };
}

/**
 * Определение типа контента
 */
function determineContentType(row) {
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
  
  return 'unknown';
}

/**
 * Проверка, является ли строка заголовком секции
 */
function isSectionHeaderRow(row) {
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
 * Извлечение данных из строки
 */
function extractRowData(row, type, mapping) {
  try {
    // Индексы колонок
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
    
    // Валидация
    if (!site && !text) {
      return null;
    }
    
    return {
      type: type,
      site: site || 'Неизвестная площадка',
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
 * Анализ результатов обработки
 */
function analyzeMayResults(processedData) {
  console.log('📊 Анализируем результаты обработки...');
  
  const reviews = processedData.filter(row => row.type === 'review');
  const comments = processedData.filter(row => row.type === 'comment');
  const discussions = processedData.filter(row => row.type === 'discussion');
  
  const results = {
    total: processedData.length,
    reviews: {
      count: reviews.length,
      percentage: Math.round((reviews.length / processedData.length) * 100)
    },
    comments: {
      count: comments.length,
      percentage: Math.round((comments.length / processedData.length) * 100)
    },
    discussions: {
      count: discussions.length,
      percentage: Math.round((discussions.length / processedData.length) * 100)
    }
  };
  
  console.log(`📊 Результаты анализа:`);
  console.log(`   - Всего записей: ${results.total}`);
  console.log(`   - Отзывы: ${results.reviews.count} (${results.reviews.percentage}%)`);
  console.log(`   - Комментарии: ${results.comments.count} (${results.comments.percentage}%)`);
  console.log(`   - Обсуждения: ${results.discussions.count} (${results.discussions.percentage}%)`);
  
  return results;
}

/**
 * Создание отчета о тестировании
 */
function createMayTestReport(sourceData, processedData, results) {
  console.log('📋 Создаем отчет о тестировании...');
  
  const report = {
    testDate: new Date().toISOString(),
    sourceInfo: {
      totalRows: sourceData.length,
      processedRows: processedData.length,
      processingRate: Math.round((processedData.length / sourceData.length) * 100)
    },
    results: results,
    quality: {
      dataIntegrity: processedData.every(row => row.site && row.text),
      typeClassification: results.reviews.count > 0 && results.comments.count > 0,
      overallScore: 0
    }
  };
  
  // Вычисляем общий балл качества
  let qualityScore = 0;
  if (report.quality.dataIntegrity) qualityScore += 50;
  if (report.quality.typeClassification) qualityScore += 30;
  if (report.sourceInfo.processingRate > 80) qualityScore += 20;
  
  report.quality.overallScore = qualityScore;
  
  console.log(`📊 Отчет создан:`);
  console.log(`   - Обработано: ${report.sourceInfo.processingRate}%`);
  console.log(`   - Качество данных: ${report.quality.dataIntegrity ? 'Хорошее' : 'Требует улучшения'}`);
  console.log(`   - Классификация типов: ${report.quality.typeClassification ? 'Работает' : 'Не работает'}`);
  console.log(`   - Общий балл: ${report.quality.overallScore}/100`);
  
  return report;
}

/**
 * Создание результирующего файла
 */
function createMayResultFile(processedData, results, testReport) {
  try {
    console.log('💾 Создаем результирующий файл...');
    
    const fileName = `Тест_Лист_Май_${new Date().toISOString().split('T')[0]}`;
    const spreadsheet = SpreadsheetApp.create(fileName);
    
    // Лист 1: Обработанные данные
    const dataSheet = spreadsheet.getActiveSheet();
    dataSheet.setName('Обработанные данные');
    
    const headers = ['Тип', 'Площадка', 'Ссылка', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Вовлечение'];
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    if (processedData.length > 0) {
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
      
      dataSheet.getRange(2, 1, dataForSheet.length, headers.length).setValues(dataForSheet);
    }
    
    // Форматирование
    dataSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    dataSheet.autoResizeColumns(1, headers.length);
    
    // Лист 2: Статистика
    const statsSheet = spreadsheet.insertSheet('Статистика');
    
    const statsData = [
      ['Параметр', 'Значение'],
      ['Всего записей', results.total],
      ['Отзывы', results.reviews.count],
      ['Комментарии', results.comments.count],
      ['Обсуждения', results.discussions.count],
      ['Общий балл качества', `${testReport.quality.overallScore}/100`],
      ['Целостность данных', testReport.quality.dataIntegrity ? 'Да' : 'Нет'],
      ['Классификация типов', testReport.quality.typeClassification ? 'Да' : 'Нет']
    ];
    
    statsSheet.getRange(1, 1, statsData.length, 2).setValues(statsData);
    statsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    statsSheet.autoResizeColumns(1, 2);
    
    console.log(`💾 Файл создан: ${fileName}`);
    console.log(`🔗 ID файла: ${spreadsheet.getId()}`);
    
    return spreadsheet.getId();
    
  } catch (error) {
    console.error('❌ Ошибка при создании файла:', error);
    throw error;
  }
}

// =============================================================================
// ФУНКЦИИ ДЛЯ ЗАПУСКА
// =============================================================================

/**
 * Основная функция для запуска тестирования
 */
function main() {
  return testMaySheetProcessor();
}

/**
 * Быстрый анализ структуры листа "Май"
 */
function quickMayAnalysis() {
  console.log('⚡ Быстрый анализ структуры листа "Май"...');
  
  try {
    const mayData = getMaySheetData();
    const structureAnalysis = analyzeMaySheetStructure(mayData);
    
    console.log('✅ Анализ завершен');
    return structureAnalysis;
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
    return { error: error.message };
  }
}

/**
 * Только обработка данных без создания файла
 */
function processMayDataOnly() {
  console.log('⚙️ Обработка данных листа "Май"...');
  
  try {
    const mayData = getMaySheetData();
    const processedData = processMaySheetData(mayData);
    const results = analyzeMayResults(processedData);
    
    console.log('✅ Обработка завершена');
    return { processedData, results };
    
  } catch (error) {
    console.error('❌ Ошибка при обработке:', error);
    return { error: error.message };
  }
}