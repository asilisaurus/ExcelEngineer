/**
 * 🚀 ФИНАЛЬНЫЙ УНИВЕРСАЛЬНЫЙ ПРОЦЕССОР V10
 * Исправлена классификация записей
 * 
 * Изменения в V10:
 * - Правильно обрабатывает "Отзывы (отзовики)", "Отзывы (аптеки)" и т.д.
 * - Улучшена логика определения типов записей
 * - Добавлена подробная статистика обработки
 */

// ==================== КОНФИГУРАЦИЯ ====================

var CONFIG = {
  STRUCTURE: {
    headerRow: 4,
    dataStartRow: 5,
    maxEmptyRows: 5
  },
  
  // Возможные колонки для просмотров
  VIEWS_COLUMNS: {
    primary: 12,    // M - Просмотров получено
    secondary: 11,  // L - Просмотры в конце месяца
    tertiary: 10    // K - Просмотры темы на старте
  },
  
  // Расширенные шаблоны для определения типов
  PATTERNS: {
    REVIEWS: ['отзыв', 'о т з ы в', 'отзовик', 'аптек'],
    COMMENTS: ['комментар', 'топ-20', 'топ 20', 'коммент'],
    DISCUSSIONS: ['обсужден', 'форум', 'сообщест', 'активные', 'мониторинг', 'дискусс']
  },
  
  FORMATTING: {
    DATE_FORMAT: 'dd.mm.yyyy',
    NUMBER_FORMAT: '#,##0'
  }
};

// ==================== ОСНОВНЫЕ ФУНКЦИИ ====================

/**
 * Главная функция обработки
 */
function processMonthlyReport() {
  try {
    console.log('🚀 FINAL PROCESSOR V10 - Начало обработки');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var sheetName = sheet.getName();
    
    // Проверяем, что это не служебный лист
    if (isServiceSheet(sheetName)) {
      showMessage('Ошибка', 'Выберите лист с данными месяца (например, Февраль25, Март25, Апрель25)');
      return;
    }
    
    console.log('📋 Обрабатывается лист: ' + sheetName);
    
    // Получение данных
    var data = sheet.getDataRange().getValues();
    console.log('📊 Загружено строк: ' + data.length);
    
    // Определение месяца
    var monthInfo = detectMonth(sheetName, data);
    console.log('📅 Определен месяц: ' + monthInfo.name + ' ' + monthInfo.year);
    
    // Автоматическое определение структуры
    var structure = analyzeStructure(data);
    console.log('🔍 Структура определена: колонка просмотров = ' + 
                (structure.hasViews ? String.fromCharCode(65 + structure.viewsColumn) : 'не найдена'));
    
    // Обработка данных
    var result = processDataFinal(data, structure);
    
    // Создание отчета
    var reportUrl = createReport(result, monthInfo);
    
    // Показ результата
    var message = 'Отчет создан успешно!\n\n' +
                  'Лист: ' + sheetName + '\n' +
                  'Ссылка: ' + reportUrl + '\n\n' +
                  'Обработано:\n' +
                  '- Отзывов: ' + result.statistics.totalReviews + '\n' +
                  '- Комментариев Топ-20: ' + result.statistics.totalCommentsTop20 + '\n' +
                  '- Активных обсуждений: ' + result.statistics.totalActiveDiscussions + '\n' +
                  '- Общие просмотры: ' + result.statistics.totalViews;
    
    if (result.statistics.totalViews === 0) {
      message += '\n\n⚠️ Внимание: просмотры не найдены или равны 0';
    }
    
    // Добавляем детали обработки
    if (result.debugInfo) {
      message += '\n\nДетали обработки:\n' + result.debugInfo;
    }
    
    console.log('✅ Обработка завершена успешно');
    showMessage('Обработка завершена', message);
    
  } catch (error) {
    console.error('❌ Ошибка: ' + error.toString());
    console.error(error.stack);
    showMessage('Ошибка', 'Произошла ошибка: ' + error.toString());
  }
}

/**
 * Анализ структуры данных
 */
function analyzeStructure(data) {
  var structure = {
    viewsColumn: -1,
    hasViews: false,
    columnMapping: {
      typeOfPlacement: 0,  // A
      platform: 1,         // B
      product: 2,          // C
      link: 3,             // D
      text: 4,             // E
      date: 6,             // G
      author: 7,           // H
      engagement: 13       // N
    }
  };
  
  console.log('🔍 Поиск колонки с просмотрами...');
  
  var possibleColumns = [
    CONFIG.VIEWS_COLUMNS.primary,
    CONFIG.VIEWS_COLUMNS.secondary,
    CONFIG.VIEWS_COLUMNS.tertiary
  ];
  
  for (var i = 0; i < possibleColumns.length; i++) {
    var colIndex = possibleColumns[i];
    var foundViews = false;
    
    for (var row = CONFIG.STRUCTURE.dataStartRow; row < Math.min(50, data.length); row++) {
      var value = data[row][colIndex];
      if (value && String(value).trim() !== '') {
        var strValue = String(value);
        if (strValue.match(/\d/) || strValue.indexOf('тыс') !== -1 || strValue.indexOf('млн') !== -1) {
          foundViews = true;
          break;
        }
      }
    }
    
    if (foundViews) {
      structure.viewsColumn = colIndex;
      structure.hasViews = true;
      console.log('✅ Найдена колонка просмотров: ' + String.fromCharCode(65 + colIndex));
      break;
    }
  }
  
  if (!structure.hasViews) {
    console.log('⚠️ Колонка просмотров не найдена, работаем без просмотров');
    structure.viewsColumn = CONFIG.VIEWS_COLUMNS.primary;
  }
  
  return structure;
}

/**
 * Финальная обработка данных с исправленной логикой
 */
function processDataFinal(data, structure) {
  var result = {
    reviews: [],
    commentsTop20: [],
    activeDiscussions: [],
    statistics: {
      totalReviews: 0,
      totalCommentsTop20: 0,
      totalActiveDiscussions: 0,
      totalViews: 0,
      engagementShare: 0
    },
    debugInfo: ''
  };
  
  var currentSection = null;
  var allTargetedRecords = [];
  var emptyRowCount = 0;
  var processedRows = 0;
  var skippedRows = 0;
  
  // Обрабатываем все строки
  for (var i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
    var row = data[i];
    
    // Проверяем пустую строку
    if (isEmptyRow(row)) {
      emptyRowCount++;
      if (emptyRowCount >= CONFIG.STRUCTURE.maxEmptyRows) {
        console.log('🛑 Остановка: найдено ' + emptyRowCount + ' пустых строк подряд');
        break;
      }
      continue;
    } else {
      emptyRowCount = 0;
    }
    
    // Останавливаемся на статистике
    if (isStatisticsRow(row)) {
      console.log('📊 Найден блок статистики, остановка обработки');
      break;
    }
    
    var firstCell = row[0] ? String(row[0]).trim() : '';
    var firstCellLower = firstCell.toLowerCase();
    
    // Проверяем, не заголовок ли это раздела
    if (isSectionHeader(firstCellLower)) {
      if (matchesPattern(firstCellLower, CONFIG.PATTERNS.REVIEWS)) {
        currentSection = 'reviews';
        console.log('📂 Заголовок раздела: Отзывы (строка ' + (i + 1) + ')');
      } else if (matchesPattern(firstCellLower, CONFIG.PATTERNS.COMMENTS)) {
        currentSection = 'comments';
        console.log('📂 Заголовок раздела: Комментарии (строка ' + (i + 1) + ')');
      } else if (matchesPattern(firstCellLower, CONFIG.PATTERNS.DISCUSSIONS)) {
        currentSection = 'discussions';
        console.log('📂 Заголовок раздела: Обсуждения (строка ' + (i + 1) + ')');
      }
      continue;
    }
    
    // Получаем данные из строки
    var typeOfPlacement = firstCell;
    var platform = row[structure.columnMapping.platform] ? String(row[structure.columnMapping.platform]).trim() : '';
    var text = row[structure.columnMapping.text] ? String(row[structure.columnMapping.text]).trim() : '';
    
    // Пропускаем строки без значимых данных
    if (!platform && !text && !typeOfPlacement) {
      skippedRows++;
      continue;
    }
    
    // Создаем запись
    var processedRow = {
      platform: platform,
      theme: row[structure.columnMapping.product] ? String(row[structure.columnMapping.product]).trim() : '',
      text: text,
      date: extractDate(row, structure.columnMapping),
      author: row[structure.columnMapping.author] ? String(row[structure.columnMapping.author]).trim() : '',
      views: extractViews(row, structure.viewsColumn),
      engagement: row[structure.columnMapping.engagement] ? String(row[structure.columnMapping.engagement]).trim() : '',
      postType: typeOfPlacement
    };
    
    processedRows++;
    
    // ИСПРАВЛЕННАЯ ЛОГИКА КЛАССИФИКАЦИИ
    // Проверяем тип по содержимому колонки A
    var isReview = false;
    var isTargeted = false;
    
    // Проверяем, является ли это отзывом
    if (currentSection === 'reviews' || 
        matchesPattern(firstCellLower, CONFIG.PATTERNS.REVIEWS) ||
        firstCellLower.indexOf('отзывы (') !== -1 ||  // "Отзывы (отзовики)", "Отзывы (аптеки)"
        firstCellLower.indexOf('отзыв') !== -1) {
      isReview = true;
    }
    // Проверяем, является ли это комментарием/обсуждением
    else if (currentSection === 'comments' || 
             currentSection === 'discussions' ||
             matchesPattern(firstCellLower, CONFIG.PATTERNS.COMMENTS) ||
             matchesPattern(firstCellLower, CONFIG.PATTERNS.DISCUSSIONS)) {
      isTargeted = true;
    }
    // Если ничего не подошло, но есть данные - считаем целевым
    else if (platform || text) {
      isTargeted = true;
    }
    
    // Распределяем записи
    if (isReview) {
      result.reviews.push(processedRow);
      result.statistics.totalReviews++;
    } else if (isTargeted) {
      allTargetedRecords.push(processedRow);
    }
  }
  
  // Сортируем целевые записи по просмотрам
  allTargetedRecords.sort(function(a, b) {
    return (b.views || 0) - (a.views || 0);
  });
  
  // Распределяем: первые 20 - топ комментарии, остальные - обсуждения
  for (var j = 0; j < allTargetedRecords.length; j++) {
    if (j < 20) {
      result.commentsTop20.push(allTargetedRecords[j]);
      result.statistics.totalCommentsTop20++;
    } else {
      result.activeDiscussions.push(allTargetedRecords[j]);
      result.statistics.totalActiveDiscussions++;
    }
  }
  
  // Считаем общие просмотры
  var allRecords = result.reviews.concat(result.commentsTop20).concat(result.activeDiscussions);
  var totalViews = 0;
  for (var k = 0; k < allRecords.length; k++) {
    if (allRecords[k].views > 0) {
      totalViews += allRecords[k].views;
    }
  }
  result.statistics.totalViews = totalViews;
  
  // Считаем долю вовлечения
  var discussionsAll = result.commentsTop20.concat(result.activeDiscussions);
  result.statistics.engagementShare = calculateEngagementRate(discussionsAll);
  
  // Формируем отладочную информацию
  result.debugInfo = 'Обработано строк: ' + processedRows + '\n' +
                     'Пропущено пустых: ' + skippedRows;
  
  console.log('📊 Итоги обработки:');
  console.log('   - Обработано строк: ' + processedRows);
  console.log('   - Отзывов: ' + result.statistics.totalReviews);
  console.log('   - Целевых записей: ' + allTargetedRecords.length);
  console.log('   - Комментариев топ-20: ' + result.statistics.totalCommentsTop20);
  console.log('   - Обсуждений: ' + result.statistics.totalActiveDiscussions);
  console.log('   - Всего просмотров: ' + result.statistics.totalViews);
  
  return result;
}

/**
 * Создание отчета
 */
function createReport(processedData, monthInfo) {
  var reportName = 'Report_' + monthInfo.name + '_' + monthInfo.year + '_' + new Date().getTime();
  var newSpreadsheet = SpreadsheetApp.create(reportName);
  var sheet = newSpreadsheet.getActiveSheet();
  sheet.setName(monthInfo.name + '_' + monthInfo.year);
  
  // Шапка отчета
  sheet.getRange('A1').setValue('Продукт');
  sheet.getRange('B1').setValue('Акрихин - Фортедетрим');
  sheet.getRange('A2').setValue('Период');
  sheet.getRange('B2').setValue(monthInfo.name + '-25');
  sheet.getRange('A3').setValue('План');
  
  // Заголовки таблицы
  var headers = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
  var row = 5;
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#3f2355').setFontColor('white');
  row++;
  
  // Записываем разделы
  row = writeSection(sheet, row, 'Отзывы', processedData.reviews, headers.length);
  row = writeSection(sheet, row, 'Комментарии Топ-20 выдачи', processedData.commentsTop20, headers.length);
  row = writeSection(sheet, row, 'Активные обсуждения (мониторинг)', processedData.activeDiscussions, headers.length);
  
  // Блок статистики
  row += 2;
  sheet.getRange(row, 1).setValue('Суммарное количество просмотров');
  sheet.getRange(row, 2).setValue(processedData.statistics.totalViews || 0);
  row++;
  sheet.getRange(row, 1).setValue('Количество карточек товара (отзывы)');
  sheet.getRange(row, 2).setValue(processedData.statistics.totalReviews || 0);
  row++;
  sheet.getRange(row, 1).setValue('Количество обсуждений (форумы, сообщества, комментарии к статьям)');
  var totalDiscussions = processedData.statistics.totalActiveDiscussions + processedData.statistics.totalCommentsTop20;
  sheet.getRange(row, 2).setValue(totalDiscussions);
  row++;
  sheet.getRange(row, 1).setValue('Доля обсуждений с вовлечением в диалог');
  sheet.getRange(row, 2).setValue(processedData.statistics.engagementShare);
  sheet.getRange(row, 2).setNumberFormat("0%");
  
  // Форматирование
  sheet.autoResizeColumns(1, headers.length);
  
  // Добавляем примечание если нет просмотров
  if (processedData.statistics.totalViews === 0) {
    sheet.getRange(row - 3, 3).setValue('(Данные о просмотрах не заполнены в исходном файле)');
  }
  
  console.log('📄 Отчет создан: ' + reportName);
  return newSpreadsheet.getUrl();
}

// ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Проверка служебного листа
 */
function isServiceSheet(sheetName) {
  var lowerName = sheetName.toLowerCase();
  return lowerName.indexOf('инструкция') !== -1 || 
         lowerName.indexOf('диагностика') !== -1 ||
         lowerName.indexOf('шаблон') !== -1 ||
         lowerName.indexOf('template') !== -1;
}

/**
 * Проверка заголовка раздела
 */
function isSectionHeader(text) {
  if (text.length > 100) return false;
  
  // Проверяем точные заголовки разделов
  if (text === 'о т з ы в ы' || 
      text === 'отзывы' ||
      text === 'комментарии топ-20 выдачи' ||
      text === 'активные обсуждения (мониторинг)') {
    return true;
  }
  
  return false;
}

/**
 * Проверка соответствия шаблону
 */
function matchesPattern(text, patterns) {
  for (var i = 0; i < patterns.length; i++) {
    if (text.indexOf(patterns[i]) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * Определение месяца
 */
function detectMonth(sheetName, data) {
  var monthFromSheet = extractMonthFromText(sheetName);
  if (monthFromSheet) {
    return monthFromSheet;
  }
  
  for (var i = 0; i < Math.min(3, data.length); i++) {
    var rowText = data[i].join(' ');
    var monthFromData = extractMonthFromText(rowText);
    if (monthFromData) {
      return monthFromData;
    }
  }
  
  var now = new Date();
  var monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  return {
    name: monthNames[now.getMonth()],
    number: now.getMonth() + 1,
    year: now.getFullYear()
  };
}

/**
 * Извлечение месяца из текста
 */
function extractMonthFromText(text) {
  var lowerText = text.toLowerCase();
  
  var months = [
    { name: 'Январь', patterns: ['январ', 'янв'], number: 1 },
    { name: 'Февраль', patterns: ['феврал', 'фев'], number: 2 },
    { name: 'Март', patterns: ['март', 'мар'], number: 3 },
    { name: 'Апрель', patterns: ['апрел', 'апр'], number: 4 },
    { name: 'Май', patterns: ['май', 'мая'], number: 5 },
    { name: 'Июнь', patterns: ['июн'], number: 6 },
    { name: 'Июль', patterns: ['июл'], number: 7 },
    { name: 'Август', patterns: ['август', 'авг'], number: 8 },
    { name: 'Сентябрь', patterns: ['сентябр', 'сен'], number: 9 },
    { name: 'Октябрь', patterns: ['октябр', 'окт'], number: 10 },
    { name: 'Ноябрь', patterns: ['ноябр', 'ноя'], number: 11 },
    { name: 'Декабрь', patterns: ['декабр', 'дек'], number: 12 }
  ];
  
  for (var i = 0; i < months.length; i++) {
    var month = months[i];
    for (var j = 0; j < month.patterns.length; j++) {
      if (lowerText.indexOf(month.patterns[j]) !== -1) {
        var year = 2025;
        var yearMatch = text.match(/20\d{2}/);
        if (yearMatch) {
          year = parseInt(yearMatch[0]);
        } else if (text.match(/\d{2}/)) {
          var shortYear = text.match(/\d{2}/)[0];
          year = 2000 + parseInt(shortYear);
        }
        
        return {
          name: month.name,
          number: month.number,
          year: year
        };
      }
    }
  }
  
  return null;
}

/**
 * Проверка пустой строки
 */
function isEmptyRow(row) {
  if (!row) return true;
  for (var i = 0; i < row.length; i++) {
    if (row[i] && String(row[i]).trim() !== '') return false;
  }
  return true;
}

/**
 * Проверка строки статистики
 */
function isStatisticsRow(row) {
  if (!row || row.length === 0) return false;
  var firstCell = String(row[0] || '').toLowerCase();
  return firstCell.indexOf('суммарное количество просмотров') !== -1 || 
         firstCell.indexOf('количество карточек товара') !== -1 ||
         firstCell.indexOf('количество обсуждений') !== -1 ||
         firstCell.indexOf('доля обсуждений') !== -1;
}

/**
 * Извлечение даты
 */
function extractDate(row, columnMapping) {
  var index = columnMapping.date;
  if (index !== undefined && row[index]) {
    var dateValue = row[index];
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), CONFIG.FORMATTING.DATE_FORMAT);
    }
    return String(dateValue);
  }
  return '';
}

/**
 * Извлечение просмотров
 */
function extractViews(row, viewsColumnIndex) {
  if (viewsColumnIndex === undefined || viewsColumnIndex < 0) return 0;
  
  var viewsValue = row[viewsColumnIndex];
  if (!viewsValue) return 0;
  
  if (typeof viewsValue === 'number' && !isNaN(viewsValue)) {
    return Math.max(0, Math.floor(viewsValue));
  }
  
  var viewsStr = String(viewsValue).trim();
  if (viewsStr === '') return 0;
  
  if (viewsStr.indexOf('тыс') !== -1) {
    var num = parseFloat(viewsStr.replace(/[^\d.,]/g, '').replace(',', '.'));
    if (!isNaN(num)) {
      return Math.floor(num * 1000);
    }
  }
  
  if (viewsStr.indexOf('млн') !== -1) {
    var numMln = parseFloat(viewsStr.replace(/[^\d.,]/g, '').replace(',', '.'));
    if (!isNaN(numMln)) {
      return Math.floor(numMln * 1000000);
    }
  }
  
  var cleanStr = viewsStr.replace(/[^\d.,]/g, '');
  var normalizedStr = cleanStr.replace(',', '.');
  var parsed = parseFloat(normalizedStr);
  
  if (!isNaN(parsed)) {
    return Math.max(0, Math.floor(parsed));
  }
  
  return 0;
}

/**
 * Расчет доли вовлечения
 */
function calculateEngagementRate(discussions) {
  if (discussions.length === 0) return 0;
  
  var withEngagement = 0;
  for (var i = 0; i < discussions.length; i++) {
    var eng = discussions[i].engagement;
    if (eng && eng !== '0' && eng !== '' && eng !== '-') {
      withEngagement++;
    }
  }
  
  return withEngagement / discussions.length;
}

/**
 * Запись раздела в отчет
 */
function writeSection(sheet, startRow, sectionName, dataArr, colCount) {
  sheet.getRange(startRow, 1).setValue(sectionName);
  sheet.getRange(startRow, 1, 1, colCount).setBackground('#b7a6c9').setFontWeight('bold');
  startRow++;
  
  if (dataArr.length > 0) {
    var rows = [];
    for (var i = 0; i < dataArr.length; i++) {
      var r = dataArr[i];
      rows.push([
        r.platform || '',
        r.theme || '',
        r.text || '',
        r.date || '',
        r.author || '',
        r.views || 0,
        r.engagement || '',
        r.postType || ''
      ]);
    }
    sheet.getRange(startRow, 1, rows.length, colCount).setValues(rows);
    startRow += rows.length;
  }
  
  console.log('📂 Раздел "' + sectionName + '": ' + dataArr.length + ' строк');
  return startRow;
}

/**
 * Показ сообщения
 */
function showMessage(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.log(title + ': ' + message);
  }
}

/**
 * Создание меню
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Финальный процессор V10')
    .addItem('🚀 Обработать текущий лист', 'processMonthlyReport')
    .addItem('🔍 Проверить типы записей', 'checkRecordTypes')
    .addItem('📋 Анализ структуры', 'analyzeCurrentSheet')
    .addToUi();
}

/**
 * Проверка типов записей в текущем листе
 */
function checkRecordTypes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var types = {};
  var examples = {};
  
  for (var i = 4; i < Math.min(100, data.length); i++) {
    var type = data[i][0] ? String(data[i][0]).trim() : '';
    if (type) {
      types[type] = (types[type] || 0) + 1;
      if (!examples[type] && data[i][1]) {
        examples[type] = {
          row: i + 1,
          platform: data[i][1],
          text: data[i][4] ? String(data[i][4]).substring(0, 30) : ''
        };
      }
    }
  }
  
  var message = 'Найденные типы записей:\n\n';
  for (var t in types) {
    message += t + ': ' + types[t] + ' записей\n';
    if (examples[t]) {
      message += '  Пример (строка ' + examples[t].row + '):\n';
      message += '  ' + examples[t].platform + '\n\n';
    }
  }
  
  showMessage('Типы записей', message);
}

/**
 * Анализ текущего листа
 */
function analyzeCurrentSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getName();
  var data = sheet.getDataRange().getValues();
  
  var structure = analyzeStructure(data);
  
  var message = 'Анализ листа: ' + sheetName + '\n\n' +
                'Строк: ' + data.length + '\n' +
                'Колонок: ' + (data[0] ? data[0].length : 0) + '\n\n' +
                'Колонка просмотров: ';
  
  if (structure.hasViews) {
    message += String.fromCharCode(65 + structure.viewsColumn) + ' (найдены данные)';
  } else {
    message += 'не найдена (будет обработано без просмотров)';
  }
  
  message += '\n\nПроцессор готов к работе!';
  
  showMessage('Анализ структуры', message);
}