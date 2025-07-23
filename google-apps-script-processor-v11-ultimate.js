/**
 * 🚀 ОКОНЧАТЕЛЬНЫЙ УНИВЕРСАЛЬНЫЙ ПРОЦЕССОР V11
 * Правильно работает со структурой где в колонке A - URL вместо типов
 * 
 * Изменения в V11:
 * - Отслеживает текущий раздел по заголовкам
 * - Правильно классифицирует записи на основе текущего раздела
 * - Работает когда в колонке A находятся URL-адреса
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
  
  // Точные заголовки разделов
  SECTION_HEADERS: {
    REVIEWS: ['отзывы', 'о т з ы в ы'],
    COMMENTS_TOP20: ['комментарии топ-20 выдачи', 'комментарии топ-20', 'комментарии топ 20'],
    DISCUSSIONS: ['активные обсуждения (мониторинг)', 'активные обсуждения', 'обсуждения']
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
    console.log('🚀 ULTIMATE PROCESSOR V11 - Начало обработки');
    
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
    
    // Обработка данных с отслеживанием разделов
    var result = processDataWithSections(data, structure);
    
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
 * Обработка данных с правильным отслеживанием разделов
 */
function processDataWithSections(data, structure) {
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
  var sectionRanges = [];
  var emptyRowCount = 0;
  var processedRows = 0;
  
  // Сначала найдем все заголовки разделов и их границы
  console.log('🔍 Поиск заголовков разделов...');
  
  for (var i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
    var row = data[i];
    
    // Проверяем пустую строку
    if (isEmptyRow(row)) {
      emptyRowCount++;
      if (emptyRowCount >= CONFIG.STRUCTURE.maxEmptyRows) {
        break;
      }
      continue;
    } else {
      emptyRowCount = 0;
    }
    
    // Проверяем на статистику
    if (isStatisticsRow(row)) {
      break;
    }
    
    var firstCell = row[0] ? String(row[0]).trim().toLowerCase() : '';
    
    // Проверяем заголовки разделов
    var sectionType = detectSectionHeader(firstCell);
    if (sectionType) {
      console.log('📂 Найден заголовок раздела "' + sectionType + '" в строке ' + (i + 1));
      if (sectionRanges.length > 0) {
        // Закрываем предыдущий раздел
        sectionRanges[sectionRanges.length - 1].endRow = i - 1;
      }
      sectionRanges.push({
        type: sectionType,
        startRow: i + 1, // Данные начинаются со следующей строки
        endRow: -1
      });
    }
  }
  
  // Закрываем последний раздел
  if (sectionRanges.length > 0) {
    sectionRanges[sectionRanges.length - 1].endRow = data.length - 1;
  }
  
  console.log('📊 Найдено разделов: ' + sectionRanges.length);
  
  // Обрабатываем каждый раздел
  for (var s = 0; s < sectionRanges.length; s++) {
    var section = sectionRanges[s];
    console.log('🔄 Обработка раздела "' + section.type + '" (строки ' + (section.startRow + 1) + '-' + (section.endRow + 1) + ')');
    
    for (var i = section.startRow; i <= section.endRow && i < data.length; i++) {
      var row = data[i];
      
      // Пропускаем пустые строки
      if (isEmptyRow(row)) continue;
      
      // Пропускаем статистику
      if (isStatisticsRow(row)) break;
      
      // Пропускаем заголовки других разделов
      var firstCellCheck = row[0] ? String(row[0]).trim().toLowerCase() : '';
      if (detectSectionHeader(firstCellCheck)) continue;
      
      // Получаем данные из строки
      var platform = row[structure.columnMapping.platform] ? String(row[structure.columnMapping.platform]).trim() : '';
      var text = row[structure.columnMapping.text] ? String(row[structure.columnMapping.text]).trim() : '';
      
      // Пропускаем строки без значимых данных
      if (!platform && !text) continue;
      
      // Создаем запись
      var processedRow = {
        platform: platform,
        theme: row[structure.columnMapping.product] ? String(row[structure.columnMapping.product]).trim() : '',
        text: text,
        date: extractDate(row, structure.columnMapping),
        author: row[structure.columnMapping.author] ? String(row[structure.columnMapping.author]).trim() : '',
        views: extractViews(row, structure.viewsColumn),
        engagement: row[structure.columnMapping.engagement] ? String(row[structure.columnMapping.engagement]).trim() : '',
        postType: row[0] ? String(row[0]).trim() : '' // Сохраняем оригинальное значение из колонки A
      };
      
      processedRows++;
      
      // Распределяем по разделам на основе текущего раздела
      if (section.type === 'reviews') {
        result.reviews.push(processedRow);
        result.statistics.totalReviews++;
      } else if (section.type === 'comments') {
        result.commentsTop20.push(processedRow);
        result.statistics.totalCommentsTop20++;
      } else if (section.type === 'discussions') {
        result.activeDiscussions.push(processedRow);
        result.statistics.totalActiveDiscussions++;
      }
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
  result.debugInfo = 'Обработано строк: ' + processedRows + '\n';
  result.debugInfo += 'Найдено разделов: ' + sectionRanges.length;
  
  console.log('📊 Итоги обработки:');
  console.log('   - Обработано строк: ' + processedRows);
  console.log('   - Отзывов: ' + result.statistics.totalReviews);
  console.log('   - Комментариев топ-20: ' + result.statistics.totalCommentsTop20);
  console.log('   - Обсуждений: ' + result.statistics.totalActiveDiscussions);
  console.log('   - Всего просмотров: ' + result.statistics.totalViews);
  
  return result;
}

/**
 * Определение типа заголовка раздела
 */
function detectSectionHeader(text) {
  // Проверяем заголовки отзывов
  for (var i = 0; i < CONFIG.SECTION_HEADERS.REVIEWS.length; i++) {
    if (text === CONFIG.SECTION_HEADERS.REVIEWS[i]) {
      return 'reviews';
    }
  }
  
  // Проверяем заголовки комментариев
  for (var j = 0; j < CONFIG.SECTION_HEADERS.COMMENTS_TOP20.length; j++) {
    if (text === CONFIG.SECTION_HEADERS.COMMENTS_TOP20[j]) {
      return 'comments';
    }
  }
  
  // Проверяем заголовки обсуждений
  for (var k = 0; k < CONFIG.SECTION_HEADERS.DISCUSSIONS.length; k++) {
    if (text === CONFIG.SECTION_HEADERS.DISCUSSIONS[k]) {
      return 'discussions';
    }
  }
  
  return null;
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
    .createMenu('📊 Процессор V11 Ultimate')
    .addItem('🚀 Обработать текущий лист', 'processMonthlyReport')
    .addItem('📑 Показать структуру разделов', 'showSectionStructure')
    .addItem('📋 Анализ структуры', 'analyzeCurrentSheet')
    .addToUi();
}

/**
 * Показать структуру разделов
 */
function showSectionStructure() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var sections = [];
  var message = 'Структура разделов в листе:\n\n';
  
  for (var i = 0; i < Math.min(200, data.length); i++) {
    var firstCell = data[i][0] ? String(data[i][0]).trim().toLowerCase() : '';
    var sectionType = detectSectionHeader(firstCell);
    
    if (sectionType) {
      sections.push({
        row: i + 1,
        type: sectionType,
        header: data[i][0]
      });
    }
  }
  
  if (sections.length > 0) {
    for (var j = 0; j < sections.length; j++) {
      var sec = sections[j];
      message += 'Строка ' + sec.row + ': ' + sec.header + ' (' + sec.type + ')\n';
    }
  } else {
    message += 'Заголовки разделов не найдены';
  }
  
  showMessage('Структура разделов', message);
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