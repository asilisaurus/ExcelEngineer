/**
 * 🚀 ПРОЦЕССОР V12 - СПЕЦИАЛЬНАЯ ВЕРСИЯ ДЛЯ АПРЕЛЯ
 * Исправляет все проблемы со структурой данных Апреля
 * 
 * Изменения в V12:
 * - Находит заголовки с пробелами (О Т З Ы В Ы)
 * - Классифицирует по значениям в колонке A
 * - Правильная обработка колонок
 * - Исправление формата даты
 * - Обработка темы из колонки D вместо C
 */

// ==================== КОНФИГУРАЦИЯ ====================

var CONFIG = {
  STRUCTURE: {
    headerRow: 4,
    dataStartRow: 5,
    maxEmptyRows: 5
  },
  
  // Структура колонок для Апреля
  APRIL_COLUMNS: {
    typeColumn: 0,       // A - Тип записи (Отзывы (отзовики), Комментарии в обсуждениях)
    platform: 1,         // B - Площадка
    wrongProduct: 2,     // C - Неправильные данные (50 тыс)
    link: 3,            // D - Ссылка (используем как тему)
    text: 4,            // E - Текст
    date: 6,            // G - Дата
    nick: 7,            // H - Ник
    author: 8,          // I - Автор
    viewsStart: 9,      // J - Просмотры на старте
    viewsEnd: 10,       // K - Просмотры в конце
    viewsReceived: 11,  // L - Просмотров получено
    engagement: 12,     // M - Вовлечение
    postType: 13        // N - Тип поста (ОС/ЦС)
  },
  
  // Типы записей в колонке A
  RECORD_TYPES: {
    REVIEWS: ['отзывы (отзовики)', 'отзывы (аптеки)', 'отзывы'],
    COMMENTS: ['комментарии в обсуждениях', 'комментарии топ-20', 'комментарии'],
    DISCUSSIONS: ['активные обсуждения', 'обсуждения']
  },
  
  FORMATTING: {
    DATE_FORMAT: 'dd.MM.yyyy',
    NUMBER_FORMAT: '#,##0'
  }
};

// ==================== ОСНОВНЫЕ ФУНКЦИИ ====================

/**
 * Главная функция обработки
 */
function processMonthlyReport() {
  try {
    console.log('🚀 PROCESSOR V12 APRIL FIX - Начало обработки');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var sheetName = sheet.getName();
    
    // Проверяем, что это не служебный лист
    if (isServiceSheet(sheetName)) {
      showMessage('Ошибка', 'Выберите лист с данными месяца (например, Апр25)');
      return;
    }
    
    console.log('📋 Обрабатывается лист: ' + sheetName);
    
    // Получение данных
    var data = sheet.getDataRange().getValues();
    console.log('📊 Загружено строк: ' + data.length);
    
    // Определение месяца
    var monthInfo = detectMonth(sheetName, data);
    console.log('📅 Определен месяц: ' + monthInfo.name + ' ' + monthInfo.year);
    
    // Обработка данных
    var result = processAprilData(data);
    
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
      message += '\n\n⚠️ Внимание: данные о просмотрах отсутствуют';
    }
    
    // Добавляем детали обработки
    if (result.debugInfo) {
      message += '\n\nДетали:\n' + result.debugInfo;
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
 * Обработка данных с учетом структуры Апреля
 */
function processAprilData(data) {
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
  
  var processedRows = 0;
  var sectionInfo = {
    reviews: { start: -1, end: -1 },
    comments: { start: -1, end: -1 },
    discussions: { start: -1, end: -1 }
  };
  
  // Ищем заголовки разделов (могут быть с пробелами)
  console.log('🔍 Поиск заголовков разделов...');
  
  for (var i = 0; i < data.length; i++) {
    var rowText = '';
    // Собираем текст из первых ячеек строки
    for (var j = 0; j < Math.min(5, data[i].length); j++) {
      if (data[i][j]) {
        rowText += String(data[i][j]).toLowerCase().replace(/\s+/g, '');
      }
    }
    
    if (rowText.indexOf('отзывы') !== -1 && sectionInfo.reviews.start === -1) {
      sectionInfo.reviews.start = i + 1;
      console.log('📂 Найден раздел ОТЗЫВЫ в строке ' + (i + 1));
    } else if (rowText.indexOf('комментарии') !== -1 && sectionInfo.comments.start === -1) {
      if (sectionInfo.reviews.end === -1 && sectionInfo.reviews.start !== -1) {
        sectionInfo.reviews.end = i - 1;
      }
      sectionInfo.comments.start = i + 1;
      console.log('📂 Найден раздел КОММЕНТАРИИ в строке ' + (i + 1));
    } else if (rowText.indexOf('обсуждения') !== -1 && sectionInfo.discussions.start === -1) {
      if (sectionInfo.comments.end === -1 && sectionInfo.comments.start !== -1) {
        sectionInfo.comments.end = i - 1;
      }
      sectionInfo.discussions.start = i + 1;
      console.log('📂 Найден раздел ОБСУЖДЕНИЯ в строке ' + (i + 1));
    }
  }
  
  // Обработка данных по типу в колонке A
  var emptyRowCount = 0;
  
  for (var i = CONFIG.STRUCTURE.dataStartRow; i < data.length; i++) {
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
    
    var recordType = row[CONFIG.APRIL_COLUMNS.typeColumn] ? 
                     String(row[CONFIG.APRIL_COLUMNS.typeColumn]).trim().toLowerCase() : '';
    
    // Пропускаем заголовки
    if (recordType.replace(/\s+/g, '') === 'отзывы' ||
        recordType === 'комментарии' ||
        recordType === 'обсуждения') {
      continue;
    }
    
    // Извлекаем данные
    var platform = row[CONFIG.APRIL_COLUMNS.platform] ? 
                   String(row[CONFIG.APRIL_COLUMNS.platform]).trim() : '';
    var text = row[CONFIG.APRIL_COLUMNS.text] ? 
               String(row[CONFIG.APRIL_COLUMNS.text]).trim() : '';
    
    // Пропускаем строки без данных
    if (!platform && !text) continue;
    
    // Создаем запись
    var processedRow = {
      platform: platform,
      // Используем ссылку (колонка D) вместо неправильной колонки C
      theme: row[CONFIG.APRIL_COLUMNS.link] ? 
             '@' + String(row[CONFIG.APRIL_COLUMNS.link]).trim() : '',
      text: text,
      date: extractDateFromValue(row[CONFIG.APRIL_COLUMNS.date]),
      author: row[CONFIG.APRIL_COLUMNS.nick] ? 
              String(row[CONFIG.APRIL_COLUMNS.nick]).trim() : '',
      views: extractViewsFromRow(row),
      engagement: row[CONFIG.APRIL_COLUMNS.engagement] ? 
                  String(row[CONFIG.APRIL_COLUMNS.engagement]).trim() : '',
      postType: row[CONFIG.APRIL_COLUMNS.postType] ? 
                String(row[CONFIG.APRIL_COLUMNS.postType]).trim() : ''
    };
    
    processedRows++;
    
    // Классифицируем по типу записи
    if (isReviewType(recordType)) {
      result.reviews.push(processedRow);
      result.statistics.totalReviews++;
    } else if (isCommentType(recordType)) {
      result.commentsTop20.push(processedRow);
      result.statistics.totalCommentsTop20++;
    } else if (isDiscussionType(recordType)) {
      result.activeDiscussions.push(processedRow);
      result.statistics.totalActiveDiscussions++;
    } else {
      // По умолчанию считаем отзывом
      result.reviews.push(processedRow);
      result.statistics.totalReviews++;
    }
  }
  
  // Сортируем комментарии по просмотрам и оставляем топ-20
  if (result.commentsTop20.length > 20) {
    result.commentsTop20.sort(function(a, b) {
      return (b.views || 0) - (a.views || 0);
    });
    result.commentsTop20 = result.commentsTop20.slice(0, 20);
    result.statistics.totalCommentsTop20 = 20;
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
  result.debugInfo += 'Найдено разделов: ' + 
                      (sectionInfo.reviews.start > -1 ? 1 : 0) +
                      (sectionInfo.comments.start > -1 ? 1 : 0) +
                      (sectionInfo.discussions.start > -1 ? 1 : 0);
  
  console.log('📊 Итоги обработки:');
  console.log('   - Обработано строк: ' + processedRows);
  console.log('   - Отзывов: ' + result.statistics.totalReviews);
  console.log('   - Комментариев топ-20: ' + result.statistics.totalCommentsTop20);
  console.log('   - Обсуждений: ' + result.statistics.totalActiveDiscussions);
  console.log('   - Всего просмотров: ' + result.statistics.totalViews);
  
  return result;
}

/**
 * Проверка типа отзыва
 */
function isReviewType(type) {
  for (var i = 0; i < CONFIG.RECORD_TYPES.REVIEWS.length; i++) {
    if (type.indexOf(CONFIG.RECORD_TYPES.REVIEWS[i]) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * Проверка типа комментария
 */
function isCommentType(type) {
  // В апреле "Комментарии в обсуждениях" это топ-20
  return type.indexOf('комментарии') !== -1;
}

/**
 * Проверка типа обсуждения
 */
function isDiscussionType(type) {
  // Только если явно указано "активные обсуждения"
  for (var i = 0; i < CONFIG.RECORD_TYPES.DISCUSSIONS.length; i++) {
    if (type.indexOf(CONFIG.RECORD_TYPES.DISCUSSIONS[i]) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * Извлечение даты из значения (может быть Date объект)
 */
function extractDateFromValue(dateValue) {
  if (!dateValue) return '';
  
  // Если это объект Date
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), CONFIG.FORMATTING.DATE_FORMAT);
  }
  
  // Если это строка
  var dateStr = String(dateValue);
  
  // Пытаемся распарсить дату из строки
  if (dateStr.match(/\d{1,2}\.\d{1,2}\.\d{4}/)) {
    return dateStr;
  }
  
  // Пытаемся создать дату из строки
  try {
    var parsed = new Date(dateStr);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), CONFIG.FORMATTING.DATE_FORMAT);
    }
  } catch (e) {
    // Игнорируем ошибки парсинга
  }
  
  return dateStr;
}

/**
 * Извлечение просмотров из строки (проверяем несколько колонок)
 */
function extractViewsFromRow(row) {
  // Проверяем колонки в порядке приоритета
  var columnsToCheck = [
    CONFIG.APRIL_COLUMNS.viewsReceived,
    CONFIG.APRIL_COLUMNS.viewsEnd,
    CONFIG.APRIL_COLUMNS.viewsStart,
    CONFIG.APRIL_COLUMNS.wrongProduct // Может там просмотры
  ];
  
  for (var i = 0; i < columnsToCheck.length; i++) {
    var value = row[columnsToCheck[i]];
    if (value) {
      var views = parseViews(value);
      if (views > 0) {
        return views;
      }
    }
  }
  
  return 0;
}

/**
 * Парсинг просмотров из различных форматов
 */
function parseViews(value) {
  if (typeof value === 'number' && !isNaN(value)) {
    return Math.max(0, Math.floor(value));
  }
  
  var str = String(value).trim();
  if (str === '') return 0;
  
  // Обработка "50 тыс", "20 млн" и т.д.
  if (str.indexOf('тыс') !== -1) {
    var num = parseFloat(str.replace(/[^\d.,]/g, '').replace(',', '.'));
    if (!isNaN(num)) {
      return Math.floor(num * 1000);
    }
  }
  
  if (str.indexOf('млн') !== -1) {
    var numMln = parseFloat(str.replace(/[^\d.,]/g, '').replace(',', '.'));
    if (!isNaN(numMln)) {
      return Math.floor(numMln * 1000000);
    }
  }
  
  var cleanStr = str.replace(/[^\d.,]/g, '');
  var normalizedStr = cleanStr.replace(',', '.');
  var parsed = parseFloat(normalizedStr);
  
  if (!isNaN(parsed)) {
    return Math.max(0, Math.floor(parsed));
  }
  
  return 0;
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
 * Расчет доли вовлечения
 */
function calculateEngagementRate(discussions) {
  if (discussions.length === 0) return 0;
  
  var withEngagement = 0;
  for (var i = 0; i < discussions.length; i++) {
    var eng = discussions[i].engagement;
    if (eng && eng !== '0' && eng !== '' && eng !== '-' && eng !== 'нет') {
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
    .createMenu('📊 Процессор V12 April')
    .addItem('🚀 Обработать текущий лист', 'processMonthlyReport')
    .addItem('🔍 Проверить структуру', 'checkDataStructure')
    .addToUi();
}

/**
 * Проверка структуры данных
 */
function checkDataStructure() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var message = 'Структура данных листа ' + sheet.getName() + ':\n\n';
  
  // Показываем заголовки
  if (data.length > 3) {
    message += 'Заголовки (строка 4):\n';
    var headers = data[3];
    for (var i = 0; i < Math.min(8, headers.length); i++) {
      if (headers[i]) {
        message += String.fromCharCode(65 + i) + ': ' + headers[i] + '\n';
      }
    }
  }
  
  message += '\nПервая строка данных:\n';
  if (data.length > 5) {
    var firstRow = data[5];
    message += 'Тип: ' + (firstRow[0] || '') + '\n';
    message += 'Площадка: ' + (firstRow[1] || '') + '\n';
    message += 'Колонка C: ' + (firstRow[2] || '') + '\n';
    message += 'Ссылка: ' + String(firstRow[3] || '').substring(0, 30) + '...\n';
  }
  
  // Ищем просмотры
  message += '\nПоиск просмотров:\n';
  var foundViews = false;
  for (var col = 2; col < Math.min(15, data[0].length); col++) {
    for (var row = 5; row < Math.min(10, data.length); row++) {
      var value = String(data[row][col] || '');
      if (value.indexOf('тыс') !== -1 || value.indexOf('млн') !== -1) {
        message += 'Найдены в колонке ' + String.fromCharCode(65 + col) + ': ' + value + '\n';
        foundViews = true;
        break;
      }
    }
    if (foundViews) break;
  }
  
  showMessage('Структура данных', message);
}