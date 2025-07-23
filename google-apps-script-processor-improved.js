/**
 * 🚀 УЛУЧШЕННЫЙ ГИБКИЙ ОБРАБОТЧИК V7
 * Google Apps Script для автоматической обработки отчетов
 * Версия: 7.0.0 - IMPROVED
 * 
 * Основано на google-apps-script-processor-final.js
 * Улучшения:
 * - Совместимость со старыми версиями Google Apps Script
 * - Автоматический выбор листа с данными
 * - Гибкая обработка любого количества данных
 * - Исправлена обработка UI контекста
 */

// ==================== КОНФИГУРАЦИЯ ====================

var CONFIG = {
  // Настройки структуры данных
  STRUCTURE: {
    headerRow: 4,        // Заголовки ВСЕГДА в строке 4
    dataStartRow: 5,     // Данные начинаются с строки 5
    infoRows: [1, 2, 3], // Мета-информация в строках 1-3
    maxRows: 10000
  },
  
  // Классификация контента
  CONTENT_TYPES: {
    REVIEWS: ['ОС', 'Отзывы Сайтов', 'ос', 'отзывы сайтов'],
    TARGETED: ['ЦС', 'Целевые Сайты', 'цс', 'целевые сайты'],
    SOCIAL: ['ПС', 'Площадки Социальные', 'пс', 'площадки социальные']
  },
  
  // Настройки форматирования
  FORMATTING: {
    DATE_FORMAT: 'dd.mm.yyyy',
    NUMBER_FORMAT: '#,##0',
    CURRENCY_FORMAT: '#,##0 ₽'
  }
};

// ==================== ОСНОВНЫЕ ФУНКЦИИ ====================

/**
 * Главная функция обработки отчета
 */
function processMonthlyReport() {
  try {
    console.log('🚀 IMPROVED PROCESSOR V7 - Начало обработки');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Автоматический выбор листа с данными
    var sheet = findDataSheet(spreadsheet);
    if (!sheet) {
      showMessage('Ошибка', 'Не найден лист с данными месяца! Убедитесь, что есть лист с названием месяца (например, Февраль25, Март25)');
      return;
    }
    
    var sheetName = sheet.getName();
    console.log('📋 Обрабатывается лист: ' + sheetName);
    
    // Получение данных
    var data = sheet.getDataRange().getValues();
    console.log('📊 Загружено строк: ' + data.length);
    
    // Определение месяца
    var monthInfo = detectMonth(sheetName, data);
    console.log('📅 Месяц: ' + monthInfo.name + ' ' + monthInfo.year);
    
    // Обработка данных
    var result = processDataFlexible(data);
    
    // Создание отчета
    var reportUrl = createReport(result, monthInfo);
    
    // Показ результата
    var message = 'Отчет создан успешно!\n\n' +
                  'Ссылка: ' + reportUrl + '\n\n' +
                  'Обработано:\n' +
                  '- Отзывов: ' + result.statistics.totalReviews + '\n' +
                  '- Комментариев Топ-20: ' + result.statistics.totalCommentsTop20 + '\n' +
                  '- Активных обсуждений: ' + result.statistics.totalActiveDiscussions + '\n' +
                  '- Общие просмотры: ' + result.statistics.totalViews;
    
    console.log('✅ Обработка завершена успешно');
    showMessage('Обработка завершена', message);
    
  } catch (error) {
    console.error('❌ Ошибка: ' + error.toString());
    console.error(error.stack);
    showMessage('Ошибка', 'Произошла ошибка: ' + error.toString());
  }
}

/**
 * Автоматический поиск листа с данными
 */
function findDataSheet(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var monthSheets = [];
  
  var monthPatterns = [
    'январ', 'феврал', 'март', 'апрел', 'май', 'мая',
    'июн', 'июл', 'август', 'сентябр', 'октябр', 'ноябр', 'декабр'
  ];
  
  // Ищем листы с названиями месяцев
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName().toLowerCase();
    
    for (var j = 0; j < monthPatterns.length; j++) {
      if (sheetName.indexOf(monthPatterns[j]) !== -1 && sheetName.indexOf('25') !== -1) {
        var monthNum = j + 1;
        if (monthPatterns[j] === 'май' || monthPatterns[j] === 'мая') monthNum = 5;
        
        monthSheets.push({
          sheet: sheets[i],
          name: sheets[i].getName(),
          monthNumber: monthNum
        });
        break;
      }
    }
  }
  
  // Сортируем по номеру месяца и возвращаем самый последний
  if (monthSheets.length > 0) {
    monthSheets.sort(function(a, b) { return b.monthNumber - a.monthNumber; });
    console.log('📂 Найдены листы с месяцами: ' + monthSheets.map(function(s) { return s.name; }).join(', '));
    console.log('✅ Выбран: ' + monthSheets[0].name);
    return monthSheets[0].sheet;
  }
  
  // Если не нашли листы с месяцами, берем текущий активный
  return spreadsheet.getActiveSheet();
}

/**
 * Гибкая обработка данных (без фиксированных лимитов)
 */
function processDataFlexible(data) {
  var result = {
    reviews: [],
    commentsTop20: [],
    activeDiscussions: [],
    statistics: {
      totalReviews: 0,
      totalCommentsTop20: 0,
      totalActiveDiscussions: 0,
      totalViews: 0,
      engagementShare: 0,
      platforms: []
    }
  };
  
  // Получаем маппинг колонок
  var columnMapping = getColumnMapping();
  
  // Массив всех записей ЦС для сортировки
  var allTargetedRecords = [];
  
  // Обрабатываем все строки данных
  for (var i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
    var row = data[i];
    
    // Пропускаем пустые строки
    if (isEmptyRow(row)) continue;
    
    // Останавливаемся на статистике
    if (isStatisticsRow(row)) break;
    
    // Извлекаем данные
    var postType = row[columnMapping.postType] ? String(row[columnMapping.postType]).trim().toUpperCase() : '';
    
    // Пропускаем строки без типа
    if (!postType) continue;
    
    // Обрабатываем строку
    var processedRow = processRow(row, columnMapping);
    
    if (processedRow) {
      // Добавляем платформу в статистику
      if (processedRow.platform && result.statistics.platforms.indexOf(processedRow.platform) === -1) {
        result.statistics.platforms.push(processedRow.platform);
      }
      
      // Классифицируем по типу поста
      if (postType === 'ОС' || postType.indexOf('ОТЗЫВ') !== -1) {
        // Отзывы - берем все
        result.reviews.push(processedRow);
        result.statistics.totalReviews++;
      } 
      else if (postType === 'ЦС' || postType.indexOf('КОММЕНТАРИЙ') !== -1 || postType.indexOf('ОБСУЖДЕНИЕ') !== -1) {
        // Целевые сайты - сохраняем для последующей сортировки
        allTargetedRecords.push(processedRow);
      }
    }
  }
  
  // Сортируем все записи ЦС по просмотрам (убывание)
  allTargetedRecords.sort(function(a, b) { 
    return (b.views || 0) - (a.views || 0); 
  });
  
  // Первые 20 - это Топ-20 комментариев
  for (var j = 0; j < allTargetedRecords.length; j++) {
    if (j < 20) {
      result.commentsTop20.push(allTargetedRecords[j]);
      result.statistics.totalCommentsTop20++;
    } else {
      // Остальные - активные обсуждения
      result.activeDiscussions.push(allTargetedRecords[j]);
      result.statistics.totalActiveDiscussions++;
    }
  }
  
  // Извлекаем статистику из исходных данных
  var sourceStats = extractStatisticsFromSourceData(data);
  result.statistics.totalViews = sourceStats.totalViews || 0;
  result.statistics.engagementShare = sourceStats.engagementShare || 0;
  
  // Если просмотры не найдены в статистике, считаем из данных
  if (result.statistics.totalViews === 0) {
    var totalViews = 0;
    var allRecords = result.reviews.concat(result.commentsTop20).concat(result.activeDiscussions);
    for (var k = 0; k < allRecords.length; k++) {
      if (allRecords[k].views && allRecords[k].views > 0) {
        totalViews += allRecords[k].views;
      }
    }
    result.statistics.totalViews = totalViews;
  }
  
  console.log('📊 Обработано: ' + result.statistics.totalReviews + ' отзывов, ' + 
              result.statistics.totalCommentsTop20 + ' комментариев топ-20, ' + 
              result.statistics.totalActiveDiscussions + ' активных обсуждений');
  console.log('👁️ Всего просмотров: ' + result.statistics.totalViews);
  
  return result;
}

/**
 * Обработка одной строки данных
 */
function processRow(row, columnMapping) {
  try {
    // Проверяем наличие данных
    var text = row[columnMapping.text] ? String(row[columnMapping.text]).trim() : '';
    var platform = row[columnMapping.platform] ? String(row[columnMapping.platform]).trim() : '';
    
    // Пропускаем строки без текста и платформы
    if (!text && !platform) {
      return null;
    }
    
    // Извлекаем все поля
    var processedRow = {
      platform: platform,
      theme: row[columnMapping.theme] ? String(row[columnMapping.theme]).trim() : '',
      text: text,
      date: extractDate(row, columnMapping),
      author: row[columnMapping.author] ? String(row[columnMapping.author]).trim() : '',
      views: extractViews(row, columnMapping),
      engagement: row[columnMapping.engagement] ? String(row[columnMapping.engagement]).trim() : '',
      postType: row[columnMapping.postType] ? String(row[columnMapping.postType]).trim() : ''
    };
    
    return processedRow;
    
  } catch (e) {
    console.error('Ошибка обработки строки: ' + e.toString());
    return null;
  }
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
  var engagementRate = calculateEngagementRate(processedData.commentsTop20.concat(processedData.activeDiscussions));
  sheet.getRange(row, 2).setValue(engagementRate);
  sheet.getRange(row, 2).setNumberFormat("0%");
  
  // Форматирование
  sheet.autoResizeColumns(1, headers.length);
  
  console.log('📄 Отчет создан: ' + reportName);
  return newSpreadsheet.getUrl();
}

/**
 * Запись раздела в отчет
 */
function writeSection(sheet, startRow, sectionName, dataArr, colCount) {
  // Заголовок раздела
  sheet.getRange(startRow, 1).setValue(sectionName);
  sheet.getRange(startRow, 1, 1, colCount).setBackground('#b7a6c9').setFontWeight('bold');
  startRow++;
  
  // Данные раздела
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

// ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Определение месяца
 */
function detectMonth(sheetName, data) {
  // Сначала пробуем по имени листа
  var monthFromSheet = extractMonthFromText(sheetName);
  if (monthFromSheet) {
    return monthFromSheet;
  }
  
  // Затем ищем в первых строках данных
  for (var i = 0; i < Math.min(3, data.length); i++) {
    var rowText = data[i].join(' ');
    var monthFromData = extractMonthFromText(rowText);
    if (monthFromData) {
      return monthFromData;
    }
  }
  
  // По умолчанию - текущий месяц
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
        return {
          name: month.name,
          number: month.number,
          year: 2025
        };
      }
    }
  }
  
  return null;
}

/**
 * Маппинг колонок (фиксированный)
 */
function getColumnMapping() {
  return {
    platform: 1,     // B - Площадка
    theme: 3,        // D - Тема
    text: 4,         // E - Текст сообщения
    date: 6,         // G - Дата
    author: 7,       // H - Ник
    views: 11,       // L - Просмотры
    engagement: 12,  // M - Вовлечение
    postType: 13     // N - Тип поста
  };
}

/**
 * Извлечение статистики из исходных данных
 */
function extractStatisticsFromSourceData(data) {
  var stats = {
    totalViews: 0,
    engagementShare: 0
  };
  
  // Ищем статистику в последних 20 строках
  var startSearch = Math.max(0, data.length - 20);
  
  for (var i = startSearch; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) continue;
    
    var firstCell = String(row[0] || '').toLowerCase().trim();
    
    // Суммарные просмотры
    if (firstCell.indexOf('суммарное количество просмотров') !== -1) {
      for (var j = 1; j < row.length; j++) {
        if (row[j]) {
          var value = parseFloat(String(row[j]).replace(/[^\d]/g, ''));
          if (!isNaN(value) && value > 0) {
            stats.totalViews = Math.floor(value);
            console.log('✅ Найдены общие просмотры: ' + stats.totalViews);
            break;
          }
        }
      }
    }
    
    // Доля вовлечения
    if (firstCell.indexOf('доля обсуждений с вовлечением') !== -1) {
      for (var k = 1; k < row.length; k++) {
        if (row[k]) {
          var cellValue = String(row[k]).trim();
          var engValue = 0;
          
          if (cellValue.indexOf('%') !== -1) {
            engValue = parseFloat(cellValue.replace('%', '')) / 100;
          } else {
            engValue = parseFloat(cellValue);
            if (!isNaN(engValue) && engValue > 1) {
              engValue = engValue / 100;
            }
          }
          
          if (!isNaN(engValue) && engValue >= 0) {
            stats.engagementShare = engValue;
            console.log('✅ Найдена доля вовлечения: ' + (engValue * 100).toFixed(0) + '%');
            break;
          }
        }
      }
    }
  }
  
  return stats;
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
function extractViews(row, columnMapping) {
  var index = columnMapping.views;
  if (index !== undefined && row[index]) {
    var viewsValue = row[index];
    
    if (typeof viewsValue === 'number' && !isNaN(viewsValue)) {
      return Math.max(0, Math.floor(viewsValue));
    }
    
    var viewsStr = String(viewsValue).trim();
    var cleanStr = viewsStr.replace(/[^\d.,]/g, '');
    var normalizedStr = cleanStr.replace(',', '.');
    var parsed = parseFloat(normalizedStr);
    
    if (!isNaN(parsed)) {
      return Math.max(0, Math.floor(parsed));
    }
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
    if (eng && eng !== '0' && eng !== '') {
      withEngagement++;
    }
  }
  
  return withEngagement / discussions.length;
}

/**
 * Показ сообщения (обработка UI контекста)
 */
function showMessage(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // Если UI недоступен, просто логируем
    console.log(title + ': ' + message);
  }
}

/**
 * Создание меню при открытии
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Обработка V7')
    .addItem('🚀 Обработать месяц', 'processMonthlyReport')
    .addItem('📋 Показать листы', 'showAvailableSheets')
    .addItem('🧪 Тест обработки', 'testProcessing')
    .addToUi();
}

/**
 * Показать доступные листы
 */
function showAvailableSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var message = 'Доступные листы:\n\n';
  
  for (var i = 0; i < sheets.length; i++) {
    message += (i + 1) + '. ' + sheets[i].getName() + '\n';
  }
  
  var dataSheet = findDataSheet(spreadsheet);
  if (dataSheet) {
    message += '\n✅ Будет обработан: ' + dataSheet.getName();
  }
  
  showMessage('Листы в таблице', message);
}

/**
 * Тестовая функция
 */
function testProcessing() {
  try {
    console.log('🧪 Запуск теста...');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findDataSheet(spreadsheet);
    
    if (!sheet) {
      showMessage('Тест', 'Лист с данными не найден');
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    var monthInfo = detectMonth(sheet.getName(), data);
    
    var message = 'Результаты теста:\n\n' +
                  'Лист: ' + sheet.getName() + '\n' +
                  'Месяц: ' + monthInfo.name + ' ' + monthInfo.year + '\n' +
                  'Строк данных: ' + data.length + '\n';
    
    // Проверяем маппинг
    var mapping = getColumnMapping();
    var headerRow = data[CONFIG.STRUCTURE.headerRow - 1];
    if (headerRow) {
      message += '\nЗаголовки колонок:\n';
      message += '- Площадка (B): ' + (headerRow[mapping.platform] || 'не найдено') + '\n';
      message += '- Текст (E): ' + (headerRow[mapping.text] || 'не найдено') + '\n';
      message += '- Тип поста (N): ' + (headerRow[mapping.postType] || 'не найдено') + '\n';
    }
    
    showMessage('Результаты теста', message);
    
  } catch (error) {
    showMessage('Ошибка теста', error.toString());
  }
}