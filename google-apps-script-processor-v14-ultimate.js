/**
 * 🚀 ОКОНЧАТЕЛЬНЫЙ ПРОЦЕССОР V14.2
 * Правильно обрабатывает ВСЕ записи без потерь
 * 
 * Изменения в V14.2:
 * - Исправлено определение границ разделов
 * - НЕ теряет записи обсуждений
 * - Правильно находит ровно 20 комментариев топ-20
 * - Учитывает пустые строки при подсчете
 * - Убрана приставка "@" из колонки "Тема"
 * - Исправлен подсчет просмотров
 */

// ==================== КОНФИГУРАЦИЯ ====================

var CONFIG = {
  STRUCTURE: {
    headerRow: 4,
    dataStartRow: 5,
    maxEmptyRows: 5
  },
  
  // Структура колонок
  COLUMNS: {
    typeColumn: 0,       // A - Тип записи
    platform: 1,         // B - Площадка
    product: 2,          // C - Продукт (НЕ просмотры!)
    link: 3,             // D - Ссылка
    text: 4,             // E - Текст
    comments: 5,         // F - Согласование/Комментарии
    date: 6,             // G - Дата
    nick: 7,             // H - Ник
    author: 8,           // I - Автор
    viewsStart: 9,       // J - Просмотры темы на старте
    viewsEnd: 10,        // K - Просмотры в конце месяца
    viewsReceived: 11,   // L - Просмотров получено
    engagement: 12,      // M - Вовлечение
    postType: 13         // N - Тип поста (ОС/ЦС)
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
         console.log('🚀 ULTIMATE PROCESSOR V14.1 - Начало обработки');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var sheetName = sheet.getName();
    
    // Проверяем, что это не служебный лист
    if (isServiceSheet(sheetName)) {
      showMessage('Ошибка', 'Выберите лист с данными месяца');
      return;
    }
    
    console.log('📋 Обрабатывается лист: ' + sheetName);
    
    // Получение данных
    var data = sheet.getDataRange().getValues();
    console.log('📊 Загружено строк: ' + data.length);
    
    // Определение месяца
    var monthInfo = detectMonth(sheetName, data);
    console.log('📅 Определен месяц: ' + monthInfo.name + ' ' + monthInfo.year);
    
    // Обработка данных с учетом специфики Апреля
    var result = processAprilDataCorrectly(data);
    
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
                  '- Общие просмотры: ' + formatNumber(result.statistics.totalViews) + '\n' +
                  '- Всего записей: ' + result.statistics.totalProcessed;
    
    if (result.statistics.recordsWithoutViews > 0) {
      message += '\n\n⚠️ Записей без просмотров: ' + result.statistics.recordsWithoutViews;
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
 * Обработка данных с правильным разделением на секции
 */
function processAprilDataCorrectly(data) {
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
      totalProcessed: 0,
      recordsWithoutViews: 0
    }
  };
  
  // Определяем границы разделов
  var sections = findSectionBoundaries(data);
  console.log('📂 Найдены разделы:', JSON.stringify(sections));
  
  var emptyRowCount = 0;
  
  // Обрабатываем данные по разделам
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
    
    // Определяем в каком разделе находимся
    var currentSection = determineSection(i, sections);
    if (!currentSection) continue;
    
    // Извлекаем данные
    var recordType = row[CONFIG.COLUMNS.typeColumn] ? 
                     String(row[CONFIG.COLUMNS.typeColumn]).trim() : '';
    var platform = row[CONFIG.COLUMNS.platform] ? 
                   String(row[CONFIG.COLUMNS.platform]).trim() : '';
    var text = row[CONFIG.COLUMNS.text] ? 
               String(row[CONFIG.COLUMNS.text]).trim() : '';
    
    // Пропускаем строки без данных или заголовки
    if ((!platform && !text) || isHeaderRow(recordType)) continue;
    
    // Извлекаем просмотры ТОЛЬКО из колонок J, K, L
    var views = extractViewsFromColumns(row);
    
         // Создаем запись
     var processedRow = {
       platform: platform,
       theme: row[CONFIG.COLUMNS.link] ? 
              String(row[CONFIG.COLUMNS.link]).trim() : '',
       text: text,
       date: extractDateFromValue(row[CONFIG.COLUMNS.date]),
       author: row[CONFIG.COLUMNS.nick] ? 
               String(row[CONFIG.COLUMNS.nick]).trim() : '',
       views: views,
       engagement: row[CONFIG.COLUMNS.engagement] ? 
                   String(row[CONFIG.COLUMNS.engagement]).trim() : '',
       postType: row[CONFIG.COLUMNS.postType] ? 
                 String(row[CONFIG.COLUMNS.postType]).trim() : ''
     };
    
    result.statistics.totalProcessed++;
    if (views === 0) {
      result.statistics.recordsWithoutViews++;
    }
    
    // Распределяем по разделам
    if (currentSection === 'reviews') {
      result.reviews.push(processedRow);
      result.statistics.totalReviews++;
    } else if (currentSection === 'commentsTop20') {
      result.commentsTop20.push(processedRow);
      result.statistics.totalCommentsTop20++;
    } else if (currentSection === 'discussions') {
      result.activeDiscussions.push(processedRow);
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
  
  console.log('📊 Итоги обработки:');
  console.log('   - Всего обработано: ' + result.statistics.totalProcessed);
  console.log('   - Отзывов: ' + result.statistics.totalReviews);
  console.log('   - Комментариев топ-20: ' + result.statistics.totalCommentsTop20);
  console.log('   - Обсуждений: ' + result.statistics.totalActiveDiscussions);
  console.log('   - Всего просмотров: ' + result.statistics.totalViews);
  
  return result;
}

/**
 * Поиск границ разделов в данных
 */
function findSectionBoundaries(data) {
  var sections = {
    reviews: { start: -1, end: -1 },
    commentsTop20: { start: -1, end: -1 },
    discussions: { start: -1, end: -1 }
  };
  
  // Ищем заголовки и границы
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var firstCell = row[0] ? String(row[0]).trim().toLowerCase() : '';
    
    // Объединяем текст из первых ячеек для поиска заголовков
    var rowText = '';
    for (var j = 0; j < Math.min(5, row.length); j++) {
      if (row[j]) rowText += String(row[j]).toLowerCase() + ' ';
    }
    rowText = rowText.replace(/\s+/g, '');
    
    // Проверяем заголовок "О Т З Ы В Ы"
    if (rowText.indexOf('отзывы') === 0 && sections.reviews.start === -1) {
      sections.reviews.start = i + 1;
      console.log('📍 Найден заголовок ОТЗЫВЫ в строке ' + (i + 1));
    }
    
    // Проверяем заголовок "ТОП-20 ВЫДАЧИ"
    if ((firstCell.indexOf('топ-20') !== -1 || firstCell.indexOf('топ 20') !== -1) && 
        sections.commentsTop20.start === -1) {
      // Закрываем раздел отзывов
      if (sections.reviews.start !== -1 && sections.reviews.end === -1) {
        sections.reviews.end = i - 1;
      }
      sections.commentsTop20.start = i + 1;
      console.log('📍 Найден заголовок ТОП-20 в строке ' + (i + 1));
    }
    
    // Проверяем заголовок "АКТИВНЫЕ ОБСУЖДЕНИЯ"
    if (rowText.indexOf('активныеобсуждения') !== -1 && sections.discussions.start === -1) {
      // Закрываем раздел комментариев топ-20
      if (sections.commentsTop20.start !== -1 && sections.commentsTop20.end === -1) {
        sections.commentsTop20.end = i - 1;
      }
      sections.discussions.start = i + 1;
      console.log('📍 Найден заголовок АКТИВНЫЕ ОБСУЖДЕНИЯ в строке ' + (i + 1));
    }
  }
  
     // Специальная логика для Апреля
   // Если нашли ТОП-20 но не нашли активные обсуждения
   if (sections.commentsTop20.start > 0 && sections.discussions.start === -1) {
     // Найдем ровно 20 непустых записей после заголовка
     var count = 0;
     var currentRow = sections.commentsTop20.start;
     
     for (var idx = sections.commentsTop20.start - 1; idx < data.length && count < 20; idx++) {
       var row = data[idx];
       // Проверяем что это не пустая строка и не заголовок
       if (row && row[1] && String(row[1]).trim() !== '' && 
           !isHeaderRow(String(row[0] || '').trim())) {
         count++;
         currentRow = idx + 1;
       }
     }
     
     sections.commentsTop20.end = currentRow;
     
     // Все остальное - активные обсуждения
     sections.discussions.start = currentRow + 1;
     sections.discussions.end = data.length;
     
     console.log('📍 Автоматическое разделение:');
     console.log('   - Комментарии топ-20: строки ' + sections.commentsTop20.start + 
                 '-' + sections.commentsTop20.end + ' (найдено ' + count + ' записей)');
     console.log('   - Активные обсуждения: с строки ' + sections.discussions.start);
   }
  
  // Если не нашли четких границ, используем стандартную логику
  if (sections.reviews.start === -1) {
    sections.reviews.start = CONFIG.STRUCTURE.dataStartRow;
    sections.reviews.end = 27; // Примерно 22-23 отзыва
  }
  
  return sections;
}

/**
 * Определение текущего раздела по номеру строки
 */
function determineSection(rowIndex, sections) {
  var rowNumber = rowIndex + 1; // Преобразуем в 1-based
  
  if (sections.reviews.start <= rowNumber && 
      (sections.reviews.end === -1 || rowNumber <= sections.reviews.end)) {
    return 'reviews';
  }
  
  if (sections.commentsTop20.start > 0 && sections.commentsTop20.start <= rowNumber && 
      (sections.commentsTop20.end === -1 || rowNumber <= sections.commentsTop20.end)) {
    return 'commentsTop20';
  }
  
  if (sections.discussions.start > 0 && sections.discussions.start <= rowNumber) {
    return 'discussions';
  }
  
  return null;
}

/**
 * Проверка является ли строка заголовком
 */
function isHeaderRow(text) {
  var cleanText = text.toLowerCase().replace(/\s+/g, '');
  return cleanText === 'топ20выдачи' || 
         cleanText === 'топ-20выдачи' ||
         cleanText === 'типразмещения' ||
         cleanText === 'активныеобсуждения';
}

/**
 * Извлечение просмотров ТОЛЬКО из правильных колонок
 */
function extractViewsFromColumns(row) {
  // Проверяем колонки в порядке приоритета
  var viewsColumns = [
    CONFIG.COLUMNS.viewsReceived,  // L - Просмотров получено
    CONFIG.COLUMNS.viewsEnd,        // K - Просмотры в конце месяца
    CONFIG.COLUMNS.viewsStart       // J - Просмотры темы на старте
  ];
  
  for (var i = 0; i < viewsColumns.length; i++) {
    var value = row[viewsColumns[i]];
    if (value !== undefined && value !== null && value !== '' && value !== '-') {
      // Если это число
      if (typeof value === 'number' && !isNaN(value)) {
        return Math.max(0, Math.floor(value));
      }
      
      // Если это строка с числом
      var strValue = String(value).trim();
      var parsed = parseFloat(strValue.replace(/[^\d.,]/g, '').replace(',', '.'));
      if (!isNaN(parsed)) {
        return Math.max(0, Math.floor(parsed));
      }
    }
  }
  
  return 0;
}

/**
 * Извлечение даты
 */
function extractDateFromValue(dateValue) {
  if (!dateValue) return '';
  
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), CONFIG.FORMATTING.DATE_FORMAT);
  }
  
  var dateStr = String(dateValue);
  
  if (dateStr.match(/\d{1,2}\.\d{1,2}\.\d{4}/)) {
    return dateStr;
  }
  
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
  sheet.getRange(row, 2).setNumberFormat('#,##0');
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
 * Форматирование числа
 */
function formatNumber(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
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
         .createMenu('📊 Ultimate процессор V14.1')
    .addItem('🚀 Обработать текущий лист', 'processMonthlyReport')
    .addItem('📊 Показать структуру данных', 'showDataStructure')
    .addToUi();
}

/**
 * Показать структуру данных
 */
function showDataStructure() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var sections = findSectionBoundaries(data);
  
  var message = 'Структура данных листа ' + sheet.getName() + ':\n\n';
  
  message += 'Раздел ОТЗЫВЫ:\n';
  if (sections.reviews.start > 0) {
    message += 'Строки ' + sections.reviews.start + '-' + 
               (sections.reviews.end > 0 ? sections.reviews.end : '?') + '\n\n';
  } else {
    message += 'Не найден\n\n';
  }
  
  message += 'Раздел КОММЕНТАРИИ ТОП-20:\n';
  if (sections.commentsTop20.start > 0) {
    message += 'Строки ' + sections.commentsTop20.start + '-' + 
               (sections.commentsTop20.end > 0 ? sections.commentsTop20.end : '?') + '\n\n';
  } else {
    message += 'Не найден\n\n';
  }
  
  message += 'Раздел АКТИВНЫЕ ОБСУЖДЕНИЯ:\n';
  if (sections.discussions.start > 0) {
    message += 'С строки ' + sections.discussions.start + '\n';
  } else {
    message += 'Не найден\n';
  }
  
  showMessage('Структура данных', message);
}