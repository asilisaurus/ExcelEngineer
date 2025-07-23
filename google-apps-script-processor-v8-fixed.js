/**
 * 🚀 ИСПРАВЛЕННЫЙ ПРОЦЕССОР V8
 * Работает с новой структурой данных Июнь 2025
 * 
 * Изменения:
 * - Обрабатывает ТЕКУЩИЙ активный лист (не переключается автоматически)
 * - Определяет тип по колонке A "Тип размещения"
 * - Адаптирован под новую структуру колонок
 */

// ==================== КОНФИГУРАЦИЯ ====================

var CONFIG = {
  STRUCTURE: {
    headerRow: 4,        // Заголовки в строке 4
    dataStartRow: 5,     // Данные начинаются с строки 5
    infoRows: [1, 2, 3]  // Мета-информация
  },
  
  // Типы контента по колонке A
  CONTENT_PATTERNS: {
    REVIEWS: ['отзыв', 'о т з ы в'],
    COMMENTS: ['комментар', 'топ-20', 'топ 20'],
    DISCUSSIONS: ['обсужден', 'форум', 'сообщест']
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
    console.log('🚀 PROCESSOR V8 - Начало обработки');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet(); // Используем ТЕКУЩИЙ активный лист
    var sheetName = sheet.getName();
    
    console.log('📋 Обрабатывается текущий лист: ' + sheetName);
    
    // Проверяем, что это лист с данными
    if (sheetName.toLowerCase().indexOf('инструкция') !== -1 || 
        sheetName.toLowerCase().indexOf('диагностика') !== -1) {
      showMessage('Ошибка', 'Выберите лист с данными месяца (например, Апрель25, Май25)');
      return;
    }
    
    // Получение данных
    var data = sheet.getDataRange().getValues();
    console.log('📊 Загружено строк: ' + data.length);
    
    // Определение месяца
    var monthInfo = detectMonth(sheetName, data);
    console.log('📅 Месяц: ' + monthInfo.name + ' ' + monthInfo.year);
    
    // Обработка данных с новой структурой
    var result = processDataNewStructure(data);
    
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
    
    console.log('✅ Обработка завершена');
    showMessage('Обработка завершена', message);
    
  } catch (error) {
    console.error('❌ Ошибка: ' + error.toString());
    console.error(error.stack);
    showMessage('Ошибка', 'Произошла ошибка: ' + error.toString());
  }
}

/**
 * Обработка данных с новой структурой
 */
function processDataNewStructure(data) {
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
    }
  };
  
  // Новый маппинг колонок для Июнь 2025
  var columnMapping = {
    typeOfPlacement: 0,  // A - Тип размещения
    platform: 1,         // B - Площадка
    product: 2,          // C - Продукт  
    link: 3,             // D - Ссылка на сообщение
    text: 4,             // E - Текст сообщения
    date: 6,             // G - Дата
    author: 7,           // H - Ник
    viewsEnd: 11,        // L - Просмотры в конце месяца
    viewsReceived: 12,   // M - Просмотров получено
    engagement: 13,      // N - Вовлечение
    postType: 14         // O - Тип поста
  };
  
  var currentSection = null;
  var allTargetedRecords = [];
  
  // Обрабатываем строки начиная с 5
  for (var i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
    var row = data[i];
    
    // Пропускаем пустые строки
    if (isEmptyRow(row)) continue;
    
    // Останавливаемся на статистике
    if (isStatisticsRow(row)) break;
    
    // Проверяем, не заголовок ли это раздела
    var firstCell = row[0] ? String(row[0]).toLowerCase() : '';
    
    // Определяем раздел по заголовку
    if (firstCell.indexOf('о т з ы в') !== -1 || firstCell.indexOf('отзывы') !== -1) {
      currentSection = 'reviews';
      console.log('📂 Найден раздел: Отзывы (строка ' + (i + 1) + ')');
      continue;
    } else if (firstCell.indexOf('комментар') !== -1 && firstCell.indexOf('топ') !== -1) {
      currentSection = 'comments';
      console.log('📂 Найден раздел: Комментарии Топ-20 (строка ' + (i + 1) + ')');
      continue;
    } else if (firstCell.indexOf('обсужден') !== -1 || firstCell.indexOf('мониторинг') !== -1) {
      currentSection = 'discussions';
      console.log('📂 Найден раздел: Активные обсуждения (строка ' + (i + 1) + ')');
      continue;
    }
    
    // Обрабатываем данные
    var typeOfPlacement = firstCell;
    var platform = row[columnMapping.platform] ? String(row[columnMapping.platform]).trim() : '';
    var text = row[columnMapping.text] ? String(row[columnMapping.text]).trim() : '';
    
    // Пропускаем строки без данных
    if (!platform && !text) continue;
    
    // Извлекаем данные
    var processedRow = {
      platform: platform,
      theme: row[columnMapping.product] ? String(row[columnMapping.product]).trim() : '',
      text: text,
      date: extractDate(row, columnMapping),
      author: row[columnMapping.author] ? String(row[columnMapping.author]).trim() : '',
      views: extractViewsReceived(row, columnMapping), // Используем "Просмотров получено"
      engagement: row[columnMapping.engagement] ? String(row[columnMapping.engagement]).trim() : '',
      postType: typeOfPlacement // Используем тип размещения как тип поста
    };
    
    // Классифицируем по текущему разделу или типу размещения
    if (currentSection === 'reviews' || 
        typeOfPlacement.indexOf('отзыв') !== -1) {
      result.reviews.push(processedRow);
      result.statistics.totalReviews++;
    } 
    else if (currentSection === 'comments' || currentSection === 'discussions' ||
             typeOfPlacement.indexOf('комментар') !== -1 || 
             typeOfPlacement.indexOf('обсужден') !== -1 ||
             typeOfPlacement.indexOf('форум') !== -1) {
      // Собираем все целевые записи для сортировки
      allTargetedRecords.push(processedRow);
    }
  }
  
  // Сортируем целевые записи по просмотрам
  allTargetedRecords.sort(function(a, b) { 
    return (b.views || 0) - (a.views || 0); 
  });
  
  // Первые 20 - комментарии топ-20
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
  var totalViews = 0;
  var allRecords = result.reviews.concat(result.commentsTop20).concat(result.activeDiscussions);
  for (var k = 0; k < allRecords.length; k++) {
    if (allRecords[k].views && allRecords[k].views > 0) {
      totalViews += allRecords[k].views;
    }
  }
  result.statistics.totalViews = totalViews;
  
  // Считаем долю вовлечения
  result.statistics.engagementShare = calculateEngagementRate(result.commentsTop20.concat(result.activeDiscussions));
  
  console.log('📊 Итоги обработки:');
  console.log('   - Отзывов: ' + result.statistics.totalReviews);
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
  
  console.log('📄 Отчет создан: ' + reportName);
  return newSpreadsheet.getUrl();
}

// ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

function detectMonth(sheetName, data) {
  var monthFromSheet = extractMonthFromText(sheetName);
  if (monthFromSheet) {
    return monthFromSheet;
  }
  
  // Проверяем первые строки
  for (var i = 0; i < Math.min(3, data.length); i++) {
    var rowText = data[i].join(' ');
    var monthFromData = extractMonthFromText(rowText);
    if (monthFromData) {
      return monthFromData;
    }
  }
  
  // По умолчанию
  var now = new Date();
  var monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  return {
    name: monthNames[now.getMonth()],
    number: now.getMonth() + 1,
    year: now.getFullYear()
  };
}

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

function isEmptyRow(row) {
  if (!row) return true;
  for (var i = 0; i < row.length; i++) {
    if (row[i] && String(row[i]).trim() !== '') return false;
  }
  return true;
}

function isStatisticsRow(row) {
  if (!row || row.length === 0) return false;
  var firstCell = String(row[0] || '').toLowerCase();
  return firstCell.indexOf('суммарное количество просмотров') !== -1 || 
         firstCell.indexOf('количество карточек товара') !== -1 ||
         firstCell.indexOf('количество обсуждений') !== -1;
}

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

function extractViewsReceived(row, columnMapping) {
  var index = columnMapping.viewsReceived; // Колонка M
  if (index !== undefined && row[index]) {
    var viewsValue = row[index];
    
    if (typeof viewsValue === 'number' && !isNaN(viewsValue)) {
      return Math.max(0, Math.floor(viewsValue));
    }
    
    var viewsStr = String(viewsValue).trim();
    
    // Обработка разных форматов: "10 тыс", "20 тыс", "4 тыс"
    if (viewsStr.indexOf('тыс') !== -1) {
      var num = parseFloat(viewsStr.replace(/[^\d.,]/g, '').replace(',', '.'));
      if (!isNaN(num)) {
        return Math.floor(num * 1000);
      }
    }
    
    // Обычные числа
    var cleanStr = viewsStr.replace(/[^\d.,]/g, '');
    var normalizedStr = cleanStr.replace(',', '.');
    var parsed = parseFloat(normalizedStr);
    
    if (!isNaN(parsed)) {
      return Math.max(0, Math.floor(parsed));
    }
  }
  return 0;
}

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
    .createMenu('📊 Обработка V8')
    .addItem('🚀 Обработать ТЕКУЩИЙ лист', 'processMonthlyReport')
    .addItem('📋 Информация о листе', 'showCurrentSheetInfo')
    .addToUi();
}

/**
 * Показать информацию о текущем листе
 */
function showCurrentSheetInfo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getName();
  var data = sheet.getDataRange().getValues();
  
  var message = 'Текущий лист: ' + sheetName + '\n' +
                'Строк: ' + data.length + '\n' +
                'Колонок: ' + (data[0] ? data[0].length : 0) + '\n\n' +
                'Этот лист будет обработан при запуске процессора.';
  
  showMessage('Информация о листе', message);
}