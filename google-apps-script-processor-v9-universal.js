/**
 * 🚀 УНИВЕРСАЛЬНЫЙ ПРОЦЕССОР V9
 * Работает с ЛЮБЫМИ месяцами и структурами данных
 * 
 * Особенности:
 * - Автоматически находит колонку с просмотрами (проверяет L, M, K)
 * - Работает даже если просмотры не заполнены
 * - Адаптируется к разным названиям разделов
 * - Обрабатывает разные форматы данных
 */

// ==================== КОНФИГУРАЦИЯ ====================

var CONFIG = {
  STRUCTURE: {
    headerRow: 4,
    dataStartRow: 5,
    maxEmptyRows: 5  // После скольких пустых строк остановиться
  },
  
  // Возможные колонки для просмотров (проверяем по порядку)
  VIEWS_COLUMNS: {
    primary: 12,    // M - Просмотров получено
    secondary: 11,  // L - Просмотры в конце месяца
    tertiary: 10    // K - Просмотры темы на старте
  },
  
  // Шаблоны для определения типов
  PATTERNS: {
    REVIEWS: ['отзыв', 'о т з ы в'],
    COMMENTS: ['комментар', 'топ-20', 'топ 20'],
    DISCUSSIONS: ['обсужден', 'форум', 'сообщест', 'активные', 'мониторинг']
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
    console.log('🚀 UNIVERSAL PROCESSOR V9 - Начало обработки');
    
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
    console.log('🔍 Структура определена: колонка просмотров = ' + structure.viewsColumn);
    
    // Обработка данных
    var result = processDataUniversal(data, structure);
    
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
  
  // Ищем колонку с просмотрами
  console.log('🔍 Поиск колонки с просмотрами...');
  
  // Проверяем каждую возможную колонку
  var possibleColumns = [
    CONFIG.VIEWS_COLUMNS.primary,
    CONFIG.VIEWS_COLUMNS.secondary,
    CONFIG.VIEWS_COLUMNS.tertiary
  ];
  
  for (var i = 0; i < possibleColumns.length; i++) {
    var colIndex = possibleColumns[i];
    var foundViews = false;
    
    // Проверяем первые 50 строк данных
    for (var row = CONFIG.STRUCTURE.dataStartRow; row < Math.min(50, data.length); row++) {
      var value = data[row][colIndex];
      if (value && String(value).trim() !== '') {
        // Проверяем, похоже ли на число или "X тыс"
        var strValue = String(value);
        if (strValue.match(/\d/) || strValue.indexOf('тыс') !== -1) {
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
    console.log('⚠️ Колонка просмотров не найдена, будем работать без просмотров');
    // Используем колонку M по умолчанию
    structure.viewsColumn = CONFIG.VIEWS_COLUMNS.primary;
  }
  
  return structure;
}

/**
 * Универсальная обработка данных
 */
function processDataUniversal(data, structure) {
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
  
  var currentSection = null;
  var allTargetedRecords = [];
  var emptyRowCount = 0;
  
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
    
    var firstCell = row[0] ? String(row[0]).toLowerCase() : '';
    
    // Определяем раздел по заголовку
    if (isSectionHeader(firstCell)) {
      if (matchesPattern(firstCell, CONFIG.PATTERNS.REVIEWS)) {
        currentSection = 'reviews';
        console.log('📂 Раздел: Отзывы (строка ' + (i + 1) + ')');
      } else if (matchesPattern(firstCell, CONFIG.PATTERNS.COMMENTS)) {
        currentSection = 'comments';
        console.log('📂 Раздел: Комментарии Топ-20 (строка ' + (i + 1) + ')');
      } else if (matchesPattern(firstCell, CONFIG.PATTERNS.DISCUSSIONS)) {
        currentSection = 'discussions';
        console.log('📂 Раздел: Активные обсуждения (строка ' + (i + 1) + ')');
      }
      continue;
    }
    
    // Обрабатываем данные
    var typeOfPlacement = firstCell;
    var platform = row[structure.columnMapping.platform] ? String(row[structure.columnMapping.platform]).trim() : '';
    var text = row[structure.columnMapping.text] ? String(row[structure.columnMapping.text]).trim() : '';
    
    // Пропускаем строки без значимых данных
    if (!platform && !text && !typeOfPlacement) continue;
    
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
    
    // Классифицируем запись
    if (currentSection === 'reviews' || matchesPattern(typeOfPlacement, CONFIG.PATTERNS.REVIEWS)) {
      result.reviews.push(processedRow);
      result.statistics.totalReviews++;
    } else {
      // Все остальное собираем для сортировки
      allTargetedRecords.push(processedRow);
    }
  }
  
  // Сортируем записи по просмотрам (даже если они 0)
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
  // Заголовки обычно короткие и содержат ключевые слова
  if (text.length > 100) return false;
  
  var patterns = CONFIG.PATTERNS.REVIEWS
    .concat(CONFIG.PATTERNS.COMMENTS)
    .concat(CONFIG.PATTERNS.DISCUSSIONS);
    
  return matchesPattern(text, patterns);
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
  // Сначала пробуем по имени листа
  var monthFromSheet = extractMonthFromText(sheetName);
  if (monthFromSheet) {
    return monthFromSheet;
  }
  
  // Затем ищем в первых строках
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
        // Определяем год
        var year = 2025; // по умолчанию
        var yearMatch = text.match(/20\d{2}/);
        if (yearMatch) {
          year = parseInt(yearMatch[0]);
        } else if (text.match(/\d{2}/)) {
          // Если есть двузначный год
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
 * Извлечение просмотров (универсальная функция)
 */
function extractViews(row, viewsColumnIndex) {
  if (viewsColumnIndex === undefined || viewsColumnIndex < 0) return 0;
  
  var viewsValue = row[viewsColumnIndex];
  if (!viewsValue) return 0;
  
  // Если это число
  if (typeof viewsValue === 'number' && !isNaN(viewsValue)) {
    return Math.max(0, Math.floor(viewsValue));
  }
  
  // Если это строка
  var viewsStr = String(viewsValue).trim();
  if (viewsStr === '') return 0;
  
  // Обработка формата "X тыс"
  if (viewsStr.indexOf('тыс') !== -1) {
    var num = parseFloat(viewsStr.replace(/[^\d.,]/g, '').replace(',', '.'));
    if (!isNaN(num)) {
      return Math.floor(num * 1000);
    }
  }
  
  // Обработка формата "X млн"
  if (viewsStr.indexOf('млн') !== -1) {
    var numMln = parseFloat(viewsStr.replace(/[^\d.,]/g, '').replace(',', '.'));
    if (!isNaN(numMln)) {
      return Math.floor(numMln * 1000000);
    }
  }
  
  // Обычные числа
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
    .createMenu('📊 Универсальный процессор V9')
    .addItem('🚀 Обработать текущий лист', 'processMonthlyReport')
    .addItem('📋 Анализ структуры', 'analyzeCurrentSheet')
    .addItem('ℹ️ О процессоре', 'showAbout')
    .addToUi();
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

/**
 * Информация о процессоре
 */
function showAbout() {
  var message = 'УНИВЕРСАЛЬНЫЙ ПРОЦЕССОР V9\n\n' +
                'Возможности:\n' +
                '✅ Работает с любыми месяцами\n' +
                '✅ Автоматически находит колонку просмотров\n' +
                '✅ Работает даже без данных о просмотрах\n' +
                '✅ Адаптируется к разным форматам\n' +
                '✅ Обрабатывает "X тыс", "X млн" и числа\n\n' +
                'Инструкция:\n' +
                '1. Откройте лист с данными месяца\n' +
                '2. Запустите "Обработать текущий лист"\n' +
                '3. Получите готовый отчет';
  
  showMessage('О процессоре', message);
}