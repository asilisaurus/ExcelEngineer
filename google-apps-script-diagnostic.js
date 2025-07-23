/**
 * ДИАГНОСТИЧЕСКАЯ ВЕРСИЯ ПРОЦЕССОРА
 * Для выявления проблем с обработкой данных
 */

function runDiagnostics() {
  try {
    console.log('🔍 ЗАПУСК ДИАГНОСТИКИ...');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findDataSheet(spreadsheet);
    
    if (!sheet) {
      showMessage('Диагностика', 'Не найден лист с данными');
      return;
    }
    
    var sheetName = sheet.getName();
    console.log('📋 Выбран лист: ' + sheetName);
    
    var data = sheet.getDataRange().getValues();
    console.log('📊 Всего строк в листе: ' + data.length);
    
    // Проверяем структуру данных
    var diagnostic = '=== ДИАГНОСТИКА ДАННЫХ ===\n\n';
    diagnostic += 'Лист: ' + sheetName + '\n';
    diagnostic += 'Всего строк: ' + data.length + '\n\n';
    
    // Показываем первые 10 строк
    diagnostic += '--- Первые 10 строк ---\n';
    for (var i = 0; i < Math.min(10, data.length); i++) {
      var row = data[i];
      diagnostic += 'Строка ' + (i + 1) + ': ';
      
      // Показываем первые 5 колонок
      for (var j = 0; j < Math.min(5, row.length); j++) {
        var cellValue = row[j] ? String(row[j]).substring(0, 20) : '[пусто]';
        diagnostic += cellValue + ' | ';
      }
      diagnostic += '\n';
    }
    
    // Проверяем заголовки в строке 4
    diagnostic += '\n--- Заголовки (строка 4) ---\n';
    if (data.length >= 4) {
      var headers = data[3]; // строка 4 (индекс 3)
      for (var k = 0; k < headers.length; k++) {
        if (headers[k]) {
          diagnostic += 'Колонка ' + String.fromCharCode(65 + k) + ': ' + headers[k] + '\n';
        }
      }
    }
    
    // Проверяем колонку "Тип поста" (колонка N = индекс 13)
    diagnostic += '\n--- Проверка типов постов ---\n';
    var postTypes = {};
    var foundRows = 0;
    
    for (var m = 4; m < Math.min(100, data.length); m++) { // проверяем первые 100 строк
      var row = data[m];
      if (row && row[13]) { // колонка N
        var postType = String(row[13]).trim();
        if (postType) {
          postTypes[postType] = (postTypes[postType] || 0) + 1;
          foundRows++;
        }
      }
    }
    
    diagnostic += 'Найдено строк с типом поста: ' + foundRows + '\n';
    diagnostic += 'Типы постов:\n';
    for (var type in postTypes) {
      diagnostic += '  - ' + type + ': ' + postTypes[type] + ' раз\n';
    }
    
    // Проверяем колонку просмотров (колонка L = индекс 11)
    diagnostic += '\n--- Проверка просмотров ---\n';
    var viewsCount = 0;
    var totalViews = 0;
    
    for (var n = 4; n < Math.min(100, data.length); n++) {
      var row = data[n];
      if (row && row[11]) { // колонка L
        var views = parseFloat(String(row[11]).replace(/[^\d]/g, ''));
        if (!isNaN(views) && views > 0) {
          viewsCount++;
          totalViews += views;
        }
      }
    }
    
    diagnostic += 'Строк с просмотрами: ' + viewsCount + '\n';
    diagnostic += 'Сумма просмотров: ' + totalViews + '\n';
    
    // Показываем примеры строк с данными
    diagnostic += '\n--- Примеры строк с данными ---\n';
    var exampleCount = 0;
    for (var p = 4; p < data.length && exampleCount < 3; p++) {
      var row = data[p];
      if (row && row[13] && row[1]) { // есть тип поста и площадка
        exampleCount++;
        diagnostic += '\nПример ' + exampleCount + ' (строка ' + (p + 1) + '):\n';
        diagnostic += '  Площадка (B): ' + (row[1] || '[пусто]') + '\n';
        diagnostic += '  Текст (E): ' + String(row[4] || '[пусто]').substring(0, 50) + '...\n';
        diagnostic += '  Просмотры (L): ' + (row[11] || '[пусто]') + '\n';
        diagnostic += '  Тип поста (N): ' + (row[13] || '[пусто]') + '\n';
      }
    }
    
    // Записываем в лог и показываем
    console.log(diagnostic);
    
    // Создаем новый лист с диагностикой
    var diagSheet = spreadsheet.insertSheet('Диагностика_' + new Date().getTime());
    diagSheet.getRange('A1').setValue('ДИАГНОСТИКА ОБРАБОТКИ ДАННЫХ');
    diagSheet.getRange('A2').setValue(new Date());
    diagSheet.getRange('A4').setValue(diagnostic);
    diagSheet.getRange('A:A').setWrap(true);
    diagSheet.setColumnWidth(1, 800);
    
    showMessage('Диагностика завершена', 'Результаты сохранены в новом листе "Диагностика_..."');
    
  } catch (error) {
    console.error('Ошибка диагностики: ' + error.toString());
    showMessage('Ошибка', 'Ошибка диагностики: ' + error.toString());
  }
}

/**
 * Быстрая проверка данных
 */
function quickDataCheck() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findDataSheet(spreadsheet);
  
  if (!sheet) {
    showMessage('Проверка', 'Лист с данными не найден');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var mapping = getColumnMapping();
  
  var counts = {
    'ОС': 0,
    'ЦС': 0,
    'ПС': 0,
    'Другое': 0,
    'Пустые': 0
  };
  
  // Считаем типы постов
  for (var i = 4; i < data.length; i++) {
    var row = data[i];
    if (isEmptyRow(row)) {
      counts['Пустые']++;
      continue;
    }
    
    if (isStatisticsRow(row)) break;
    
    var postType = row[mapping.postType] ? String(row[mapping.postType]).trim().toUpperCase() : '';
    
    if (!postType) continue;
    
    if (postType === 'ОС') {
      counts['ОС']++;
    } else if (postType === 'ЦС') {
      counts['ЦС']++;
    } else if (postType === 'ПС') {
      counts['ПС']++;
    } else {
      counts['Другое']++;
    }
  }
  
  var message = 'Быстрая проверка данных:\n\n' +
                'Лист: ' + sheet.getName() + '\n' +
                'Всего строк: ' + data.length + '\n\n' +
                'Типы постов:\n' +
                '- ОС (отзывы): ' + counts['ОС'] + '\n' +
                '- ЦС (целевые): ' + counts['ЦС'] + '\n' +
                '- ПС (соцсети): ' + counts['ПС'] + '\n' +
                '- Другое: ' + counts['Другое'] + '\n' +
                '- Пустые строки: ' + counts['Пустые'];
  
  showMessage('Результаты проверки', message);
}

// ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

function findDataSheet(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var monthSheets = [];
  
  var monthPatterns = [
    'январ', 'феврал', 'март', 'апрел', 'май', 'мая',
    'июн', 'июл', 'август', 'сентябр', 'октябр', 'ноябр', 'декабр'
  ];
  
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
  
  if (monthSheets.length > 0) {
    monthSheets.sort(function(a, b) { return b.monthNumber - a.monthNumber; });
    return monthSheets[0].sheet;
  }
  
  return spreadsheet.getActiveSheet();
}

function getColumnMapping() {
  return {
    platform: 1,     // B
    theme: 3,        // D  
    text: 4,         // E
    date: 6,         // G
    author: 7,       // H
    views: 11,       // L
    engagement: 12,  // M
    postType: 13     // N
  };
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

function showMessage(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.log(title + ': ' + message);
  }
}

/**
 * Меню
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Диагностика')
    .addItem('📊 Полная диагностика', 'runDiagnostics')
    .addItem('⚡ Быстрая проверка', 'quickDataCheck')
    .addToUi();
}