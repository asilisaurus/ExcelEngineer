/**
 * ДЕТАЛЬНАЯ ДИАГНОСТИКА ДЛЯ ЛИСТА АПРЕЛЬ
 */

function diagnoseAprilSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var sheetName = sheet.getName();
    
    if (!sheetName.toLowerCase().includes('апр')) {
      showMessage('Внимание', 'Пожалуйста, переключитесь на лист Апрель и запустите снова');
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    var diagnostic = '=== ДЕТАЛЬНАЯ ДИАГНОСТИКА ЛИСТА ' + sheetName + ' ===\n\n';
    
    // 1. Ищем заголовки разделов
    diagnostic += '--- ПОИСК ЗАГОЛОВКОВ РАЗДЕЛОВ ---\n';
    var sectionHeaders = [];
    
    for (var i = 0; i < Math.min(100, data.length); i++) {
      var row = data[i];
      if (row && row[0]) {
        var firstCell = String(row[0]).toLowerCase();
        
        // Проверяем на заголовки разделов
        if (firstCell.indexOf('отзыв') !== -1 || 
            firstCell.indexOf('о т з ы в') !== -1 ||
            firstCell.indexOf('комментар') !== -1 ||
            firstCell.indexOf('обсужден') !== -1 ||
            firstCell.indexOf('топ') !== -1) {
          
          sectionHeaders.push({
            row: i + 1,
            text: row[0],
            fullRow: row.slice(0, 5).join(' | ')
          });
        }
      }
    }
    
    diagnostic += 'Найдено заголовков: ' + sectionHeaders.length + '\n';
    for (var j = 0; j < sectionHeaders.length; j++) {
      var header = sectionHeaders[j];
      diagnostic += '  Строка ' + header.row + ': "' + header.text + '"\n';
      diagnostic += '    Полная строка: ' + header.fullRow + '\n';
    }
    
    // 2. Анализируем данные между заголовками
    diagnostic += '\n--- АНАЛИЗ ДАННЫХ МЕЖДУ ЗАГОЛОВКАМИ ---\n';
    
    for (var k = 0; k < sectionHeaders.length; k++) {
      var startRow = sectionHeaders[k].row;
      var endRow = (k < sectionHeaders.length - 1) ? sectionHeaders[k + 1].row - 1 : Math.min(startRow + 50, data.length);
      
      diagnostic += '\nРаздел "' + sectionHeaders[k].text + '" (строки ' + startRow + '-' + endRow + '):\n';
      
      var dataCount = 0;
      var examples = [];
      
      for (var m = startRow; m < endRow && examples.length < 3; m++) {
        var dataRow = data[m];
        if (dataRow && dataRow[1] && String(dataRow[1]).trim() !== '') {
          dataCount++;
          if (examples.length < 3) {
            examples.push({
              row: m + 1,
              platform: dataRow[1] || '',
              text: dataRow[4] ? String(dataRow[4]).substring(0, 50) : '',
              views: dataRow[12] || dataRow[11] || ''
            });
          }
        }
      }
      
      diagnostic += '  Строк с данными: ' + dataCount + '\n';
      for (var n = 0; n < examples.length; n++) {
        var ex = examples[n];
        diagnostic += '  Пример ' + (n + 1) + ' (строка ' + ex.row + '):\n';
        diagnostic += '    Площадка: ' + ex.platform + '\n';
        diagnostic += '    Текст: ' + ex.text + '...\n';
        diagnostic += '    Просмотры: ' + ex.views + '\n';
      }
    }
    
    // 3. Проверяем колонку просмотров
    diagnostic += '\n--- АНАЛИЗ ПРОСМОТРОВ ---\n';
    var viewsColumns = [11, 12]; // L и M
    var viewsFound = {};
    
    for (var col = 0; col < viewsColumns.length; col++) {
      var colIndex = viewsColumns[col];
      var colLetter = String.fromCharCode(65 + colIndex);
      viewsFound[colLetter] = [];
      
      for (var p = 5; p < Math.min(50, data.length); p++) {
        var viewValue = data[p][colIndex];
        if (viewValue && String(viewValue).trim() !== '') {
          viewsFound[colLetter].push({
            row: p + 1,
            value: viewValue,
            type: typeof viewValue
          });
          
          if (viewsFound[colLetter].length >= 5) break;
        }
      }
    }
    
    for (var colKey in viewsFound) {
      diagnostic += '\nКолонка ' + colKey + ':\n';
      var examples = viewsFound[colKey];
      for (var q = 0; q < examples.length; q++) {
        var vEx = examples[q];
        diagnostic += '  Строка ' + vEx.row + ': "' + vEx.value + '" (тип: ' + vEx.type + ')\n';
      }
    }
    
    // 4. Структура листа
    diagnostic += '\n--- ОБЩАЯ СТРУКТУРА ---\n';
    diagnostic += 'Всего строк: ' + data.length + '\n';
    diagnostic += 'Всего колонок: ' + (data[0] ? data[0].length : 0) + '\n';
    
    // Заголовки в строке 4
    if (data.length >= 4) {
      diagnostic += '\nЗаголовки (строка 4):\n';
      var headers = data[3];
      for (var r = 0; r < Math.min(15, headers.length); r++) {
        if (headers[r]) {
          diagnostic += '  ' + String.fromCharCode(65 + r) + ': ' + headers[r] + '\n';
        }
      }
    }
    
    // Сохраняем диагностику
    var diagSheet = spreadsheet.insertSheet('Диагностика_Апрель_' + new Date().getTime());
    diagSheet.getRange('A1').setValue(diagnostic);
    diagSheet.getRange('A:A').setWrap(true);
    diagSheet.setColumnWidth(1, 800);
    
    console.log(diagnostic);
    showMessage('Диагностика завершена', 'Результаты сохранены в листе "Диагностика_Апрель_..."');
    
  } catch (error) {
    console.error('Ошибка: ' + error.toString());
    showMessage('Ошибка', error.toString());
  }
}

/**
 * Проверка конкретных строк
 */
function checkSpecificRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var message = 'Проверка строк 5-15:\n\n';
  
  for (var i = 4; i < Math.min(15, data.length); i++) {
    var row = data[i];
    message += 'Строка ' + (i + 1) + ':\n';
    message += '  A (Тип): ' + (row[0] || '[пусто]') + '\n';
    message += '  B (Площадка): ' + (row[1] || '[пусто]') + '\n';
    message += '  E (Текст): ' + (row[4] ? String(row[4]).substring(0, 30) + '...' : '[пусто]') + '\n';
    message += '  M (Просмотры): ' + (row[12] || '[пусто]') + '\n\n';
  }
  
  showMessage('Данные строк', message);
}

function showMessage(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.log(title + ': ' + message);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Диагностика Апрель')
    .addItem('📊 Полная диагностика', 'diagnoseAprilSheet')
    .addItem('🔎 Проверить строки 5-15', 'checkSpecificRows')
    .addToUi();
}