/**
 * Расширенная диагностика листа Апрель для проверки заголовков и структуры
 */

function runAprilDiagnosticV2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  console.log('=== ДИАГНОСТИКА АПРЕЛЬ V2 ===');
  console.log('Лист: ' + sheet.getName());
  console.log('Всего строк: ' + data.length);
  
  // 1. Проверка заголовков разделов
  console.log('\n=== ПОИСК ЗАГОЛОВКОВ РАЗДЕЛОВ ===');
  var sectionHeaders = [];
  
  for (var i = 0; i < Math.min(100, data.length); i++) {
    var row = data[i];
    var firstCell = row[0] ? String(row[0]).trim() : '';
    
    // Проверяем различные варианты написания
    if (firstCell.toLowerCase().indexOf('отзыв') !== -1 ||
        firstCell.toLowerCase().indexOf('коммент') !== -1 ||
        firstCell.toLowerCase().indexOf('обсужден') !== -1 ||
        firstCell.toLowerCase() === 'отзывы' ||
        firstCell.toLowerCase() === 'комментарии топ-20 выдачи' ||
        firstCell.toLowerCase() === 'активные обсуждения (мониторинг)') {
      
      sectionHeaders.push({
        row: i + 1,
        text: firstCell,
        fullRow: row.slice(0, 5).map(function(cell) { return String(cell || '').substring(0, 30); })
      });
      
      console.log('Строка ' + (i + 1) + ': "' + firstCell + '"');
    }
  }
  
  if (sectionHeaders.length === 0) {
    console.log('❌ Заголовки разделов НЕ НАЙДЕНЫ!');
    
    // Показываем что находится в колонке A
    console.log('\n=== ПЕРВЫЕ 10 ЗНАЧЕНИЙ В КОЛОНКЕ A ===');
    for (var j = 4; j < Math.min(14, data.length); j++) {
      console.log('Строка ' + (j + 1) + ': "' + (data[j][0] || '') + '"');
    }
  }
  
  // 2. Анализ строки заголовков таблицы
  console.log('\n=== СТРОКА ЗАГОЛОВКОВ (строка 4) ===');
  if (data.length > 3) {
    var headerRow = data[3]; // строка 4
    for (var k = 0; k < Math.min(15, headerRow.length); k++) {
      if (headerRow[k]) {
        console.log('Колонка ' + String.fromCharCode(65 + k) + ': "' + headerRow[k] + '"');
      }
    }
  }
  
  // 3. Проверка первой строки данных
  console.log('\n=== ПЕРВАЯ СТРОКА ДАННЫХ (строка 6) ===');
  if (data.length > 5) {
    var firstDataRow = data[5]; // строка 6
    for (var m = 0; m < Math.min(15, firstDataRow.length); m++) {
      var value = firstDataRow[m] || '';
      var displayValue = String(value).substring(0, 50);
      console.log('Колонка ' + String.fromCharCode(65 + m) + ': "' + displayValue + '"');
    }
  }
  
  // 4. Поиск даты в неправильном формате
  console.log('\n=== ПРОВЕРКА ФОРМАТА ДАТ ===');
  var dateCount = 0;
  for (var n = 5; n < Math.min(20, data.length); n++) {
    for (var col = 0; col < data[n].length; col++) {
      var cellValue = String(data[n][col] || '');
      if (cellValue.match(/\d{1,2}\.\d{2}\.\d{4}/)) {
        dateCount++;
        if (dateCount <= 3) {
          console.log('Найдена дата в строке ' + (n + 1) + ', колонка ' + 
                      String.fromCharCode(65 + col) + ': "' + cellValue + '"');
        }
      }
    }
  }
  
  // 5. Поиск колонок с ОС/ЦС
  console.log('\n=== ПОИСК КОЛОНОК С ОС/ЦС ===');
  var oscsFound = false;
  for (var p = 0; p < Math.min(15, data[0].length); p++) {
    for (var q = 5; q < Math.min(20, data.length); q++) {
      var val = String(data[q][p] || '').trim().toUpperCase();
      if (val === 'ОС' || val === 'ЦС' || val === 'ПС') {
        console.log('Найдено "' + val + '" в колонке ' + String.fromCharCode(65 + p) + 
                    ', строка ' + (q + 1));
        oscsFound = true;
        break;
      }
    }
    if (oscsFound) break;
  }
  
  // 6. Проверка структуры по строке 5
  console.log('\n=== ПОЛНАЯ СТРОКА 5 (заголовок раздела?) ===');
  if (data.length > 4) {
    var row5 = data[4];
    console.log('Первые 5 ячеек:');
    for (var r = 0; r < Math.min(5, row5.length); r++) {
      console.log('  [' + String.fromCharCode(65 + r) + ']: "' + (row5[r] || '') + '"');
    }
  }
  
  console.log('\n=== ДИАГНОСТИКА ЗАВЕРШЕНА ===');
}

// Создаем меню
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Диагностика V2')
    .addItem('Запустить диагностику Апрель', 'runAprilDiagnosticV2')
    .addToUi();
}