/**
 * Диагностика проблем с топ-20 и просмотрами
 */

function checkTop20Problem() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  console.log('=== ДИАГНОСТИКА ТОП-20 ===');
  
  // Ищем строку с "ТОП-20"
  var top20Row = -1;
  for (var i = 0; i < data.length; i++) {
    var firstCell = data[i][0] ? String(data[i][0]).toLowerCase() : '';
    if (firstCell.indexOf('топ-20') !== -1 || firstCell.indexOf('топ 20') !== -1) {
      top20Row = i + 1;
      console.log('Найден заголовок ТОП-20 в строке ' + top20Row);
      break;
    }
  }
  
  if (top20Row === -1) {
    console.log('❌ Заголовок ТОП-20 не найден!');
    return;
  }
  
  // Проверяем следующие 25 строк после заголовка
  console.log('\n=== СТРОКИ ПОСЛЕ ТОП-20 ===');
  var commentCount = 0;
  var totalViews = 0;
  
  for (var j = top20Row; j < Math.min(top20Row + 25, data.length); j++) {
    var row = data[j - 1]; // Преобразуем в 0-based индекс
    var typeCol = row[0] ? String(row[0]).trim() : '';
    var platform = row[1] ? String(row[1]).trim() : '';
    var viewsL = row[11]; // Колонка L
    var viewsK = row[10]; // Колонка K
    var viewsJ = row[9];  // Колонка J
    
    // Определяем просмотры
    var views = 0;
    if (viewsL && viewsL !== '-') views = parseFloat(String(viewsL)) || 0;
    else if (viewsK && viewsK !== '-') views = parseFloat(String(viewsK)) || 0;
    else if (viewsJ && viewsJ !== '-') views = parseFloat(String(viewsJ)) || 0;
    
    // Проверяем, это данные или пустая строка
    var isEmpty = !typeCol && !platform;
    
    console.log('Строка ' + j + ': ' + 
                'Тип="' + typeCol.substring(0, 30) + '", ' +
                'Площадка="' + platform.substring(0, 30) + '", ' +
                'Просмотры=' + views + 
                (isEmpty ? ' [ПУСТАЯ]' : ''));
    
    if (!isEmpty && typeCol !== 'Тип размещения') {
      commentCount++;
      totalViews += views;
    }
    
    // Если нашли 20 записей или встретили новый заголовок
    if (commentCount >= 20 || 
        (typeCol && (typeCol.toLowerCase().indexOf('активные') !== -1 || 
                     typeCol.toLowerCase().indexOf('обсуждения') !== -1))) {
      break;
    }
  }
  
  console.log('\n=== ИТОГИ ===');
  console.log('Найдено комментариев после ТОП-20: ' + commentCount);
  console.log('Сумма просмотров в топ-20: ' + totalViews);
  
  // Проверяем общие просмотры
  console.log('\n=== ОБЩИЕ ПРОСМОТРЫ ===');
  var allViews = 0;
  var recordsWithViews = 0;
  
  for (var k = 4; k < data.length; k++) {
    var rowData = data[k];
    var viewsVal = 0;
    
    if (rowData[11] && rowData[11] !== '-') viewsVal = parseFloat(String(rowData[11])) || 0;
    else if (rowData[10] && rowData[10] !== '-') viewsVal = parseFloat(String(rowData[10])) || 0;
    else if (rowData[9] && rowData[9] !== '-') viewsVal = parseFloat(String(rowData[9])) || 0;
    
    if (viewsVal > 0) {
      allViews += viewsVal;
      recordsWithViews++;
    }
  }
  
  console.log('Всего просмотров в данных: ' + allViews);
  console.log('Записей с просмотрами: ' + recordsWithViews);
  
  // Ищем записи с большими просмотрами
  console.log('\n=== ЗАПИСИ С ПРОСМОТРАМИ > 1000 ===');
  for (var m = 4; m < Math.min(100, data.length); m++) {
    var rowCheck = data[m];
    var viewsCheck = 0;
    
    if (rowCheck[11] && rowCheck[11] !== '-') viewsCheck = parseFloat(String(rowCheck[11])) || 0;
    else if (rowCheck[10] && rowCheck[10] !== '-') viewsCheck = parseFloat(String(rowCheck[10])) || 0;
    else if (rowCheck[9] && rowCheck[9] !== '-') viewsCheck = parseFloat(String(rowCheck[9])) || 0;
    
    if (viewsCheck > 1000) {
      console.log('Строка ' + (m + 1) + ': ' + viewsCheck + ' просмотров, тип: "' + 
                  (rowCheck[0] || '') + '"');
    }
  }
}

// Меню
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Диагностика Топ-20')
    .addItem('Проверить топ-20 и просмотры', 'checkTop20Problem')
    .addToUi();
}