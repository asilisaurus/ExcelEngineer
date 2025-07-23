/**
 * Проверка где находятся данные о просмотрах в листе Апрель
 */

function checkViewsInApril() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  console.log('=== ПОИСК ПРОСМОТРОВ В АПРЕЛЕ ===');
  console.log('Лист: ' + sheet.getName());
  
  // Колонки где должны быть просмотры
  var viewColumns = {
    J: 9,  // Просмотры темы на старте
    K: 10, // Просмотры в конце месяца
    L: 11  // Просмотров получено
  };
  
  // Проверяем заголовки
  console.log('\n=== ЗАГОЛОВКИ КОЛОНОК ПРОСМОТРОВ ===');
  if (data.length > 3) {
    var headers = data[3];
    console.log('Колонка J: "' + (headers[9] || '') + '"');
    console.log('Колонка K: "' + (headers[10] || '') + '"');
    console.log('Колонка L: "' + (headers[11] || '') + '"');
  }
  
  // Ищем непустые значения в колонках просмотров
  console.log('\n=== ПОИСК ДАННЫХ В КОЛОНКАХ J, K, L ===');
  var foundData = {
    J: [],
    K: [],
    L: []
  };
  
  for (var row = 5; row < Math.min(100, data.length); row++) {
    // Колонка J
    if (data[row][9] && String(data[row][9]).trim() !== '') {
      foundData.J.push({
        row: row + 1,
        value: String(data[row][9]).substring(0, 30)
      });
    }
    
    // Колонка K
    if (data[row][10] && String(data[row][10]).trim() !== '') {
      foundData.K.push({
        row: row + 1,
        value: String(data[row][10]).substring(0, 30)
      });
    }
    
    // Колонка L
    if (data[row][11] && String(data[row][11]).trim() !== '') {
      foundData.L.push({
        row: row + 1,
        value: String(data[row][11]).substring(0, 30)
      });
    }
  }
  
  // Выводим результаты
  console.log('\nКолонка J (Просмотры на старте): ' + 
              (foundData.J.length > 0 ? 'найдено ' + foundData.J.length + ' значений' : 'ПУСТО'));
  if (foundData.J.length > 0) {
    for (var i = 0; i < Math.min(3, foundData.J.length); i++) {
      console.log('  Строка ' + foundData.J[i].row + ': "' + foundData.J[i].value + '"');
    }
  }
  
  console.log('\nКолонка K (Просмотры в конце): ' + 
              (foundData.K.length > 0 ? 'найдено ' + foundData.K.length + ' значений' : 'ПУСТО'));
  if (foundData.K.length > 0) {
    for (var j = 0; j < Math.min(3, foundData.K.length); j++) {
      console.log('  Строка ' + foundData.K[j].row + ': "' + foundData.K[j].value + '"');
    }
  }
  
  console.log('\nКолонка L (Просмотров получено): ' + 
              (foundData.L.length > 0 ? 'найдено ' + foundData.L.length + ' значений' : 'ПУСТО'));
  if (foundData.L.length > 0) {
    for (var k = 0; k < Math.min(3, foundData.L.length); k++) {
      console.log('  Строка ' + foundData.L[k].row + ': "' + foundData.L[k].value + '"');
    }
  }
  
  // Проверяем другие колонки на наличие чисел
  console.log('\n=== ПОИСК ЧИСЕЛ В ДРУГИХ КОЛОНКАХ ===');
  for (var col = 12; col < Math.min(20, data[0].length); col++) {
    var numbersFound = 0;
    var examples = [];
    
    for (var r = 5; r < Math.min(50, data.length); r++) {
      var val = data[r][col];
      if (val && typeof val === 'number' && val > 0) {
        numbersFound++;
        if (examples.length < 3) {
          examples.push({
            row: r + 1,
            value: val
          });
        }
      }
    }
    
    if (numbersFound > 0) {
      console.log('\nКолонка ' + String.fromCharCode(65 + col) + ': найдено ' + numbersFound + ' чисел');
      for (var e = 0; e < examples.length; e++) {
        console.log('  Строка ' + examples[e].row + ': ' + examples[e].value);
      }
    }
  }
  
  console.log('\n=== ПРОВЕРКА ЗАВЕРШЕНА ===');
}

// Меню
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Проверка просмотров')
    .addItem('Проверить просмотры в Апреле', 'checkViewsInApril')
    .addToUi();
}