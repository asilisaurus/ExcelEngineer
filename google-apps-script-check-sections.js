/**
 * Проверка точного распределения типов записей в Апреле
 */

function checkRecordTypes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  console.log('=== АНАЛИЗ ТИПОВ ЗАПИСЕЙ ===');
  console.log('Лист: ' + sheet.getName());
  
  var types = {};
  var sections = [];
  var recordsByRow = [];
  
  // Поиск заголовков разделов и подсчет типов
  for (var i = 4; i < data.length; i++) {
    var row = data[i];
    var typeValue = row[0] ? String(row[0]).trim() : '';
    
    // Проверяем заголовок раздела
    var rowText = '';
    for (var j = 0; j < Math.min(5, row.length); j++) {
      if (row[j]) rowText += String(row[j]) + ' ';
    }
    
    if (rowText.toLowerCase().replace(/\s+/g, '').indexOf('отзывы') === 0) {
      sections.push({ row: i + 1, type: 'ЗАГОЛОВОК ОТЗЫВЫ' });
      continue;
    }
    
    if (rowText.toLowerCase().indexOf('комментарии') === 0 && 
        rowText.toLowerCase().indexOf('топ') !== -1) {
      sections.push({ row: i + 1, type: 'ЗАГОЛОВОК КОММЕНТАРИИ ТОП-20' });
      continue;
    }
    
    if (rowText.toLowerCase().indexOf('активные обсуждения') === 0) {
      sections.push({ row: i + 1, type: 'ЗАГОЛОВОК АКТИВНЫЕ ОБСУЖДЕНИЯ' });
      continue;
    }
    
    // Подсчитываем типы
    if (typeValue && typeValue !== '-') {
      if (!types[typeValue]) {
        types[typeValue] = { count: 0, firstRow: i + 1, lastRow: i + 1 };
      }
      types[typeValue].count++;
      types[typeValue].lastRow = i + 1;
      
      // Сохраняем для анализа
      if (recordsByRow.length < 100) {
        recordsByRow.push({
          row: i + 1,
          type: typeValue,
          hasViews: (row[11] && row[11] !== '-') || (row[10] && row[10] !== '-')
        });
      }
    }
  }
  
  console.log('\n=== НАЙДЕННЫЕ ЗАГОЛОВКИ РАЗДЕЛОВ ===');
  for (var s = 0; s < sections.length; s++) {
    console.log('Строка ' + sections[s].row + ': ' + sections[s].type);
  }
  
  console.log('\n=== ТИПЫ ЗАПИСЕЙ И ИХ КОЛИЧЕСТВО ===');
  var sortedTypes = Object.keys(types).sort(function(a, b) {
    return types[b].count - types[a].count;
  });
  
  for (var t = 0; t < sortedTypes.length; t++) {
    var typeName = sortedTypes[t];
    var typeInfo = types[typeName];
    console.log('"' + typeName + '": ' + typeInfo.count + 
                ' записей (строки ' + typeInfo.firstRow + '-' + typeInfo.lastRow + ')');
  }
  
  console.log('\n=== ПЕРВЫЕ ЗАПИСИ ПО СТРОКАМ ===');
  for (var r = 0; r < Math.min(20, recordsByRow.length); r++) {
    var rec = recordsByRow[r];
    console.log('Строка ' + rec.row + ': "' + rec.type + '"' + 
                (rec.hasViews ? ' [есть просмотры]' : ' [нет просмотров]'));
  }
  
  // Анализ проблемы
  console.log('\n=== АНАЛИЗ СТРУКТУРЫ ===');
  console.log('Похоже, что:');
  if (types['Комментарии в обсуждениях'] && types['Комментарии в обсуждениях'].count > 100) {
    console.log('- "Комментарии в обсуждениях" (' + types['Комментарии в обсуждениях'].count + 
                ' записей) - это НЕ топ-20, а активные обсуждения');
  }
  if (sections.length === 0) {
    console.log('- Заголовки разделов отсутствуют или написаны иначе');
  }
  
  console.log('\n=== ПРОВЕРКА ЗАВЕРШЕНА ===');
}

// Меню
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Проверка типов')
    .addItem('Проверить типы записей', 'checkRecordTypes')
    .addToUi();
}