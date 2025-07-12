/**
 * 🧪 ТЕСТ ОБНОВЛЕННЫХ СКРИПТОВ
 * Проверка работы с эталонными листами в той же таблице
 * 
 * Автор: AI Assistant
 * Дата: 2025
 */

/**
 * Тест поиска эталонных листов
 */
function testReferenceSheetDetection() {
  console.log('🧪 ТЕСТ ПОИСКА ЭТАЛОННЫХ ЛИСТОВ');
  console.log('================================');
  
  try {
    const spreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing');
    const sheets = spreadsheet.getSheets();
    
    console.log(`📊 Всего листов в таблице: ${sheets.length}`);
    
    const sourceSheets = [];
    const referenceSheets = [];
    
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      console.log(`📋 Лист: "${sheetName}"`);
      
      // Проверяем, является ли эталонным
      const isReference = isReferenceSheet(sheetName);
      
      if (isReference) {
        referenceSheets.push(sheetName);
        console.log(`  ✅ Эталонный лист`);
      } else {
        // Проверяем, является ли исходным (содержит месяц)
        const monthInfo = detectMonthFromSheetName(sheetName);
        if (monthInfo) {
          sourceSheets.push({ name: sheetName, monthInfo: monthInfo });
          console.log(`  📅 Исходный лист: ${monthInfo.name} ${monthInfo.year}`);
          
          // Проверяем, есть ли соответствующий эталон
          const expectedReferenceName = `${monthInfo.name} ${monthInfo.year} (эталон)`;
          const hasReference = sheets.some(s => s.getName() === expectedReferenceName);
          console.log(`  ${hasReference ? '✅' : '❌'} Эталон: ${expectedReferenceName}`);
        }
      }
    }
    
    console.log('\n📊 РЕЗУЛЬТАТЫ:');
    console.log(`📅 Исходных листов: ${sourceSheets.length}`);
    console.log(`📋 Эталонных листов: ${referenceSheets.length}`);
    
    // Проверяем соответствие
    for (const source of sourceSheets) {
      const expectedReferenceName = `${source.monthInfo.name} ${source.monthInfo.year} (эталон)`;
      const hasReference = referenceSheets.includes(expectedReferenceName);
      console.log(`${hasReference ? '✅' : '❌'} ${source.name} -> ${expectedReferenceName}`);
    }
    
  } catch (error) {
    console.error('❌ Ошибка тестирования:', error);
  }
}

/**
 * Проверка, является ли лист эталонным
 */
function isReferenceSheet(sheetName) {
  const suffixes = [' (эталон)', ' (reference)', ' (etalon)'];
  return suffixes.some(suffix => sheetName.includes(suffix));
}

/**
 * Определение месяца из названия листа
 */
function detectMonthFromSheetName(sheetName) {
  const lowerSheetName = sheetName.toLowerCase();
  
  const months = [
    { name: 'Январь', short: 'Янв', number: 1 },
    { name: 'Февраль', short: 'Фев', number: 2 },
    { name: 'Март', short: 'Мар', number: 3 },
    { name: 'Апрель', short: 'Апр', number: 4 },
    { name: 'Май', short: 'Май', number: 5 },
    { name: 'Июнь', short: 'Июн', number: 6 },
    { name: 'Июль', short: 'Июл', number: 7 },
    { name: 'Август', short: 'Авг', number: 8 },
    { name: 'Сентябрь', short: 'Сен', number: 9 },
    { name: 'Октябрь', short: 'Окт', number: 10 },
    { name: 'Ноябрь', short: 'Ноя', number: 11 },
    { name: 'Декабрь', short: 'Дек', number: 12 }
  ];
  
  for (const month of months) {
    const monthVariants = [
      month.name.toLowerCase(),
      month.short.toLowerCase(),
      `${month.short}25`,
      `${month.name}25`,
      `${month.short}2025`,
      `${month.name}2025`
    ];
    
    if (monthVariants.some(variant => lowerSheetName.includes(variant))) {
      return {
        key: `${month.short}${month.year || 2025}`,
        name: month.name,
        short: month.short,
        number: month.number,
        year: 2025
      };
    }
  }
  
  return null;
}

/**
 * Тест автоопределения итоговой строки (УДАЛЕН - не нужен)
 * Итоговая строка определяется только тестером на основе эталона
 */

/**
 * Запуск всех тестов
 */
function runAllTests() {
  console.log('🚀 ЗАПУСК ВСЕХ ТЕСТОВ ОБНОВЛЕННЫХ СКРИПТОВ');
  console.log('==========================================');
  
  testReferenceSheetDetection();
  
  console.log('\n✅ Все тесты завершены');
} 