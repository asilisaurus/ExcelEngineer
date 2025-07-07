const ExcelJS = require('exceljs');

async function compareDataReading() {
  console.log('🔍 СРАВНЕНИЕ ЧТЕНИЯ ДАННЫХ');
  console.log('='.repeat(50));
  
  try {
    const sourceFile = 'uploads/Fortedetrim ORM report source.xlsx';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourceFile);
    
    const worksheet = workbook.getWorksheet(1);
    
    console.log('\n📊 МЕТОД 1: Чтение как в отладочном скрипте (eachRow)');
    const dataMethod1 = [];
    
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && rowNumber <= 10) { // Первые 10 строк
        const rowData = [];
        for (let i = 1; i <= 15; i++) {
          const cell = row.getCell(i);
          rowData.push(cell.value);
        }
        dataMethod1.push(rowData);
        
        if (rowNumber >= 6 && rowNumber <= 8) { // Строки с данными
          console.log(`Строка ${rowNumber}:`);
          console.log(`  Колонка G (${6}): value="${rowData[6]}", type="${typeof rowData[6]}"`);
          console.log(`  Колонка H (${7}): value="${rowData[7]}", type="${typeof rowData[7]}"`);
        }
      }
    });
    
    console.log('\n📊 МЕТОД 2: Чтение как в реальном процессоре (getRows)');
    const dataMethod2 = [];
    
    // Имитируем чтение как в процессоре
    for (let rowIndex = 2; rowIndex <= 10; rowIndex++) {
      const row = worksheet.getRow(rowIndex);
      const rowData = [];
      
      for (let colIndex = 1; colIndex <= 15; colIndex++) {
        const cell = row.getCell(colIndex);
        rowData.push(cell.value);
      }
      
      dataMethod2.push(rowData);
      
      if (rowIndex >= 6 && rowIndex <= 8) { // Строки с данными
        console.log(`Строка ${rowIndex}:`);
        console.log(`  Колонка G (${6}): value="${rowData[6]}", type="${typeof rowData[6]}"`);
        console.log(`  Колонка H (${7}): value="${rowData[7]}", type="${typeof rowData[7]}"`);
      }
    }
    
    console.log('\n🔍 СРАВНЕНИЕ РЕЗУЛЬТАТОВ:');
    
    // Сравниваем данные из строки 6 (первая строка с данными)
    const row6Method1 = dataMethod1[4]; // индекс 4 = строка 6
    const row6Method2 = dataMethod2[4];
    
    console.log('\nСтрока 6 - Колонка G (дата):');
    console.log(`  Метод 1: "${row6Method1[6]}" (тип: ${typeof row6Method1[6]})`);
    console.log(`  Метод 2: "${row6Method2[6]}" (тип: ${typeof row6Method2[6]})`);
    console.log(`  Одинаковые: ${row6Method1[6] === row6Method2[6]}`);
    
    console.log('\nСтрока 6 - Колонка H (автор):');
    console.log(`  Метод 1: "${row6Method1[7]}" (тип: ${typeof row6Method1[7]})`);
    console.log(`  Метод 2: "${row6Method2[7]}" (тип: ${typeof row6Method2[7]})`);
    console.log(`  Одинаковые: ${row6Method1[7] === row6Method2[7]}`);
    
    // Проверяем как работает наша логика поиска дат с двумя методами
    console.log('\n🧪 ТЕСТИРОВАНИЕ ЛОГИКИ ПОИСКА ДАТ:');
    
    const testRow1 = row6Method1;
    const testRow2 = row6Method2;
    
    console.log('\nМетод 1 (отладочный скрипт):');
    testDateDetection(testRow1, 'Метод 1');
    
    console.log('\nМетод 2 (как в процессоре):');
    testDateDetection(testRow2, 'Метод 2');
    
  } catch (error) {
    console.error('❌ Ошибка:', error.message);
    console.error(error.stack);
  }
}

function testDateDetection(row, methodName) {
  const potentialDateColumns = [6, 3, 5]; // G, D, F
  
  for (const colIndex of potentialDateColumns) {
    const cellValue = row[colIndex];
    if (cellValue) {
      const cellStr = cellValue.toString();
      console.log(`  ${methodName} - Колонка ${String.fromCharCode(65 + colIndex)}: "${cellStr}"`);
      
      // Тестируем все проверки
      const isExcelNumber = typeof cellValue === 'number' && cellValue > 40000 && cellValue < 50000;
      const isDateSlash = cellStr.match(/\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4}/);
      const isDateFormat = cellStr.match(/(Mon|Tue|Wed|Thu|Fri|Sat|Sun).+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec).+\d{4}/);
      const isDateObject = cellValue instanceof Date || (typeof cellValue === 'object' && cellValue.toString().includes('GMT'));
      
      console.log(`    Excel number: ${isExcelNumber}`);
      console.log(`    Date slash: ${!!isDateSlash}`);
      console.log(`    Date format: ${!!isDateFormat}`);
      console.log(`    Date object: ${isDateObject}`);
      
      if (isExcelNumber || isDateSlash || isDateFormat || isDateObject) {
        console.log(`    ✅ ДАТА НАЙДЕНА!`);
        break;
      }
    }
  }
}

compareDataReading(); 