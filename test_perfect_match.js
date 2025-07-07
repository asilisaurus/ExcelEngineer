import ExcelJS from 'exceljs';
import path from 'path';
import { simpleProcessor } from './server/services/excel-processor-simple.ts';

async function testPerfectMatch() {
  try {
    console.log('=== ТЕСТ ИДЕАЛЬНОГО СООТВЕТСТВИЯ С ЭТАЛОНОМ ===\n');
    
    // Тестовый файл
    const testFile = 'attached_assets/Фортедетрим_ORM_отчет_исходник_1751040742705.xlsx';
    
    console.log('🔄 Обрабатываем исходный файл...');
    const result = await simpleProcessor.processExcelFile(testFile);
    
    console.log('\n📊 РЕЗУЛЬТАТ ОБРАБОТКИ:');
    console.log(`📁 Выходной файл: ${result.outputPath}`);
    console.log('📈 Статистика:');
    console.log(`  - Всего записей: ${result.statistics.totalRows}`);
    console.log(`  - Отзывы (карточки товара): ${result.statistics.reviewsCount}`);
    console.log(`  - Комментарии (форумы, сообщества): ${result.statistics.commentsCount}`);
    console.log(`  - Активные обсуждения: ${result.statistics.activeDiscussionsCount}`);
    console.log(`  - Суммарные просмотры: ${result.statistics.totalViews.toLocaleString()}`);
    console.log(`  - Доля вовлечений: ${result.statistics.engagementRate}%`);
    console.log(`  - Площадки с данными: ${result.statistics.platformsWithData}%`);
    
    // Проверим структуру созданного файла
    console.log('\n🔍 ПРОВЕРКА СТРУКТУРЫ ФАЙЛА:');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(result.outputPath);
    const worksheet = workbook.getWorksheet(1);
    
    console.log(`📋 Всего строк: ${worksheet.rowCount}`);
    console.log(`📋 Всего столбцов: ${worksheet.columnCount}`);
    
    // Проверим заголовки
    console.log('\n📄 ЗАГОЛОВКИ:');
    console.log(`Строка 1: ${worksheet.getCell('A1').value} | ${worksheet.getCell('C1').value}`);
    console.log(`Строка 2: ${worksheet.getCell('A2').value} | ${worksheet.getCell('C2').value}`);
    console.log(`Строка 3: ${worksheet.getCell('A3').value} | ${worksheet.getCell('C3').value}`);
    
    // Проверим таблицу заголовков
    console.log('\n📋 ЗАГОЛОВКИ ТАБЛИЦЫ:');
    for (let col = 1; col <= 8; col++) {
      const cell = worksheet.getCell(4, col);
      console.log(`Столбец ${col}: ${cell.value}`);
    }
    
    // Проверим статистику в правой части
    console.log('\n📊 СТАТИСТИКА В ПРАВОЙ ЧАСТИ:');
    console.log(`I4 (Отзыв): ${worksheet.getCell('I4').value}`);
    console.log(`J4 (Упоминан): ${worksheet.getCell('J4').value}`);
    console.log(`K4 (Поддерживающ): ${worksheet.getCell('K4').value}`);
    console.log(`L4 (Всего): ${worksheet.getCell('L4').value}`);
    
    // Найдем строку "Итого"
    console.log('\n🔍 ПОИСК СТРОКИ "ИТОГО":');
    for (let row = 1; row <= worksheet.rowCount; row++) {
      const cellA = worksheet.getCell(row, 1);
      if (cellA.value && cellA.value.toString().toLowerCase().includes('итого')) {
        console.log(`Найдена строка "Итого" в строке ${row}:`);
        console.log(`  I${row}: ${worksheet.getCell(row, 9).value}`);
        console.log(`  J${row}: ${worksheet.getCell(row, 10).value}`);
        console.log(`  K${row}: ${worksheet.getCell(row, 11).value}`);
        console.log(`  L${row}: ${worksheet.getCell(row, 12).value}`);
        break;
      }
    }
    
    // Найдем и проверим финальную статистику
    console.log('\n🔍 ПРОВЕРКА ФИНАЛЬНОЙ СТАТИСТИКИ:');
    for (let row = worksheet.rowCount - 10; row <= worksheet.rowCount; row++) {
      const cellA = worksheet.getCell(row, 1);
      if (cellA.value) {
        const label = cellA.value.toString();
        const value = worksheet.getCell(row, 6).value;
        
        if (label.includes('Суммарное количество') || 
            label.includes('Количество карточек') || 
            label.includes('Количество обсуждений') || 
            label.includes('Доля обсуждений') || 
            label.includes('Площадки со статистикой')) {
          console.log(`Строка ${row}: ${label} = ${value}`);
        }
      }
    }
    
    console.log('\n✅ ТЕСТ ЗАВЕРШЕН!');
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
  }
}

testPerfectMatch(); 