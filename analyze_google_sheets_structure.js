/**
 * 🔍 АНАЛИЗ СТРУКТУРЫ GOOGLE SHEETS
 * Скрипт для анализа реальной структуры данных
 * URL: https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing
 */

// Извлекаем ID таблицы из URL
const SHEET_ID = '1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI';
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing';

console.log('🔍 АНАЛИЗ СТРУКТУРЫ GOOGLE SHEETS');
console.log('================================');
console.log(`📊 ID таблицы: ${SHEET_ID}`);
console.log(`🔗 URL: ${SHEET_URL}`);

try {
  // Открываем таблицу
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  
  console.log(`✅ Таблица открыта: ${spreadsheet.getName()}`);
  
  // Получаем все листы
  const sheets = spreadsheet.getSheets();
  console.log(`📋 Найдено листов: ${sheets.length}`);
  
  // Анализируем каждый лист
  sheets.forEach((sheet, index) => {
    const sheetName = sheet.getName();
    const maxRows = sheet.getMaxRows();
    const maxColumns = sheet.getMaxColumns();
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    console.log(`\n📄 ЛИСТ ${index + 1}: "${sheetName}"`);
    console.log(`   Размер: ${maxRows}x${maxColumns}`);
    console.log(`   Данные: ${lastRow}x${lastColumn}`);
    
    // Получаем данные листа
    if (lastRow > 0 && lastColumn > 0) {
      const data = sheet.getRange(1, 1, Math.min(lastRow, 20), lastColumn).getValues();
      
      console.log(`   Первые ${Math.min(data.length, 20)} строк:`);
      
      data.forEach((row, rowIndex) => {
        const rowData = row.map(cell => {
          if (cell === null || cell === undefined) return '';
          const str = cell.toString();
          return str.length > 30 ? str.substring(0, 30) + '...' : str;
        });
        console.log(`     ${rowIndex + 1}: [${rowData.join(', ')}]`);
      });
      
      // Анализируем структуру для поиска заголовков
      console.log(`\n   🔍 АНАЛИЗ ЗАГОЛОВКОВ для листа "${sheetName}"`);
      for (let i = 0; i < Math.min(10, data.length); i++) {
        const row = data[i];
        const rowText = row.join(' ').toLowerCase();
        
        // Поиск характерных паттернов
        const patterns = [
          'тип размещения',
          'площадка',
          'текст сообщения',
          'дата',
          'автор',
          'просмотры',
          'вовлечение',
          'тип поста'
        ];
        
        const foundPatterns = patterns.filter(pattern => rowText.includes(pattern));
        
        if (foundPatterns.length > 0) {
          console.log(`     Строка ${i + 1}: ВОЗМОЖНЫЕ ЗАГОЛОВКИ`);
          console.log(`     Найденные паттерны: ${foundPatterns.join(', ')}`);
          
          // Показываем маппинг колонок
          row.forEach((cell, colIndex) => {
            if (cell && cell.toString().trim()) {
              console.log(`       Колонка ${colIndex + 1}: "${cell}"`);
            }
          });
        }
      }
      
      // Поиск разделов данных
      console.log(`\n   📊 АНАЛИЗ РАЗДЕЛОВ для листа "${sheetName}"`);
      const sections = [];
      
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row.length > 0) {
          const firstCell = (row[0] || '').toString().toLowerCase().trim();
          
          if (firstCell.includes('отзыв') || firstCell.includes('ос') ||
              firstCell.includes('целевые') || firstCell.includes('цс') ||
              firstCell.includes('площадки') || firstCell.includes('пс') ||
              firstCell.includes('обсуждения') || firstCell.includes('комментарии')) {
            
            sections.push({
              row: i + 1,
              name: firstCell,
              type: firstCell.includes('отзыв') || firstCell.includes('ос') ? 'reviews' :
                    firstCell.includes('целевые') || firstCell.includes('цс') ? 'targeted' :
                    firstCell.includes('площадки') || firstCell.includes('пс') ? 'social' : 'other'
            });
          }
        }
      }
      
      console.log(`     Найдено разделов: ${sections.length}`);
      sections.forEach(section => {
        console.log(`       Строка ${section.row}: "${section.name}" (${section.type})`);
      });
    }
  });
  
  console.log('\n✅ АНАЛИЗ ЗАВЕРШЕН');
  console.log('Используйте эту информацию для настройки процессора!');
  
} catch (error) {
  console.error('❌ Ошибка при анализе:', error);
  console.error('Проверьте доступ к таблице!');
}