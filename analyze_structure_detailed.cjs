const XLSX = require('xlsx');
const fs = require('fs');

console.log('🔍 ДЕТАЛЬНЫЙ АНАЛИЗ СТРУКТУРЫ ДЛЯ УЛУЧШЕНИЯ ПРОЦЕССОРА');
console.log('📊 Цель: Понять требования к гибкой обработке "грязных" данных\n');

function analyzeFile(filename, description) {
  console.log(`${'='.repeat(80)}`);
  console.log(`📋 АНАЛИЗ ${description.toUpperCase()}: ${filename}`);
  console.log(`${'='.repeat(80)}`);
  
  if (!fs.existsSync(filename)) {
    console.log(`❌ Файл ${filename} не найден!`);
    return null;
  }
  
  const buffer = fs.readFileSync(filename);
  const workbook = XLSX.read(buffer, { 
    type: 'buffer',
    cellDates: true,
    raw: false
  });
  
  console.log(`📑 Листы: ${workbook.SheetNames.join(', ')}`);
  
  // Анализируем каждый лист
  workbook.SheetNames.forEach(sheetName => {
    console.log(`\n📊 ЛИСТ: "${sheetName}"`);
    
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { 
      header: 1, 
      defval: '',
      raw: false
    });
    
    console.log(`   Строк: ${data.length}`);
    
    if (data.length > 0) {
      // Ищем заголовки в первых 10 строках
      console.log('\n   🔍 ПОИСК ЗАГОЛОВКОВ:');
      
      for (let i = 0; i < Math.min(10, data.length); i++) {
        const row = data[i];
        if (row && row.length > 0) {
          const hasHeaders = row.some(cell => {
            const cellStr = (cell || '').toString().toLowerCase();
            return cellStr.includes('тип размещения') || 
                   cellStr.includes('площадка') || 
                   cellStr.includes('текст сообщения') ||
                   cellStr.includes('тип поста');
          });
          
          if (hasHeaders) {
            console.log(`   ✅ ЗАГОЛОВКИ НАЙДЕНЫ В СТРОКЕ ${i + 1}:`);
            row.forEach((header, index) => {
              if (header) {
                console.log(`      ${String.fromCharCode(65 + index)} (${index}): "${header}"`);
              }
            });
            
            // Анализируем данные после заголовков
            console.log(`\n   📊 ПЕРВЫЕ 5 СТРОК ДАННЫХ:`);
            for (let j = i + 1; j < Math.min(i + 6, data.length); j++) {
              const dataRow = data[j];
              if (dataRow && dataRow.length > 0) {
                console.log(`   Строка ${j + 1}:`);
                console.log(`      Тип размещения: "${dataRow[0] || 'пусто'}"`);
                console.log(`      Площадка: "${dataRow[1] || 'пусто'}"`);
                console.log(`      Текст: "${(dataRow[4] || '').toString().substring(0, 50)}..."`);
                console.log(`      Тип поста: "${dataRow[14] || 'пусто'}"`);
              }
            }
            
            // Анализируем уникальные значения в ключевых колонках
            console.log(`\n   🎯 АНАЛИЗ СОДЕРЖИМОГО:`);
            
            // Анализ типов постов
            const postTypes = new Set();
            const placementTypes = new Set();
            let textCount = 0;
            let validRowCount = 0;
            
            for (let k = i + 1; k < data.length; k++) {
              const row = data[k];
              if (row && row.length > 0) {
                const postType = row[14];
                const placementType = row[0];
                const text = row[4];
                
                if (postType) postTypes.add(postType.toString().trim());
                if (placementType) placementTypes.add(placementType.toString().trim());
                if (text && text.toString().length > 10) textCount++;
                if ((text && text.toString().length > 0) || (row[1] && row[1].toString().length > 0)) {
                  validRowCount++;
                }
              }
            }
            
            console.log(`      Типы постов: ${Array.from(postTypes).join(', ')}`);
            console.log(`      Типы размещения: ${Array.from(placementTypes).slice(0, 5).join(', ')}${placementTypes.size > 5 ? '...' : ''}`);
            console.log(`      Строк с текстом: ${textCount}`);
            console.log(`      Всего валидных строк: ${validRowCount}`);
            
            return; // Нашли заголовки, переходим к следующему листу
          }
        }
      }
      
      console.log('   ⚠️ Заголовки не найдены в первых 10 строках');
    }
  });
  
  return workbook;
}

// Анализируем исходный файл
const sourceWorkbook = analyzeFile('source_file.xlsx', 'ИСХОДНЫЙ ФАЙЛ');

console.log('\n');

// Анализируем эталонный файл
const referenceWorkbook = analyzeFile('reference_file.xlsx', 'ЭТАЛОННЫЙ ФАЙЛ');

console.log('\n' + '='.repeat(80));
console.log('💡 ВЫВОДЫ ДЛЯ УЛУЧШЕНИЯ ПРОЦЕССОРА:');
console.log('='.repeat(80));

console.log(`
🎯 КЛЮЧЕВЫЕ ТРЕБОВАНИЯ К ГИБКОСТИ:

1. 📋 ДИНАМИЧЕСКОЕ ОПРЕДЕЛЕНИЕ ЗАГОЛОВКОВ:
   - НЕ полагаться на фиксированные позиции (избегать usecols='B:S')
   - Искать заголовки в разных строках (не только в первой)
   - Поддерживать вариации названий колонок

2. 🔍 ГИБКАЯ КЛАССИФИКАЦИЯ:
   - Анализировать "Тип поста" для определения ОС/ЦС/ПС
   - Использовать "Тип размещения" как резервный критерий
   - Поддерживать разные варианты написания

3. 📊 АДАПТИВНАЯ ОБРАБОТКА:
   - Обрабатывать строки с разной структурой
   - Пропускать пустые/служебные строки
   - Извлекать данные из любых позиций колонок

4. 🎨 СООТВЕТСТВИЕ ЭТАЛОНУ:
   - Проверить формат вывода против эталонного файла
   - Обеспечить правильную структуру результата
   - Сохранить все необходимые поля
`);

console.log('\n✅ АНАЛИЗ ЗАВЕРШЕН - ГОТОВ К УЛУЧШЕНИЮ ПРОЦЕССОРА!');