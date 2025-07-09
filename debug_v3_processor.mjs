import * as XLSX from 'xlsx';
import fs from 'fs';

console.log('🔍 ДИАГНОСТИКА ПРОЦЕССОРА V3 - АНАЛИЗ ПРОБЛЕМЫ');
console.log('📊 Поиск причины 0 записей\n');

const sourceFile = 'source_structure_analysis.xlsx';

if (!fs.existsSync(sourceFile)) {
  console.log(`❌ Файл ${sourceFile} не найден!`);
  process.exit(1);
}

const buffer = fs.readFileSync(sourceFile);
const workbook = XLSX.read(buffer, { 
  type: 'buffer',
  cellDates: true,
  raw: false
});

console.log(`📋 Доступные листы: ${workbook.SheetNames.join(', ')}`);

// Анализируем лист Июнь25
const sheetName = 'Июнь25';
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { 
  header: 1, 
  defval: '',
  raw: false
});

console.log(`\n📊 Анализ листа "${sheetName}"`);
console.log(`📝 Всего строк: ${data.length}`);

// Проверяем заголовки в строке 4
console.log('\n🔍 ЗАГОЛОВКИ В СТРОКЕ 4:');
if (data[3]) {
  data[3].forEach((header, index) => {
    if (header) {
      console.log(`   ${index}: "${header}"`);
    }
  });
}

// Проверяем данные в первых 10 строках после заголовков
console.log('\n🔍 ПЕРВЫЕ 10 СТРОК ДАННЫХ (с строки 5):');
for (let i = 4; i < Math.min(14, data.length); i++) {
  const row = data[i];
  if (row && row.length > 0) {
    console.log(`\n   СТРОКА ${i + 1}:`);
    console.log(`     Тип размещения (A): "${row[0] || 'пусто'}"`);
    console.log(`     Площадка (B): "${row[1] || 'пусто'}"`);
    console.log(`     Текст (E): "${(row[4] || '').substring(0, 50)}..."`);
    console.log(`     Просмотры (L): "${row[11] || 'пусто'}"`);
    console.log(`     Тип поста (O): "${row[14] || 'пусто'}"`);
  }
}

// Анализируем уникальные значения в колонке "Тип поста"
console.log('\n🔍 УНИКАЛЬНЫЕ ЗНАЧЕНИЯ В КОЛОНКЕ "ТИП ПОСТА" (O):');
const postTypes = new Set();
for (let i = 4; i < data.length; i++) {
  const row = data[i];
  if (row && row[14]) {
    postTypes.add(row[14].toString().trim());
  }
}

console.log(`   Найдено уникальных типов: ${postTypes.size}`);
Array.from(postTypes).forEach(type => {
  console.log(`   - "${type}"`);
});

// Подсчитываем количество каждого типа
console.log('\n📊 КОЛИЧЕСТВО ЗАПИСЕЙ ПО ТИПАМ:');
const typeCounts = {};
for (let i = 4; i < data.length; i++) {
  const row = data[i];
  if (row && row[14]) {
    const type = row[14].toString().trim();
    typeCounts[type] = (typeCounts[type] || 0) + 1;
  }
}

Object.entries(typeCounts).forEach(([type, count]) => {
  console.log(`   "${type}": ${count} записей`);
});

// Проверяем записи с непустым текстом
console.log('\n🔍 АНАЛИЗ ЗАПИСЕЙ С ТЕКСТОМ:');
let withText = 0;
let withViews = 0;
let withValidData = 0;

for (let i = 4; i < data.length; i++) {
  const row = data[i];
  if (row) {
    const text = row[4] || '';
    const views = parseInt(row[11]) || 0;
    const platform = row[1] || '';
    const postType = row[14] || '';
    
    if (text.length > 0) withText++;
    if (views > 0) withViews++;
    if (text.length > 0 && platform.length > 0) withValidData++;
  }
}

console.log(`   Записей с текстом: ${withText}`);
console.log(`   Записей с просмотрами: ${withViews}`);
console.log(`   Записей с валидными данными: ${withValidData}`);

// Проверяем примеры записей ОС, ЦС, ПС
console.log('\n🔍 ПРИМЕРЫ ЗАПИСЕЙ ПО ТИПАМ:');
const examples = { 'ОС': [], 'ЦС': [], 'ПС': [] };

for (let i = 4; i < data.length; i++) {
  const row = data[i];
  if (row && row[14]) {
    const type = row[14].toString().trim().toUpperCase();
    if (examples[type] && examples[type].length < 3) {
      examples[type].push({
        platform: row[1] || '',
        text: (row[4] || '').substring(0, 100),
        views: parseInt(row[11]) || 0,
        type: row[14] || ''
      });
    }
  }
}

Object.entries(examples).forEach(([type, records]) => {
  console.log(`\n   ${type} (${records.length} примеров):`);
  records.forEach((record, index) => {
    console.log(`     ${index + 1}. Площадка: "${record.platform}"`);
    console.log(`        Текст: "${record.text}..."`);
    console.log(`        Просмотры: ${record.views}`);
    console.log(`        Тип: "${record.type}"`);
  });
});

console.log('\n🎯 ДИАГНОСТИКА ЗАВЕРШЕНА');