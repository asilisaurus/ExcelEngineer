const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function detailedValidation() {
  console.log('🔍 ДЕТАЛЬНАЯ ПРОВЕРКА РЕЗУЛЬТАТА АГЕНТА');
  console.log('=====================================\n');
  
  try {
    // Находим самый новый файл результата
    const uploadsDir = 'uploads';
    const files = fs.readdirSync(uploadsDir);
    
    const resultFiles = files.filter(file => 
      file.includes('результат') && file.endsWith('.xlsx') && !file.includes('temp_google_sheets')
    );
    
    const fileStats = resultFiles.map(file => ({
      name: file,
      path: path.join(uploadsDir, file),
      stats: fs.statSync(path.join(uploadsDir, file))
    })).sort((a, b) => b.stats.mtime - a.stats.mtime);
    
    if (fileStats.length === 0) {
      console.log('❌ Файлы результата не найдены');
      return;
    }
    
    const latestFile = fileStats[0];
    console.log(`📁 Анализируем: ${latestFile.name}`);
    console.log(`📊 Размер: ${latestFile.stats.size} байт`);
    console.log(`🕒 Создан: ${latestFile.stats.mtime.toISOString()}\n`);
    
    const workbook = XLSX.readFile(latestFile.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`📊 Лист: "${sheetName}"`);
    console.log(`📋 Всего строк: ${data.length}\n`);
    
    // Проверяем структуру заголовков
    console.log('🎯 ПРОВЕРКА СТРУКТУРЫ:');
    if (data[0] && data[0][0] === 'Продукт') {
      console.log('✅ Заголовок "Продукт" найден');
    } else {
      console.log('❌ ПРОБЛЕМА: Заголовок "Продукт" не найден!');
    }
    
    if (data[1] && data[1][0] === 'Период') {
      console.log('✅ Заголовок "Период" найден');
    } else {
      console.log('❌ ПРОБЛЕМА: Заголовок "Период" не найден!');
    }
    
    if (data[2] && data[2][0] === 'План') {
      console.log('✅ Заголовок "План" найден');
    } else {
      console.log('❌ ПРОБЛЕМА: Заголовок "План" не найден!');
    }
    
    if (data[3] && data[3][0] === 'Площадка') {
      console.log('✅ Заголовок "Площадка" найден');
    } else {
      console.log('❌ ПРОБЛЕМА: Заголовок "Площадка" не найден!');
    }
    
    console.log('\n📊 ПРОВЕРКА СЕКЦИЙ:');
    
    let reviewsFound = false;
    let commentsFound = false;
    let discussionsFound = false;
    let reviewsCount = 0;
    let commentsCount = 0;
    let discussionsCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        
        if (cellValue.includes('Отзывы') && !cellValue.includes('Количество')) {
          reviewsFound = true;
          console.log(`✅ Секция "Отзывы" найдена в строке ${i + 1}`);
        }
        
        if (cellValue.includes('Комментарии') && !cellValue.includes('Количество')) {
          commentsFound = true;
          console.log(`✅ Секция "Комментарии" найдена в строке ${i + 1}`);
        }
        
        if (cellValue.includes('Активные обсуждения') || cellValue.includes('Обсуждения')) {
          discussionsFound = true;
          console.log(`✅ Секция "Активные обсуждения" найдена в строке ${i + 1}`);
        }
      }
    }
    
    if (!reviewsFound) console.log('❌ ПРОБЛЕМА: Секция "Отзывы" не найдена!');
    if (!commentsFound) console.log('❌ ПРОБЛЕМА: Секция "Комментарии" не найдена!');
    if (!discussionsFound) console.log('❌ ПРОБЛЕМА: Секция "Активные обсуждения" не найдена!');
    
    console.log('\n📈 ПРОВЕРКА СТАТИСТИКИ:');
    
    // Ищем итоговые строки
    for (let i = data.length - 10; i < data.length; i++) {
      if (data[i] && data[i][0]) {
        const cellValue = data[i][0].toString();
        
        if (cellValue.includes('Количество отзывов')) {
          console.log(`✅ Найдена статистика отзывов: ${data[i][5] || 'N/A'}`);
        }
        
        if (cellValue.includes('Количество комментариев')) {
          console.log(`✅ Найдена статистика комментариев: ${data[i][5] || 'N/A'}`);
        }
        
        if (cellValue.includes('Количество обсуждений')) {
          console.log(`✅ Найдена статистика обсуждений: ${data[i][5] || 'N/A'}`);
        }
      }
    }
    
    console.log('\n🎯 ИТОГОВАЯ ОЦЕНКА:');
    
    const structureScore = (reviewsFound ? 1 : 0) + (commentsFound ? 1 : 0) + (discussionsFound ? 1 : 0);
    const structurePercentage = (structureScore / 3) * 100;
    
    console.log(`📊 Структура: ${structurePercentage.toFixed(1)}% (${structureScore}/3 секций)`);
    
    if (structurePercentage >= 95) {
      console.log('🎉 ОТЛИЧНО! Структура соответствует требованиям 95%+');
    } else if (structurePercentage >= 80) {
      console.log('⚠️ ХОРОШО, но нужны улучшения');
    } else {
      console.log('❌ КРИТИЧЕСКИЕ ПРОБЛЕМЫ! Нужны срочные исправления');
    }
    
    // Проверяем наличие данных
    let hasData = false;
    for (let i = 5; i < Math.min(data.length, 20); i++) {
      if (data[i] && data[i][0] && data[i][0].toString().includes('.')) {
        hasData = true;
        break;
      }
    }
    
    if (hasData) {
      console.log('✅ Данные присутствуют в файле');
    } else {
      console.log('❌ ПРОБЛЕМА: Данные отсутствуют!');
    }
    
    console.log('\n🔧 РЕКОМЕНДАЦИИ:');
    if (!reviewsFound) console.log('- Добавить секцию "Отзывы"');
    if (!commentsFound) console.log('- Добавить секцию "Комментарии"');
    if (!discussionsFound) console.log('- Добавить секцию "Активные обсуждения"');
    if (!hasData) console.log('- Проверить логику извлечения данных');
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
  }
}

detailedValidation();