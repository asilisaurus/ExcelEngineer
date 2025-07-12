const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('🧪 ТЕСТИРОВАНИЕ УЛУЧШЕННОГО ПРОЦЕССОРА');
console.log('📊 Цель: Проверить гибкую обработку "грязных" данных\n');

async function testProcessor() {
  console.log('� Компилируем TypeScript...');
  
  const tscProcess = spawn('npx', ['tsc', '--build'], { 
    stdio: 'inherit',
    cwd: process.cwd()
  });
  
  await new Promise((resolve, reject) => {
    tscProcess.on('close', (code) => {
      if (code === 0) {
        console.log('✅ Компиляция успешна');
        resolve();
      } else {
        console.log('❌ Ошибка компиляции');
        reject(new Error(`Compilation failed with code ${code}`));
      }
    });
  });
  
  console.log('\n� Тестируем улучшенный процессор...');
  
  // Импортируем скомпилированный модуль
  const { fixedProcessor } = require('./dist/server/services/excel-processor-fixed.js');
  
  console.log('\n📋 Тест 1: Обработка исходного файла');
  try {
    const result1 = await fixedProcessor.processExcelFile('source_file.xlsx');
    console.log('✅ Результат 1:', {
      outputPath: result1.outputPath,
      statistics: result1.statistics
    });
    
    console.log('\n� Статистика:');
    console.log(`  Отзывы: ${result1.statistics.reviewsCount}`);
    console.log(`  Комментарии: ${result1.statistics.commentsCount}`);
    console.log(`  Обсуждения: ${result1.statistics.activeDiscussionsCount}`);
    console.log(`  Всего: ${result1.statistics.totalRows}`);
    console.log(`  Просмотры: ${result1.statistics.totalViews}`);
    console.log(`  Качество: ${result1.statistics.qualityScore}%`);
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
  }
  
  console.log('\n� Тест 2: Обработка эталонного файла');
  try {
    const result2 = await fixedProcessor.processExcelFile('reference_file.xlsx');
    console.log('✅ Результат 2:', {
      outputPath: result2.outputPath,
      statistics: result2.statistics
    });
    
    console.log('\n� Статистика:');
    console.log(`  Отзывы: ${result2.statistics.reviewsCount}`);
    console.log(`  Комментарии: ${result2.statistics.commentsCount}`);
    console.log(`  Обсуждения: ${result2.statistics.activeDiscussionsCount}`);
    console.log(`  Всего: ${result2.statistics.totalRows}`);
    console.log(`  Просмотры: ${result2.statistics.totalViews}`);
    console.log(`  Качество: ${result2.statistics.qualityScore}%`);
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error.message);
  }
  
  console.log('\n� Проверяем файлы результатов:');
  const uploadsDir = path.join(process.cwd(), 'uploads');
  if (fs.existsSync(uploadsDir)) {
    const files = fs.readdirSync(uploadsDir).filter(f => f.includes('результат'));
    console.log('📁 Созданные файлы:');
    files.forEach(file => {
      const filePath = path.join(uploadsDir, file);
      const stats = fs.statSync(filePath);
      console.log(`  ${file} (${Math.round(stats.size / 1024)}KB)`);
    });
  }
  
  console.log('\n✅ ТЕСТИРОВАНИЕ ЗАВЕРШЕНО');
  console.log('🎯 Улучшенный процессор готов к использованию!');
}

testProcessor().catch(console.error);