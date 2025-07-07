const fs = require('fs');

// Поскольку это CommonJS, а процессор в TypeScript, нужно компилировать или использовать ts-node
const { spawn } = require('child_process');

// Создаем простой тестовый скрипт через ts-node
const testScript = `
import fs from 'fs';
import { perfectProcessor } from './server/services/excel-processor-perfect';

async function testPerfectProcessor() {
  try {
    console.log('🚀 ТЕСТИРОВАНИЕ СОВЕРШЕННОГО ПРОЦЕССОРА');
    console.log('=====================================');
    
    // Читаем исходный файл
    const buffer = fs.readFileSync('uploads/Fortedetrim ORM report source.xlsx');
    
    console.log('📁 Исходный файл загружен:', buffer.length, 'байт');
    
    // Обрабатываем файл
    const result = await perfectProcessor.processExcelFile(buffer, 'Fortedetrim ORM report source.xlsx');
    
    console.log('✅ РЕЗУЛЬТАТЫ ОБРАБОТКИ:');
    console.log('=======================');
    console.log('📊 Статистика:');
    console.log('  - Общее количество строк:', result.statistics.totalRows);
    console.log('  - Количество отзывов:', result.statistics.reviewsCount);
    console.log('  - Количество комментариев:', result.statistics.commentsCount);
    console.log('  - Активные обсуждения:', result.statistics.activeDiscussionsCount);
    console.log('  - Общие просмотры:', result.statistics.totalViews);
    console.log('  - Процент вовлечения:', result.statistics.engagementRate + '%');
    console.log('  - Площадки с данными:', result.statistics.platformsWithData + '%');
    
    // Сохраняем результат
    const outputBuffer = await result.workbook.xlsx.writeBuffer();
    fs.writeFileSync('test_perfect_output.xlsx', outputBuffer);
    
    console.log('💾 Результат сохранен в test_perfect_output.xlsx');
    console.log('📈 Размер выходного файла:', outputBuffer.length, 'байт');
    
    console.log('\\n🎯 ТЕСТ ЗАВЕРШЕН УСПЕШНО!');
    
  } catch (error) {
    console.error('❌ ОШИБКА ТЕСТИРОВАНИЯ:', error);
    console.error('Stack:', error.stack);
  }
}

testPerfectProcessor();
`;

// Записываем временный TypeScript файл
fs.writeFileSync('temp_test.ts', testScript);

console.log('Запуск теста через ts-node...');

// Запускаем через ts-node
const tsNode = spawn('npx', ['ts-node', 'temp_test.ts'], { 
  stdio: 'inherit',
  shell: true 
});

tsNode.on('close', (code) => {
  console.log(`Тест завершен с кодом ${code}`);
  
  // Очищаем временный файл
  try {
    fs.unlinkSync('temp_test.ts');
  } catch (e) {
    // Игнорируем ошибки удаления
  }
});

tsNode.on('error', (error) => {
  console.error('Ошибка запуска теста:', error);
  
  // Очищаем временный файл
  try {
    fs.unlinkSync('temp_test.ts');
  } catch (e) {
    // Игнорируем ошибки удаления
  }
}); 