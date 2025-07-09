import { ExcelProcessorImprovedV2 } from './server/services/excel-processor-improved-v2.ts';
import fs from 'fs';
import path from 'path';

console.log('🔥 ТЕСТ: Критически обновленный процессор V2');
console.log('📋 Конфигурация: заголовки в строке 4, данные с строки 5');
console.log('🎯 Ожидаемые результаты: ~13 отзывов, ~15 комментариев, ~42 обсуждения\n');

async function testCriticalUpdateProcessor() {
  try {
    const processor = new ExcelProcessorImprovedV2();
    
    // Используем скачанный файл для анализа
    const sourceFile = 'source_structure_analysis.xlsx';
    
    if (!fs.existsSync(sourceFile)) {
      console.log(`❌ Файл ${sourceFile} не найден!`);
      return;
    }
    
    console.log(`📂 Тестируем файл: ${sourceFile}`);
    console.log(`📊 Размер файла: ${(fs.statSync(sourceFile).size / 1024 / 1024).toFixed(2)} MB\n`);
    
    const startTime = Date.now();
    
    // Обрабатываем файл с критическими обновлениями
    const result = await processor.processExcelFile(sourceFile);
    
    const processingTime = Date.now() - startTime;
    
    console.log('\n🎯 РЕЗУЛЬТАТЫ КРИТИЧЕСКОГО ОБНОВЛЕНИЯ:');
    console.log(`📄 Выходной файл: ${result.outputPath}`);
    console.log(`⏱️ Время обработки: ${processingTime}ms`);
    console.log('\n📊 СТАТИСТИКА:');
    console.log(`   📝 Всего строк: ${result.statistics.totalRows}`);
    console.log(`   ⭐ Отзывов: ${result.statistics.reviewsCount} (ожидалось: ~13)`);
    console.log(`   💬 Комментариев: ${result.statistics.commentsCount} (ожидалось: ~15)`);
    console.log(`   🔥 Активных обсуждений: ${result.statistics.activeDiscussionsCount} (ожидалось: ~42)`);
    console.log(`   👀 Общие просмотры: ${result.statistics.totalViews.toLocaleString()}`);
    console.log(`   📈 Вовлеченность: ${result.statistics.engagementRate.toFixed(2)}%`);
    console.log(`   🏆 Оценка качества: ${result.statistics.qualityScore.toFixed(1)}/10`);
    
    // Проверяем соответствие ожиданиям
    console.log('\n🔍 ПРОВЕРКА СООТВЕТСТВИЯ ОЖИДАНИЯМ:');
    
    const reviewsMatch = Math.abs(result.statistics.reviewsCount - 13) <= 3;
    const commentsMatch = Math.abs(result.statistics.commentsCount - 15) <= 5;
    const discussionsMatch = Math.abs(result.statistics.activeDiscussionsCount - 42) <= 10;
    
    console.log(`   ⭐ Отзывы: ${reviewsMatch ? '✅ СООТВЕТСТВУЕТ' : '❌ НЕ СООТВЕТСТВУЕТ'} (${result.statistics.reviewsCount} vs ~13)`);
    console.log(`   💬 Комментарии: ${commentsMatch ? '✅ СООТВЕТСТВУЕТ' : '❌ НЕ СООТВЕТСТВУЕТ'} (${result.statistics.commentsCount} vs ~15)`);
    console.log(`   🔥 Обсуждения: ${discussionsMatch ? '✅ СООТВЕТСТВУЕТ' : '❌ НЕ СООТВЕТСТВУЕТ'} (${result.statistics.activeDiscussionsCount} vs ~42)`);
    
    const overallMatch = reviewsMatch && commentsMatch && discussionsMatch;
    console.log(`\n🎯 ОБЩИЙ РЕЗУЛЬТАТ: ${overallMatch ? '✅ КРИТИЧЕСКОЕ ОБНОВЛЕНИЕ УСПЕШНО!' : '⚠️ ТРЕБУЕТСЯ ДОПОЛНИТЕЛЬНАЯ НАСТРОЙКА'}`);
    
    if (fs.existsSync(result.outputPath)) {
      console.log(`\n📄 Выходной файл создан: ${result.outputPath}`);
      console.log(`📊 Размер выходного файла: ${(fs.statSync(result.outputPath).size / 1024).toFixed(2)} KB`);
    }
    
  } catch (error) {
    console.error('❌ ОШИБКА ПРИ ТЕСТИРОВАНИИ:', error);
    console.error('📋 Детали ошибки:', error.message);
  }
}

// Запускаем тест
testCriticalUpdateProcessor();