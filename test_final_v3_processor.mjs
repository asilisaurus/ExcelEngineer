import { ExcelProcessorFinalV3 } from './server/services/excel-processor-final-v3.ts';
import fs from 'fs';

console.log('🚨 КРИТИЧЕСКАЯ МИССИЯ: ТЕСТИРОВАНИЕ ФИНАЛЬНОГО ПРОЦЕССОРА V3');
console.log('🎯 Цель: ТОЧНОЕ соответствие 13/15/42 записей для достижения 95%+ точности');
console.log('⚡ Максимальный приоритет!\n');

async function testFinalV3Processor() {
  try {
    const processor = new ExcelProcessorFinalV3();
    
    const sourceFile = 'source_structure_analysis.xlsx';
    
    if (!fs.existsSync(sourceFile)) {
      console.log(`❌ Файл ${sourceFile} не найден!`);
      return;
    }
    
    console.log(`📂 Тестируем файл: ${sourceFile}`);
    console.log(`📊 Размер файла: ${(fs.statSync(sourceFile).size / 1024 / 1024).toFixed(2)} MB\n`);
    
    const startTime = Date.now();
    
    // Обрабатываем файл с критическими настройками V3
    const result = await processor.processExcelFile(sourceFile, 'source_structure_analysis.xlsx', 'Июнь25');
    
    const processingTime = Date.now() - startTime;
    
    console.log('\n🎯 РЕЗУЛЬТАТЫ КРИТИЧЕСКОЙ МИССИИ V3:');
    console.log(`📄 Выходной файл: ${result.outputPath}`);
    console.log(`⏱️ Время обработки: ${processingTime}ms`);
    console.log('\n📊 КРИТИЧЕСКИЕ МЕТРИКИ:');
    console.log(`   📝 Всего строк: ${result.statistics.totalRows}`);
    console.log(`   ⭐ Отзывов: ${result.statistics.reviewsCount} (цель: ТОЧНО 13)`);
    console.log(`   💬 Комментариев: ${result.statistics.commentsCount} (цель: ТОЧНО 15)`);
    console.log(`   🔥 Активных обсуждений: ${result.statistics.activeDiscussionsCount} (цель: ТОЧНО 42)`);
    console.log(`   👀 Общие просмотры: ${result.statistics.totalViews.toLocaleString()}`);
    console.log(`   📈 Вовлеченность: ${result.statistics.engagementRate.toFixed(2)}%`);
    console.log(`   🏆 Оценка качества: ${result.statistics.qualityScore.toFixed(1)}/100`);
    
    // Проверяем КРИТИЧЕСКИЕ МЕТРИКИ
    console.log('\n🔍 ПРОВЕРКА КРИТИЧЕСКИХ МЕТРИК:');
    
    const reviewsMatch = result.statistics.reviewsCount === 13;
    const commentsMatch = result.statistics.commentsCount === 15;
    const discussionsMatch = result.statistics.activeDiscussionsCount === 42;
    
    console.log(`   ⭐ Отзывы: ${reviewsMatch ? '✅ ТОЧНОЕ СООТВЕТСТВИЕ' : '❌ ОТКЛОНЕНИЕ'} (${result.statistics.reviewsCount}/13)`);
    console.log(`   💬 Комментарии: ${commentsMatch ? '✅ ТОЧНОЕ СООТВЕТСТВИЕ' : '❌ ОТКЛОНЕНИЕ'} (${result.statistics.commentsCount}/15)`);
    console.log(`   🔥 Обсуждения: ${discussionsMatch ? '✅ ТОЧНОЕ СООТВЕТСТВИЕ' : '❌ ОТКЛОНЕНИЕ'} (${result.statistics.activeDiscussionsCount}/42)`);
    
    // Подсчет точности
    const accuracy = result.statistics.qualityScore;
    console.log(`\n🎯 ОБЩАЯ ТОЧНОСТЬ: ${accuracy}%`);
    
    const missionSuccess = accuracy >= 95;
    console.log(`\n${missionSuccess ? '✅' : '❌'} РЕЗУЛЬТАТ КРИТИЧЕСКОЙ МИССИИ: ${missionSuccess ? '🏆 УСПЕХ - ЦЕЛЬ ДОСТИГНУТА!' : '⚠️ ТРЕБУЕТСЯ ДОПОЛНИТЕЛЬНАЯ НАСТРОЙКА'}`);
    
    // Дополнительная аналитика
    const exactMatch = reviewsMatch && commentsMatch && discussionsMatch;
    if (exactMatch) {
      console.log('🚀 ИДЕАЛЬНОЕ СООТВЕТСТВИЕ: ВСЕ КРИТИЧЕСКИЕ МЕТРИКИ ТОЧНЫ!');
    } else {
      console.log('⚠️ ЧАСТИЧНОЕ СООТВЕТСТВИЕ: ТРЕБУЕТСЯ ДОПОЛНИТЕЛЬНАЯ КАЛИБРОВКА');
    }
    
    if (fs.existsSync(result.outputPath)) {
      console.log(`\n📄 Выходной файл создан: ${result.outputPath}`);
      console.log(`📊 Размер выходного файла: ${(fs.statSync(result.outputPath).size / 1024).toFixed(2)} KB`);
    }
    
  } catch (error) {
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА ПРИ ТЕСТИРОВАНИИ:', error);
    console.error('📋 Детали ошибки:', error.message);
  }
}

// Запускаем критическую миссию
testFinalV3Processor();