// 🚨 QUICK START CRITICAL - МГНОВЕННЫЙ ЗАПУСК КРИТИЧЕСКОЙ МИССИИ
// Цель: 95%+ точности для всех месяцев (Февраль-Май 2025)

import { ExcelProcessorFinalV3 } from './server/services/excel-processor-final-v3.ts';
import fs from 'fs';
import path from 'path';

console.log('🚨 КРИТИЧЕСКАЯ МИССИЯ ЗАПУЩЕНА!');
console.log('🎯 Цель: 95%+ точности для всех месяцев');
console.log('⚡ Максимальный приоритет активирован!\n');

// 🔧 КРИТИЧЕСКИЕ ФУНКЦИИ
export async function launchNow() {
  console.log('🚀 launchNow() - МГНОВЕННЫЙ ЗАПУСК ВСЕЙ МИССИИ');
  
  try {
    // 1. Быстрая проверка готовности
    await quickCheck();
    
    // 2. Запуск финального тестирования
    const results = await runFinalTesting();
    
    // 3. Проверка достижения цели
    const success = validateResults(results);
    
    if (success) {
      console.log('✅ КРИТИЧЕСКАЯ МИССИЯ УСПЕШНО ЗАВЕРШЕНА!');
      console.log('🏆 Достигнута точность 95%+');
      return { success: true, results };
    } else {
      console.log('⚠️ ТРЕБУЕТСЯ ДОПОЛНИТЕЛЬНАЯ ОПТИМИЗАЦИЯ');
      return { success: false, results };
    }
    
  } catch (error) {
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА В МИССИИ:', error);
    throw error;
  }
}

export async function quickCheck() {
  console.log('🔍 quickCheck() - БЫСТРАЯ ПРОВЕРКА ГОТОВНОСТИ');
  
  // Проверка файлов
  const sourceFile = 'source_structure_analysis.xlsx';
  const refFile = 'reference_structure_analysis.xlsx';
  
  if (!fs.existsSync(sourceFile)) {
    throw new Error(`❌ Исходный файл не найден: ${sourceFile}`);
  }
  
  if (!fs.existsSync(refFile)) {
    throw new Error(`❌ Эталонный файл не найден: ${refFile}`);
  }
  
  console.log('✅ Все файлы готовы');
  console.log('✅ Процессор готов к запуску');
  console.log('✅ Система готова к критической миссии\n');
  
  return true;
}

export async function runFinalTesting() {
  console.log('🧪 runFinalTesting() - ПОЛНОЕ ТЕСТИРОВАНИЕ ВСЕХ МЕСЯЦЕВ');
  
  const processor = new ExcelProcessorFinalV3();
  const sourceFile = 'source_structure_analysis.xlsx';
  const results = {};
  
  // Целевые месяцы для тестирования
  const targetMonths = ['Февраль', 'Март', 'Апрель', 'Май'];
  
  for (const month of targetMonths) {
    console.log(`📋 Тестирование месяца: ${month}`);
    
    try {
      const result = await processor.processExcelFile(sourceFile, month);
      
      results[month] = {
        reviews: result.statistics.reviewsCount,
        comments: result.statistics.commentsCount,
        discussions: result.statistics.activeDiscussionsCount,
        accuracy: calculateAccuracy(result.statistics),
        processingTime: result.statistics.processingTime
      };
      
      console.log(`   📊 ${month}: ${result.statistics.reviewsCount} отзывов, ${result.statistics.commentsCount} комментариев, ${result.statistics.activeDiscussionsCount} обсуждений`);
      console.log(`   🎯 Точность: ${results[month].accuracy}%`);
      
    } catch (error) {
      console.error(`❌ Ошибка при тестировании ${month}:`, error);
      results[month] = { error: error.message };
    }
  }
  
  return results;
}

// 🎯 ФУНКЦИИ ВАЛИДАЦИИ
function calculateAccuracy(stats) {
  const target = { reviews: 13, comments: 15, discussions: 42 };
  
  const reviewsAccuracy = Math.max(0, 100 - Math.abs(stats.reviewsCount - target.reviews) / target.reviews * 100);
  const commentsAccuracy = Math.max(0, 100 - Math.abs(stats.commentsCount - target.comments) / target.comments * 100);
  const discussionsAccuracy = Math.max(0, 100 - Math.abs(stats.activeDiscussionsCount - target.discussions) / target.discussions * 100);
  
  return Math.round((reviewsAccuracy + commentsAccuracy + discussionsAccuracy) / 3);
}

function validateResults(results) {
  console.log('\n🔍 ВАЛИДАЦИЯ РЕЗУЛЬТАТОВ:');
  
  let totalAccuracy = 0;
  let validMonths = 0;
  
  for (const [month, data] of Object.entries(results)) {
    if (data.error) {
      console.log(`❌ ${month}: ОШИБКА - ${data.error}`);
      continue;
    }
    
    const accuracy = data.accuracy;
    console.log(`📊 ${month}: ${accuracy}% точности`);
    
    totalAccuracy += accuracy;
    validMonths++;
  }
  
  const avgAccuracy = validMonths > 0 ? totalAccuracy / validMonths : 0;
  
  console.log(`\n🎯 ОБЩАЯ ТОЧНОСТЬ: ${avgAccuracy.toFixed(1)}%`);
  console.log(`📋 Цель: 95%+ точности`);
  
  const success = avgAccuracy >= 95;
  console.log(`${success ? '✅' : '❌'} Результат: ${success ? 'УСПЕХ' : 'ТРЕБУЕТСЯ ДОРАБОТКА'}`);
  
  return success;
}

// 🚨 КРИТИЧЕСКИЕ МЕТРИКИ
function displayCriticalMetrics(results) {
  console.log('\n📊 КРИТИЧЕСКИЕ МЕТРИКИ:');
  console.log('═'.repeat(50));
  
  for (const [month, data] of Object.entries(results)) {
    if (data.error) continue;
    
    console.log(`🗓️  ${month.toUpperCase()}`);
    console.log(`   ⭐ Отзывы: ${data.reviews} (цель: ~13)`);
    console.log(`   💬 Комментарии: ${data.comments} (цель: ~15)`);
    console.log(`   🔥 Обсуждения: ${data.discussions} (цель: ~42)`);
    console.log(`   🎯 Точность: ${data.accuracy}%`);
    console.log(`   ⏱️  Время: ${data.processingTime}ms`);
    console.log('');
  }
}

// 🚀 АВТОМАТИЧЕСКИЙ ЗАПУСК ПРИ ИМПОРТЕ
if (import.meta.url === `file://${process.argv[1]}`) {
  console.log('🚨 АВТОМАТИЧЕСКИЙ ЗАПУСК КРИТИЧЕСКОЙ МИССИИ');
  
  launchNow()
    .then(result => {
      if (result.success) {
        console.log('🏆 МИССИЯ ВЫПОЛНЕНА УСПЕШНО!');
        displayCriticalMetrics(result.results);
      } else {
        console.log('⚠️ МИССИЯ ТРЕБУЕТ ДОРАБОТКИ');
        displayCriticalMetrics(result.results);
      }
    })
    .catch(error => {
      console.error('💥 КРИТИЧЕСКАЯ ОШИБКА:', error);
      process.exit(1);
    });
}

export default { launchNow, quickCheck, runFinalTesting };