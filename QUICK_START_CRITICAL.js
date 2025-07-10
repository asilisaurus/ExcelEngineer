/**
 * ⚡ БЫСТРЫЙ СТАРТ КРИТИЧЕСКОЙ МИССИИ
 * Мгновенный запуск тестирования Google Apps Script
 * 
 * Автор: AI Assistant
 * Приоритет: КРИТИЧЕСКИ ВЫСОКИЙ
 * Время: МАКСИМАЛЬНО БЫСТРО
 */

// ==================== КРИТИЧЕСКИЙ КОНФИГ ====================

const CRITICAL_CONFIG = {
  // URL данных (КРИТИЧЕСКИ ВАЖНЫЕ)
  SOURCE_URL: 'https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing',
  REFERENCE_URL: 'https://docs.google.com/spreadsheets/d/1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk/edit?',
  
  // Критические настройки
  TARGET_SIMILARITY: 0.95, // 95% - КРИТИЧЕСКИ ВАЖНО
  MAX_RETRIES: 5,
  TIMEOUT_SECONDS: 300,
  
  // Структура данных (ОСНОВАНА НА АНАЛИЗЕ)
  STRUCTURE: {
    headerRow: 4,        // КРИТИЧЕСКИ ВАЖНО
    dataStartRow: 5,     // КРИТИЧЕСКИ ВАЖНО
    infoRows: [1, 2, 3]
  }
};

// ==================== КРИТИЧЕСКИЙ ЗАПУСК ====================

/**
 * ⚡ МГНОВЕННЫЙ ЗАПУСК КРИТИЧЕСКОЙ МИССИИ
 */
function launchCriticalMission() {
  console.log('🚨 ЗАПУСК КРИТИЧЕСКОЙ МИССИИ!');
  console.log('🎯 ЦЕЛЬ: 95%+ СОВПАДЕНИЕ');
  console.log('⏰ ВРЕМЯ: МАКСИМАЛЬНО БЫСТРО');
  console.log('💪 РЕСУРСЫ: МАКСИМУМ');
  
  const startTime = Date.now();
  
  try {
    // 1. Быстрая проверка доступа
    console.log('🔍 Проверка доступа к данным...');
    const sourceAccess = testDataAccess(CRITICAL_CONFIG.SOURCE_URL);
    const referenceAccess = testDataAccess(CRITICAL_CONFIG.REFERENCE_URL);
    
    if (!sourceAccess || !referenceAccess) {
      throw new Error('КРИТИЧЕСКАЯ ОШИБКА: Нет доступа к данным!');
    }
    
    console.log('✅ Доступ к данным подтвержден');
    
    // 2. Запуск процессора
    console.log('🚀 Запуск финального процессора...');
    const processor = new FinalMonthlyReportProcessor();
    
    // 3. Быстрое тестирование всех месяцев
    const months = ['Февраль', 'Март', 'Апрель', 'Май'];
    const results = {};
    
    for (const month of months) {
      console.log(`📅 КРИТИЧЕСКОЕ ТЕСТИРОВАНИЕ: ${month} 2025`);
      
      const monthResult = testMonthCritical(processor, month);
      results[month] = monthResult;
      
      // Критическая проверка
      if (monthResult.similarity < CRITICAL_CONFIG.TARGET_SIMILARITY) {
        console.log(`⚠️ КРИТИЧЕСКОЕ ПРЕДУПРЕЖДЕНИЕ: ${month} - ${(monthResult.similarity * 100).toFixed(1)}% < 95%`);
        
        // Попытка исправления
        const fixedResult = attemptCriticalFix(processor, month);
        if (fixedResult.similarity >= CRITICAL_CONFIG.TARGET_SIMILARITY) {
          console.log(`✅ ИСПРАВЛЕНО: ${month} - ${(fixedResult.similarity * 100).toFixed(1)}%`);
          results[month] = fixedResult;
        }
      } else {
        console.log(`✅ УСПЕХ: ${month} - ${(monthResult.similarity * 100).toFixed(1)}%`);
      }
    }
    
    // 4. Критический анализ результатов
    const analysis = analyzeCriticalResults(results);
    
    // 5. Генерация критического отчета
    const reportUrl = generateCriticalReport(results, analysis, startTime);
    
    console.log('🎉 КРИТИЧЕСКАЯ МИССИЯ ЗАВЕРШЕНА!');
    console.log(`📊 Общий результат: ${(analysis.overallSimilarity * 100).toFixed(1)}%`);
    console.log(`📄 Отчет: ${reportUrl}`);
    
    return {
      success: analysis.overallSimilarity >= CRITICAL_CONFIG.TARGET_SIMILARITY,
      overallSimilarity: analysis.overallSimilarity,
      results: results,
      reportUrl: reportUrl,
      processingTime: Date.now() - startTime
    };
    
  } catch (error) {
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА:', error);
    return {
      success: false,
      error: error.toString(),
      processingTime: Date.now() - startTime
    };
  }
}

/**
 * ⚡ Критическое тестирование месяца
 */
function testMonthCritical(processor, month) {
  try {
    console.log(`🔍 Тестирование ${month} 2025...`);
    
    // Быстрое получение данных
    const sourceData = getMonthDataCritical(CRITICAL_CONFIG.SOURCE_URL, month);
    const referenceData = getMonthDataCritical(CRITICAL_CONFIG.REFERENCE_URL, month);
    
    if (!sourceData || !referenceData) {
      return {
        month: month,
        similarity: 0,
        error: 'Данные не найдены',
        status: 'FAILED'
      };
    }
    
    // Обработка процессором
    const processedResult = processor.processReport(sourceData.spreadsheetId, sourceData.sheetName);
    
    if (!processedResult.success) {
      return {
        month: month,
        similarity: 0,
        error: processedResult.error,
        status: 'FAILED'
      };
    }
    
    // Критическое сравнение
    const comparison = compareCritical(processedResult, referenceData);
    
    return {
      month: month,
      similarity: comparison.similarity,
      details: comparison.details,
      status: comparison.similarity >= CRITICAL_CONFIG.TARGET_SIMILARITY ? 'PASSED' : 'FAILED',
      processedData: processedResult.statistics,
      referenceData: referenceData.statistics
    };
    
  } catch (error) {
    return {
      month: month,
      similarity: 0,
      error: error.toString(),
      status: 'FAILED'
    };
  }
}

/**
 * ⚡ Получение данных месяца (критически быстро)
 */
function getMonthDataCritical(spreadsheetUrl, month) {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    const sheets = spreadsheet.getSheets();
    
    // Поиск листа с нужным месяцем
    const targetSheet = sheets.find(sheet => {
      const sheetName = sheet.getName().toLowerCase();
      const monthLower = month.toLowerCase();
      return sheetName.includes(monthLower) || 
             sheetName.includes(monthLower.substring(0, 3));
    });
    
    if (!targetSheet) {
      console.log(`⚠️ Лист для ${month} не найден`);
      return null;
    }
    
    return {
      spreadsheetId: spreadsheet.getId(),
      sheetName: targetSheet.getName(),
      data: targetSheet.getDataRange().getValues()
    };
    
  } catch (error) {
    console.error(`❌ Ошибка получения данных для ${month}:`, error);
    return null;
  }
}

/**
 * ⚡ Критическое сравнение результатов
 */
function compareCritical(processedResult, referenceData) {
  try {
    // Быстрое сравнение статистики
    const processedStats = processedResult.statistics;
    const referenceStats = extractStatsCritical(referenceData.data);
    
    let similarity = 0;
    let totalChecks = 0;
    
    // Сравнение количества записей
    if (referenceStats.reviewsCount > 0) {
      const reviewAccuracy = Math.min(processedStats.reviewsCount / referenceStats.reviewsCount, 1);
      similarity += reviewAccuracy;
      totalChecks++;
    }
    
    if (referenceStats.targetedCount > 0) {
      const targetedAccuracy = Math.min(processedStats.targetedCount / referenceStats.targetedCount, 1);
      similarity += targetedAccuracy;
      totalChecks++;
    }
    
    if (referenceStats.socialCount > 0) {
      const socialAccuracy = Math.min(processedStats.socialCount / referenceStats.socialCount, 1);
      similarity += socialAccuracy;
      totalChecks++;
    }
    
    if (referenceStats.totalViews > 0) {
      const viewsAccuracy = Math.min(processedStats.totalViews / referenceStats.totalViews, 1);
      similarity += viewsAccuracy;
      totalChecks++;
    }
    
    const finalSimilarity = totalChecks > 0 ? similarity / totalChecks : 0;
    
    return {
      similarity: finalSimilarity,
      details: {
        processedStats: processedStats,
        referenceStats: referenceStats,
        checks: totalChecks
      }
    };
    
  } catch (error) {
    return {
      similarity: 0,
      error: error.toString()
    };
  }
}

/**
 * ⚡ Извлечение статистики (критически быстро)
 */
function extractStatsCritical(data) {
  let reviews = 0;
  let targeted = 0;
  let social = 0;
  let totalViews = 0;
  
  let inReviewsSection = false;
  let inTargetedSection = false;
  let inSocialSection = false;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row.length === 0) continue;
    
    const firstCell = String(row[0]).toLowerCase();
    
    // Определение секций
    if (firstCell.includes('отзывы') || firstCell.includes('ос')) {
      inReviewsSection = true;
      inTargetedSection = false;
      inSocialSection = false;
      continue;
    }
    
    if (firstCell.includes('целевые') || firstCell.includes('цс')) {
      inReviewsSection = false;
      inTargetedSection = true;
      inSocialSection = false;
      continue;
    }
    
    if (firstCell.includes('площадки') || firstCell.includes('пс')) {
      inReviewsSection = false;
      inTargetedSection = false;
      inSocialSection = true;
      continue;
    }
    
    // Подсчет записей
    if (inReviewsSection && isDataRowCritical(row)) {
      reviews++;
      totalViews += extractViewsCritical(row);
    }
    
    if (inTargetedSection && isDataRowCritical(row)) {
      targeted++;
      totalViews += extractViewsCritical(row);
    }
    
    if (inSocialSection && isDataRowCritical(row)) {
      social++;
      totalViews += extractViewsCritical(row);
    }
  }
  
  return { reviews, targeted, social, totalViews };
}

/**
 * ⚡ Проверка строки данных (критически быстро)
 */
function isDataRowCritical(row) {
  if (row.length < 3) return false;
  const textCell = row[2] || row[1] || row[0];
  return textCell && String(textCell).trim().length > 10;
}

/**
 * ⚡ Извлечение просмотров (критически быстро)
 */
function extractViewsCritical(row) {
  if (row.length < 6) return 0;
  const viewsCell = row[5];
  if (!viewsCell) return 0;
  const viewsStr = String(viewsCell).replace(/[^\d]/g, '');
  const views = parseInt(viewsStr);
  return isNaN(views) ? 0 : views;
}

/**
 * ⚡ Попытка критического исправления
 */
function attemptCriticalFix(processor, month) {
  console.log(`🔧 Попытка исправления для ${month}...`);
  
  // Здесь можно добавить логику автоматического исправления
  // Например, настройка параметров, изменение логики и т.д.
  
  // Пока возвращаем исходный результат
  return {
    month: month,
    similarity: 0,
    status: 'FAILED',
    error: 'Автоматическое исправление не реализовано'
  };
}

/**
 * ⚡ Критический анализ результатов
 */
function analyzeCriticalResults(results) {
  let totalSimilarity = 0;
  let passedTests = 0;
  let totalTests = 0;
  
  for (const month in results) {
    const result = results[month];
    totalSimilarity += result.similarity;
    totalTests++;
    
    if (result.status === 'PASSED') {
      passedTests++;
    }
  }
  
  const overallSimilarity = totalTests > 0 ? totalSimilarity / totalTests : 0;
  const successRate = totalTests > 0 ? passedTests / totalTests : 0;
  
  return {
    overallSimilarity: overallSimilarity,
    successRate: successRate,
    passedTests: passedTests,
    totalTests: totalTests,
    targetAchieved: overallSimilarity >= CRITICAL_CONFIG.TARGET_SIMILARITY
  };
}

/**
 * ⚡ Генерация критического отчета
 */
function generateCriticalReport(results, analysis, startTime) {
  const reportData = [
    ['🚨 КРИТИЧЕСКИЙ ОТЧЕТ ТЕСТИРОВАНИЯ'],
    [''],
    [`Дата: ${new Date().toLocaleDateString('ru-RU')}`],
    [`Время: ${new Date().toLocaleTimeString('ru-RU')}`],
    [`Длительность: ${((Date.now() - startTime) / 1000).toFixed(2)} сек`],
    [''],
    ['ОБЩИЕ РЕЗУЛЬТАТЫ:'],
    [`Общее совпадение: ${(analysis.overallSimilarity * 100).toFixed(1)}%`],
    [`Цель достигнута: ${analysis.targetAchieved ? '✅ ДА' : '❌ НЕТ'}`],
    [`Успешных тестов: ${analysis.passedTests}/${analysis.totalTests}`],
    [''],
    ['ДЕТАЛЬНЫЕ РЕЗУЛЬТАТЫ:'],
    ['Месяц', 'Статус', 'Совпадение', 'Отзывы', 'Целевые', 'Социальные', 'Просмотры']
  ];
  
  for (const month in results) {
    const result = results[month];
    const statusIcon = result.status === 'PASSED' ? '✅' : '❌';
    
    reportData.push([
      month,
      `${statusIcon} ${result.status}`,
      `${(result.similarity * 100).toFixed(1)}%`,
      result.processedData?.reviewsCount || 'N/A',
      result.processedData?.targetedCount || 'N/A',
      result.processedData?.socialCount || 'N/A',
      result.processedData?.totalViews || 'N/A'
    ]);
  }
  
  // Создание отчета
  const reportSpreadsheet = SpreadsheetApp.create(`КРИТИЧЕСКИЙ_ОТЧЕТ_${new Date().toISOString().split('T')[0]}`);
  const reportSheet = reportSpreadsheet.getActiveSheet();
  
  reportSheet.getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);
  reportSheet.autoResizeColumns(1, reportData[0].length);
  
  return reportSpreadsheet.getUrl();
}

/**
 * ⚡ Тест доступа к данным
 */
function testDataAccess(url) {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(url);
    return spreadsheet !== null;
  } catch (error) {
    return false;
  }
}

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * ⚡ МГНОВЕННЫЙ ЗАПУСК
 */
function launchNow() {
  return launchCriticalMission();
}

/**
 * ⚡ Быстрая проверка
 */
function quickCheck() {
  console.log('🔍 Быстрая проверка конфигурации...');
  console.log(`📊 Исходные данные: ${CRITICAL_CONFIG.SOURCE_URL}`);
  console.log(`📊 Эталонные данные: ${CRITICAL_CONFIG.REFERENCE_URL}`);
  console.log(`🎯 Целевое совпадение: ${CRITICAL_CONFIG.TARGET_SIMILARITY * 100}%`);
  
  const sourceAccess = testDataAccess(CRITICAL_CONFIG.SOURCE_URL);
  const referenceAccess = testDataAccess(CRITICAL_CONFIG.REFERENCE_URL);
  
  console.log(`✅ Доступ к исходным данным: ${sourceAccess ? 'ДА' : 'НЕТ'}`);
  console.log(`✅ Доступ к эталонным данным: ${referenceAccess ? 'ДА' : 'НЕТ'}`);
  
  return {
    sourceAccess: sourceAccess,
    referenceAccess: referenceAccess,
    ready: sourceAccess && referenceAccess
  };
} 