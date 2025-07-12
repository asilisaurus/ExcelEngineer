/**
 * 🧪 ТЕСТ ИСПРАВЛЕННОГО ПРОЦЕССОРА
 * Проверка исправлений в google-apps-script-processor-final.js
 * 
 * Автор: AI Assistant
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ ТЕСТИРОВАНИЯ ====================

const TEST_CONFIG = {
  // URL исходных данных
  SOURCE_URL: 'https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing',
  
  // Настройки тестирования
  TESTING: {
    ENABLE_DEBUG: true,
    LOG_DETAILS: true,
    VALIDATE_RESULTS: true
  }
};

// ==================== КЛАСС ТЕСТИРОВАНИЯ ====================

/**
 * Класс для тестирования исправленного процессора
 */
class FixedProcessorTester {
  constructor() {
    this.testResults = {
      totalTests: 0,
      passedTests: 0,
      failedTests: 0,
      details: [],
      processingTime: 0
    };
    
    this.processor = new FinalMonthlyReportProcessor();
  }

  /**
   * Запуск тестирования исправленного процессора
   */
  async runTesting() {
    const startTime = Date.now();
    
    console.log('🚀 ТЕСТИРОВАНИЕ ИСПРАВЛЕННОГО ПРОЦЕССОРА');
    console.log('==========================================');
    console.log(`📊 Исходные данные: ${TEST_CONFIG.SOURCE_URL}`);
    
    try {
      // 1. Подготовка данных
      console.log('\n📋 ПОДГОТОВКА ДАННЫХ...');
      const sourceData = await this.prepareSourceData();
      
      // 2. Тестирование обработки
      console.log('\n🧪 ТЕСТИРОВАНИЕ ОБРАБОТКИ...');
      await this.testProcessing(sourceData);
      
      // 3. Анализ результатов
      this.analyzeResults();
      
      this.testResults.processingTime = Date.now() - startTime;
      
    } catch (error) {
      console.error('❌ Критическая ошибка тестирования:', error);
      this.testResults.details.push({
        test: 'Общее тестирование',
        status: 'FAILED',
        error: error.toString()
      });
    }
  }

  /**
   * Подготовка исходных данных
   */
  async prepareSourceData() {
    console.log('📊 Загрузка исходных данных...');
    
    try {
      const spreadsheet = SpreadsheetApp.openByUrl(TEST_CONFIG.SOURCE_URL);
      const sheets = spreadsheet.getSheets();
      
      const sourceData = {};
      
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        
        // Пропускаем эталонные листы
        if (this.isReferenceSheet(sheetName)) {
          console.log(`⏭️ Пропускаем эталонный лист "${sheetName}"`);
          continue;
        }
        
        // Пропускаем не-месячные листы
        if (this.isNonMonthlySheet(sheetName)) {
          console.log(`⏭️ Пропускаем не-месячный лист "${sheetName}"`);
          continue;
        }
        
        const data = sheet.getDataRange().getValues();
        
        // Определяем месяц для каждого листа
        const monthInfo = this.detectMonthFromSheet(sheetName, data);
        if (monthInfo) {
          sourceData[monthInfo.key] = {
            sheet: sheet,
            data: data,
            monthInfo: monthInfo
          };
          
          console.log(`✅ Лист "${sheetName}" -> ${monthInfo.name} ${monthInfo.year}`);
        }
      }
      
      console.log(`📊 Загружено ${Object.keys(sourceData).length} листов с данными`);
      return sourceData;
      
    } catch (error) {
      throw new Error(`Ошибка загрузки исходных данных: ${error.message}`);
    }
  }

  /**
   * Проверка, является ли лист эталонным
   */
  isReferenceSheet(sheetName) {
    const referencePatterns = [
      ' (эталон)',
      ' (reference)', 
      ' (etalon)'
    ];
    
    return referencePatterns.some(pattern => 
      sheetName.includes(pattern)
    );
  }

  /**
   * Проверка, является ли лист не-месячным
   */
  isNonMonthlySheet(sheetName) {
    const excludedPatterns = [
      'бриф',
      'репутация',
      'медиаплан',
      'эталон',
      'reference',
      'etalon'
    ];
    
    const lowerSheetName = sheetName.toLowerCase();
    return excludedPatterns.some(pattern => 
      lowerSheetName.includes(pattern.toLowerCase())
    );
  }

  /**
   * Определение месяца из названия листа
   */
  detectMonthFromSheet(sheetName, data) {
    const lowerSheetName = sheetName.toLowerCase();
    
    // Определяем год из названия листа
    let detectedYear = 2025; // по умолчанию
    
    if (lowerSheetName.includes('2024') || lowerSheetName.match(/\b24\b/)) {
      detectedYear = 2024;
    } else if (lowerSheetName.includes('2023') || lowerSheetName.match(/\b23\b/)) {
      detectedYear = 2023;
    } else if (lowerSheetName.includes('2022') || lowerSheetName.match(/\b22\b/)) {
      detectedYear = 2022;
    }
    
    const months = [
      { name: 'Январь', short: 'Янв', number: 1 },
      { name: 'Февраль', short: 'Фев', number: 2 },
      { name: 'Март', short: 'Мар', number: 3 },
      { name: 'Апрель', short: 'Апр', number: 4 },
      { name: 'Май', short: 'Май', number: 5 },
      { name: 'Июнь', short: 'Июн', number: 6 },
      { name: 'Июль', short: 'Июл', number: 7 },
      { name: 'Август', short: 'Авг', number: 8 },
      { name: 'Сентябрь', short: 'Сен', number: 9 },
      { name: 'Октябрь', short: 'Окт', number: 10 },
      { name: 'Ноябрь', short: 'Ноя', number: 11 },
      { name: 'Декабрь', short: 'Дек', number: 12 }
    ];
    
    // Поиск месяца в названии листа
    for (const month of months) {
      const exactMatches = [
        month.name.toLowerCase(),
        month.short.toLowerCase(),
        `${month.short}${detectedYear.toString().slice(-2)}`,
        `${month.name}${detectedYear.toString().slice(-2)}`,
        `${month.short}${detectedYear}`,
        `${month.name}${detectedYear}`
      ];
      
      for (const exactMatch of exactMatches) {
        if (lowerSheetName === exactMatch || lowerSheetName.includes(exactMatch)) {
          return {
            key: `${month.short}${detectedYear}`,
            name: month.name,
            short: month.short,
            number: month.number,
            year: detectedYear,
            detectedFrom: 'sheet'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Тестирование обработки данных
   */
  async testProcessing(sourceData) {
    console.log('🔄 Тестирование обработки данных...');
    
    // Тестируем каждый найденный месяц
    const testMonths = Object.keys(sourceData);
    
    for (const monthKey of testMonths) {
      const sourceInfo = sourceData[monthKey];
      
      console.log(`\n📅 ТЕСТИРОВАНИЕ: ${sourceInfo.monthInfo.name} ${sourceInfo.monthInfo.year}`);
      console.log('='.repeat(60));
      
      await this.testMonthProcessing(sourceInfo);
    }
  }

  /**
   * Тестирование обработки конкретного месяца
   */
  async testMonthProcessing(sourceDataInfo) {
    try {
      const { sheet, data, monthInfo } = sourceDataInfo;
      
      console.log(`🔄 Обработка ${monthInfo.name} ${monthInfo.year}...`);
      
      // Создаем временную таблицу для тестирования
      const testSpreadsheet = SpreadsheetApp.create(`Тест_${monthInfo.name}_${monthInfo.year}_исправленный`);
      const testSheet = testSpreadsheet.getActiveSheet();
      
      // Копируем данные
      testSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      
      // Переименовываем лист для правильного определения месяца
      testSheet.setName(`${monthInfo.name} ${monthInfo.year}`);
      
      // Запускаем наш процессор
      const result = this.processor.processReport(testSpreadsheet.getId(), testSheet.getName());
      
      // Анализируем результат
      this.analyzeProcessingResult(monthInfo, result);
      
      // Удаляем временную таблицу
      try {
        DriveApp.getFileById(testSpreadsheet.getId()).setTrashed(true);
      } catch (error) {
        console.warn('⚠️ Не удалось удалить временную таблицу:', error.message);
      }
      
    } catch (error) {
      console.error(`❌ Ошибка тестирования ${sourceDataInfo.monthInfo.name} ${sourceDataInfo.monthInfo.year}:`, error);
      this.recordTestResult(sourceDataInfo.monthInfo, null, error.toString());
    }
  }

  /**
   * Анализ результата обработки
   */
  analyzeProcessingResult(monthInfo, result) {
    if (!result.success) {
      console.error(`❌ Ошибка обработки: ${result.error}`);
      this.recordTestResult(monthInfo, null, result.error);
      return;
    }
    
    const stats = result.statistics;
    
    console.log(`✅ Обработка завершена успешно`);
    console.log(`📊 Статистика:`);
    console.log(`   - Всего строк: ${stats.totalRows}`);
    console.log(`   - Отзывы: ${stats.reviewsCount}`);
    console.log(`   - Комментарии Топ-20: ${stats.commentsTop20Count}`);
    console.log(`   - Активные обсуждения: ${stats.activeDiscussionsCount}`);
    console.log(`   - Общие просмотры: ${stats.totalViews}`);
    console.log(`   - Время обработки: ${stats.processingTime}мс`);
    
    // Проверяем качество обработки
    const qualityScore = this.calculateQualityScore(stats);
    
    console.log(`📈 Качество обработки: ${(qualityScore * 100).toFixed(1)}%`);
    
    // Записываем результат
    this.recordTestResult(monthInfo, stats, null, qualityScore);
  }

  /**
   * Расчет качества обработки
   */
  calculateQualityScore(stats) {
    let score = 0;
    let totalChecks = 0;
    
    // Проверяем наличие данных
    if (stats.totalRows > 0) {
      score += 0.2;
      totalChecks++;
    }
    
    // Проверяем наличие отзывов
    if (stats.reviewsCount > 0) {
      score += 0.3;
      totalChecks++;
    }
    
    // Проверяем наличие комментариев Топ-20
    if (stats.commentsTop20Count > 0) {
      score += 0.25;
      totalChecks++;
    }
    
    // Проверяем наличие активных обсуждений
    if (stats.activeDiscussionsCount > 0) {
      score += 0.25;
      totalChecks++;
    }
    
    return totalChecks > 0 ? score / totalChecks : 0;
  }

  /**
   * Запись результата теста
   */
  recordTestResult(month, stats, error = null, qualityScore = 0) {
    this.testResults.totalTests++;
    
    const testDetail = {
      month: `${month.name} ${month.year}`,
      status: error ? 'FAILED' : 'PASSED',
      qualityScore: qualityScore,
      statistics: stats,
      error: error
    };
    
    this.testResults.details.push(testDetail);
    
    if (testDetail.status === 'PASSED') {
      this.testResults.passedTests++;
      console.log(`✅ ${month.name} ${month.year}: ПРОЙДЕН (${(qualityScore * 100).toFixed(1)}%)`);
    } else {
      this.testResults.failedTests++;
      console.log(`❌ ${month.name} ${month.year}: ПРОВАЛЕН`);
      if (error) console.log(`   Ошибка: ${error}`);
    }
  }

  /**
   * Анализ результатов
   */
  analyzeResults() {
    console.log('\n📊 АНАЛИЗ РЕЗУЛЬТАТОВ ТЕСТИРОВАНИЯ');
    console.log('=====================================');
    
    const successRate = this.testResults.totalTests > 0 ? 
      (this.testResults.passedTests / this.testResults.totalTests) * 100 : 0;
    
    console.log(`📈 Общий результат: ${successRate.toFixed(1)}%`);
    console.log(`✅ Пройдено тестов: ${this.testResults.passedTests}/${this.testResults.totalTests}`);
    console.log(`❌ Провалено тестов: ${this.testResults.failedTests}/${this.testResults.totalTests}`);
    console.log(`⏱️ Время тестирования: ${(this.testResults.processingTime / 1000).toFixed(2)} сек`);
    
    // Детальный анализ
    console.log('\n📋 ДЕТАЛЬНЫЕ РЕЗУЛЬТАТЫ:');
    this.testResults.details.forEach(detail => {
      const statusIcon = detail.status === 'PASSED' ? '✅' : '❌';
      const quality = detail.qualityScore ? `(${(detail.qualityScore * 100).toFixed(1)}%)` : '';
      console.log(`${statusIcon} ${detail.month}: ${detail.status} ${quality}`);
      
      if (detail.error) {
        console.log(`   Ошибка: ${detail.error}`);
      }
      
      if (detail.statistics) {
        console.log(`   Статистика: ${detail.statistics.reviewsCount} отзывов, ${detail.statistics.commentsTop20Count} топ-20, ${detail.statistics.activeDiscussionsCount} обсуждений`);
      }
    });
  }
}

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Запуск тестирования исправленного процессора
 */
function testFixedProcessor() {
  const tester = new FixedProcessorTester();
  return tester.runTesting();
}

/**
 * Обновление меню с функциями тестирования
 */
function updateMenuWithTesting() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧪 Тестирование исправленного процессора')
    .addItem('Запустить тестирование', 'testFixedProcessor')
    .addToUi();
} 