/**
 * 🧪 ФИНАЛЬНЫЙ ТЕСТОВЫЙ ФРЕЙМВОРК НА ОСНОВЕ АНАЛИЗА БЭКАГЕНТА 1
 * Тестирование Google Apps Script решения на реальных данных
 * 
 * Автор: AI Assistant + Background Agent bc-851d0563-ea94-47b9-ba36-0f832bafdb25
 * Версия: 2.0.0 - ОСНОВАНА НА РЕАЛЬНЫХ ДАННЫХ
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ ТЕСТИРОВАНИЯ ====================

const TEST_CONFIG = {
  // URL правильных данных (ОСНОВАНЫ НА АНАЛИЗЕ БЭКАГЕНТА 1)
  SOURCE_URL: 'https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing',
  
  // URL эталонных результатов
  REFERENCE_URL: 'https://docs.google.com/spreadsheets/d/1pxUF5HnII7hVnaw077mE0FHqGp-TN1Rk/edit?',
  
  // Настройки тестирования
  TESTING: {
    MAX_RETRIES: 3,
    TIMEOUT_SECONDS: 300,
    COMPARISON_THRESHOLD: 0.95, // 95% совпадение считается успешным
    ENABLE_DETAILED_LOGGING: true
  },
  
  // Структура данных (ОСНОВАНА НА АНАЛИЗЕ БЭКАГЕНТА 1)
  DATA_STRUCTURE: {
    headerRow: 4,        // Заголовки в строке 4
    dataStartRow: 5,     // Данные с строки 5
    infoRows: [1, 2, 3]  // Мета-информация в строках 1-3
  }
};

// ==================== КЛАСС ПРОЦЕССОРА ====================

/**
 * Класс-обёртка для финального процессора
 */
class FinalMonthlyReportProcessor {
  processReport(spreadsheetId) {
    try {
      // Устанавливаем ID таблицы во временную переменную
      const originalSheetId = '1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI';
      
      // Вызываем главную функцию обработки
      const result = processGoogleSheets();
      
      return {
        success: result.success,
        statistics: result.statistics || {
          totalRows: result.processedRows || 0,
          reviewsCount: result.reviewsCount || 0,
          targetedCount: result.targetedCount || 0,
          socialCount: result.socialCount || 0,
          totalViews: result.totalViews || 0,
          totalEngagement: result.totalEngagement || 0
        },
        resultFileId: result.resultFileId
      };
      
    } catch (error) {
      return {
        success: false,
        error: error.toString(),
        statistics: {
          totalRows: 0,
          reviewsCount: 0,
          targetedCount: 0,
          socialCount: 0,
          totalViews: 0,
          totalEngagement: 0
        }
      };
    }
  }
}

// ==================== КЛАСС ТЕСТИРОВАНИЯ ====================

/**
 * Финальный класс для тестирования на основе анализа Бэкагента 1
 */
class FinalGoogleAppsScriptTester {
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
   * Запуск полного тестирования
   */
  async runFullTesting() {
    const startTime = Date.now();
    
    console.log('🚀 ЗАПУСК ФИНАЛЬНОГО ТЕСТИРОВАНИЯ НА ОСНОВЕ АНАЛИЗА БЭКАГЕНТА 1');
    console.log('================================================================');
    console.log(`📊 Исходные данные: ${TEST_CONFIG.SOURCE_URL}`);
    console.log(`📊 Эталонные результаты: ${TEST_CONFIG.REFERENCE_URL}`);
    console.log(`📋 Структура: заголовки в строке ${TEST_CONFIG.DATA_STRUCTURE.headerRow}, данные с строки ${TEST_CONFIG.DATA_STRUCTURE.dataStartRow}`);
    
    try {
      // 1. Подготовка данных
      console.log('\n📋 ПОДГОТОВКА ДАННЫХ...');
      const sourceData = await this.prepareSourceData();
      const referenceData = await this.prepareReferenceData();
      
      // 2. Тестирование обработки
      console.log('\n🧪 ТЕСТИРОВАНИЕ ОБРАБОТКИ...');
      await this.testProcessing(sourceData, referenceData);
      
      // 3. Анализ результатов
      this.analyzeResults();
      
      // 4. Генерация отчета
      await this.generateTestReport();
      
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
   * Подготовка эталонных данных
   */
  async prepareReferenceData() {
    console.log('📊 Загрузка эталонных данных...');
    
    try {
      const spreadsheet = SpreadsheetApp.openByUrl(TEST_CONFIG.REFERENCE_URL);
      const sheets = spreadsheet.getSheets();
      
      const referenceData = {};
      
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        const data = sheet.getDataRange().getValues();
        
        // Определяем месяц для каждого листа
        const monthInfo = this.detectMonthFromSheet(sheetName, data);
        if (monthInfo) {
          referenceData[monthInfo.key] = {
            sheet: sheet,
            data: data,
            monthInfo: monthInfo
          };
          console.log(`✅ Эталонный лист "${sheetName}" -> ${monthInfo.name} ${monthInfo.year}`);
        }
      }
      
      console.log(`📊 Загружено ${Object.keys(referenceData).length} эталонных листов`);
      return referenceData;
      
    } catch (error) {
      throw new Error(`Ошибка загрузки эталонных данных: ${error.message}`);
    }
  }

  /**
   * Определение месяца из названия листа или данных
   */
  detectMonthFromSheet(sheetName, data) {
    const lowerSheetName = sheetName.toLowerCase();
    
    // Поиск в названии листа
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
    
    for (const month of months) {
      const monthVariants = [
        month.name.toLowerCase(),
        month.short.toLowerCase(),
        `${month.short}25`,
        `${month.name}25`,
        `${month.short}2025`,
        `${month.name}2025`
      ];
      
      if (monthVariants.some(variant => lowerSheetName.includes(variant))) {
        return {
          key: `${month.short}${month.year || 2025}`,
          name: month.name,
          short: month.short,
          number: month.number,
          year: 2025,
          detectedFrom: 'sheet'
        };
      }
    }
    
    // Поиск в мета-информации (строки 1-3)
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      
      for (const month of months) {
        const monthVariants = [
          month.name.toLowerCase(),
          month.short.toLowerCase(),
          `${month.short}25`,
          `${month.name}25`
        ];
        
        if (monthVariants.some(variant => rowText.includes(variant))) {
          return {
            key: `${month.short}${month.year || 2025}`,
            name: month.name,
            short: month.short,
            number: month.number,
            year: 2025,
            detectedFrom: 'content'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Тестирование обработки данных
   */
  async testProcessing(sourceData, referenceData) {
    console.log('🔄 Тестирование обработки данных...');
    
    // Тестируем каждый найденный месяц
    const testMonths = Object.keys(sourceData);
    
    for (const monthKey of testMonths) {
      const sourceInfo = sourceData[monthKey];
      const referenceInfo = referenceData[monthKey];
      
      if (!referenceInfo) {
        console.log(`⚠️ Для ${sourceInfo.monthInfo.name} ${sourceInfo.monthInfo.year} нет эталонных данных`);
        continue;
      }
      
      console.log(`\n📅 ТЕСТИРОВАНИЕ: ${sourceInfo.monthInfo.name} ${sourceInfo.monthInfo.year}`);
      console.log('='.repeat(60));
      
      await this.testMonthProcessing(sourceInfo, referenceInfo);
    }
  }

  /**
   * Тестирование обработки конкретного месяца
   */
  async testMonthProcessing(sourceInfo, referenceInfo) {
    try {
      console.log(`🔄 Обработка ${sourceInfo.monthInfo.name} ${sourceInfo.monthInfo.year}...`);
      
      // 1. Обрабатываем данные нашим процессором
      const processedResult = await this.processMonthData(sourceInfo);
      
      // 2. Сравниваем с эталоном
      const comparisonResult = this.compareWithReference(processedResult, referenceInfo);
      
      // 3. Записываем результат
      this.recordTestResult(sourceInfo.monthInfo, processedResult, comparisonResult);
      
      // 4. Если результат неудовлетворительный, пробуем исправить
      if (comparisonResult.similarity < TEST_CONFIG.TESTING.COMPARISON_THRESHOLD) {
        console.log(`⚠️ Низкое совпадение (${(comparisonResult.similarity * 100).toFixed(1)}%), пробуем исправить...`);
        await this.attemptFix(sourceInfo.monthInfo, sourceInfo, referenceInfo);
      }
      
    } catch (error) {
      console.error(`❌ Ошибка тестирования ${sourceInfo.monthInfo.name} ${sourceInfo.monthInfo.year}:`, error);
      this.recordTestResult(sourceInfo.monthInfo, null, null, error.toString());
    }
  }

  /**
   * Обработка данных месяца нашим процессором
   */
  async processMonthData(sourceDataInfo) {
    const { sheet, data, monthInfo } = sourceDataInfo;
    
    // Создаем временную таблицу для тестирования
    const testSpreadsheet = SpreadsheetApp.create(`Тест_${monthInfo.name}_${monthInfo.year}`);
    const testSheet = testSpreadsheet.getActiveSheet();
    
    // Копируем данные
    testSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Переименовываем лист для правильного определения месяца
    testSheet.setName(`${monthInfo.name} ${monthInfo.year}`);
    
    // Запускаем наш процессор
    const result = this.processor.processReport(testSpreadsheet.getId());
    
    // Получаем результат
    const resultSheet = testSpreadsheet.getSheetByName(`Отчет_${monthInfo.name}_${new Date().getFullYear()}`);
    if (!resultSheet) {
      throw new Error('Отчет не был создан');
    }
    
    const resultData = resultSheet.getDataRange().getValues();
    
    // Удаляем временную таблицу
    DriveApp.getFileById(testSpreadsheet.getId()).setTrashed(true);
    
    return {
      data: resultData,
      statistics: result.statistics,
      monthInfo: monthInfo
    };
  }

  /**
   * Сравнение с эталонными данными
   */
  compareWithReference(processedResult, referenceDataInfo) {
    const { data: processedData } = processedResult;
    const { data: referenceData } = referenceDataInfo;
    
    console.log(`📊 Сравнение: обработано ${processedData.length} строк, эталон ${referenceData.length} строк`);
    
    // Анализ структуры
    const structureComparison = this.compareStructure(processedData, referenceData);
    
    // Анализ содержания
    const contentComparison = this.compareContent(processedData, referenceData);
    
    // Анализ статистики
    const statsComparison = this.compareStatistics(processedResult.statistics, referenceDataInfo);
    
    const overallSimilarity = (structureComparison + contentComparison + statsComparison) / 3;
    
    return {
      similarity: overallSimilarity,
      structure: structureComparison,
      content: contentComparison,
      statistics: statsComparison,
      details: {
        processedRows: processedData.length,
        referenceRows: referenceData.length,
        structureMatch: structureComparison,
        contentMatch: contentComparison,
        statsMatch: statsComparison
      }
    };
  }

  /**
   * Сравнение структуры данных
   */
  compareStructure(processedData, referenceData) {
    // Проверяем наличие основных разделов
    const processedSections = this.findSections(processedData);
    const referenceSections = this.findSections(referenceData);
    
    let structureScore = 0;
    let totalChecks = 0;
    
    // Проверка заголовков
    if (processedData.length > 0 && referenceData.length > 0) {
      const processedHeaders = processedData[0].join(' ').toLowerCase();
      const referenceHeaders = referenceData[0].join(' ').toLowerCase();
      
      const headerSimilarity = this.calculateTextSimilarity(processedHeaders, referenceHeaders);
      structureScore += headerSimilarity;
      totalChecks++;
    }
    
    // Проверка разделов
    const sectionSimilarity = this.calculateSectionSimilarity(processedSections, referenceSections);
    structureScore += sectionSimilarity;
    totalChecks++;
    
    // Проверка итоговой строки
    const totalRowSimilarity = this.compareTotalRows(processedData, referenceData);
    structureScore += totalRowSimilarity;
    totalChecks++;
    
    return totalChecks > 0 ? structureScore / totalChecks : 0;
  }

  /**
   * Сравнение содержания данных
   */
  compareContent(processedData, referenceData) {
    let contentScore = 0;
    let totalComparisons = 0;
    
    // Сравниваем количество записей по типам
    const processedCounts = this.countRecordsByType(processedData);
    const referenceCounts = this.countRecordsByType(referenceData);
    
    // Сравнение отзывов (ОС)
    if (referenceCounts.reviews > 0) {
      const reviewAccuracy = Math.min(processedCounts.reviews / referenceCounts.reviews, 1);
      contentScore += reviewAccuracy;
      totalComparisons++;
    }
    
    // Сравнение целевых сайтов (ЦС)
    if (referenceCounts.targeted > 0) {
      const targetedAccuracy = Math.min(processedCounts.targeted / referenceCounts.targeted, 1);
      contentScore += targetedAccuracy;
      totalComparisons++;
    }
    
    // Сравнение социальных площадок (ПС)
    if (referenceCounts.social > 0) {
      const socialAccuracy = Math.min(processedCounts.social / referenceCounts.social, 1);
      contentScore += socialAccuracy;
      totalComparisons++;
    }
    
    // Сравнение общих просмотров
    if (referenceCounts.totalViews > 0) {
      const viewsAccuracy = Math.min(processedCounts.totalViews / referenceCounts.totalViews, 1);
      contentScore += viewsAccuracy;
      totalComparisons++;
    }
    
    return totalComparisons > 0 ? contentScore / totalComparisons : 0;
  }

  /**
   * Сравнение статистики
   */
  compareStatistics(processedStats, referenceDataInfo) {
    // Извлекаем статистику из эталонных данных
    const referenceStats = this.extractStatisticsFromData(referenceDataInfo.data);
    
    let statsScore = 0;
    let totalChecks = 0;
    
    // Сравнение количества строк
    if (referenceStats.totalRows > 0) {
      const rowsAccuracy = Math.min(processedStats.totalRows / referenceStats.totalRows, 1);
      statsScore += rowsAccuracy;
      totalChecks++;
    }
    
    // Сравнение отзывов
    if (referenceStats.reviewsCount > 0) {
      const reviewsAccuracy = Math.min(processedStats.reviewsCount / referenceStats.reviewsCount, 1);
      statsScore += reviewsAccuracy;
      totalChecks++;
    }
    
    // Сравнение целевых сайтов
    if (referenceStats.targetedCount > 0) {
      const targetedAccuracy = Math.min(processedStats.targetedCount / referenceStats.targetedCount, 1);
      statsScore += targetedAccuracy;
      totalChecks++;
    }
    
    // Сравнение социальных площадок
    if (referenceStats.socialCount > 0) {
      const socialAccuracy = Math.min(processedStats.socialCount / referenceStats.socialCount, 1);
      statsScore += socialAccuracy;
      totalChecks++;
    }
    
    return totalChecks > 0 ? statsScore / totalChecks : 0;
  }

  /**
   * Поиск разделов в данных
   */
  findSections(data) {
    const sections = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length > 0) {
        const firstCell = String(row[0]).toLowerCase();
        if (firstCell.includes('отзывы') || firstCell.includes('целевые') || 
            firstCell.includes('площадки') || firstCell.includes('ос') || 
            firstCell.includes('цс') || firstCell.includes('пс')) {
          sections.push({
            name: firstCell,
            row: i + 1
          });
        }
      }
    }
    
    return sections;
  }

  /**
   * Подсчет записей по типам
   */
  countRecordsByType(data) {
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
      
      // Определяем секции
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
      
      // Подсчитываем записи
      if (inReviewsSection && this.isDataRow(row)) {
        reviews++;
        totalViews += this.extractViews(row);
      }
      
      if (inTargetedSection && this.isDataRow(row)) {
        targeted++;
        totalViews += this.extractViews(row);
      }
      
      if (inSocialSection && this.isDataRow(row)) {
        social++;
        totalViews += this.extractViews(row);
      }
    }
    
    return { reviews, targeted, social, totalViews };
  }

  /**
   * Проверка, является ли строка данными
   */
  isDataRow(row) {
    if (row.length < 3) return false;
    
    // Проверяем, что есть значимый текст
    const textCell = row[2] || row[1] || row[0];
    return textCell && String(textCell).trim().length > 10;
  }

  /**
   * Извлечение просмотров из строки
   */
  extractViews(row) {
    if (row.length < 6) return 0;
    
    const viewsCell = row[5];
    if (!viewsCell) return 0;
    
    const viewsStr = String(viewsCell).replace(/[^\d]/g, '');
    const views = parseInt(viewsStr);
    
    return isNaN(views) ? 0 : views;
  }

  /**
   * Извлечение статистики из данных
   */
  extractStatisticsFromData(data) {
    const counts = this.countRecordsByType(data);
    
    return {
      totalRows: data.length,
      reviewsCount: counts.reviews,
      targetedCount: counts.targeted,
      socialCount: counts.social,
      totalViews: counts.totalViews
    };
  }

  /**
   * Сравнение итоговых строк
   */
  compareTotalRows(processedData, referenceData) {
    const processedTotal = this.findTotalRow(processedData);
    const referenceTotal = this.findTotalRow(referenceData);
    
    if (!processedTotal && !referenceTotal) return 1;
    if (!processedTotal || !referenceTotal) return 0;
    
    const processedText = processedTotal.join(' ').toLowerCase();
    const referenceText = referenceTotal.join(' ').toLowerCase();
    
    return this.calculateTextSimilarity(processedText, referenceText);
  }

  /**
   * Поиск итоговой строки
   */
  findTotalRow(data) {
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      if (row.length > 0) {
        const firstCell = String(row[0]).toLowerCase();
        if (firstCell.includes('итого') || firstCell.includes('total')) {
          return row;
        }
      }
    }
    return null;
  }

  /**
   * Расчет схожести текста
   */
  calculateTextSimilarity(text1, text2) {
    const words1 = text1.split(/\s+/).filter(w => w.length > 0);
    const words2 = text2.split(/\s+/).filter(w => w.length > 0);
    
    const commonWords = words1.filter(word => words2.includes(word));
    const totalWords = Math.max(words1.length, words2.length);
    
    return totalWords > 0 ? commonWords.length / totalWords : 0;
  }

  /**
   * Расчет схожести разделов
   */
  calculateSectionSimilarity(sections1, sections2) {
    if (sections1.length === 0 && sections2.length === 0) return 1;
    if (sections1.length === 0 || sections2.length === 0) return 0;
    
    const sectionNames1 = sections1.map(s => s.name);
    const sectionNames2 = sections2.map(s => s.name);
    
    const commonSections = sectionNames1.filter(name => 
      sectionNames2.some(name2 => this.calculateTextSimilarity(name, name2) > 0.5)
    );
    
    const totalSections = Math.max(sections1.length, sections2.length);
    return totalSections > 0 ? commonSections.length / totalSections : 0;
  }

  /**
   * Запись результата теста
   */
  recordTestResult(month, processedResult, comparisonResult, error = null) {
    this.testResults.totalTests++;
    
    const testDetail = {
      month: `${month.name} ${month.year}`,
      status: error ? 'FAILED' : (comparisonResult.similarity >= TEST_CONFIG.TESTING.COMPARISON_THRESHOLD ? 'PASSED' : 'FAILED'),
      similarity: comparisonResult ? comparisonResult.similarity : 0,
      details: comparisonResult ? comparisonResult.details : null,
      error: error
    };
    
    this.testResults.details.push(testDetail);
    
    if (testDetail.status === 'PASSED') {
      this.testResults.passedTests++;
      console.log(`✅ ${month.name} ${month.year}: ПРОЙДЕН (${(comparisonResult.similarity * 100).toFixed(1)}%)`);
    } else {
      this.testResults.failedTests++;
      console.log(`❌ ${month.name} ${month.year}: ПРОВАЛЕН (${(comparisonResult ? (comparisonResult.similarity * 100).toFixed(1) : 0)}%)`);
      if (error) console.log(`   Ошибка: ${error}`);
    }
  }

  /**
   * Попытка исправления
   */
  async attemptFix(month, sourceDataInfo, referenceDataInfo) {
    console.log(`🔧 Попытка исправления для ${month.name} ${month.year}...`);
    
    // Здесь можно добавить логику автоматического исправления
    // Например, настройка параметров обработки, изменение логики определения типов и т.д.
    
    console.log(`⚠️ Автоматическое исправление не реализовано для ${month.name} ${month.year}`);
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
      const similarity = detail.similarity ? `(${(detail.similarity * 100).toFixed(1)}%)` : '';
      console.log(`${statusIcon} ${detail.month}: ${detail.status} ${similarity}`);
      
      if (detail.error) {
        console.log(`   Ошибка: ${detail.error}`);
      }
    });
  }

  /**
   * Генерация отчета о тестировании
   */
  async generateTestReport() {
    console.log('\n📄 ГЕНЕРАЦИЯ ОТЧЕТА О ТЕСТИРОВАНИИ...');
    
    const reportData = [
      ['ОТЧЕТ О ФИНАЛЬНОМ ТЕСТИРОВАНИИ GOOGLE APPS SCRIPT'],
      ['ОСНОВАН НА АНАЛИЗЕ БЭКАГЕНТА 1'],
      [''],
      [`Дата тестирования: ${new Date().toLocaleDateString('ru-RU')}`],
      [`Время тестирования: ${new Date().toLocaleTimeString('ru-RU')}`],
      [`Общее время: ${(this.testResults.processingTime / 1000).toFixed(2)} сек`],
      [''],
      ['ОБЩИЕ РЕЗУЛЬТАТЫ:'],
      [`Всего тестов: ${this.testResults.totalTests}`],
      [`Пройдено: ${this.testResults.passedTests}`],
      [`Провалено: ${this.testResults.failedTests}`],
      [`Успешность: ${this.testResults.totalTests > 0 ? (this.testResults.passedTests / this.testResults.totalTests * 100).toFixed(1) : 0}%`],
      [''],
      ['ДЕТАЛЬНЫЕ РЕЗУЛЬТАТЫ:'],
      ['Месяц', 'Статус', 'Схожесть', 'Обработано строк', 'Эталон строк', 'Отзывы', 'Целевые', 'Социальные', 'Просмотры']
    ];
    
    this.testResults.details.forEach(detail => {
      const details = detail.details || {};
      reportData.push([
        detail.month,
        detail.status,
        detail.similarity ? `${(detail.similarity * 100).toFixed(1)}%` : 'N/A',
        details.processedRows || 'N/A',
        details.referenceRows || 'N/A',
        details.reviews || 'N/A',
        details.targeted || 'N/A',
        details.social || 'N/A',
        details.views || 'N/A'
      ]);
    });
    
    // Создаем отчет
    const reportSpreadsheet = SpreadsheetApp.create(`Финальный_отчет_тестирования_${new Date().toISOString().split('T')[0]}`);
    const reportSheet = reportSpreadsheet.getActiveSheet();
    
    reportSheet.getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);
    reportSheet.autoResizeColumns(1, reportData[0].length);
    
    console.log(`✅ Финальный отчет создан: ${reportSpreadsheet.getUrl()}`);
    
    return reportSpreadsheet.getUrl();
  }
}

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Запуск финального тестирования
 */
function runFinalTesting() {
  const tester = new FinalGoogleAppsScriptTester();
  return tester.runFullTesting();
}

/**
 * Просмотр конфигурации тестирования
 */
function showFinalTestConfig() {
  console.log('⚙️ КОНФИГУРАЦИЯ ФИНАЛЬНОГО ТЕСТИРОВАНИЯ:');
  console.log('==========================================');
  console.log(`📊 Исходные данные: ${TEST_CONFIG.SOURCE_URL}`);
  console.log(`📊 Эталонные данные: ${TEST_CONFIG.REFERENCE_URL}`);
  console.log(`📋 Структура: заголовки в строке ${TEST_CONFIG.DATA_STRUCTURE.headerRow}, данные с строки ${TEST_CONFIG.DATA_STRUCTURE.dataStartRow}`);
  console.log(`🎯 Порог успешности: ${TEST_CONFIG.TESTING.COMPARISON_THRESHOLD * 100}%`);
  console.log(`🔄 Максимум попыток: ${TEST_CONFIG.TESTING.MAX_RETRIES}`);
  console.log(`📝 Детальное логирование: ${TEST_CONFIG.TESTING.ENABLE_DETAILED_LOGGING ? 'Включено' : 'Выключено'}`);
}

/**
 * Обновление меню с функциями финального тестирования
 */
function updateMenuWithFinalTesting() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧪 Финальное тестирование')
    .addItem('Запустить финальное тестирование', 'runFinalTesting')
    .addItem('Быстрый тест процессора', 'quickTestProcessor')
    .addSeparator()
    .addItem('Показать конфигурацию', 'showFinalTestConfig')
    .addToUi();
}

/**
 * Быстрое тестирование процессора
 */
function quickTestProcessor() {
  console.log('🧪 БЫСТРОЕ ТЕСТИРОВАНИЕ ПРОЦЕССОРА');
  console.log('==================================');
  
  try {
    // Запускаем процессор напрямую
    const result = processGoogleSheets();
    
    console.log('\n📊 РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ:');
    console.log(`✅ Успех: ${result.success}`);
    
    if (result.success) {
      console.log(`📊 Исходных строк: ${result.sourceRows}`);
      console.log(`✅ Обработано строк: ${result.processedRows}`);
      console.log(`📝 Отзывов: ${result.reviewsCount || 0}`);
      console.log(`🎯 Целевых: ${result.targetedCount || 0}`);
      console.log(`📱 Социальных: ${result.socialCount || 0}`);
      console.log(`👁️ Всего просмотров: ${result.totalViews || 0}`);
      console.log(`💬 Всего вовлечений: ${result.totalEngagement || 0}`);
      console.log(`🔗 ID результата: ${result.resultFileId}`);
      
      // Анализ качества
      const quality = analyzeProcessorQuality(result);
      console.log(`\n📈 ОЦЕНКА КАЧЕСТВА: ${quality.score.toFixed(1)}%`);
      
      if (quality.issues.length > 0) {
        console.log('\n⚠️ НАЙДЕННЫЕ ПРОБЛЕМЫ:');
        quality.issues.forEach((issue, index) => {
          console.log(`${index + 1}. ${issue}`);
        });
      }
      
      if (quality.recommendations.length > 0) {
        console.log('\n💡 РЕКОМЕНДАЦИИ:');
        quality.recommendations.forEach((rec, index) => {
          console.log(`${index + 1}. ${rec}`);
        });
      }
      
    } else {
      console.log(`❌ Ошибка: ${result.error}`);
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Критическая ошибка тестирования:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Анализ качества работы процессора
 */
function analyzeProcessorQuality(result) {
  const analysis = {
    score: 0,
    issues: [],
    recommendations: []
  };
  
  let scoreComponents = 0;
  let maxScore = 0;
  
  // Проверка успешности
  if (result.success) {
    scoreComponents += 20;
  } else {
    analysis.issues.push('Процессор завершился с ошибкой');
    analysis.recommendations.push('Исправить основную ошибку обработки');
  }
  maxScore += 20;
  
  // Проверка количества обработанных строк
  const processedRows = result.processedRows || 0;
  if (processedRows >= 10) {
    scoreComponents += 15;
  } else if (processedRows > 0) {
    scoreComponents += 10;
    analysis.issues.push(`Обработано мало строк: ${processedRows}`);
    analysis.recommendations.push('Проверить логику определения строк данных');
  } else {
    analysis.issues.push('Не обработано ни одной строки');
    analysis.recommendations.push('Исправить логику обработки данных');
  }
  maxScore += 15;
  
  // Проверка отзывов
  const reviewsCount = result.reviewsCount || 0;
  if (reviewsCount >= 3) {
    scoreComponents += 15;
  } else if (reviewsCount > 0) {
    scoreComponents += 10;
    analysis.issues.push(`Найдено мало отзывов: ${reviewsCount}`);
    analysis.recommendations.push('Улучшить определение раздела отзывов');
  } else {
    analysis.issues.push('Не найдено отзывов');
    analysis.recommendations.push('Исправить поиск раздела "Отзывы сайтов (ОС)"');
  }
  maxScore += 15;
  
  // Проверка целевых сайтов
  const targetedCount = result.targetedCount || 0;
  if (targetedCount >= 2) {
    scoreComponents += 15;
  } else if (targetedCount > 0) {
    scoreComponents += 10;
    analysis.issues.push(`Найдено мало целевых сайтов: ${targetedCount}`);
    analysis.recommendations.push('Улучшить определение раздела целевых сайтов');
  } else {
    analysis.issues.push('Не найдено целевых сайтов');
    analysis.recommendations.push('Исправить поиск раздела "Целевые сайты (ЦС)"');
  }
  maxScore += 15;
  
  // Проверка социальных площадок
  const socialCount = result.socialCount || 0;
  if (socialCount >= 2) {
    scoreComponents += 15;
  } else if (socialCount > 0) {
    scoreComponents += 10;
    analysis.issues.push(`Найдено мало социальных площадок: ${socialCount}`);
    analysis.recommendations.push('Улучшить определение раздела социальных площадок');
  } else {
    analysis.issues.push('Не найдено социальных площадок');
    analysis.recommendations.push('Исправить поиск раздела "Площадки социальные (ПС)"');
  }
  maxScore += 15;
  
  // Проверка просмотров
  const totalViews = result.totalViews || 0;
  if (totalViews > 100) {
    scoreComponents += 10;
  } else if (totalViews > 0) {
    scoreComponents += 5;
    analysis.issues.push(`Найдено мало просмотров: ${totalViews}`);
    analysis.recommendations.push('Проверить извлечение просмотров из данных');
  } else {
    analysis.issues.push('Не найдено просмотров');
    analysis.recommendations.push('Исправить извлечение просмотров');
  }
  maxScore += 10;
  
  // Проверка вовлечения
  const totalEngagement = result.totalEngagement || 0;
  if (totalEngagement > 10) {
    scoreComponents += 10;
  } else if (totalEngagement > 0) {
    scoreComponents += 5;
    analysis.issues.push(`Найдено мало вовлечений: ${totalEngagement}`);
    analysis.recommendations.push('Проверить извлечение вовлечений из данных');
  } else {
    analysis.issues.push('Не найдено вовлечений');
    analysis.recommendations.push('Исправить извлечение вовлечений');
  }
  maxScore += 10;
  
  analysis.score = maxScore > 0 ? (scoreComponents / maxScore) * 100 : 0;
  
  return analysis;
}

/**
 * Автоматическое исправление процессора
 */
function autoFixProcessor() {
  console.log('🔧 АВТОМАТИЧЕСКОЕ ИСПРАВЛЕНИЕ ПРОЦЕССОРА');
  console.log('=========================================');
  
  // Сначала запускаем тест
  const result = quickTestProcessor();
  
  if (!result.success) {
    console.log('❌ Невозможно исправить - процессор не запускается');
    return false;
  }
  
  const quality = analyzeProcessorQuality(result);
  
  if (quality.score >= 90) {
    console.log('✅ Процессор работает отлично, исправления не требуются');
    return true;
  }
  
  console.log('\n🔧 Применение исправлений...');
  console.log('Примечание: Автоматические исправления требуют ручной настройки кода');
  
  quality.recommendations.forEach((rec, index) => {
    console.log(`${index + 1}. TODO: ${rec}`);
  });
  
  return false;
} 