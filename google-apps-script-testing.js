/**
 * 🧪 ПОЛНОЦЕННОЕ ТЕСТИРОВАНИЕ GOOGLE APPS SCRIPT
 * Тестирование на основе данных фонового агента
 * 
 * Автор: AI Assistant + Background Agent
 * Версия: 1.0.0
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ ТЕСТИРОВАНИЯ ====================

const TEST_CONFIG = {
  // URL исходных данных фонового агента
  SOURCE_URL: 'https://docs.google.com/spreadsheets/d/1z4KJfXSNaV8Zw5Qi4hKdVsoKvt9GTSrweWv45URzB4I/edit?gid=1258324199#gid=1258324199',
  
  // URL эталонных результатов
  REFERENCE_URL: 'https://docs.google.com/spreadsheets/d/18jkeKNIn5QJpzQrDN3RAT3mEYRXlSNKOnNvA3pxoBx8/edit?gid=535445992#gid=535445992',
  
  // Месяцы для тестирования (февраль-май 2025)
  TEST_MONTHS: [
    { name: 'Февраль', short: 'Фев', number: 2, year: 2025 },
    { name: 'Март', short: 'Мар', number: 3, year: 2025 },
    { name: 'Апрель', short: 'Апр', number: 4, year: 2025 },
    { name: 'Май', short: 'Май', number: 5, year: 2025 }
  ],
  
  // Настройки тестирования
  TESTING: {
    MAX_RETRIES: 3,
    TIMEOUT_SECONDS: 300,
    COMPARISON_THRESHOLD: 0.95 // 95% совпадение считается успешным
  }
};

// ==================== КЛАСС ТЕСТИРОВАНИЯ ====================

/**
 * Класс для полноценного тестирования
 */
class GoogleAppsScriptTester {
  constructor() {
    this.testResults = {
      totalTests: 0,
      passedTests: 0,
      failedTests: 0,
      details: []
    };
    
    this.processor = new MonthlyReportProcessor();
  }

  /**
   * Запуск полного тестирования
   */
  async runFullTesting() {
    console.log('🚀 ЗАПУСК ПОЛНОЦЕННОГО ТЕСТИРОВАНИЯ');
    console.log('=====================================');
    console.log(`📊 Тестируем месяцы: ${TEST_CONFIG.TEST_MONTHS.map(m => m.name).join(', ')}`);
    console.log(`📋 Исходные данные: ${TEST_CONFIG.SOURCE_URL}`);
    console.log(`📋 Эталонные результаты: ${TEST_CONFIG.REFERENCE_URL}`);
    
    try {
      // 1. Подготовка данных
      console.log('\n📋 ПОДГОТОВКА ДАННЫХ...');
      const sourceData = await this.prepareSourceData();
      const referenceData = await this.prepareReferenceData();
      
      // 2. Тестирование каждого месяца
      for (const month of TEST_CONFIG.TEST_MONTHS) {
        console.log(`\n📅 ТЕСТИРОВАНИЕ МЕСЯЦА: ${month.name} ${month.year}`);
        console.log('='.repeat(50));
        
        await this.testMonth(month, sourceData, referenceData);
      }
      
      // 3. Анализ результатов
      this.analyzeResults();
      
      // 4. Генерация отчета
      await this.generateTestReport();
      
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
    for (const month of TEST_CONFIG.TEST_MONTHS) {
      const monthVariants = [
        month.name.toLowerCase(),
        month.short.toLowerCase(),
        `${month.short}${month.year}`.toLowerCase(),
        `${month.name}${month.year}`.toLowerCase()
      ];
      
      if (monthVariants.some(variant => lowerSheetName.includes(variant))) {
        return {
          key: `${month.short}${month.year}`,
          name: month.name,
          short: month.short,
          number: month.number,
          year: month.year,
          detectedFrom: 'sheet'
        };
      }
    }
    
    // Поиск в данных (первые 10 строк)
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      
      for (const month of TEST_CONFIG.TEST_MONTHS) {
        const monthVariants = [
          month.name.toLowerCase(),
          month.short.toLowerCase(),
          `${month.short}${month.year}`.toLowerCase()
        ];
        
        if (monthVariants.some(variant => rowText.includes(variant))) {
          return {
            key: `${month.short}${month.year}`,
            name: month.name,
            short: month.short,
            number: month.number,
            year: month.year,
            detectedFrom: 'content'
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Тестирование конкретного месяца
   */
  async testMonth(month, sourceData, referenceData) {
    const monthKey = `${month.short}${month.year}`;
    
    try {
      // Проверяем наличие исходных данных
      if (!sourceData[monthKey]) {
        throw new Error(`Исходные данные для ${month.name} ${month.year} не найдены`);
      }
      
      // Проверяем наличие эталонных данных
      if (!referenceData[monthKey]) {
        throw new Error(`Эталонные данные для ${month.name} ${month.year} не найдены`);
      }
      
      console.log(`🔄 Обработка ${month.name} ${month.year}...`);
      
      // 1. Обрабатываем данные нашим скриптом
      const processedResult = await this.processMonthData(sourceData[monthKey]);
      
      // 2. Сравниваем с эталоном
      const comparisonResult = this.compareWithReference(processedResult, referenceData[monthKey]);
      
      // 3. Записываем результат
      this.recordTestResult(month, processedResult, comparisonResult);
      
      // 4. Если результат неудовлетворительный, пробуем исправить
      if (comparisonResult.similarity < TEST_CONFIG.TESTING.COMPARISON_THRESHOLD) {
        console.log(`⚠️ Низкое совпадение (${(comparisonResult.similarity * 100).toFixed(1)}%), пробуем исправить...`);
        await this.attemptFix(month, sourceData[monthKey], referenceData[monthKey]);
      }
      
    } catch (error) {
      console.error(`❌ Ошибка тестирования ${month.name} ${month.year}:`, error);
      this.recordTestResult(month, null, null, error.toString());
    }
  }

  /**
   * Обработка данных месяца нашим скриптом
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
    
    return totalChecks > 0 ? structureScore / totalChecks : 0;
  }

  /**
   * Сравнение содержания данных
   */
  compareContent(processedData, referenceData) {
    let contentScore = 0;
    let totalComparisons = 0;
    
    // Сравниваем количество отзывов и комментариев
    const processedCounts = this.countRecords(processedData);
    const referenceCounts = this.countRecords(referenceData);
    
    // Сравнение количества отзывов
    if (referenceCounts.reviews > 0) {
      const reviewAccuracy = Math.min(processedCounts.reviews / referenceCounts.reviews, 1);
      contentScore += reviewAccuracy;
      totalComparisons++;
    }
    
    // Сравнение количества комментариев
    if (referenceCounts.comments > 0) {
      const commentAccuracy = Math.min(processedCounts.comments / referenceCounts.comments, 1);
      contentScore += commentAccuracy;
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
    
    // Сравнение комментариев
    if (referenceStats.commentsCount > 0) {
      const commentsAccuracy = Math.min(processedStats.commentsCount / referenceStats.commentsCount, 1);
      statsScore += commentsAccuracy;
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
        if (firstCell.includes('отзывы') || firstCell.includes('комментарии') || 
            firstCell.includes('reviews') || firstCell.includes('comments')) {
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
   * Подсчет записей в данных
   */
  countRecords(data) {
    let reviews = 0;
    let comments = 0;
    let totalViews = 0;
    
    let inReviewsSection = false;
    let inCommentsSection = false;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length === 0) continue;
      
      const firstCell = String(row[0]).toLowerCase();
      
      // Определяем секции
      if (firstCell.includes('отзывы') || firstCell.includes('reviews')) {
        inReviewsSection = true;
        inCommentsSection = false;
        continue;
      }
      
      if (firstCell.includes('комментарии') || firstCell.includes('comments')) {
        inReviewsSection = false;
        inCommentsSection = true;
        continue;
      }
      
      // Подсчитываем записи
      if (inReviewsSection && this.isDataRow(row)) {
        reviews++;
        totalViews += this.extractViews(row);
      }
      
      if (inCommentsSection && this.isDataRow(row)) {
        comments++;
        totalViews += this.extractViews(row);
      }
    }
    
    return { reviews, comments, totalViews };
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
    const counts = this.countRecords(data);
    
    return {
      totalRows: data.length,
      reviewsCount: counts.reviews,
      commentsCount: counts.comments,
      totalViews: counts.totalViews
    };
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
      ['ОТЧЕТ О ТЕСТИРОВАНИИ GOOGLE APPS SCRIPT'],
      [''],
      [`Дата тестирования: ${new Date().toLocaleDateString('ru-RU')}`],
      [`Время тестирования: ${new Date().toLocaleTimeString('ru-RU')}`],
      [''],
      ['ОБЩИЕ РЕЗУЛЬТАТЫ:'],
      [`Всего тестов: ${this.testResults.totalTests}`],
      [`Пройдено: ${this.testResults.passedTests}`],
      [`Провалено: ${this.testResults.failedTests}`],
      [`Успешность: ${this.testResults.totalTests > 0 ? (this.testResults.passedTests / this.testResults.totalTests * 100).toFixed(1) : 0}%`],
      [''],
      ['ДЕТАЛЬНЫЕ РЕЗУЛЬТАТЫ:'],
      ['Месяц', 'Статус', 'Схожесть', 'Обработано строк', 'Эталон строк', 'Отзывы', 'Комментарии', 'Просмотры']
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
        details.comments || 'N/A',
        details.views || 'N/A'
      ]);
    });
    
    // Создаем отчет
    const reportSpreadsheet = SpreadsheetApp.create(`Отчет_тестирования_${new Date().toISOString().split('T')[0]}`);
    const reportSheet = reportSpreadsheet.getActiveSheet();
    
    reportSheet.getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);
    reportSheet.autoResizeColumns(1, reportData[0].length);
    
    console.log(`✅ Отчет создан: ${reportSpreadsheet.getUrl()}`);
    
    return reportSpreadsheet.getUrl();
  }
}

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Запуск полного тестирования
 */
function runFullTesting() {
  const tester = new GoogleAppsScriptTester();
  return tester.runFullTesting();
}

/**
 * Тестирование одного месяца
 */
function testSingleMonth(monthName) {
  const month = TEST_CONFIG.TEST_MONTHS.find(m => 
    m.name.toLowerCase() === monthName.toLowerCase() || 
    m.short.toLowerCase() === monthName.toLowerCase()
  );
  
  if (!month) {
    console.error(`❌ Месяц "${monthName}" не найден в списке тестирования`);
    return;
  }
  
  const tester = new GoogleAppsScriptTester();
  tester.testMonth(month, {}, {});
}

/**
 * Просмотр конфигурации тестирования
 */
function showTestConfig() {
  console.log('⚙️ КОНФИГУРАЦИЯ ТЕСТИРОВАНИЯ:');
  console.log('==============================');
  console.log(`📊 Исходные данные: ${TEST_CONFIG.SOURCE_URL}`);
  console.log(`📊 Эталонные данные: ${TEST_CONFIG.REFERENCE_URL}`);
  console.log(`📅 Тестируемые месяцы: ${TEST_CONFIG.TEST_MONTHS.map(m => m.name).join(', ')}`);
  console.log(`🎯 Порог успешности: ${TEST_CONFIG.TESTING.COMPARISON_THRESHOLD * 100}%`);
  console.log(`🔄 Максимум попыток: ${TEST_CONFIG.TESTING.MAX_RETRIES}`);
}

/**
 * Обновление меню с функциями тестирования
 */
function updateMenuWithTesting() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧪 Тестирование')
    .addItem('Полное тестирование', 'runFullTesting')
    .addSeparator()
    .addItem('Тест Февраль', 'testFebruary')
    .addItem('Тест Март', 'testMarch')
    .addItem('Тест Апрель', 'testApril')
    .addItem('Тест Май', 'testMay')
    .addSeparator()
    .addItem('Показать конфигурацию', 'showTestConfig')
    .addToUi();
}

// Функции для тестирования отдельных месяцев
function testFebruary() { testSingleMonth('Февраль'); }
function testMarch() { testSingleMonth('Март'); }
function testApril() { testSingleMonth('Апрель'); }
function testMay() { testSingleMonth('Май'); } 