/**
 * 🧪 ФИНАЛЬНЫЙ ТЕСТОВЫЙ ФРЕЙМВОРК НА ОСНОВЕ АНАЛИЗА БЭКАГЕНТА 1
 * Тестирование Google Apps Script решения на реальных данных
 * 
 * Автор: AI Assistant + Background Agent bc-851d0563-ea94-47b9-ba36-0f832bafdb25
 * Версия: 2.1.0 - ЭТАЛОННЫЕ ЛИСТЫ В ТОЙ ЖЕ ТАБЛИЦЕ
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ ТЕСТИРОВАНИЯ ====================

const TEST_CONFIG = {
  // URL исходных данных (ОСНОВАНЫ НА АНАЛИЗЕ БЭКАГЕНТА 1)
  SOURCE_URL: 'https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing',
  
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
  },
  
  // Шаблоны для поиска эталонных листов
  REFERENCE_PATTERNS: {
    suffix: ' (эталон)',
    alternativeSuffixes: [' (эталон)', ' (reference)', ' (etalon)']
  }
};

// ==================== КЛАСС ТЕСТИРОВАНИЯ ====================

/**
 * Финальный класс для тестирования на основе анализа Бэкагента 1
 * ОБНОВЛЕН: эталонные листы в той же таблице
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
    this.currentSpreadsheet = null;
  }

  /**
   * Запуск полного тестирования
   */
  async runFullTesting() {
    const startTime = Date.now();
    
    console.log('🚀 ЗАПУСК ФИНАЛЬНОГО ТЕСТИРОВАНИЯ НА ОСНОВЕ АНАЛИЗА БЭКАГЕНТА 1');
    console.log('================================================================');
    console.log(`📊 Исходные данные: ${TEST_CONFIG.SOURCE_URL}`);
    console.log(`📊 Эталонные листы: в той же таблице (шаблон: "Месяц (эталон)")`);
    console.log(`📋 Структура: заголовки в строке ${TEST_CONFIG.DATA_STRUCTURE.headerRow}, данные с строки ${TEST_CONFIG.DATA_STRUCTURE.dataStartRow}`);
    
    try {
      // 1. Подготовка данных
      console.log('\n📋 ПОДГОТОВКА ДАННЫХ...');
      const sourceData = await this.prepareSourceData();
      
      // 2. Тестирование обработки
      console.log('\n🧪 ТЕСТИРОВАНИЕ ОБРАБОТКИ...');
      await this.testProcessing(sourceData);
      
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
   * Подготовка исходных данных (ИСПРАВЛЕНО: только целевые месяцы)
   */
  async prepareSourceData() {
    console.log('📊 Загрузка исходных данных...');
    
    try {
      this.currentSpreadsheet = SpreadsheetApp.openByUrl(TEST_CONFIG.SOURCE_URL);
      const sheets = this.currentSpreadsheet.getSheets();
      
      const sourceData = {};
      
      // ИСПРАВЛЕНИЕ: Определяем только целевые месяцы 2025 года
      const targetMonths = ['Февраль', 'Март', 'Апрель', 'Май'];
      const targetYear = 2025;
      
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        
        // Пропускаем эталонные листы - они не должны обрабатываться как исходные данные
        if (this.isReferenceSheet(sheetName)) {
          console.log(`⏭️ Пропускаем эталонный лист "${sheetName}"`);
          continue;
        }
        
        const data = sheet.getDataRange().getValues();
        
        // Определяем месяц для каждого листа
        const monthInfo = this.detectMonthFromSheet(sheetName, data);
        
        // ИСПРАВЛЕНИЕ: Обрабатываем только целевые месяцы 2025 года
        if (monthInfo && monthInfo.year === targetYear && targetMonths.includes(monthInfo.name)) {
          // Проверяем, есть ли соответствующий эталонный лист
          const referenceSheet = this.findReferenceSheet(monthInfo);
          
          sourceData[monthInfo.key] = {
            sheet: sheet,
            data: data,
            monthInfo: monthInfo,
            referenceSheet: referenceSheet
          };
          
          if (referenceSheet) {
            console.log(`✅ Лист "${sheetName}" -> ${monthInfo.name} ${monthInfo.year} (есть эталон)`);
          } else {
            console.log(`⚠️ Лист "${sheetName}" -> ${monthInfo.name} ${monthInfo.year} (нет эталона)`);
          }
        } else if (monthInfo) {
          console.log(`⏭️ Пропускаем лист "${sheetName}" -> ${monthInfo.name} ${monthInfo.year} (не целевой месяц)`);
        }
      }
      
      console.log(`📊 Загружено ${Object.keys(sourceData).length} целевых листов с данными (февраль-май 2025)`);
      return sourceData;
      
    } catch (error) {
      throw new Error(`Ошибка загрузки исходных данных: ${error.message}`);
    }
  }

  /**
   * Поиск эталонного листа для месяца
   */
  findReferenceSheet(monthInfo) {
    if (!this.currentSpreadsheet) return null;
    
    // ЭТАЛОНЫ ТОЛЬКО ДЛЯ 2025 ГОДА
    if (monthInfo.year !== 2025) {
      return null;
    }
    
    const sheets = this.currentSpreadsheet.getSheets();
    const referenceName = this.getReferenceSheetName(monthInfo);
    
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (sheetName === referenceName) {
        return sheet;
      }
    }
    
    return null;
  }

  /**
   * Генерация имени эталонного листа
   */
  getReferenceSheetName(monthInfo) {
    return `${monthInfo.name} ${monthInfo.year}${TEST_CONFIG.REFERENCE_PATTERNS.suffix}`;
  }

  /**
   * Подготовка эталонных данных (ОБНОВЛЕНО)
   */
  async prepareReferenceData() {
    console.log('📊 Загрузка эталонных данных из текущей таблицы...');
    
    try {
      if (!this.currentSpreadsheet) {
        this.currentSpreadsheet = SpreadsheetApp.openByUrl(TEST_CONFIG.SOURCE_URL);
      }
      
      const sheets = this.currentSpreadsheet.getSheets();
      const referenceData = {};
      
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        
        // Проверяем, является ли лист эталонным
        if (this.isReferenceSheet(sheetName)) {
          const data = sheet.getDataRange().getValues();
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
      }
      
      console.log(`📊 Загружено ${Object.keys(referenceData).length} эталонных листов`);
      return referenceData;
      
    } catch (error) {
      throw new Error(`Ошибка загрузки эталонных данных: ${error.message}`);
    }
  }

  /**
   * Проверка, является ли лист эталонным
   */
  isReferenceSheet(sheetName) {
    // Более точная проверка эталонных листов
    const referencePatterns = [
      ' (эталон)',
      ' (reference)', 
      ' (etalon)',
      ' (эталон)',
      ' (эталон)',
      ' (эталон)'
    ];
    
    return referencePatterns.some(pattern => 
      sheetName.includes(pattern)
    );
  }

  /**
   * Определение месяца из названия листа или данных
   */
  detectMonthFromSheet(sheetName, data) {
    const lowerSheetName = sheetName.toLowerCase();
    
    // ИСКЛЮЧАЕМ листы, которые не являются месячными данными
    const excludedPatterns = [
      'бриф',
      'репутация',
      'медиаплан',
      'эталон',
      'reference',
      'etalon',
      'инструкция'
    ];
    
    // Проверяем исключения (ИСПРАВЛЕНО: более точная проверка)
    for (const pattern of excludedPatterns) {
      if (lowerSheetName.includes(pattern.toLowerCase())) {
        console.log(`⏭️ Исключен лист "${sheetName}" (содержит "${pattern}")`);
        return null;
      }
    }
    
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
    
    // Определяем год из названия листа (ИСПРАВЛЕНО: более точное определение)
    let detectedYear = 2025; // по умолчанию
    
    // Ищем год в названии (приоритет точным совпадениям)
    if (lowerSheetName.includes('2024') || lowerSheetName.match(/\b24\b/)) {
      detectedYear = 2024;
    } else if (lowerSheetName.includes('2023') || lowerSheetName.match(/\b23\b/)) {
      detectedYear = 2023;
    } else if (lowerSheetName.includes('2022') || lowerSheetName.match(/\b22\b/)) {
      detectedYear = 2022;
    }
    
    console.log(`🔍 Определение года для "${sheetName}": ${detectedYear}`);
    
    // Более точный поиск с приоритетом точных совпадений
    for (const month of months) {
      // Точные совпадения (высший приоритет)
      const exactMatches = [
        month.name.toLowerCase(),
        month.short.toLowerCase(),
        `${month.short}${detectedYear.toString().slice(-2)}`,
        `${month.name}${detectedYear.toString().slice(-2)}`,
        `${month.short}${detectedYear}`,
        `${month.name}${detectedYear}`
      ];
      
      // Проверяем точные совпадения
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
    
    // Поиск в мета-информации (строки 1-3)
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      
      for (const month of months) {
        const monthVariants = [
          month.name.toLowerCase(),
          month.short.toLowerCase(),
          `${month.short}${detectedYear.toString().slice(-2)}`,
          `${month.name}${detectedYear.toString().slice(-2)}`
        ];
        
        if (monthVariants.some(variant => rowText.includes(variant))) {
          return {
            key: `${month.short}${detectedYear}`,
            name: month.name,
            short: month.short,
            number: month.number,
            year: detectedYear,
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
  async testProcessing(sourceData) {
    console.log('🔄 Тестирование обработки данных...');
    
    // Тестируем каждый найденный месяц
    const testMonths = Object.keys(sourceData);
    
    for (const monthKey of testMonths) {
      const sourceInfo = sourceData[monthKey];
      
      // Проверяем, есть ли эталонный лист для этого месяца
      if (!sourceInfo.referenceSheet) {
        console.log(`⚠️ Для ${sourceInfo.monthInfo.name} ${sourceInfo.monthInfo.year} нет эталонных данных`);
        continue;
      }
      
      // Получаем данные эталонного листа
      const referenceData = sourceInfo.referenceSheet.getDataRange().getValues();
      const referenceInfo = {
        sheet: sourceInfo.referenceSheet,
        data: referenceData,
        monthInfo: sourceInfo.monthInfo
      };
      
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
      if (!comparisonResult.match) {
        console.log(`⚠️ Низкое совпадение, пробуем исправить...`);
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
    
    // Запускаем наш процессор с передачей имени листа
    const result = this.processor.processReport(testSpreadsheet.getId(), testSheet.getName());
    
    // ИСПРАВЛЕНИЕ: Процессор создает отчет в отдельной временной таблице
    // Нужно найти эту таблицу и получить из неё данные
    
    // Ищем временную таблицу с результатом (процессор возвращает URL)
    let resultSpreadsheet = null;
    let resultSheet = null;
    
    // Пытаемся найти таблицу по шаблону имени
    const tempSpreadsheetName = `temp_google_sheets_*_${monthInfo.name}_${monthInfo.year}_результат`;
    
    try {
      // Ищем в Drive по шаблону имени
      const files = DriveApp.getFilesByName(tempSpreadsheetName);
      
      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          resultSpreadsheet = SpreadsheetApp.openById(file.getId());
          console.log(`✅ Найдена временная таблица: ${file.getName()}`);
          break;
        }
      }
      
      // Если не нашли по шаблону, ищем по времени создания (последние 5 минут)
      if (!resultSpreadsheet) {
        const fiveMinutesAgo = new Date(Date.now() - 5 * 60 * 1000);
        const files = DriveApp.getFiles();
        
        while (files.hasNext()) {
          const file = files.next();
          if (file.getMimeType() === MimeType.GOOGLE_SHEETS && 
              file.getName().includes('temp_google_sheets') &&
              file.getName().includes('результат') &&
              file.getDateCreated() > fiveMinutesAgo) {
            resultSpreadsheet = SpreadsheetApp.openById(file.getId());
            console.log(`✅ Найдена временная таблица по времени: ${file.getName()}`);
            break;
          }
        }
      }
      
      if (resultSpreadsheet) {
        console.log(`🔗 Ссылка на временную таблицу: ${resultSpreadsheet.getUrl()}`);
        // Гибкий поиск листа с отчётом
        const possibleNames = [
          `${monthInfo.name}_${monthInfo.year}`,
          `${monthInfo.name} ${monthInfo.year}`,
          `${monthInfo.name}_${monthInfo.year}_результат`,
          `${monthInfo.name} ${monthInfo.year} результат`,
          `${monthInfo.name}_${monthInfo.year} результат`,
          `${monthInfo.name} ${monthInfo.year}_результат`
        ];
        resultSheet = null;
        for (const name of possibleNames) {
          resultSheet = resultSpreadsheet.getSheetByName(name);
          if (resultSheet) {
            console.log(`✅ Найден лист отчёта: ${name}`);
            break;
          }
        }
        if (!resultSheet) {
          // Показываем все доступные листы
          const allSheets = resultSpreadsheet.getSheets();
          const sheetNames = allSheets.map(s => s.getName());
          console.log(`❌ Лист отчёта не найден. Доступные листы: ${sheetNames.join(', ')}`);
          console.log(`🔍 Проверяем таблицу: ${resultSpreadsheet.getUrl()}`);
          throw new Error('Лист отчёта не найден!');
        }
      }
      
    } catch (error) {
      console.error('❌ Ошибка поиска временной таблицы:', error.message);
    }
    
    if (!resultSheet) {
      // Показываем все доступные листы для отладки
      const allSheets = testSpreadsheet.getSheets();
      const sheetNames = allSheets.map(s => s.getName());
      console.log(`❌ Отчет не найден. Доступные листы: ${sheetNames.join(', ')}`);
      
      // НЕ удаляем таблицу сразу, чтобы можно было проверить
      console.log(`🔍 Проверяем таблицу: ${testSpreadsheet.getUrl()}`);
      
      throw new Error(`Отчет не был создан. Доступные листы: ${sheetNames.join(', ')}`);
    }
    
    const resultData = resultSheet.getDataRange().getValues();
    
    // Удаляем временные таблицы только после успешного получения данных
    try {
      DriveApp.getFileById(testSpreadsheet.getId()).setTrashed(true);
      if (resultSpreadsheet) {
        DriveApp.getFileById(resultSpreadsheet.getId()).setTrashed(true);
      }
    } catch (error) {
      console.warn('⚠️ Не удалось удалить временные таблицы:', error.message);
    }
    
    return {
      data: resultData,
      statistics: result.statistics,
      monthInfo: monthInfo
    };
  }

  /**
   * Сравнение с эталонными данными (только нужные разделы и метрики)
   */
  compareWithReference(processedResult, referenceDataInfo) {
    const processedStats = this.extractStatisticsFromData(processedResult.data);
    const referenceStats = this.extractStatisticsFromData(referenceDataInfo.data);

    // Сравнение по разделам
    const sectionResults = [
      {
        name: 'Отзывы',
        processed: processedStats.reviews,
        reference: referenceStats.reviews,
        match: processedStats.reviews === referenceStats.reviews
      },
      {
        name: 'Комментарии Топ-20 выдачи',
        processed: processedStats.commentsTop20,
        reference: referenceStats.commentsTop20,
        match: processedStats.commentsTop20 === referenceStats.commentsTop20
      },
      {
        name: 'Активные обсуждения (мониторинг)',
        processed: processedStats.activeDiscussions,
        reference: referenceStats.activeDiscussions,
        match: processedStats.activeDiscussions === referenceStats.activeDiscussions
      }
    ];

    // Сравнение по метрикам
    const statsResults = [
      {
        name: 'Суммарное количество просмотров',
        processed: processedStats.totalViews,
        reference: referenceStats.totalViews,
        match: processedStats.totalViews === referenceStats.totalViews
      },
      {
        name: 'Количество карточек товара (отзывов)',
        processed: processedStats.productCards,
        reference: referenceStats.productCards,
        match: processedStats.productCards === referenceStats.productCards
      },
      {
        name: 'Количество обсуждений',
        processed: processedStats.discussions,
        reference: referenceStats.discussions,
        match: processedStats.discussions === referenceStats.discussions
      },
      {
        name: 'Доля обсуждений с вовлечением',
        processed: processedStats.engagementShare,
        reference: referenceStats.engagementShare,
        match: processedStats.engagementShare === referenceStats.engagementShare
      }
    ];

    // Итоговое совпадение: все разделы и метрики совпадают
    const allSectionsMatch = sectionResults.every(s => s.match);
    const allStatsMatch = statsResults.every(s => s.match);
    const overallMatch = allSectionsMatch && allStatsMatch;

    // Логируем различия
    sectionResults.forEach(s => {
      if (!s.match) {
        console.log(`❌ Раздел "${s.name}": ${s.processed} (обработано) vs ${s.reference} (эталон)`);
      } else {
        console.log(`✅ Раздел "${s.name}": совпадает (${s.processed})`);
      }
    });
    statsResults.forEach(s => {
      if (!s.match) {
        console.log(`❌ Метрика "${s.name}": ${s.processed} (обработано) vs ${s.reference} (эталон)`);
      } else {
        console.log(`✅ Метрика "${s.name}": совпадает (${s.processed})`);
      }
    });

    return {
      match: overallMatch,
      sectionResults,
      statsResults,
      details: {
        processedStats,
        referenceStats,
        sectionResults,
        statsResults
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
   * Извлечение статистики из блока статистики и разделов
   */
  extractStatisticsFromData(data) {
    console.log(`🔍 Анализ структуры файла (${data.length} строк)`);
    
    // Ищем 4 ключевые метрики в последних 50 строках
    let totalViews = null;
    let productCards = null;
    let discussions = null;
    let engagementShare = null;
    
    const N = Math.min(50, data.length);
    for (let idx = 0; idx < N; idx++) {
      const i = data.length - N + idx;
      const row = data[i].map(cell => String(cell).replace(/\s+/g, ' ').trim());
      const joined = row.join(' ').toLowerCase();
      
      if (idx >= N - 10) {
        console.log(`📋 Строка ${i + 1}: "${joined.substring(0, 100)}..."`);
      }
      
      // Суммарное количество просмотров (ТОЧНОЕ совпадение)
      if (totalViews === null && joined.includes('суммарное количество просмотров')) {
        for (const cell of row) {
          const match = cell.match(/(\d{4,})/);
          if (match) {
            totalViews = parseInt(match[1]);
            console.log(`✅ Найдены просмотры: ${totalViews} в строке ${i + 1}`);
            break;
          }
        }
      }
      
      // Количество карточек товара (отзывы) (ТОЧНОЕ совпадение)
      if (productCards === null && joined.includes('количество карточек товара') && joined.includes('отзыв')) {
        for (const cell of row) {
          const match = cell.match(/(\d{1,})/);
          if (match) {
            productCards = parseInt(match[1]);
            console.log(`✅ Найдены карточки товара: ${productCards} в строке ${i + 1}`);
            break;
          }
        }
      }
      
      // Количество обсуждений (ТОЧНОЕ совпадение)
      if (discussions === null && joined.includes('количество обсуждений') && joined.includes('форумы')) {
        for (const cell of row) {
          const match = cell.match(/(\d{1,})/);
          if (match) {
            discussions = parseInt(match[1]);
            console.log(`✅ Найдены обсуждения: ${discussions} в строке ${i + 1}`);
            break;
          }
        }
      }
      
      // Доля обсуждений с вовлечением (ТОЧНОЕ совпадение)
      if (engagementShare === null && joined.includes('доля обсуждений с вовлечением')) {
        for (const cell of row) {
          const match = cell.match(/(\d+[\.,]\d+)/);
          if (match) {
            engagementShare = Math.round(parseFloat(match[1].replace(',', '.')) * 100);
            console.log(`✅ Найдена доля вовлечения: ${engagementShare}% в строке ${i + 1}`);
            break;
          }
        }
      }
    }
    
    // Подсчёт строк в разделах (ИСПРАВЛЕНО: точное определение разделов)
    let reviews = 0, commentsTop20 = 0, activeDiscussions = 0;
    let currentSection = '';
    let sectionStartRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length === 0) continue;
      
      const firstCell = String(row[0]).toLowerCase().trim();
      
      // ИСПРАВЛЕНИЕ: Точное определение разделов
      if (firstCell === 'отзывы') {
        currentSection = 'reviews';
        sectionStartRow = i;
        console.log(`📂 Найден раздел "Отзывы" в строке ${i + 1}`);
        continue;
      }
      
      if (firstCell === 'комментарии топ-20 выдачи' || firstCell.includes('комментарии топ-20')) {
        currentSection = 'commentsTop20';
        sectionStartRow = i;
        console.log(`📂 Найден раздел "Комментарии Топ-20" в строке ${i + 1}`);
        continue;
      }
      
      if (firstCell === 'активные обсуждения (мониторинг)' || firstCell.includes('активные обсуждения')) {
        currentSection = 'activeDiscussions';
        sectionStartRow = i;
        console.log(`📂 Найден раздел "Активные обсуждения" в строке ${i + 1}`);
        continue;
      }
      
      // Подсчитываем строки данных в разделах
      if (currentSection && sectionStartRow !== -1 && i > sectionStartRow) {
        const hasData = row.some(cell => String(cell).trim().length > 0);
        const isHeader = row.some(cell => String(cell).toLowerCase().includes('площадка') || 
                                        String(cell).toLowerCase().includes('тема') ||
                                        String(cell).toLowerCase().includes('продукт'));
        const isStatistics = firstCell.includes('суммарное количество') || 
                           firstCell.includes('количество карточек') ||
                           firstCell.includes('количество обсуждений') ||
                           firstCell.includes('доля обсуждений');
        
        if (hasData && !isHeader && !isStatistics) {
          if (currentSection === 'reviews') reviews++;
          if (currentSection === 'commentsTop20') commentsTop20++;
          if (currentSection === 'activeDiscussions') activeDiscussions++;
        }
      }
    }
    
    console.log(`📊 Результат подсчёта разделов:`);
    console.log(`   - Отзывы: ${reviews} строк`);
    console.log(`   - Комментарии Топ-20: ${commentsTop20} строк`);
    console.log(`   - Активные обсуждения: ${activeDiscussions} строк`);
    
    return {
      totalViews,
      productCards,
      discussions,
      engagementShare,
      reviews,
      commentsTop20,
      activeDiscussions
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
      status: error ? 'FAILED' : (comparisonResult.match ? 'PASSED' : 'FAILED'),
      similarity: comparisonResult ? (comparisonResult.match ? 1 : 0) : 0,
      details: comparisonResult ? {
        processedStats: comparisonResult.details.processedStats,
        referenceStats: comparisonResult.details.referenceStats,
        sectionResults: comparisonResult.sectionResults,
        statsResults: comparisonResult.statsResults
      } : null,
      error: error
    };
    
    this.testResults.details.push(testDetail);
    
    if (testDetail.status === 'PASSED') {
      this.testResults.passedTests++;
      console.log(`✅ ${month.name} ${month.year}: ПРОЙДЕН`);
    } else {
      this.testResults.failedTests++;
      console.log(`❌ ${month.name} ${month.year}: ПРОВАЛЕН`);
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
        details.processedStats ? details.processedStats.totalRows : 'N/A',
        details.referenceStats ? details.referenceStats.totalRows : 'N/A',
        details.processedStats ? details.processedStats.reviews : 'N/A',
        details.processedStats ? details.processedStats.targeted : 'N/A',
        details.processedStats ? details.processedStats.social : 'N/A',
        details.processedStats ? details.processedStats.totalViews : 'N/A'
      ]);
    });
    
    // ИСПРАВЛЕНИЕ: Определяем максимальное количество колонок
    const maxColumns = Math.max(...reportData.map(row => row.length));
    
    // Дополняем все строки до максимального количества колонок
    const normalizedData = reportData.map(row => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxColumns) {
        normalizedRow.push('');
      }
      return normalizedRow;
    });
    
    // Создаем отчет
    const reportSpreadsheet = SpreadsheetApp.create(`Финальный_отчет_тестирования_${new Date().toISOString().split('T')[0]}`);
    const reportSheet = reportSpreadsheet.getActiveSheet();
    
    // ИСПРАВЛЕНИЕ: Записываем данные с правильным количеством колонок
    reportSheet.getRange(1, 1, normalizedData.length, maxColumns).setValues(normalizedData);
    reportSheet.autoResizeColumns(1, maxColumns);
    
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
  console.log(`📊 Эталонные листы: в той же таблице (шаблон: "Месяц (эталон)")`);
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
    .addSeparator()
    .addItem('Показать конфигурацию', 'showFinalTestConfig')
    .addToUi();
} 