/**
 * 🧪 ЛОКАЛЬНОЕ ТЕСТИРОВАНИЕ ПРОЦЕССОРА
 * Node.js скрипт для отладки логики перед загрузкой в Google Apps Script
 * 
 * Автор: AI Assistant
 * Версия: 1.0.0
 * Дата: 2025
 */

const fs = require('fs');
const path = require('path');

// ==================== КОНФИГУРАЦИЯ ====================

const LOCAL_CONFIG = {
  // Путь к тестовым данным
  TEST_DATA_PATH: './test_data.json',
  
  // Настройки тестирования
  TESTING: {
    MAX_ITERATIONS: 5,
    SUCCESS_THRESHOLD: 0.9,
    VERBOSE_LOGGING: true
  },
  
  // Критерии успешности
  SUCCESS_CRITERIA: {
    MIN_PROCESSED_ROWS: 10,
    MIN_REVIEWS_COUNT: 3,
    MIN_TARGETED_COUNT: 2,
    MIN_SOCIAL_COUNT: 2,
    REQUIRED_TOTAL_ROW: true
  }
};

// ==================== МОКИРОВАННЫЕ ДАННЫЕ ====================

// Симулируем структуру Google Sheets
const MOCK_GOOGLE_SHEETS_DATA = [
  // Мета-информация
  ['Отчет по размещениям', 'Март 2025', '', ''],
  ['Период:', '01.03.2025 - 31.03.2025', '', ''],
  ['Статус:', 'Завершен', '', ''],
  
  // Заголовки
  ['Тип размещения', 'Площадка', 'Продукт', 'Ссылка на сообщение', 'Текст сообщения', 'Согласование/комментарии', 'Дата', 'Ник', 'Автор', 'Просмотры темы на старте', 'Просмотры в конце месяца', 'Просмотров получено', 'Вовлечение', 'Тип поста'],
  
  // Раздел отзывов
  ['Отзывы Сайтов (ОС)', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['Основное сообщение', 'Отзовик', 'Фортедетрим', 'https://otzovik.com/review_123', 'Отличный препарат, помогает быстро восстановиться после болезни...', 'Согласовано', '15.03.2025', 'user123', 'Анна П.', '45', '78', '33', '5', 'ОС'],
  ['Основное сообщение', 'IRecommend', 'Фортедетрим', 'https://irecommend.ru/review_456', 'Рекомендую этот препарат всем, кто заботится о своем здоровье...', 'Согласовано', '18.03.2025', 'healthlover', 'Мария К.', '23', '67', '44', '8', 'ОС'],
  ['Основное сообщение', 'Флап', 'Фортедетрим', 'https://flap.ru/review_789', 'Превосходное средство для поддержания иммунитета...', 'Согласовано', '22.03.2025', 'immunefan', 'Олег С.', '12', '89', '77', '12', 'ОС'],
  
  // Раздел целевых сайтов
  ['Целевые Сайты (ЦС)', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['Целевое сообщение', 'Медпортал', 'Фортедетрим', 'https://medportal.ru/forum_111', 'Кто-нибудь принимал Фортедетрим? Интересует ваш опыт...', 'Согласовано', '10.03.2025', 'meduser', 'Врач И.', '89', '134', '45', '15', 'ЦС'],
  ['Целевое сообщение', 'Здоровье.ру', 'Фортедетрим', 'https://zdravie.ru/topic_222', 'Поделитесь опытом использования иммуномодуляторов...', 'Согласовано', '25.03.2025', 'healthdoc', 'Доктор М.', '67', '98', '31', '9', 'ЦС'],
  
  // Раздел социальных площадок
  ['Площадки Социальные (ПС)', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['Социальный пост', 'ВКонтакте', 'Фортедетрим', 'https://vk.com/post_333', 'Делюсь своим опытом приема витаминов для иммунитета...', 'Согласовано', '05.03.2025', 'vkuser', 'Елена Р.', '156', '203', '47', '23', 'ПС'],
  ['Социальный пост', 'Одноклассники', 'Фортедетрим', 'https://ok.ru/post_444', 'Рекомендую всем мамам этот препарат для детей...', 'Согласовано', '28.03.2025', 'okuser', 'Светлана Т.', '89', '142', '53', '18', 'ПС'],
  ['Социальный пост', 'Telegram', 'Фортедетрим', 'https://t.me/channel_555', 'Важная информация о поддержке иммунитета в весенний период...', 'Согласовано', '30.03.2025', 'tguser', 'Иван К.', '234', '287', '53', '31', 'ПС']
];

// ==================== ЛОКАЛЬНЫЙ ПРОЦЕССОР ====================

class LocalProcessor {
  constructor() {
    this.reset();
  }
  
  reset() {
    this.diagnostics = {
      startTime: new Date(),
      steps: [],
      errors: [],
      warnings: [],
      statistics: {
        totalRows: 0,
        processedRows: 0,
        reviewsCount: 0,
        targetedCount: 0,
        socialCount: 0,
        totalViews: 0,
        totalEngagement: 0
      }
    };
    
    this.processedData = {
      reviews: [],
      targeted: [],
      social: []
    };
    
    this.columnMapping = {};
    this.monthInfo = null;
  }
  
  processData(data) {
    try {
      this.log('🚀 ЗАПУСК ЛОКАЛЬНОГО ПРОЦЕССОРА');
      this.log('==============================');
      
      // Анализ структуры данных
      this.analyzeDataStructure(data);
      
      // Обработка данных
      this.processRows(data);
      
      // Финальная диагностика
      this.finalizeDiagnostics();
      
      return {
        success: true,
        statistics: this.diagnostics.statistics,
        processedData: this.processedData,
        diagnostics: this.diagnostics
      };
      
    } catch (error) {
      this.logError('❌ КРИТИЧЕСКАЯ ОШИБКА', error);
      return {
        success: false,
        error: error.toString(),
        diagnostics: this.diagnostics
      };
    }
  }
  
  analyzeDataStructure(data) {
    this.log('🔍 АНАЛИЗ СТРУКТУРЫ ДАННЫХ');
    this.log('=========================');
    
    this.diagnostics.statistics.totalRows = data.length;
    this.log(`📊 Всего строк: ${data.length}`);
    
    // Поиск заголовков
    let headerRow = -1;
    let maxMatches = 0;
    
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const row = data[i];
      const rowText = row.join(' ').toLowerCase();
      
      const keyWords = ['тип размещения', 'площадка', 'текст сообщения', 'дата', 'автор'];
      const matches = keyWords.filter(keyword => rowText.includes(keyword)).length;
      
      if (matches > maxMatches) {
        maxMatches = matches;
        headerRow = i;
      }
      
      this.log(`  Строка ${i + 1}: ${matches} совпадений - [${row.slice(0, 3).join(', ')}...]`);
    }
    
    if (headerRow >= 0) {
      this.log(`✅ Заголовки найдены в строке ${headerRow + 1}`);
      this.createColumnMapping(data[headerRow]);
    } else {
      this.logWarning('⚠️ Заголовки не найдены, используется стандартный маппинг');
      this.createDefaultMapping();
    }
    
    // Определение месяца
    this.detectMonth(data);
  }
  
  createColumnMapping(headerRow) {
    this.log('📊 Создание маппинга колонок...');
    
    headerRow.forEach((cell, index) => {
      if (cell && cell.toString().trim()) {
        const header = cell.toString().toLowerCase().trim();
        this.columnMapping[header] = index;
        this.log(`  Колонка ${index + 1}: "${header}"`);
      }
    });
    
    this.log(`✅ Создан маппинг для ${Object.keys(this.columnMapping).length} колонок`);
  }
  
  createDefaultMapping() {
    this.columnMapping = {
      'тип размещения': 0,
      'площадка': 1,
      'продукт': 2,
      'ссылка на сообщение': 3,
      'текст сообщения': 4,
      'согласование/комментарии': 5,
      'дата': 6,
      'ник': 7,
      'автор': 8,
      'просмотры темы на старте': 9,
      'просмотры в конце месяца': 10,
      'просмотров получено': 11,
      'вовлечение': 12,
      'тип поста': 13
    };
  }
  
  detectMonth(data) {
    this.log('📅 Определение месяца...');
    
    const monthNames = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь'];
    const monthShort = ['янв', 'фев', 'мар', 'апр', 'май', 'июн'];
    
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      
      for (let j = 0; j < monthNames.length; j++) {
        if (rowText.includes(monthNames[j]) || rowText.includes(monthShort[j])) {
          this.monthInfo = {
            name: monthNames[j],
            short: monthShort[j],
            number: j + 1,
            year: 2025
          };
          
          this.log(`✅ Определен месяц: ${this.monthInfo.name} ${this.monthInfo.year}`);
          return;
        }
      }
    }
    
    this.monthInfo = { name: 'март', short: 'мар', number: 3, year: 2025 };
    this.logWarning('⚠️ Месяц не найден, используется март');
  }
  
  processRows(data) {
    this.log('🔄 ОБРАБОТКА ДАННЫХ');
    this.log('==================');
    
    let currentSection = null;
    let processedRows = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      if (this.isEmptyRow(row)) {
        continue;
      }
      
      const firstCell = (row[0] || '').toString().toLowerCase().trim();
      
      // Определение секции
      if (this.isSectionHeader(firstCell)) {
        currentSection = this.determineSectionType(firstCell);
        this.log(`📋 Переход к разделу: ${currentSection}`);
        continue;
      }
      
      // Пропуск заголовков таблицы
      if (this.isTableHeader(row)) {
        continue;
      }
      
      // Обработка строк данных
      if (currentSection && this.isDataRow(row)) {
        const processedRow = this.processRow(row, currentSection, i + 1);
        
        if (processedRow) {
          this.processedData[currentSection].push(processedRow);
          processedRows++;
          
          this.updateStatistics(processedRow, currentSection);
          
          if (LOCAL_CONFIG.TESTING.VERBOSE_LOGGING) {
            this.log(`  ✅ Обработана строка ${i + 1}: ${processedRow.platform} - ${processedRow.views} просмотров`);
          }
        }
      }
    }
    
    this.diagnostics.statistics.processedRows = processedRows;
    this.log(`✅ Обработано ${processedRows} строк данных`);
  }
  
  isEmptyRow(row) {
    return !row || row.every(cell => !cell || cell.toString().trim() === '');
  }
  
  isSectionHeader(text) {
    const patterns = [
      'отзыв', 'ос', 'основное сообщение',
      'целевые', 'цс', 'целевое сообщение',
      'площадки', 'пс', 'социальные'
    ];
    
    return patterns.some(pattern => text.includes(pattern));
  }
  
  determineSectionType(text) {
    if (text.includes('отзыв') || text.includes('ос')) return 'reviews';
    if (text.includes('целевые') || text.includes('цс')) return 'targeted';
    if (text.includes('площадки') || text.includes('пс')) return 'social';
    return 'other';
  }
  
  isTableHeader(row) {
    if (!row || row.length === 0) return false;
    
    const rowText = row.join(' ').toLowerCase();
    const headerPatterns = ['тип размещения', 'площадка', 'текст сообщения'];
    
    return headerPatterns.some(pattern => rowText.includes(pattern));
  }
  
  isDataRow(row) {
    if (!row || row.length < 3) return false;
    
    const platform = this.getColumnValue(row, 'площадка');
    const text = this.getColumnValue(row, 'текст сообщения');
    
    return platform && text && text.length > 10;
  }
  
  processRow(row, sectionType, rowNumber) {
    try {
      const processedRow = {
        section: sectionType,
        rowNumber: rowNumber,
        placementType: this.getColumnValue(row, 'тип размещения'),
        platform: this.getColumnValue(row, 'площадка'),
        product: this.getColumnValue(row, 'продукт'),
        messageLink: this.getColumnValue(row, 'ссылка на сообщение'),
        messageText: this.getColumnValue(row, 'текст сообщения'),
        date: this.getColumnValue(row, 'дата'),
        author: this.getColumnValue(row, 'автор'),
        views: this.extractNumber(this.getColumnValue(row, 'просмотров получено')),
        engagement: this.extractNumber(this.getColumnValue(row, 'вовлечение')),
        postType: this.getColumnValue(row, 'тип поста')
      };
      
      if (!processedRow.platform || !processedRow.messageText) {
        return null;
      }
      
      return processedRow;
      
    } catch (error) {
      this.logError(`❌ Ошибка обработки строки ${rowNumber}`, error);
      return null;
    }
  }
  
  getColumnValue(row, columnName) {
    const columnIndex = this.columnMapping[columnName];
    if (columnIndex === undefined || columnIndex >= row.length) {
      return '';
    }
    
    const value = row[columnIndex];
    return value ? value.toString().trim() : '';
  }
  
  extractNumber(value) {
    if (!value) return 0;
    
    const str = value.toString().replace(/[^\d.,]/g, '');
    const num = parseFloat(str.replace(',', '.'));
    
    return isNaN(num) ? 0 : num;
  }
  
  updateStatistics(row, sectionType) {
    const stats = this.diagnostics.statistics;
    
    switch (sectionType) {
      case 'reviews':
        stats.reviewsCount++;
        break;
      case 'targeted':
        stats.targetedCount++;
        break;
      case 'social':
        stats.socialCount++;
        break;
    }
    
    stats.totalViews += row.views || 0;
    stats.totalEngagement += row.engagement || 0;
  }
  
  finalizeDiagnostics() {
    this.diagnostics.endTime = new Date();
    this.diagnostics.totalTime = this.diagnostics.endTime - this.diagnostics.startTime;
    
    const stats = this.diagnostics.statistics;
    
    this.log('📊 ФИНАЛЬНАЯ ДИАГНОСТИКА');
    this.log('========================');
    this.log(`⏱️ Общее время: ${this.diagnostics.totalTime}ms`);
    this.log(`📊 Всего строк: ${stats.totalRows}`);
    this.log(`✅ Обработано: ${stats.processedRows}`);
    this.log(`📝 Отзывов: ${stats.reviewsCount}`);
    this.log(`🎯 Целевых: ${stats.targetedCount}`);
    this.log(`📱 Социальных: ${stats.socialCount}`);
    this.log(`👁️ Всего просмотров: ${stats.totalViews}`);
    this.log(`💬 Всего вовлечений: ${stats.totalEngagement}`);
  }
  
  log(message) {
    console.log(message);
    this.diagnostics.steps.push(`${new Date().toISOString()}: ${message}`);
  }
  
  logError(message, error) {
    const errorMsg = `${message}: ${error.toString()}`;
    console.error(errorMsg);
    this.diagnostics.errors.push(errorMsg);
  }
  
  logWarning(message) {
    console.warn(message);
    this.diagnostics.warnings.push(message);
  }
}

// ==================== ТЕСТЕР ====================

class LocalTester {
  constructor() {
    this.processor = new LocalProcessor();
  }
  
  runTest() {
    console.log('🧪 ЗАПУСК ЛОКАЛЬНОГО ТЕСТИРОВАНИЯ');
    console.log('=================================');
    
    // Тестирование на мокированных данных
    const result = this.processor.processData(MOCK_GOOGLE_SHEETS_DATA);
    
    // Анализ результатов
    const analysis = this.analyzeResult(result);
    
    // Отчет
    this.generateReport(result, analysis);
    
    return {
      result,
      analysis,
      recommendations: this.generateRecommendations(analysis)
    };
  }
  
  analyzeResult(result) {
    const analysis = {
      success: result.success,
      score: 0,
      issues: [],
      positives: []
    };
    
    if (!result.success) {
      analysis.issues.push('Процессор завершился с ошибкой');
      return analysis;
    }
    
    const stats = result.statistics;
    
    // Анализ количества обработанных строк
    if (stats.processedRows >= LOCAL_CONFIG.SUCCESS_CRITERIA.MIN_PROCESSED_ROWS) {
      analysis.score += 0.2;
      analysis.positives.push(`Обработано достаточно строк: ${stats.processedRows}`);
    } else {
      analysis.issues.push(`Обработано мало строк: ${stats.processedRows}`);
    }
    
    // Анализ отзывов
    if (stats.reviewsCount >= LOCAL_CONFIG.SUCCESS_CRITERIA.MIN_REVIEWS_COUNT) {
      analysis.score += 0.2;
      analysis.positives.push(`Найдено отзывов: ${stats.reviewsCount}`);
    } else {
      analysis.issues.push(`Найдено мало отзывов: ${stats.reviewsCount}`);
    }
    
    // Анализ целевых сайтов
    if (stats.targetedCount >= LOCAL_CONFIG.SUCCESS_CRITERIA.MIN_TARGETED_COUNT) {
      analysis.score += 0.2;
      analysis.positives.push(`Найдено целевых сайтов: ${stats.targetedCount}`);
    } else {
      analysis.issues.push(`Найдено мало целевых сайтов: ${stats.targetedCount}`);
    }
    
    // Анализ социальных площадок
    if (stats.socialCount >= LOCAL_CONFIG.SUCCESS_CRITERIA.MIN_SOCIAL_COUNT) {
      analysis.score += 0.2;
      analysis.positives.push(`Найдено социальных площадок: ${stats.socialCount}`);
    } else {
      analysis.issues.push(`Найдено мало социальных площадок: ${stats.socialCount}`);
    }
    
    // Анализ просмотров
    if (stats.totalViews > 0) {
      analysis.score += 0.1;
      analysis.positives.push(`Общие просмотры: ${stats.totalViews}`);
    } else {
      analysis.issues.push('Не найдено просмотров');
    }
    
    // Анализ вовлечения
    if (stats.totalEngagement > 0) {
      analysis.score += 0.1;
      analysis.positives.push(`Общее вовлечение: ${stats.totalEngagement}`);
    } else {
      analysis.issues.push('Не найдено вовлечений');
    }
    
    analysis.success = analysis.score >= LOCAL_CONFIG.TESTING.SUCCESS_THRESHOLD;
    
    return analysis;
  }
  
  generateReport(result, analysis) {
    console.log('\n📊 ОТЧЕТ О ТЕСТИРОВАНИИ');
    console.log('=======================');
    
    console.log(`🎯 Общий результат: ${analysis.success ? '✅ УСПЕХ' : '❌ НЕУДАЧА'}`);
    console.log(`📈 Оценка: ${(analysis.score * 100).toFixed(1)}%`);
    
    if (analysis.positives.length > 0) {
      console.log('\n✅ Положительные результаты:');
      analysis.positives.forEach(positive => console.log(`  • ${positive}`));
    }
    
    if (analysis.issues.length > 0) {
      console.log('\n❌ Найденные проблемы:');
      analysis.issues.forEach(issue => console.log(`  • ${issue}`));
    }
    
    const stats = result.statistics;
    console.log('\n📊 Детальная статистика:');
    console.log(`  • Всего строк: ${stats.totalRows}`);
    console.log(`  • Обработано строк: ${stats.processedRows}`);
    console.log(`  • Отзывов: ${stats.reviewsCount}`);
    console.log(`  • Целевых сайтов: ${stats.targetedCount}`);
    console.log(`  • Социальных площадок: ${stats.socialCount}`);
    console.log(`  • Общие просмотры: ${stats.totalViews}`);
    console.log(`  • Общее вовлечение: ${stats.totalEngagement}`);
    
    if (result.diagnostics.errors.length > 0) {
      console.log('\n❌ Ошибки:');
      result.diagnostics.errors.forEach(error => console.log(`  • ${error}`));
    }
    
    if (result.diagnostics.warnings.length > 0) {
      console.log('\n⚠️ Предупреждения:');
      result.diagnostics.warnings.forEach(warning => console.log(`  • ${warning}`));
    }
  }
  
  generateRecommendations(analysis) {
    const recommendations = [];
    
    analysis.issues.forEach(issue => {
      if (issue.includes('мало строк')) {
        recommendations.push('Проверить логику определения строк данных');
      } else if (issue.includes('мало отзывов')) {
        recommendations.push('Улучшить определение раздела отзывов');
      } else if (issue.includes('мало целевых')) {
        recommendations.push('Улучшить определение раздела целевых сайтов');
      } else if (issue.includes('мало социальных')) {
        recommendations.push('Улучшить определение раздела социальных площадок');
      } else if (issue.includes('просмотров')) {
        recommendations.push('Исправить извлечение просмотров из данных');
      } else if (issue.includes('вовлечений')) {
        recommendations.push('Исправить извлечение вовлечений из данных');
      }
    });
    
    return recommendations;
  }
}

// ==================== ЗАПУСК ТЕСТИРОВАНИЯ ====================

function runLocalTest() {
  const tester = new LocalTester();
  const testResult = tester.runTest();
  
  console.log('\n🎯 РЕКОМЕНДАЦИИ ПО УЛУЧШЕНИЮ:');
  console.log('=============================');
  
  if (testResult.recommendations.length > 0) {
    testResult.recommendations.forEach((rec, index) => {
      console.log(`${index + 1}. ${rec}`);
    });
  } else {
    console.log('✅ Все тесты пройдены успешно!');
  }
  
  // Сохранение результатов в файл
  const reportPath = './test_results.json';
  try {
    fs.writeFileSync(reportPath, JSON.stringify(testResult, null, 2));
    console.log(`\n📄 Результаты сохранены в: ${reportPath}`);
  } catch (error) {
    console.error(`❌ Ошибка сохранения результатов: ${error.message}`);
  }
  
  return testResult;
}

// Запуск тестирования
if (require.main === module) {
  runLocalTest();
}

module.exports = { LocalProcessor, LocalTester, runLocalTest };