/**
 * 🚀 ULTIMATE GOOGLE APPS SCRIPT PROCESSOR
 * Финальная версия с полной диагностикой и системой тестирования
 * 
 * Автор: AI Assistant
 * Версия: 3.0.0 - ULTIMATE
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ ====================

const ULTIMATE_CONFIG = {
  // ID тестовой таблицы
  TEST_SHEET_ID: '1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI',
  
  // Настройки диагностики
  DIAGNOSTICS: {
    ENABLE_DETAILED_LOGGING: true,
    LOG_EVERY_ROW: false,
    VALIDATE_EACH_STEP: true,
    EXPORT_DEBUG_DATA: true
  },
  
  // Структура данных (уточненная)
  DATA_STRUCTURE: {
    HEADER_ROW: 4,        // Заголовки в 4-й строке
    DATA_START_ROW: 5,    // Данные начинаются с 5-й строки
    INFO_ROWS: [1, 2, 3]  // Мета-информация в строках 1-3
  },
  
  // Маппинг колонок (гибкий)
  COLUMN_MAPPING: {
    'тип размещения': 'placement_type',
    'площадка': 'platform',
    'продукт': 'product',
    'ссылка на сообщение': 'message_link',
    'текст сообщения': 'message_text',
    'согласование/комментарии': 'approval_comments',
    'дата': 'date',
    'ник': 'nickname',
    'автор': 'author',
    'просмотры темы на старте': 'start_views',
    'просмотры в конце месяца': 'end_views',
    'просмотров получено': 'views_gained',
    'вовлечение': 'engagement',
    'тип поста': 'post_type'
  },
  
  // Типы контента (точная классификация)
  CONTENT_TYPES: {
    REVIEWS: ['ос', 'основное сообщение', 'отзыв', 'отзывы'],
    TARGETED: ['цс', 'целевое сообщение', 'целевые', 'целевые сайты'],
    SOCIAL: ['пс', 'площадки социальные', 'социальные', 'площадки'],
    DISCUSSIONS: ['обсуждения', 'дискуссии', 'форумы']
  },
  
  // Настройки вывода
  OUTPUT: {
    CREATE_SEPARATE_SHEET: true,
    SHEET_NAME_TEMPLATE: 'Отчет_{month}_{year}',
    ADD_TOTAL_ROW: true,
    FORMAT_NUMBERS: true,
    ADD_STATISTICS: true
  }
};

// ==================== КЛАСС ULTIMATE PROCESSOR ====================

class UltimateGoogleAppsScriptProcessor {
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
      social: [],
      discussions: []
    };
    
    this.columnMapping = {};
    this.monthInfo = null;
  }
  
  // ==================== ГЛАВНЫЙ МЕТОД ====================
  
  /**
   * Главная функция обработки
   */
  processSheet(sheetId = null, sheetName = null) {
    try {
      this.log('🚀 ЗАПУСК ULTIMATE PROCESSOR');
      this.log('============================');
      
      // Используем тестовую таблицу, если ID не указан
      const targetSheetId = sheetId || ULTIMATE_CONFIG.TEST_SHEET_ID;
      
      this.log(`📊 Обработка таблицы: ${targetSheetId}`);
      
      // 1. Открытие и анализ таблицы
      const spreadsheet = this.openSpreadsheet(targetSheetId);
      const sheet = this.selectSheet(spreadsheet, sheetName);
      
      // 2. Получение и анализ данных
      const rawData = this.getData(sheet);
      
      // 3. Диагностика структуры данных
      this.analyzeDataStructure(rawData);
      
      // 4. Обработка данных
      this.processData(rawData);
      
      // 5. Создание отчета
      const resultSheet = this.createReport(spreadsheet);
      
      // 6. Финальная диагностика
      this.finalizeDiagnostics();
      
      this.log('✅ ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО');
      
      return {
        success: true,
        resultSheetId: spreadsheet.getId(),
        resultSheetName: resultSheet.getName(),
        statistics: this.diagnostics.statistics,
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
  
  // ==================== РАБОТА С ТАБЛИЦАМИ ====================
  
  openSpreadsheet(sheetId) {
    this.log(`📋 Открытие таблицы: ${sheetId}`);
    
    try {
      const spreadsheet = SpreadsheetApp.openById(sheetId);
      this.log(`✅ Таблица открыта: "${spreadsheet.getName()}"`);
      return spreadsheet;
    } catch (error) {
      throw new Error(`Не удалось открыть таблицу: ${error.message}`);
    }
  }
  
  selectSheet(spreadsheet, sheetName) {
    const sheets = spreadsheet.getSheets();
    this.log(`📄 Найдено листов: ${sheets.length}`);
    
    let targetSheet = null;
    
    if (sheetName) {
      // Поиск конкретного листа
      targetSheet = sheets.find(sheet => sheet.getName() === sheetName);
      if (!targetSheet) {
        this.logWarning(`⚠️ Лист "${sheetName}" не найден`);
      }
    }
    
    if (!targetSheet) {
      // Поиск листа с данными для текущего месяца
      targetSheet = this.findCurrentMonthSheet(sheets);
    }
    
    if (!targetSheet) {
      // Берем первый лист
      targetSheet = sheets[0];
      this.logWarning('⚠️ Используется первый лист');
    }
    
    this.log(`✅ Выбран лист: "${targetSheet.getName()}"`);
    return targetSheet;
  }
  
  findCurrentMonthSheet(sheets) {
    const currentMonth = new Date().getMonth();
    const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 
                       'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    
    // Поиск листа с названием месяца
    for (const sheet of sheets) {
      const sheetName = sheet.getName().toLowerCase();
      if (sheetName.includes(monthNames[currentMonth])) {
        return sheet;
      }
    }
    
    return null;
  }
  
  getData(sheet) {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    this.log(`📊 Размер данных: ${lastRow}x${lastColumn}`);
    
    if (lastRow === 0 || lastColumn === 0) {
      throw new Error('Нет данных для обработки');
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    this.log(`✅ Загружено ${data.length} строк данных`);
    
    this.diagnostics.statistics.totalRows = data.length;
    
    return data;
  }
  
  // ==================== АНАЛИЗ СТРУКТУРЫ ====================
  
  analyzeDataStructure(data) {
    this.log('🔍 АНАЛИЗ СТРУКТУРЫ ДАННЫХ');
    this.log('=========================');
    
    // Анализируем первые 10 строк для поиска заголовков
    this.log('📋 Поиск заголовков...');
    
    let headerRow = -1;
    let maxMatches = 0;
    
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const row = data[i];
      const rowText = row.join(' ').toLowerCase();
      
      // Подсчитываем совпадения с ключевыми словами
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
    
    // Анализируем разделы данных
    this.analyzeSections(data);
    
    // Определяем месяц
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
    
    this.log('✅ Использован стандартный маппинг колонок');
  }
  
  analyzeSections(data) {
    this.log('🔍 Анализ разделов данных...');
    
    const sections = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length === 0) continue;
      
      const firstCell = (row[0] || '').toString().toLowerCase().trim();
      
      // Поиск заголовков разделов
      if (this.isSectionHeader(firstCell)) {
        const sectionType = this.determineSectionType(firstCell);
        sections.push({
          row: i + 1,
          name: firstCell,
          type: sectionType
        });
        
        this.log(`  Раздел найден: строка ${i + 1} - "${firstCell}" (${sectionType})`);
      }
    }
    
    this.log(`✅ Найдено разделов: ${sections.length}`);
    this.sections = sections;
  }
  
  isSectionHeader(text) {
    const sectionPatterns = [
      'отзыв', 'ос', 'основное сообщение',
      'целевые', 'цс', 'целевое сообщение',
      'площадки', 'пс', 'социальные',
      'обсуждения', 'дискуссии', 'комментарии'
    ];
    
    return sectionPatterns.some(pattern => text.includes(pattern));
  }
  
  determineSectionType(text) {
    if (text.includes('отзыв') || text.includes('ос')) return 'reviews';
    if (text.includes('целевые') || text.includes('цс')) return 'targeted';
    if (text.includes('площадки') || text.includes('пс')) return 'social';
    if (text.includes('обсуждения') || text.includes('дискуссии')) return 'discussions';
    return 'other';
  }
  
  detectMonth(data) {
    this.log('📅 Определение месяца...');
    
    const monthNames = [
      'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
      'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'
    ];
    
    const monthShort = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 
                       'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    
    // Поиск в первых 5 строках
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      
      for (let j = 0; j < monthNames.length; j++) {
        if (rowText.includes(monthNames[j]) || rowText.includes(monthShort[j])) {
          this.monthInfo = {
            name: monthNames[j],
            short: monthShort[j],
            number: j + 1,
            year: new Date().getFullYear()
          };
          
          this.log(`✅ Определен месяц: ${this.monthInfo.name} ${this.monthInfo.year}`);
          return;
        }
      }
    }
    
    // Если не найдено, используем текущий месяц
    const currentMonth = new Date().getMonth();
    this.monthInfo = {
      name: monthNames[currentMonth],
      short: monthShort[currentMonth],
      number: currentMonth + 1,
      year: new Date().getFullYear()
    };
    
    this.logWarning(`⚠️ Месяц не найден, используется текущий: ${this.monthInfo.name}`);
  }
  
  // ==================== ОБРАБОТКА ДАННЫХ ====================
  
  processData(data) {
    this.log('🔄 ОБРАБОТКА ДАННЫХ');
    this.log('==================');
    
    let currentSection = null;
    let processedRows = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Пропускаем пустые строки
      if (this.isEmptyRow(row)) {
        continue;
      }
      
      // Проверяем, является ли строка заголовком раздела
      const firstCell = (row[0] || '').toString().toLowerCase().trim();
      if (this.isSectionHeader(firstCell)) {
        currentSection = this.determineSectionType(firstCell);
        this.log(`📋 Переход к разделу: ${currentSection}`);
        continue;
      }
      
      // Пропускаем строки заголовков таблицы
      if (this.isTableHeader(row)) {
        this.log(`📊 Пропуск заголовка таблицы: строка ${i + 1}`);
        continue;
      }
      
      // Обрабатываем строки данных
      if (currentSection && this.isDataRow(row)) {
        const processedRow = this.processRow(row, currentSection, i + 1);
        
        if (processedRow) {
          this.processedData[currentSection].push(processedRow);
          processedRows++;
          
          // Обновляем статистику
          this.updateStatistics(processedRow, currentSection);
          
          if (ULTIMATE_CONFIG.DIAGNOSTICS.LOG_EVERY_ROW) {
            this.log(`  ✅ Обработана строка ${i + 1}: ${processedRow.platform} - ${processedRow.views} просмотров`);
          }
        }
      }
    }
    
    this.diagnostics.statistics.processedRows = processedRows;
    this.log(`✅ Обработано ${processedRows} строк данных`);
    
    // Выводим статистику по разделам
    Object.keys(this.processedData).forEach(section => {
      const count = this.processedData[section].length;
      if (count > 0) {
        this.log(`  📊 ${section}: ${count} записей`);
      }
    });
  }
  
  isEmptyRow(row) {
    return !row || row.every(cell => !cell || cell.toString().trim() === '');
  }
  
  isTableHeader(row) {
    if (!row || row.length === 0) return false;
    
    const rowText = row.join(' ').toLowerCase();
    const headerPatterns = ['тип размещения', 'площадка', 'текст сообщения'];
    
    return headerPatterns.some(pattern => rowText.includes(pattern));
  }
  
  isDataRow(row) {
    if (!row || row.length < 3) return false;
    
    // Проверяем, что есть основные данные
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
        approvalComments: this.getColumnValue(row, 'согласование/комментарии'),
        date: this.getColumnValue(row, 'дата'),
        nickname: this.getColumnValue(row, 'ник'),
        author: this.getColumnValue(row, 'автор'),
        startViews: this.extractNumber(this.getColumnValue(row, 'просмотры темы на старте')),
        endViews: this.extractNumber(this.getColumnValue(row, 'просмотры в конце месяца')),
        views: this.extractNumber(this.getColumnValue(row, 'просмотров получено')),
        engagement: this.extractNumber(this.getColumnValue(row, 'вовлечение')),
        postType: this.getColumnValue(row, 'тип поста')
      };
      
      // Валидация обязательных полей
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
  
  // ==================== СОЗДАНИЕ ОТЧЕТА ====================
  
  createReport(spreadsheet) {
    this.log('📋 СОЗДАНИЕ ОТЧЕТА');
    this.log('==================');
    
    const sheetName = ULTIMATE_CONFIG.OUTPUT.SHEET_NAME_TEMPLATE
      .replace('{month}', this.monthInfo.name)
      .replace('{year}', this.monthInfo.year);
    
    // Создаем новый лист для отчета
    const reportSheet = spreadsheet.insertSheet(sheetName);
    this.log(`✅ Создан лист отчета: "${sheetName}"`);
    
    // Создаем заголовки
    const headers = [
      'Раздел', 'Тип размещения', 'Площадка', 'Продукт', 'Ссылка', 
      'Текст сообщения', 'Дата', 'Автор', 'Просмотры получено', 'Вовлечение'
    ];
    
    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Форматируем заголовки
    const headerRange = reportSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    // Добавляем данные
    let currentRow = 2;
    
    Object.keys(this.processedData).forEach(section => {
      const sectionData = this.processedData[section];
      
      if (sectionData.length > 0) {
        // Добавляем заголовок раздела
        reportSheet.getRange(currentRow, 1).setValue(section.toUpperCase());
        reportSheet.getRange(currentRow, 1, 1, headers.length).setBackground('#f0f0f0');
        currentRow++;
        
        // Добавляем данные раздела
        sectionData.forEach(row => {
          const rowData = [
            section,
            row.placementType,
            row.platform,
            row.product,
            row.messageLink,
            row.messageText.substring(0, 100) + (row.messageText.length > 100 ? '...' : ''),
            row.date,
            row.author,
            row.views,
            row.engagement
          ];
          
          reportSheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
          currentRow++;
        });
        
        currentRow++; // Пустая строка между разделами
      }
    });
    
    // Добавляем итоговую строку
    if (ULTIMATE_CONFIG.OUTPUT.ADD_TOTAL_ROW) {
      this.addTotalRow(reportSheet, currentRow, headers.length);
    }
    
    // Автоматическое изменение размера колонок
    reportSheet.autoResizeColumns(1, headers.length);
    
    this.log(`✅ Отчет создан: ${sectionData.length} записей`);
    
    return reportSheet;
  }
  
  addTotalRow(sheet, row, columnsCount) {
    const stats = this.diagnostics.statistics;
    
    const totalData = [
      'ИТОГО',
      '',
      `Платформ: ${Object.keys(this.processedData).length}`,
      '',
      '',
      `Всего записей: ${stats.processedRows}`,
      '',
      '',
      stats.totalViews,
      stats.totalEngagement
    ];
    
    // Добавляем пустую строку
    sheet.getRange(row, 1, 1, columnsCount).setValues([Array(columnsCount).fill('')]);
    
    // Добавляем итоговую строку
    const totalRange = sheet.getRange(row + 1, 1, 1, totalData.length);
    totalRange.setValues([totalData]);
    totalRange.setFontWeight('bold');
    totalRange.setBackground('#e8f0fe');
    
    this.log('✅ Добавлена итоговая строка');
  }
  
  // ==================== ДИАГНОСТИКА ====================
  
  finalizeDiagnostics() {
    this.diagnostics.endTime = new Date();
    this.diagnostics.totalTime = this.diagnostics.endTime - this.diagnostics.startTime;
    
    this.log('📊 ФИНАЛЬНАЯ ДИАГНОСТИКА');
    this.log('========================');
    this.log(`⏱️ Общее время: ${this.diagnostics.totalTime}ms`);
    this.log(`📊 Всего строк: ${this.diagnostics.statistics.totalRows}`);
    this.log(`✅ Обработано: ${this.diagnostics.statistics.processedRows}`);
    this.log(`📝 Отзывов: ${this.diagnostics.statistics.reviewsCount}`);
    this.log(`🎯 Целевых: ${this.diagnostics.statistics.targetedCount}`);
    this.log(`📱 Социальных: ${this.diagnostics.statistics.socialCount}`);
    this.log(`👁️ Всего просмотров: ${this.diagnostics.statistics.totalViews}`);
    this.log(`💬 Всего вовлечений: ${this.diagnostics.statistics.totalEngagement}`);
    
    if (this.diagnostics.errors.length > 0) {
      this.log(`❌ Ошибок: ${this.diagnostics.errors.length}`);
      this.diagnostics.errors.forEach(error => this.log(`   ${error}`));
    }
    
    if (this.diagnostics.warnings.length > 0) {
      this.log(`⚠️ Предупреждений: ${this.diagnostics.warnings.length}`);
      this.diagnostics.warnings.forEach(warning => this.log(`   ${warning}`));
    }
  }
  
  log(message) {
    if (ULTIMATE_CONFIG.DIAGNOSTICS.ENABLE_DETAILED_LOGGING) {
      console.log(message);
    }
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

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Главная функция для запуска обработки
 */
function processUltimateGoogleSheets() {
  const processor = new UltimateGoogleAppsScriptProcessor();
  return processor.processSheet();
}

/**
 * Тестирование на конкретной таблице
 */
function testUltimateProcessor(sheetId, sheetName) {
  const processor = new UltimateGoogleAppsScriptProcessor();
  return processor.processSheet(sheetId, sheetName);
}

/**
 * Анализ структуры данных
 */
function analyzeUltimateDataStructure(sheetId) {
  const processor = new UltimateGoogleAppsScriptProcessor();
  
  try {
    const spreadsheet = processor.openSpreadsheet(sheetId || ULTIMATE_CONFIG.TEST_SHEET_ID);
    const sheet = processor.selectSheet(spreadsheet);
    const data = processor.getData(sheet);
    
    processor.analyzeDataStructure(data);
    
    return {
      success: true,
      diagnostics: processor.diagnostics,
      columnMapping: processor.columnMapping,
      monthInfo: processor.monthInfo,
      sections: processor.sections
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Получение диагностической информации
 */
function getUltimateDiagnostics() {
  const processor = new UltimateGoogleAppsScriptProcessor();
  return processor.diagnostics;
}

// Функция для тестирования в Google Apps Script
function main() {
  return processUltimateGoogleSheets();
}