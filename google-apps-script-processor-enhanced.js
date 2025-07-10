/**
 * 🚀 УЛУЧШЕННЫЙ ГИБКИЙ ОБРАБОТЧИК ДЛЯ ЕЖЕМЕСЯЧНОЙ РАБОТЫ
 * Google Apps Script для автоматической обработки отчетов
 * Версия для тестирования с эталонными данными
 * 
 * Автор: AI Assistant + Background Agent
 * Версия: 2.0.0
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ ====================

const CONFIG = {
  // Настройки обработки
  PROCESSING: {
    MAX_ROWS: 10000,
    BATCH_SIZE: 100,
    TIMEOUT_SECONDS: 300
  },
  
  // Настройки колонок (гибкие)
  COLUMNS: {
    // Основные колонки (будут определяться автоматически)
    PLATFORM: ['площадка', 'platform', 'site', 'платформа'],
    TEXT: ['текст сообщения', 'текст', 'message', 'content', 'сообщение'],
    DATE: ['дата', 'date', 'created', 'время'],
    AUTHOR: ['автор', 'ник', 'author', 'nickname', 'пользователь'],
    VIEWS: ['просмотры', 'просмотров получено', 'views', 'просмотров'],
    ENGAGEMENT: ['вовлечение', 'engagement', 'взаимодействие'],
    POST_TYPE: ['тип поста', 'тип размещения', 'post_type', 'тип'],
    
    // Дополнительные колонки
    THEME: ['тема', 'theme', 'subject', 'заголовок'],
    LINK: ['ссылка', 'link', 'url', 'адрес']
  },
  
  // Типы постов (расширенные)
  POST_TYPES: {
    REVIEW: ['ос', 'основное сообщение', 'review', 'отзыв', 'основной'],
    COMMENT: ['цс', 'целевое сообщение', 'comment', 'комментарий', 'целевой']
  },
  
  // Настройки форматирования
  FORMATTING: {
    DATE_FORMAT: 'dd.mm.yyyy',
    NUMBER_FORMAT: '#,##0',
    CURRENCY_FORMAT: '#,##0 ₽'
  },
  
  // Настройки тестирования
  TESTING: {
    ENABLE_DEBUG: true,
    LOG_DETAILS: true,
    VALIDATE_RESULTS: true
  }
};

// ==================== КЛАСС ОБРАБОТЧИКА ====================

/**
 * Улучшенный класс обработки ежемесячных отчетов
 */
class EnhancedMonthlyReportProcessor {
  constructor() {
    this.stats = {
      totalRows: 0,
      reviewsCount: 0,
      commentsCount: 0,
      totalViews: 0,
      processingTime: 0,
      errors: []
    };
    
    this.columnMapping = {};
    this.monthInfo = null;
  }

  /**
   * Главный метод обработки отчета
   */
  processReport(spreadsheetId, sheetName = null) {
    const startTime = Date.now();
    
    try {
      console.log('🚀 ENHANCED PROCESSOR - Начало обработки');
      
      // 1. Получение данных
      const sourceData = this.getSourceData(spreadsheetId, sheetName);
      
      // 2. Определение месяца
      this.monthInfo = this.detectMonth(sourceData);
      console.log(`📅 Определен месяц: ${this.monthInfo.name} ${this.monthInfo.year}`);
      
      // 3. Анализ структуры данных
      this.analyzeDataStructure(sourceData);
      
      // 4. Обработка данных
      const processedData = this.processData(sourceData);
      
      // 5. Создание отчета
      const reportUrl = this.createReport(processedData);
      
      // 6. Обновление статистики
      this.stats.processingTime = Date.now() - startTime;
      
      console.log('✅ Обработка завершена успешно');
      console.log('📊 Статистика:', this.stats);
      
      return {
        success: true,
        reportUrl: reportUrl,
        statistics: this.stats,
        monthInfo: this.monthInfo
      };
      
    } catch (error) {
      console.error('❌ Ошибка обработки:', error);
      this.stats.errors.push(error.toString());
      
      return {
        success: false,
        error: error.toString(),
        statistics: this.stats
      };
    }
  }

  /**
   * Получение исходных данных
   */
  getSourceData(spreadsheetId, sheetName) {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Лист "${sheetName}" не найден`);
    }
    
    const data = sheet.getDataRange().getValues();
    this.stats.totalRows = data.length;
    
    console.log(`📊 Загружено ${data.length} строк данных`);
    return data;
  }

  /**
   * Улучшенное определение месяца
   */
  detectMonth(data) {
    // Поиск в названии листа
    const sheetName = SpreadsheetApp.getActiveSheet().getName();
    const monthFromSheet = this.extractMonthFromText(sheetName);
    if (monthFromSheet) {
      return { ...monthFromSheet, detectedFrom: 'sheet' };
    }
    
    // Поиск в данных (первые 20 строк)
    for (let i = 0; i < Math.min(20, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      const monthFromData = this.extractMonthFromText(rowText);
      if (monthFromData) {
        return { ...monthFromData, detectedFrom: 'content' };
      }
    }
    
    // Поиск в имени файла
    const fileName = SpreadsheetApp.getActiveSpreadsheet().getName();
    const monthFromFile = this.extractMonthFromText(fileName);
    if (monthFromFile) {
      return { ...monthFromFile, detectedFrom: 'filename' };
    }
    
    // По умолчанию - текущий месяц
    const now = new Date();
    return {
      name: this.getMonthName(now.getMonth()),
      short: this.getMonthShort(now.getMonth()),
      number: now.getMonth() + 1,
      year: now.getFullYear(),
      detectedFrom: 'default'
    };
  }

  /**
   * Извлечение месяца из текста
   */
  extractMonthFromText(text) {
    const lowerText = text.toLowerCase();
    
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
      const variants = [
        month.name.toLowerCase(),
        month.short.toLowerCase(),
        `${month.short}25`,
        `${month.name}25`,
        `${month.short}2025`,
        `${month.name}2025`
      ];
      
      if (variants.some(variant => lowerText.includes(variant))) {
        return {
          name: month.name,
          short: month.short,
          number: month.number,
          year: 2025
        };
      }
    }
    
    return null;
  }

  /**
   * Анализ структуры данных
   */
  analyzeDataStructure(data) {
    if (data.length === 0) {
      throw new Error('Данные пусты');
    }
    
    const headers = data[0];
    console.log('📋 Заголовки:', headers);
    
    // Определяем маппинг колонок
    this.columnMapping = {};
    
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();
      
      // Площадка
      if (CONFIG.COLUMNS.PLATFORM.some(keyword => header.includes(keyword))) {
        this.columnMapping.platform = i;
      }
      
      // Текст
      if (CONFIG.COLUMNS.TEXT.some(keyword => header.includes(keyword))) {
        this.columnMapping.text = i;
      }
      
      // Дата
      if (CONFIG.COLUMNS.DATE.some(keyword => header.includes(keyword))) {
        this.columnMapping.date = i;
      }
      
      // Автор
      if (CONFIG.COLUMNS.AUTHOR.some(keyword => header.includes(keyword))) {
        this.columnMapping.author = i;
      }
      
      // Просмотры
      if (CONFIG.COLUMNS.VIEWS.some(keyword => header.includes(keyword))) {
        this.columnMapping.views = i;
      }
      
      // Тип поста
      if (CONFIG.COLUMNS.POST_TYPE.some(keyword => header.includes(keyword))) {
        this.columnMapping.postType = i;
      }
    }
    
    console.log('🗺️ Маппинг колонок:', this.columnMapping);
    
    // Проверяем обязательные колонки
    const requiredColumns = ['platform', 'text', 'date'];
    const missingColumns = requiredColumns.filter(col => this.columnMapping[col] === undefined);
    
    if (missingColumns.length > 0) {
      console.warn(`⚠️ Не найдены колонки: ${missingColumns.join(', ')}`);
    }
  }

  /**
   * Обработка данных
   */
  processData(data) {
    const processedData = {
      reviews: [],
      comments: [],
      statistics: {
        totalReviews: 0,
        totalComments: 0,
        totalViews: 0,
        platforms: new Set()
      }
    };
    
    let currentSection = null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Пропускаем пустые строки
      if (this.isEmptyRow(row)) continue;
      
      // Определяем секцию
      const section = this.detectSection(row);
      if (section) {
        currentSection = section;
        continue;
      }
      
      // Обрабатываем данные
      if (currentSection && this.isDataRow(row)) {
        const record = this.processRow(row, currentSection);
        
        if (record) {
          if (currentSection === 'reviews') {
            processedData.reviews.push(record);
            processedData.statistics.totalReviews++;
          } else if (currentSection === 'comments') {
            processedData.comments.push(record);
            processedData.statistics.totalComments++;
          }
          
          // Обновляем статистику
          processedData.statistics.totalViews += record.views || 0;
          if (record.platform) {
            processedData.statistics.platforms.add(record.platform);
          }
        }
      }
    }
    
    // Обновляем глобальную статистику
    this.stats.reviewsCount = processedData.statistics.totalReviews;
    this.stats.commentsCount = processedData.statistics.totalComments;
    this.stats.totalViews = processedData.statistics.totalViews;
    
    console.log(`📊 Обработано: ${processedData.statistics.totalReviews} отзывов, ${processedData.statistics.totalComments} комментариев`);
    
    return processedData;
  }

  /**
   * Определение секции
   */
  detectSection(row) {
    if (row.length === 0) return null;
    
    const firstCell = String(row[0]).toLowerCase();
    
    if (firstCell.includes('отзывы') || firstCell.includes('reviews')) {
      return 'reviews';
    }
    
    if (firstCell.includes('комментарии') || firstCell.includes('comments')) {
      return 'comments';
    }
    
    return null;
  }

  /**
   * Проверка на пустую строку
   */
  isEmptyRow(row) {
    return !row || row.every(cell => !cell || String(cell).trim() === '');
  }

  /**
   * Проверка на строку с данными
   */
  isDataRow(row) {
    if (row.length < 3) return false;
    
    // Проверяем наличие значимого текста
    const textIndex = this.columnMapping.text || 2;
    const text = row[textIndex];
    
    return text && String(text).trim().length > 10;
  }

  /**
   * Обработка строки данных
   */
  processRow(row, section) {
    try {
      const record = {
        section: section,
        platform: this.extractPlatform(row),
        text: this.extractText(row),
        date: this.extractDate(row),
        author: this.extractAuthor(row),
        views: this.extractViews(row),
        postType: this.extractPostType(row),
        theme: this.extractTheme(row),
        link: this.extractLink(row)
      };
      
      // Валидация записи
      if (!record.text || record.text.length < 10) {
        return null;
      }
      
      return record;
      
    } catch (error) {
      console.warn(`⚠️ Ошибка обработки строки ${row}:`, error);
      return null;
    }
  }

  /**
   * Извлечение платформы
   */
  extractPlatform(row) {
    const index = this.columnMapping.platform;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return 'Неизвестно';
  }

  /**
   * Извлечение текста
   */
  extractText(row) {
    const index = this.columnMapping.text;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    
    // Поиск в других колонках
    for (let i = 0; i < row.length; i++) {
      const cell = row[i];
      if (cell && String(cell).length > 20) {
        return String(cell).trim();
      }
    }
    
    return '';
  }

  /**
   * Извлечение даты
   */
  extractDate(row) {
    const index = this.columnMapping.date;
    if (index !== undefined && row[index]) {
      const dateValue = row[index];
      
      if (dateValue instanceof Date) {
        return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), CONFIG.FORMATTING.DATE_FORMAT);
      }
      
      return String(dateValue);
    }
    
    return '';
  }

  /**
   * Извлечение автора
   */
  extractAuthor(row) {
    const index = this.columnMapping.author;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return 'Аноним';
  }

  /**
   * Извлечение просмотров
   */
  extractViews(row) {
    const index = this.columnMapping.views;
    if (index !== undefined && row[index]) {
      const viewsStr = String(row[index]).replace(/[^\d]/g, '');
      const views = parseInt(viewsStr);
      return isNaN(views) ? 0 : views;
    }
    return 0;
  }

  /**
   * Извлечение типа поста
   */
  extractPostType(row) {
    const index = this.columnMapping.postType;
    if (index !== undefined && row[index]) {
      const type = String(row[index]).toLowerCase();
      
      if (CONFIG.POST_TYPES.REVIEW.some(keyword => type.includes(keyword))) {
        return 'Отзыв';
      }
      
      if (CONFIG.POST_TYPES.COMMENT.some(keyword => type.includes(keyword))) {
        return 'Комментарий';
      }
    }
    
    return 'Неизвестно';
  }

  /**
   * Извлечение темы
   */
  extractTheme(row) {
    const index = this.columnMapping.theme;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение ссылки
   */
  extractLink(row) {
    const index = this.columnMapping.link;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Создание отчета
   */
  createReport(processedData) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheetName = `Отчет_${this.monthInfo.name}_${this.monthInfo.year}`;
    
    // Удаляем старый отчет если есть
    const existingSheet = spreadsheet.getSheetByName(reportSheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const reportSheet = spreadsheet.insertSheet(reportSheetName);
    
    // Заголовок отчета
    reportSheet.getRange('A1').setValue(`ОТЧЕТ ЗА ${this.monthInfo.name.toUpperCase()} ${this.monthInfo.year}`);
    reportSheet.getRange('A1:H1').merge();
    reportSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
    
    // Статистика
    reportSheet.getRange('A3').setValue('СТАТИСТИКА:');
    reportSheet.getRange('A3').setFontWeight('bold');
    
    const statsData = [
      ['Всего отзывов:', this.stats.reviewsCount],
      ['Всего комментариев:', this.stats.commentsCount],
      ['Общие просмотры:', this.stats.totalViews],
      ['Платформ:', Array.from(processedData.statistics.platforms).length],
      ['Время обработки:', `${(this.stats.processingTime / 1000).toFixed(2)} сек`]
    ];
    
    reportSheet.getRange(4, 1, statsData.length, 2).setValues(statsData);
    
    // Отзывы
    let currentRow = statsData.length + 6;
    reportSheet.getRange(`A${currentRow}`).setValue('ОТЗЫВЫ:');
    reportSheet.getRange(`A${currentRow}`).setFontWeight('bold').setFontSize(12);
    currentRow++;
    
    const reviewHeaders = ['Платформа', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Тип', 'Тема', 'Ссылка'];
    reportSheet.getRange(currentRow, 1, 1, reviewHeaders.length).setValues([reviewHeaders]);
    reportSheet.getRange(currentRow, 1, 1, reviewHeaders.length).setFontWeight('bold');
    currentRow++;
    
    if (processedData.reviews.length > 0) {
      const reviewData = processedData.reviews.map(review => [
        review.platform,
        review.text,
        review.date,
        review.author,
        review.views,
        review.postType,
        review.theme,
        review.link
      ]);
      
      reportSheet.getRange(currentRow, 1, reviewData.length, reviewHeaders.length).setValues(reviewData);
      currentRow += reviewData.length;
    }
    
    // Комментарии
    currentRow += 2;
    reportSheet.getRange(`A${currentRow}`).setValue('КОММЕНТАРИИ:');
    reportSheet.getRange(`A${currentRow}`).setFontWeight('bold').setFontSize(12);
    currentRow++;
    
    reportSheet.getRange(currentRow, 1, 1, reviewHeaders.length).setValues([reviewHeaders]);
    reportSheet.getRange(currentRow, 1, 1, reviewHeaders.length).setFontWeight('bold');
    currentRow++;
    
    if (processedData.comments.length > 0) {
      const commentData = processedData.comments.map(comment => [
        comment.platform,
        comment.text,
        comment.date,
        comment.author,
        comment.views,
        comment.postType,
        comment.theme,
        comment.link
      ]);
      
      reportSheet.getRange(currentRow, 1, commentData.length, reviewHeaders.length).setValues(commentData);
      currentRow += commentData.length;
    }
    
    // Итоговая строка
    currentRow += 2;
    const totalRow = [
      'ИТОГО:',
      '',
      '',
      '',
      this.stats.totalViews,
      '',
      '',
      ''
    ];
    reportSheet.getRange(currentRow, 1, 1, totalRow.length).setValues([totalRow]);
    reportSheet.getRange(currentRow, 1, 1, totalRow.length).setFontWeight('bold');
    
    // Форматирование
    reportSheet.autoResizeColumns(1, reviewHeaders.length);
    
    // Применяем фильтры
    const dataRange = reportSheet.getDataRange();
    reportSheet.getRange(1, 1, dataRange.getNumRows(), dataRange.getNumColumns()).createFilter();
    
    console.log(`📄 Отчет создан: ${reportSheetName}`);
    
    return spreadsheet.getUrl();
  }

  /**
   * Вспомогательные методы
   */
  getMonthName(monthIndex) {
    const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                   'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    return months[monthIndex];
  }

  getMonthShort(monthIndex) {
    const months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                   'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
    return months[monthIndex];
  }
}

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Главная функция обработки
 */
function processMonthlyReport(spreadsheetId = null, sheetName = null) {
  const processor = new EnhancedMonthlyReportProcessor();
  
  if (!spreadsheetId) {
    spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }
  
  return processor.processReport(spreadsheetId, sheetName);
}

/**
 * Быстрая обработка текущего листа
 */
function quickProcess() {
  return processMonthlyReport();
}

/**
 * Обработка с выбором листа
 */
function processWithSheetSelection() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  const sheetNames = sheets.map(sheet => sheet.getName());
  const response = ui.prompt(
    'Выбор листа',
    `Доступные листы:\n${sheetNames.join('\n')}\n\nВведите название листа для обработки:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const selectedSheet = response.getResponseText().trim();
    return processMonthlyReport(spreadsheet.getId(), selectedSheet);
  }
}

/**
 * Обновление меню
 */
function updateMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Обработка отчетов')
    .addItem('Быстрая обработка', 'quickProcess')
    .addItem('Обработка с выбором листа', 'processWithSheetSelection')
    .addSeparator()
    .addItem('Настройки', 'showSettings')
    .addToUi();
}

/**
 * Показать настройки
 */
function showSettings() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Настройки процессора',
    `Конфигурация:\n` +
    `- Максимум строк: ${CONFIG.PROCESSING.MAX_ROWS}\n` +
    `- Размер пакета: ${CONFIG.PROCESSING.BATCH_SIZE}\n` +
    `- Таймаут: ${CONFIG.PROCESSING.TIMEOUT_SECONDS} сек\n` +
    `- Отладка: ${CONFIG.TESTING.ENABLE_DEBUG ? 'Включена' : 'Выключена'}`,
    ui.ButtonSet.OK
  );
}

/**
 * Инициализация при открытии
 */
function onOpen() {
  updateMenu();
} 