/**
 * 🚀 ИДЕАЛЬНЫЙ ГИБКИЙ ОБРАБОТЧИК ДЛЯ ЕЖЕМЕСЯЧНОЙ РАБОТЫ
 * Google Apps Script для автоматической обработки отчетов
 * 
 * Автор: AI Assistant + Background Agent
 * Версия: 1.0.0
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
    PLATFORM: ['площадка', 'platform', 'site'],
    TEXT: ['текст сообщения', 'текст', 'message', 'content'],
    DATE: ['дата', 'date', 'created'],
    AUTHOR: ['автор', 'ник', 'author', 'nickname'],
    VIEWS: ['просмотры', 'просмотров получено', 'views'],
    ENGAGEMENT: ['вовлечение', 'engagement'],
    POST_TYPE: ['тип поста', 'тип размещения', 'post_type'],
    
    // Дополнительные колонки
    THEME: ['тема', 'theme', 'subject'],
    LINK: ['ссылка', 'link', 'url']
  },
  
  // Типы постов
  POST_TYPES: {
    REVIEW: ['ос', 'основное сообщение', 'review', 'отзыв'],
    COMMENT: ['цс', 'целевое сообщение', 'comment', 'комментарий']
  },
  
  // Месяцы
  MONTHS: {
    'январь': { name: 'Январь', short: 'Янв', number: 1 },
    'янв': { name: 'Январь', short: 'Янв', number: 1 },
    'февраль': { name: 'Февраль', short: 'Фев', number: 2 },
    'фев': { name: 'Февраль', short: 'Фев', number: 2 },
    'март': { name: 'Март', short: 'Мар', number: 3 },
    'мар': { name: 'Март', short: 'Мар', number: 3 },
    'апрель': { name: 'Апрель', short: 'Апр', number: 4 },
    'апр': { name: 'Апрель', short: 'Апр', number: 4 },
    'май': { name: 'Май', short: 'Май', number: 5 },
    'июнь': { name: 'Июнь', short: 'Июн', number: 6 },
    'июн': { name: 'Июнь', short: 'Июн', number: 6 },
    'июль': { name: 'Июль', short: 'Июл', number: 7 },
    'июл': { name: 'Июль', short: 'Июл', number: 7 },
    'август': { name: 'Август', short: 'Авг', number: 8 },
    'авг': { name: 'Август', short: 'Авг', number: 8 },
    'сентябрь': { name: 'Сентябрь', short: 'Сен', number: 9 },
    'сен': { name: 'Сентябрь', short: 'Сен', number: 9 },
    'октябрь': { name: 'Октябрь', short: 'Окт', number: 10 },
    'окт': { name: 'Октябрь', short: 'Окт', number: 10 },
    'ноябрь': { name: 'Ноябрь', short: 'Ноя', number: 11 },
    'ноя': { name: 'Ноябрь', short: 'Ноя', number: 11 },
    'декабрь': { name: 'Декабрь', short: 'Дек', number: 12 },
    'дек': { name: 'Декабрь', short: 'Дек', number: 12 }
  }
};

// ==================== ОСНОВНЫЕ КЛАССЫ ====================

/**
 * Главный класс обработчика
 */
class MonthlyReportProcessor {
  constructor() {
    this.stats = {
      totalRows: 0,
      reviewsCount: 0,
      commentsCount: 0,
      totalViews: 0,
      processingTime: 0,
      qualityScore: 0,
      platformsWithData: 0
    };
    
    this.columnMapping = {};
    this.monthInfo = null;
    this.processedData = {
      reviews: [],
      comments: [],
      discussions: []
    };
  }

  /**
   * Основной метод обработки
   */
  processReport(sourceSheetId, sourceSheetName = null) {
    const startTime = new Date();
    
    try {
      console.log('🚀 Начало обработки ежемесячного отчета...');
      
      // 1. Получаем исходные данные
      const sourceData = this.getSourceData(sourceSheetId, sourceSheetName);
      
      // 2. Определяем месяц
      this.monthInfo = this.detectMonth(sourceData);
      console.log(`📅 Определен месяц: ${this.monthInfo.name}`);
      
      // 3. Анализируем структуру данных
      this.analyzeStructure(sourceData);
      
      // 4. Обрабатываем данные
      this.processData(sourceData);
      
      // 5. Создаем отчет
      const reportSheetId = this.createReport();
      
      // 6. Рассчитываем статистику
      this.calculateStatistics(startTime);
      
      console.log('✅ Обработка завершена успешно!');
      
      return {
        success: true,
        reportSheetId: reportSheetId,
        statistics: this.stats,
        monthInfo: this.monthInfo
      };
      
    } catch (error) {
      console.error('❌ Ошибка при обработке:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  }

  /**
   * Получение исходных данных
   */
  getSourceData(sheetId, sheetName) {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
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
   * Умное определение месяца
   */
  detectMonth(data) {
    // 1. Поиск в названиях листов
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    for (const sheetName of spreadsheet.getSheetNames()) {
      const monthInfo = this.findMonthInText(sheetName);
      if (monthInfo) {
        return { ...monthInfo, detectedFrom: 'sheet' };
      }
    }
    
    // 2. Поиск в данных
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      const monthInfo = this.findMonthInText(rowText);
      if (monthInfo) {
        return { ...monthInfo, detectedFrom: 'content' };
      }
    }
    
    // 3. Текущий месяц по умолчанию
    const currentMonth = new Date().getMonth();
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    const shortNames = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                       'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
    
    return {
      name: monthNames[currentMonth],
      shortName: shortNames[currentMonth],
      number: currentMonth + 1,
      detectedFrom: 'default'
    };
  }

  /**
   * Поиск месяца в тексте
   */
  findMonthInText(text) {
    const lowerText = text.toLowerCase();
    for (const [key, value] of Object.entries(CONFIG.MONTHS)) {
      if (lowerText.includes(key)) {
        return value;
      }
    }
    return null;
  }

  /**
   * Анализ структуры данных
   */
  analyzeStructure(data) {
    console.log('🔍 Анализ структуры данных...');
    
    // Ищем строку с заголовками
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      if (this.isHeaderRow(rowText)) {
        headerRowIndex = i;
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      throw new Error('Не найдена строка с заголовками');
    }
    
    const headers = data[headerRowIndex];
    this.createColumnMapping(headers);
    
    console.log('📋 Найдены колонки:', this.columnMapping);
  }

  /**
   * Проверка, является ли строка заголовком
   */
  isHeaderRow(rowText) {
    const headerKeywords = ['тип размещения', 'площадка', 'текст сообщения', 'дата', 'автор'];
    return headerKeywords.some(keyword => rowText.includes(keyword));
  }

  /**
   * Создание маппинга колонок
   */
  createColumnMapping(headers) {
    this.columnMapping = {};
    
    headers.forEach((header, index) => {
      if (!header) return;
      
      const cleanHeader = this.cleanText(header);
      
      // Ищем соответствия для каждой категории колонок
      for (const [category, keywords] of Object.entries(CONFIG.COLUMNS)) {
        if (keywords.some(keyword => cleanHeader.includes(keyword))) {
          this.columnMapping[category] = index;
          break;
        }
      }
    });
    
    // Проверяем обязательные колонки
    const requiredColumns = ['PLATFORM', 'TEXT', 'DATE'];
    const missingColumns = requiredColumns.filter(col => !(col in this.columnMapping));
    
    if (missingColumns.length > 0) {
      console.warn(`⚠️ Не найдены колонки: ${missingColumns.join(', ')}`);
    }
  }

  /**
   * Обработка данных
   */
  processData(data) {
    console.log('🔄 Обработка данных...');
    
    let processedRows = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Пропускаем заголовки и пустые строки
      if (this.isEmptyRow(row) || this.isHeaderRow(row.join(' '))) {
        continue;
      }
      
      // Определяем тип записи
      const recordType = this.determineRecordType(row);
      
      if (recordType === 'review') {
        const reviewData = this.extractReviewData(row, i);
        if (reviewData) {
          this.processedData.reviews.push(reviewData);
          this.stats.reviewsCount++;
        }
      } else if (recordType === 'comment') {
        const commentData = this.extractCommentData(row, i);
        if (commentData) {
          this.processedData.comments.push(commentData);
          this.stats.commentsCount++;
        }
      }
      
      processedRows++;
      
      // Показываем прогресс каждые 100 строк
      if (processedRows % 100 === 0) {
        console.log(`📊 Обработано ${processedRows} строк...`);
      }
    }
    
    console.log(`✅ Обработано ${processedRows} строк данных`);
  }

  /**
   * Определение типа записи
   */
  determineRecordType(row) {
    const postTypeCol = this.columnMapping.POST_TYPE;
    if (postTypeCol !== undefined) {
      const postType = this.cleanText(row[postTypeCol]);
      
      if (CONFIG.POST_TYPES.REVIEW.some(type => postType.includes(type))) {
        return 'review';
      }
      if (CONFIG.POST_TYPES.COMMENT.some(type => postType.includes(type))) {
        return 'comment';
      }
    }
    
    // Fallback: анализ по тексту
    const textCol = this.columnMapping.TEXT;
    if (textCol !== undefined) {
      const text = this.cleanText(row[textCol]);
      if (text.length > 50) { // Длинный текст - скорее всего отзыв
        return 'review';
      }
    }
    
    return 'comment'; // По умолчанию
  }

  /**
   * Извлечение данных отзыва
   */
  extractReviewData(row, index) {
    try {
      const platform = this.getColumnValue(row, 'PLATFORM');
      const text = this.getColumnValue(row, 'TEXT');
      const date = this.getColumnValue(row, 'DATE');
      const author = this.getColumnValue(row, 'AUTHOR');
      const views = this.getColumnValue(row, 'VIEWS');
      const engagement = this.getColumnValue(row, 'ENGAGEMENT');
      
      if (!text || text.length < 10) {
        return null; // Пропускаем записи без значимого текста
      }
      
      return {
        площадка: platform,
        тема: this.extractTheme(text),
        текст: text,
        дата: this.formatDate(date),
        ник: author,
        просмотры: this.parseViews(views),
        вовлечение: engagement,
        типПоста: 'Отзыв',
        section: 'reviews',
        originalRow: row,
        rowIndex: index + 1
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения отзыва из строки ${index + 1}:`, error);
      return null;
    }
  }

  /**
   * Извлечение данных комментария
   */
  extractCommentData(row, index) {
    try {
      const platform = this.getColumnValue(row, 'PLATFORM');
      const text = this.getColumnValue(row, 'TEXT');
      const date = this.getColumnValue(row, 'DATE');
      const author = this.getColumnValue(row, 'AUTHOR');
      const views = this.getColumnValue(row, 'VIEWS');
      const engagement = this.getColumnValue(row, 'ENGAGEMENT');
      
      if (!text || text.length < 5) {
        return null;
      }
      
      return {
        площадка: platform,
        тема: this.extractTheme(text),
        текст: text,
        дата: this.formatDate(date),
        ник: author,
        просмотры: this.parseViews(views),
        вовлечение: engagement,
        типПоста: 'Комментарий',
        section: 'comments',
        originalRow: row,
        rowIndex: index + 1
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения комментария из строки ${index + 1}:`, error);
      return null;
    }
  }

  /**
   * Создание отчета
   */
  createReport() {
    console.log('📊 Создание отчета...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheetName = `Отчет_${this.monthInfo.name}_${new Date().getFullYear()}`;
    
    // Удаляем старый отчет если есть
    const existingSheet = spreadsheet.getSheetByName(reportSheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // Создаем новый лист
    const reportSheet = spreadsheet.insertSheet(reportSheetName);
    
    // Создаем заголовок
    this.createReportHeader(reportSheet);
    
    // Добавляем данные
    this.addReportData(reportSheet);
    
    // Добавляем статистику
    this.addReportStatistics(reportSheet);
    
    // Форматируем
    this.formatReport(reportSheet);
    
    console.log(`✅ Отчет создан: ${reportSheetName}`);
    return spreadsheet.getId();
  }

  /**
   * Создание заголовка отчета
   */
  createReportHeader(sheet) {
    const headerData = [
      ['ЕЖЕМЕСЯЧНЫЙ ОТЧЕТ ПО СОЦИАЛЬНЫМ МЕДИА'],
      [''],
      [`Месяц: ${this.monthInfo.name} ${new Date().getFullYear()}`],
      [`Дата создания: ${new Date().toLocaleDateString('ru-RU')}`],
      [''],
      ['СТАТИСТИКА ОБРАБОТКИ:'],
      [`Всего записей: ${this.stats.totalRows}`],
      [`Отзывов: ${this.stats.reviewsCount}`],
      [`Комментариев: ${this.stats.commentsCount}`],
      [`Общие просмотры: ${this.stats.totalViews.toLocaleString()}`],
      ['']
    ];
    
    sheet.getRange(1, 1, headerData.length, 1).setValues(headerData);
    
    // Форматирование заголовка
    sheet.getRange(1, 1).setFontWeight('bold').setFontSize(16);
    sheet.getRange(3, 1).setFontWeight('bold');
    sheet.getRange(6, 1).setFontWeight('bold');
  }

  /**
   * Добавление данных в отчет
   */
  addReportData(sheet) {
    let currentRow = 12;
    
    // Заголовки таблицы
    const tableHeaders = [
      'Площадка', 'Тема', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Вовлечение', 'Тип'
    ];
    
    sheet.getRange(currentRow, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(currentRow, 1, 1, 1, tableHeaders.length).setFontWeight('bold');
    currentRow++;
    
    // Добавляем отзывы
    if (this.processedData.reviews.length > 0) {
      sheet.getRange(currentRow, 1).setValue('ОТЗЫВЫ:');
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
      
      for (const review of this.processedData.reviews) {
        const rowData = [
          review.площадка,
          review.тема,
          review.текст,
          review.дата,
          review.ник,
          review.просмотры,
          review.вовлечение,
          review.типПоста
        ];
        
        sheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
        currentRow++;
      }
      
      currentRow++; // Пустая строка
    }
    
    // Добавляем комментарии
    if (this.processedData.comments.length > 0) {
      sheet.getRange(currentRow, 1).setValue('КОММЕНТАРИИ:');
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
      
      for (const comment of this.processedData.comments) {
        const rowData = [
          comment.площадка,
          comment.тема,
          comment.текст,
          comment.дата,
          comment.ник,
          comment.просмотры,
          comment.вовлечение,
          comment.типПоста
        ];
        
        sheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
        currentRow++;
      }
    }
  }

  /**
   * Добавление статистики
   */
  addReportStatistics(sheet) {
    const lastRow = sheet.getLastRow();
    const statsRow = lastRow + 3;
    
    const statsData = [
      ['ИТОГО:'],
      [`Всего записей: ${this.stats.totalRows}`],
      [`Отзывов: ${this.stats.reviewsCount}`],
      [`Комментариев: ${this.stats.commentsCount}`],
      [`Общие просмотры: ${this.stats.totalViews.toLocaleString()}`],
      [`Время обработки: ${this.stats.processingTime} сек`],
      [`Качество данных: ${this.stats.qualityScore}%`]
    ];
    
    sheet.getRange(statsRow, 1, statsData.length, 1).setValues(statsData);
    sheet.getRange(statsRow, 1).setFontWeight('bold');
  }

  /**
   * Форматирование отчета
   */
  formatReport(sheet) {
    // Автоподбор ширины колонок
    sheet.autoResizeColumns(1, 8);
    
    // Границы для таблицы
    const dataRange = sheet.getDataRange();
    dataRange.setBorder(true, true, true, true, true, true);
    
    // Чередующиеся цвета строк
    const tableRange = sheet.getRange(13, 1, sheet.getLastRow() - 12, 8);
    tableRange.setAlternatingRowColors(true);
  }

  /**
   * Расчет статистики
   */
  calculateStatistics(startTime) {
    const endTime = new Date();
    this.stats.processingTime = Math.round((endTime - startTime) / 1000);
    
    // Общие просмотры
    this.stats.totalViews = this.processedData.reviews.reduce((sum, r) => sum + (r.просмотры || 0), 0) +
                           this.processedData.comments.reduce((sum, c) => sum + (c.просмотры || 0), 0);
    
    // Качество данных
    const totalProcessed = this.stats.reviewsCount + this.stats.commentsCount;
    this.stats.qualityScore = totalProcessed > 0 ? Math.round((totalProcessed / this.stats.totalRows) * 100) : 0;
    
    // Площадки с данными
    const platforms = new Set();
    [...this.processedData.reviews, ...this.processedData.comments].forEach(item => {
      if (item.площадка) platforms.add(item.площадка);
    });
    this.stats.platformsWithData = platforms.size;
  }

  // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================

  /**
   * Получение значения из колонки
   */
  getColumnValue(row, columnType) {
    const columnIndex = this.columnMapping[columnType];
    if (columnIndex === undefined || columnIndex >= row.length) {
      return '';
    }
    return this.cleanText(row[columnIndex]);
  }

  /**
   * Очистка текста
   */
  cleanText(text) {
    if (!text) return '';
    return String(text).trim();
  }

  /**
   * Проверка пустой строки
   */
  isEmptyRow(row) {
    return !row.some(cell => cell && String(cell).trim().length > 0);
  }

  /**
   * Извлечение темы из текста
   */
  extractTheme(text) {
    if (!text || text.length < 10) return 'Общая тема';
    
    // Простая логика извлечения темы
    const words = text.split(' ').slice(0, 5);
    return words.join(' ') + (text.length > 50 ? '...' : '');
  }

  /**
   * Форматирование даты
   */
  formatDate(dateValue) {
    if (!dateValue) return '';
    
    if (dateValue instanceof Date) {
      return dateValue.toLocaleDateString('ru-RU');
    }
    
    const dateStr = String(dateValue);
    if (dateStr.includes('/') || dateStr.includes('.')) {
      return dateStr;
    }
    
    return dateStr;
  }

  /**
   * Парсинг просмотров
   */
  parseViews(viewsValue) {
    if (!viewsValue) return 0;
    
    const viewsStr = String(viewsValue).replace(/[^\d]/g, '');
    const views = parseInt(viewsStr);
    
    return isNaN(views) ? 0 : views;
  }
}

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Основная функция для запуска обработки
 */
function processMonthlyReport() {
  const processor = new MonthlyReportProcessor();
  
  // Получаем активную таблицу
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = spreadsheet.getId();
  
  console.log('🚀 Запуск обработки ежемесячного отчета...');
  console.log(`📊 Таблица: ${spreadsheet.getName()}`);
  
  const result = processor.processReport(sheetId);
  
  if (result.success) {
    console.log('✅ Обработка завершена успешно!');
    console.log('📊 Статистика:', result.statistics);
    
    // Показываем уведомление
    SpreadsheetApp.getUi().alert(
      'Обработка завершена!',
      `Отчет создан успешно!\n\nСтатистика:\n- Отзывов: ${result.statistics.reviewsCount}\n- Комментариев: ${result.statistics.commentsCount}\n- Общие просмотры: ${result.statistics.totalViews.toLocaleString()}\n- Время обработки: ${result.statistics.processingTime} сек`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    console.error('❌ Ошибка обработки:', result.error);
    
    SpreadsheetApp.getUi().alert(
      'Ошибка обработки',
      `Произошла ошибка при обработке:\n${result.error}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  return result;
}

/**
 * Функция для создания меню
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Обработчик отчетов')
    .addItem('Обработать ежемесячный отчет', 'processMonthlyReport')
    .addSeparator()
    .addItem('Настройки', 'showSettings')
    .addToUi();
}

/**
 * Показать настройки
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(`
    <h3>Настройки обработчика</h3>
    <p>Текущая конфигурация:</p>
    <ul>
      <li>Максимум строк: ${CONFIG.PROCESSING.MAX_ROWS}</li>
      <li>Размер пакета: ${CONFIG.PROCESSING.BATCH_SIZE}</li>
      <li>Таймаут: ${CONFIG.PROCESSING.TIMEOUT_SECONDS} сек</li>
    </ul>
    <p>Для изменения настроек отредактируйте файл скрипта.</p>
  `)
    .setWidth(400)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройки');
}

/**
 * Тестовая функция
 */
function testProcessor() {
  console.log('🧪 Тестирование процессора...');
  
  const processor = new MonthlyReportProcessor();
  console.log('✅ Процессор инициализирован');
  
  // Тест определения месяца
  const testText = 'Отчет за июнь 2025';
  const monthInfo = processor.findMonthInText(testText);
  console.log('📅 Тест определения месяца:', monthInfo);
  
  // Тест очистки текста
  const cleanText = processor.cleanText('  Тестовый текст  ');
  console.log('🧹 Тест очистки текста:', cleanText);
  
  console.log('✅ Тестирование завершено');
} 