/**
 * 🚀 ФИНАЛЬНЫЙ ГИБКИЙ ОБРАБОТЧИК НА ОСНОВЕ АНАЛИЗА БЭКАГЕНТА 1
 * Google Apps Script для автоматической обработки отчетов
 * Версия: 3.0.0 - ОСНОВАНА НА РЕАЛЬНЫХ ДАННЫХ
 * 
 * Автор: AI Assistant + Background Agent bc-851d0563-ea94-47b9-ba36-0f832bafdb25
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ НА ОСНОВЕ АНАЛИЗА ====================

const CONFIG = {
  // Настройки структуры данных (ОСНОВАНЫ НА АНАЛИЗЕ БЭКАГЕНТА 1)
  STRUCTURE: {
    headerRow: 4,        // Заголовки ВСЕГДА в строке 4
    dataStartRow: 5,     // Данные начинаются с строки 5
    infoRows: [1, 2, 3], // Мета-информация в строках 1-3
    maxRows: 10000
  },
  
  // Классификация контента (ОСНОВАНА НА АНАЛИЗЕ БЭКАГЕНТА 1)
  CONTENT_TYPES: {
    REVIEWS: ['ОС', 'Отзывы Сайтов', 'ос', 'отзывы сайтов'],
    TARGETED: ['ЦС', 'Целевые Сайты', 'цс', 'целевые сайты'],
    SOCIAL: ['ПС', 'Площадки Социальные', 'пс', 'площадки социальные']
  },
  
  // Настройки колонок (гибкие)
  COLUMNS: {
    PLATFORM: ['площадка', 'platform', 'site', 'платформа', 'источник'],
    TEXT: ['текст сообщения', 'текст', 'message', 'content', 'сообщение', 'описание'],
    DATE: ['дата', 'date', 'created', 'время', 'дата создания'],
    AUTHOR: ['автор', 'ник', 'author', 'nickname', 'пользователь', 'имя'],
    VIEWS: ['просмотры', 'просмотров получено', 'views', 'просмотров', 'количество просмотров'],
    ENGAGEMENT: ['вовлечение', 'engagement', 'взаимодействие'],
    POST_TYPE: ['тип поста', 'тип размещения', 'post_type', 'тип', 'категория'],
    
    // Дополнительные колонки
    THEME: ['тема', 'theme', 'subject', 'заголовок'],
    LINK: ['ссылка', 'link', 'url', 'адрес'],
    RATING: ['оценка', 'rating', 'рейтинг', 'звезды']
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
 * Финальный класс обработки ежемесячных отчетов
 * ОСНОВАН НА АНАЛИЗЕ БЭКАГЕНТА 1
 */
class FinalMonthlyReportProcessor {
  constructor() {
    this.stats = {
      totalRows: 0,
      reviewsCount: 0,
      targetedCount: 0,
      socialCount: 0,
      totalViews: 0,
      processingTime: 0,
      errors: []
    };
    
    this.columnMapping = {};
    this.monthInfo = null;
    this.contentSections = {
      reviews: [],
      targeted: [],
      social: []
    };
  }

  /**
   * Главный метод обработки отчета
   */
  processReport(spreadsheetId, sheetName = null) {
    const startTime = Date.now();
    
    try {
      console.log('🚀 FINAL PROCESSOR - Начало обработки на основе анализа Бэкагента 1');
      
      // 1. Получение данных
      const sourceData = this.getSourceData(spreadsheetId, sheetName);
      
      // 2. Определение месяца
      this.monthInfo = this.detectMonth(sourceData);
      console.log(`📅 Определен месяц: ${this.monthInfo.name} ${this.monthInfo.year}`);
      
      // 3. Анализ структуры данных (ОСНОВАН НА АНАЛИЗЕ БЭКАГЕНТА 1)
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
   * Получение исходных данных (ОБНОВЛЕНО НА ОСНОВЕ АНАЛИЗА)
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
    console.log(`📋 Структура: заголовки в строке ${CONFIG.STRUCTURE.headerRow}, данные с строки ${CONFIG.STRUCTURE.dataStartRow}`);
    
    return data;
  }

  /**
   * Определение месяца из мета-информации (строки 1-3)
   */
  detectMonth(data) {
    // Поиск в мета-информации (строки 1-3)
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      const monthFromMeta = this.extractMonthFromText(rowText);
      if (monthFromMeta) {
        return { ...monthFromMeta, detectedFrom: 'meta' };
      }
    }
    
    // Поиск в названии листа
    const sheetName = SpreadsheetApp.getActiveSheet().getName();
    const monthFromSheet = this.extractMonthFromText(sheetName);
    if (monthFromSheet) {
      return { ...monthFromSheet, detectedFrom: 'sheet' };
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
   * Анализ структуры данных (ИСПРАВЛЕНО НА ОСНОВЕ РЕАЛЬНЫХ ДАННЫХ)
   */
  analyzeDataStructure(data) {
    if (data.length < CONFIG.STRUCTURE.headerRow) {
      throw new Error('Недостаточно строк для анализа структуры');
    }
    
    // Заголовки в строке 4 (ОСНОВАНО НА АНАЛИЗЕ)
    const headers = data[CONFIG.STRUCTURE.headerRow - 1];
    console.log('📋 Заголовки (строка 4):', headers);
    
    // ИСПРАВЛЕНИЕ: Устанавливаем маппинг на основе реальной структуры данных
    this.columnMapping = {
      platform: 0,    // Колонка A - сайт (например, "otzovik.com")
      link: 1,        // Колонка B - ссылка
      text: 2,        // Колонка C - текст отзыва/комментария
      date: 3,        // Колонка D - дата
      author: 4,      // Колонка E - автор
      // Колонки F-K: "Нет данных"
      views: 11,      // Колонка L - просмотры (если есть)
      postType: headers.length - 1  // Последняя колонка - тип ("ОС", "ЦС")
    };
    
    console.log('🗺️ Маппинг колонок (исправленный):', this.columnMapping);
    
    // Проверяем обязательные колонки
    const requiredColumns = ['platform', 'text', 'date'];
    const missingColumns = requiredColumns.filter(col => this.columnMapping[col] === undefined);
    
    if (missingColumns.length > 0) {
      console.warn(`⚠️ Не найдены колонки: ${missingColumns.join(', ')}`);
    }
  }

  /**
   * Обработка данных (ОБНОВЛЕНО НА ОСНОВЕ АНАЛИЗА)
   */
  processData(data) {
    const processedData = {
      reviews: [],
      targeted: [],
      social: [],
      statistics: {
        totalReviews: 0,
        totalTargeted: 0,
        totalSocial: 0,
        totalViews: 0,
        platforms: new Set()
      }
    };
    
    // Обрабатываем данные начиная с строки 5 (ОСНОВАНО НА АНАЛИЗЕ)
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      
      // Пропускаем пустые строки
      if (this.isEmptyRow(row)) continue;
      
      // Определяем тип контента
      const contentType = this.detectContentType(row);
      
      // Обрабатываем данные
      if (contentType && this.isDataRow(row)) {
        const record = this.processRow(row, contentType);
        
        if (record) {
          if (contentType === 'reviews') {
            processedData.reviews.push(record);
            processedData.statistics.totalReviews++;
          } else if (contentType === 'targeted') {
            processedData.targeted.push(record);
            processedData.statistics.totalTargeted++;
          } else if (contentType === 'social') {
            processedData.social.push(record);
            processedData.statistics.totalSocial++;
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
    this.stats.targetedCount = processedData.statistics.totalTargeted;
    this.stats.socialCount = processedData.statistics.totalSocial;
    this.stats.totalViews = processedData.statistics.totalViews;
    
    console.log(`📊 Обработано: ${processedData.statistics.totalReviews} отзывов, ${processedData.statistics.totalTargeted} целевых, ${processedData.statistics.totalSocial} социальных`);
    
    return processedData;
  }

  /**
   * Определение типа контента (ИСПРАВЛЕНО НА ОСНОВЕ РЕАЛЬНЫХ ДАННЫХ)
   */
  detectContentType(row) {
    if (row.length === 0) return null;
    
    // ИСПРАВЛЕНИЕ: Тип находится в последней колонке строки
    const lastColumnIndex = row.length - 1;
    const postType = String(row[lastColumnIndex] || '').toLowerCase().trim();
    
    console.log(`🔍 Проверка типа в последней колонке (${lastColumnIndex}): "${postType}"`);
    
    // Проверяем тип по последней колонке
    if (CONFIG.CONTENT_TYPES.REVIEWS.some(type => postType.includes(type))) {
      return 'reviews';
    }
    
    if (CONFIG.CONTENT_TYPES.TARGETED.some(type => postType.includes(type))) {
      return 'targeted';
    }
    
    if (CONFIG.CONTENT_TYPES.SOCIAL.some(type => postType.includes(type))) {
      return 'social';
    }
    
    // Альтернативная проверка по тексту (если тип не найден)
    const textIndex = this.columnMapping.text || 2;
    if (textIndex !== undefined && row[textIndex]) {
      const text = String(row[textIndex]).toLowerCase();
      
      // Простые эвристики для определения типа
      if (text.includes('отзыв') || text.includes('review')) {
        return 'reviews';
      }
      
      if (text.includes('целевой') || text.includes('target')) {
        return 'targeted';
      }
      
      if (text.includes('социальн') || text.includes('social')) {
        return 'social';
      }
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
   * Проверка на строку с данными (ИСПРАВЛЕНО)
   */
  isDataRow(row) {
    if (row.length < 3) return false;
    
    // ИСПРАВЛЕНИЕ: Пропускаем строки-заголовки секций
    const firstCell = String(row[0] || '').toLowerCase().trim();
    if (firstCell === 'отзывы' || firstCell === 'комментарии' || firstCell === 'обсуждения') {
      console.log(`⏭️ Пропускаем заголовок секции: "${firstCell}"`);
      return false;
    }
    
    // Проверяем наличие значимого текста в колонке C (индекс 2)
    const text = row[2];
    return text && String(text).trim().length > 10;
  }

  /**
   * Обработка строки данных
   */
  processRow(row, contentType) {
    try {
      const record = {
        contentType: contentType,
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
   * Извлечение платформы (ИСПРАВЛЕНО)
   */
  extractPlatform(row) {
    const index = this.columnMapping.platform;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return 'Неизвестно';
  }

  /**
   * Извлечение текста (ИСПРАВЛЕНО)
   */
  extractText(row) {
    const index = this.columnMapping.text;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение даты (ИСПРАВЛЕНО)
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
   * Извлечение автора (ИСПРАВЛЕНО)
   */
  extractAuthor(row) {
    const index = this.columnMapping.author;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return 'Аноним';
  }

  /**
   * Извлечение просмотров (ИСПРАВЛЕНО)
   */
  extractViews(row) {
    const index = this.columnMapping.views;
    if (index !== undefined && row[index]) {
      const viewsValue = row[index];
      
      if (typeof viewsValue === 'number') {
        return viewsValue;
      }
      
      const parsed = parseInt(String(viewsValue).replace(/\D/g, ''));
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  }

  /**
   * Извлечение типа поста (ИСПРАВЛЕНО)
   */
  extractPostType(row) {
    const index = this.columnMapping.postType;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
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
   * Создание отчета (ОБНОВЛЕНО НА ОСНОВЕ АНАЛИЗА)
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
      ['Отзывы Сайтов (ОС):', this.stats.reviewsCount],
      ['Целевые Сайты (ЦС):', this.stats.targetedCount],
      ['Площадки Социальные (ПС):', this.stats.socialCount],
      ['Общие просмотры:', this.stats.totalViews],
      ['Платформ:', Array.from(processedData.statistics.platforms).length],
      ['Время обработки:', `${(this.stats.processingTime / 1000).toFixed(2)} сек`]
    ];
    
    reportSheet.getRange(4, 1, statsData.length, 2).setValues(statsData);
    
    // Отзывы Сайтов (ОС)
    let currentRow = statsData.length + 6;
    reportSheet.getRange(`A${currentRow}`).setValue('ОТЗЫВЫ САЙТОВ (ОС):');
    reportSheet.getRange(`A${currentRow}`).setFontWeight('bold').setFontSize(12);
    currentRow++;
    
    const headers = ['Платформа', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Тип', 'Тема', 'Ссылка'];
    reportSheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
    reportSheet.getRange(currentRow, 1, 1, headers.length).setFontWeight('bold');
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
      
      reportSheet.getRange(currentRow, 1, reviewData.length, headers.length).setValues(reviewData);
      currentRow += reviewData.length;
    }
    
    // Целевые Сайты (ЦС)
    currentRow += 2;
    reportSheet.getRange(`A${currentRow}`).setValue('ЦЕЛЕВЫЕ САЙТЫ (ЦС):');
    reportSheet.getRange(`A${currentRow}`).setFontWeight('bold').setFontSize(12);
    currentRow++;
    
    reportSheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
    reportSheet.getRange(currentRow, 1, 1, headers.length).setFontWeight('bold');
    currentRow++;
    
    if (processedData.targeted.length > 0) {
      const targetedData = processedData.targeted.map(targeted => [
        targeted.platform,
        targeted.text,
        targeted.date,
        targeted.author,
        targeted.views,
        targeted.postType,
        targeted.theme,
        targeted.link
      ]);
      
      reportSheet.getRange(currentRow, 1, targetedData.length, headers.length).setValues(targetedData);
      currentRow += targetedData.length;
    }
    
    // Площадки Социальные (ПС)
    currentRow += 2;
    reportSheet.getRange(`A${currentRow}`).setValue('ПЛОЩАДКИ СОЦИАЛЬНЫЕ (ПС):');
    reportSheet.getRange(`A${currentRow}`).setFontWeight('bold').setFontSize(12);
    currentRow++;
    
    reportSheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
    reportSheet.getRange(currentRow, 1, 1, headers.length).setFontWeight('bold');
    currentRow++;
    
    if (processedData.social.length > 0) {
      const socialData = processedData.social.map(social => [
        social.platform,
        social.text,
        social.date,
        social.author,
        social.views,
        social.postType,
        social.theme,
        social.link
      ]);
      
      reportSheet.getRange(currentRow, 1, socialData.length, headers.length).setValues(socialData);
      currentRow += socialData.length;
    }
    
    // Итоговая строка (ОБЯЗАТЕЛЬНО!)
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
    reportSheet.autoResizeColumns(1, headers.length);
    
    // Применяем фильтры
    const dataRange = reportSheet.getDataRange();
    reportSheet.getRange(1, 1, dataRange.getNumRows(), dataRange.getNumColumns()).createFilter();
    
    console.log(`📄 Отчет создан: ${reportSheetName}`);
    
    return spreadsheet.getUrl();
  }

  /**
   * Вспомогательные методы
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
  const processor = new FinalMonthlyReportProcessor();
  
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
  ui.createMenu('📊 Обработка отчетов (ФИНАЛЬНАЯ)')
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
    'Настройки процессора (ОСНОВАНЫ НА АНАЛИЗЕ БЭКАГЕНТА 1)',
    `Конфигурация:\n` +
    `- Заголовки в строке: ${CONFIG.STRUCTURE.headerRow}\n` +
    `- Данные с строки: ${CONFIG.STRUCTURE.dataStartRow}\n` +
    `- Мета-информация: строки ${CONFIG.STRUCTURE.infoRows.join(', ')}\n` +
    `- Максимум строк: ${CONFIG.STRUCTURE.maxRows}\n` +
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