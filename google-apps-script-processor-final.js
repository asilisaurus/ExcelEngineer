/**
 * 🚀 ФИНАЛЬНЫЙ ГИБКИЙ ОБРАБОТЧИК НА ОСНОВЕ АНАЛИЗА БЭКАГЕНТА 1
 * Google Apps Script для автоматической обработки отчетов
 * Версия: 3.1.0 - ЭТАЛОННЫЕ ЛИСТЫ В ТОЙ ЖЕ ТАБЛИЦЕ
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
 * УНИВЕРСАЛЬНЫЙ: работает только с исходниками, не зависит от эталонов
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
      
      // 2. Определение месяца (sheetName приоритетно)
      this.monthInfo = this.detectMonth(sourceData, sheetName);
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
  detectMonth(data, sheetName = null) {
    // 1. Пробуем определить месяц из sheetName (если передан)
    if (sheetName) {
      const monthFromSheet = this.extractMonthFromText(sheetName);
      if (monthFromSheet) {
        console.log(`📅 Месяц определен по имени листа (аргумент): ${monthFromSheet.name} ${monthFromSheet.year}`);
        return { ...monthFromSheet, detectedFrom: 'sheet' };
      }
    }
    // 2. Пробуем определить месяц из имени активного листа
    try {
      const activeSheet = SpreadsheetApp.getActiveSheet();
      if (activeSheet) {
        const activeSheetName = activeSheet.getName();
        const monthFromSheet = this.extractMonthFromText(activeSheetName);
        if (monthFromSheet) {
          console.log(`📅 Месяц определен по имени активного листа: ${monthFromSheet.name} ${monthFromSheet.year}`);
          return { ...monthFromSheet, detectedFrom: 'sheet' };
        }
      }
    } catch (error) {
      console.warn('⚠️ Не удалось получить название активного листа:', error.message);
    }
    // 3. Если не найдено — ищем в мета-информации (строки 1-3)
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      const monthFromMeta = this.extractMonthFromText(rowText);
      if (monthFromMeta) {
        console.log(`📅 Месяц определен по мета-информации: ${monthFromMeta.name} ${monthFromMeta.year}`);
        return { ...monthFromMeta, detectedFrom: 'meta' };
      }
    }
    // 4. По умолчанию — текущий месяц
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
    // Восстановленный фиксированный маппинг колонок для всех месяцев
    this.columnMapping = this.getColumnMapping();
    
    console.log('🗺️ Маппинг колонок (исправленный):', this.columnMapping);
    
    // Проверяем обязательные колонки
    const requiredColumns = ['platform', 'text', 'date'];
    const missingColumns = requiredColumns.filter(col => this.columnMapping[col] === undefined);
    
    if (missingColumns.length > 0) {
      console.warn(`⚠️ Не найдены колонки: ${missingColumns.join(', ')}`);
    }
  }

  /**
   * Обработка данных (ИСПРАВЛЕНО НА ОСНОВЕ АНАЛИЗА)
   */
  processData(data) {
    const processedData = {
      reviews: [],
      commentsTop20: [],
      activeDiscussions: [],
      statistics: {
        totalReviews: 0,
        totalCommentsTop20: 0,
        totalActiveDiscussions: 0,
        totalViews: 0,
        platforms: new Set()
      }
    };
    
    let processedRows = 0;
    let skippedRows = 0;
    let debugSkip = 0;
    
    // Получаем фиксированный маппинг
    const columnMapping = this.getColumnMapping();
    
    // ИСПРАВЛЕНИЕ: Правильное определение разделов и их границ
    let currentSection = null;
    let sectionStartRow = -1;
    let sectionEndRow = -1;
    
    // Сначала находим все разделы и их границы
    const sections = this.findSectionBoundaries(data);
    console.log('📂 Найденные разделы:', sections);
    
    // Обрабатываем каждый раздел отдельно
    for (const section of sections) {
      currentSection = section.type;
      console.log(`🔄 Обработка раздела "${section.name}" (строки ${section.startRow + 1}-${section.endRow + 1})`);
      
      // Обрабатываем строки в пределах раздела
      for (let i = section.startRow; i <= section.endRow; i++) {
        const row = data[i];
        
        // Пропускаем заголовки разделов и пустые строки
        if (this.isSectionHeader(row) || this.isEmptyRow(row)) {
          skippedRows++;
          continue;
        }
        
        // Пропускаем строки статистики
        if (this.isStatisticsRow(row)) {
          skippedRows++;
          continue;
        }
        
        // Обрабатываем строку данных
        const record = this.processRow(row, currentSection, columnMapping);
        if (record) {
          if (currentSection === 'reviews') {
            processedData.reviews.push(record);
            processedData.statistics.totalReviews++;
          } else if (currentSection === 'commentsTop20') {
            processedData.commentsTop20.push(record);
            processedData.statistics.totalCommentsTop20++;
          } else if (currentSection === 'activeDiscussions') {
            processedData.activeDiscussions.push(record);
            processedData.statistics.totalActiveDiscussions++;
          }
          processedData.statistics.totalViews += record.views || 0;
          if (record.platform) {
            processedData.statistics.platforms.add(record.platform);
          }
          processedRows++;
        } else {
          if (debugSkip < 10) {
            console.log(`[SKIP] processRow вернул null для строки ${i + 1}:`, row);
            debugSkip++;
          }
          skippedRows++;
        }
      }
    }
    
    // Обновляем глобальную статистику
    this.stats.reviewsCount = processedData.statistics.totalReviews;
    this.stats.commentsTop20Count = processedData.statistics.totalCommentsTop20;
    this.stats.activeDiscussionsCount = processedData.statistics.totalActiveDiscussions;
    this.stats.totalViews = processedData.statistics.totalViews;
    
    console.log(`📊 Обработано: ${processedRows} строк данных, пропущено: ${skippedRows} строк`);
    console.log(`📊 Результат: ${processedData.statistics.totalReviews} отзывов, ${processedData.statistics.totalCommentsTop20} топ-20, ${processedData.statistics.totalActiveDiscussions} обсуждений`);
    
    return processedData;
  }

  /**
   * Определение типа контента (ИСПРАВЛЕНО НА ОСНОВЕ РЕАЛЬНЫХ ДАННЫХ)
   */
  detectContentType(row) {
    if (row.length === 0) return null;
    
    // ИСПРАВЛЕНИЕ: Ищем тип в колонке "Тип поста" (индекс 13)
    const postTypeIndex = 13; // Колонка N - "Тип поста" (индекс 13)
    let postType = '';
    
    if (row.length > postTypeIndex) {
      postType = String(row[postTypeIndex] || '').toLowerCase().trim();
    }
    
    // Если не найден в колонке "Тип поста", проверяем последнюю колонку
    if (!postType) {
      const lastColumnIndex = row.length - 1;
      postType = String(row[lastColumnIndex] || '').toLowerCase().trim();
    }
    
    // Удаляем избыточные логи (оставляем только для ошибок)
    if (postType) {
      if (CONFIG.CONTENT_TYPES.REVIEWS.some(type => postType.includes(type))) {
        return 'reviews';
      }
      if (CONFIG.CONTENT_TYPES.TARGETED.some(type => postType.includes(type))) {
        return 'targeted';
      }
      if (CONFIG.CONTENT_TYPES.SOCIAL.some(type => postType.includes(type))) {
        return 'social';
      }
    }
    
    // Альтернативная проверка по тексту (если тип не найден)
    const columnMapping = this.getColumnMapping();
    const textIndex = columnMapping.text; // Колонка E - "Текст сообщения" (индекс 4)
    if (textIndex !== undefined && row[textIndex]) {
      const text = String(row[textIndex]).toLowerCase();
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
    
    // Можно оставить только для редкой диагностики
    // console.log(`❌ Тип контента не определен для строки`);
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
    if (row.length < 5) return false; // Минимум 5 колонок для данных
    
    // ИСПРАВЛЕНИЕ: Пропускаем строки-заголовки секций
    const firstCell = String(row[0] || '').toLowerCase().trim();
    if (firstCell === 'отзывы' || firstCell === 'комментарии' || firstCell === 'обсуждения') {
      console.log(`⏭️ Пропускаем заголовок секции: "${firstCell}"`);
      return false;
    }
    
    // ИСПРАВЛЕНИЕ: Получаем columnMapping в контексте метода
    const columnMapping = this.getColumnMapping();
    
    // ИСПРАВЛЕНИЕ: Проверяем наличие значимого текста в колонке E (индекс 4) - "Текст сообщения"
    const textIndex = columnMapping.text; // Колонка E - "Текст сообщения" (индекс 4)
    const text = row[textIndex];
    
    if (text && String(text).trim().length > 10) {
      return true;
    }
    
    // Альтернативная проверка: наличие платформы и даты
    const platformIndex = columnMapping.platform; // Колонка B - "Площадка" (индекс 1)
    const dateIndex = columnMapping.date; // Колонка G - "Дата" (индекс 6)
    
    const hasPlatform = row[platformIndex] && String(row[platformIndex]).trim().length > 0;
    const hasDate = row[dateIndex] && String(row[dateIndex]).trim().length > 0;
    
    return hasPlatform && hasDate;
  }

  /**
   * Обработка строки данных (ИСПРАВЛЕНО)
   */
  processRow(row, currentSection, columnMapping) {
    try {
      // Проверяем, что строка содержит данные
      if (!row || row.length === 0) {
        return null;
      }

      // ИСПРАВЛЕНИЕ: Более гибкая проверка наличия данных
      const textIndex = columnMapping.text;
      const platformIndex = columnMapping.platform;
      const dateIndex = columnMapping.date;
      
      const text = row[textIndex] ? String(row[textIndex]).trim() : '';
      const platform = row[platformIndex] ? String(row[platformIndex]).trim() : '';
      const date = row[dateIndex] ? String(row[dateIndex]).trim() : '';
      
      // Проверяем наличие хотя бы одного из: текста, платформы, даты или ссылки
      const hasText = text.length > 5;
      const hasPlatform = platform.length > 0;
      const hasDate = date.length > 0;
      const hasLink = row.some(cell => String(cell).includes('http'));
      
      if (!hasText && !hasPlatform && !hasDate && !hasLink) {
        return null;
      }

      // Извлекаем данные по маппингу
      const extractedPlatform = this.extractPlatform(row, columnMapping);
      const theme = this.extractTheme(row, columnMapping);
      const textContent = this.extractText(row, columnMapping);
      const extractedDate = this.extractDate(row, columnMapping);
      const author = this.extractAuthor(row, columnMapping);
      const views = this.extractViews(row, columnMapping);
      const engagement = this.extractEngagement(row, columnMapping);
      const postType = this.extractPostType(row, columnMapping);
      const link = this.extractLink(row, columnMapping);

      // ИСПРАВЛЕНИЕ: Определяем тип поста на основе раздела и данных
      const type = this.determinePostTypeBySection(row, textContent, postType, currentSection, columnMapping);

      return {
        platform: extractedPlatform,
        theme,
        text: textContent,
        date: extractedDate,
        author,
        views,
        engagement,
        type,
        link,
        section: currentSection
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка обработки строки: [${error}]`);
      return null;
    }
  }

  /**
   * Определение типа поста на основе раздела (ИСПРАВЛЕНО)
   */
  determinePostTypeBySection(row, text, postType, currentSection, columnMapping) {
    // ИСПРАВЛЕНИЕ: Определяем тип на основе раздела
    if (currentSection === 'reviews') {
      return 'ОС'; // Отзывы сайтов
    } else if (currentSection === 'commentsTop20') {
      return 'ЦС'; // Целевые сайты
    } else if (currentSection === 'activeDiscussions') {
      return 'ПС'; // Площадки социальные
    }
    
    // Альтернативная проверка по колонке "Тип поста"
    const postTypeIndex = columnMapping.postType;
    if (postTypeIndex !== undefined && row[postTypeIndex]) {
      const type = String(row[postTypeIndex]).trim().toLowerCase();
      if (type === 'ос' || type === 'о.с.') {
        return 'ОС';
      } else if (type === 'цс' || type === 'ц.с.') {
        return 'ЦС';
      } else if (type === 'пс' || type === 'п.с.') {
        return 'ПС';
      }
    }

    // Альтернативная проверка по тексту
    if (text) {
      const lowerText = text.toLowerCase();
      if (lowerText.includes('отзыв') || lowerText.includes('рекомендую') || lowerText.includes('покупала')) {
        return 'ОС';
      } else if (lowerText.includes('комментарий') || lowerText.includes('ответ') || lowerText.includes('обсуждение')) {
        return 'ЦС';
      } else if (lowerText.includes('социальн') || lowerText.includes('форум') || lowerText.includes('сообщество')) {
        return 'ПС';
      }
    }

    // По умолчанию на основе раздела
    return currentSection === 'reviews' ? 'ОС' : 'ЦС';
  }

  /**
   * Извлечение платформы (ИСПРАВЛЕНО)
   */
  extractPlatform(row, columnMapping) {
    const index = columnMapping.platform;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение текста (ИСПРАВЛЕНО)
   */
  extractText(row, columnMapping) {
    const index = columnMapping.text;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение даты (ИСПРАВЛЕНО)
   */
  extractDate(row, columnMapping) {
    const index = columnMapping.date;
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
  extractAuthor(row, columnMapping) {
    const index = columnMapping.author;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение просмотров (ИСПРАВЛЕНО)
   */
  extractViews(row, columnMapping) {
    const index = columnMapping.views;
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
   * Извлечение вовлечения (ИСПРАВЛЕНО)
   */
  extractEngagement(row, columnMapping) {
    const index = columnMapping.engagement;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение типа поста (ИСПРАВЛЕНО)
   */
  extractPostType(row, columnMapping) {
    const index = columnMapping.postType;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Определение типа поста (ИСПРАВЛЕНО)
   */
  determinePostType(row, text, postType, columnMapping) {
    // Проверяем колонку "Тип поста" (колонка N - индекс 13)
    const postTypeIndex = columnMapping.postType;
    if (postTypeIndex !== undefined && row[postTypeIndex]) {
      const type = String(row[postTypeIndex]).trim().toLowerCase();
      if (type === 'ос' || type === 'о.с.') {
        return 'ОС';
      } else if (type === 'цс' || type === 'ц.с.') {
        return 'ЦС';
      }
    }

    // Альтернативная проверка по тексту (если тип не найден)
    const textIndex = columnMapping.text; // Колонка E - "Текст сообщения" (индекс 4)
    if (textIndex !== undefined && row[textIndex]) {
      const text = String(row[textIndex]).toLowerCase();
      if (text.includes('отзыв') || text.includes('рекомендую') || text.includes('покупала')) {
        return 'ОС';
      } else if (text.includes('комментарий') || text.includes('ответ') || text.includes('обсуждение')) {
        return 'ЦС';
      }
    }

    return 'ОС'; // По умолчанию
  }

  /**
   * Извлечение темы
   */
  extractTheme(row, columnMapping) {
    const index = columnMapping.theme;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение ссылки
   */
  extractLink(row, columnMapping) {
    const index = columnMapping.link;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Создание отчета (ИСПРАВЛЕНО)
   * @param {Object} processedData - обработанные данные
   */
  createReport(processedData) {
    // Создаем временную таблицу и лист с нужным именем
    const tempSpreadsheet = SpreadsheetApp.create(`temp_google_sheets_${Date.now()}_${this.monthInfo.name}_${this.monthInfo.year}_результат`);
    const reportSheetName = `${this.monthInfo.name}_${this.monthInfo.year}`;
    const sheet = tempSpreadsheet.getActiveSheet();
    sheet.setName(reportSheetName);
    
    // Шапка: Продукт, Период, План
    sheet.getRange('A1').setValue('Продукт');
    sheet.getRange('B1').setValue('Акрихин - Фортедетрим');
    sheet.getRange('A2').setValue('Период');
    sheet.getRange('B2').setValue(`${this.monthInfo.name}-25`);
    sheet.getRange('A3').setValue('План');
    
    // 2. Заголовки таблицы
    const tableHeaders = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
    let row = 5;
    sheet.getRange(row, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(row, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#3f2355').setFontColor('white');
    row++;
    
    // 3. Разделы и данные
    function writeSection(sectionName, dataArr) {
      sheet.getRange(row, 1).setValue(sectionName);
      sheet.getRange(row, 1, 1, tableHeaders.length).setBackground('#b7a6c9').setFontWeight('bold');
      row++;
      if (dataArr.length) {
        // ИСПРАВЛЕНИЕ: Правильное отображение типа поста
        const safeData = dataArr.map(r => {
          const arr = [r.platform, r.theme, r.text, r.date, r.author, r.views, r.engagement, r.type];
          while (arr.length < tableHeaders.length) arr.push('');
          return arr.slice(0, tableHeaders.length);
        });
        sheet.getRange(row, 1, safeData.length, tableHeaders.length).setValues(safeData);
        row += safeData.length;
      }
      console.log(`📂 Раздел "${sectionName}": ${dataArr.length} строк`);
    }
    
    writeSection('Отзывы', processedData.reviews);
    writeSection('Комментарии Топ-20 выдачи', processedData.commentsTop20);
    writeSection('Активные обсуждения (мониторинг)', processedData.activeDiscussions);
    
    // 4. Блок статистики внизу
    row += 2;
    sheet.getRange(row, 1).setValue('Суммарное количество просмотров');
    sheet.getRange(row, 2).setValue(this.stats.totalViews);
    row++;
    sheet.getRange(row, 1).setValue('Количество карточек товара (отзывы)');
    sheet.getRange(row, 2).setValue(processedData.reviews.length);
    row++;
    sheet.getRange(row, 1).setValue('Количество обсуждений (форумы, сообщества, комментарии к статьям)');
    sheet.getRange(row, 2).setValue(processedData.activeDiscussions.length);
    row++;
    sheet.getRange(row, 1).setValue('Доля обсуждений с вовлечением в диалог');
    sheet.getRange(row, 2).setValue(this.stats.engagementShare || 0);
    sheet.getRange(row - 3, 1, 4, 2).setFontWeight('bold');
    
    // Оформление
    sheet.autoResizeColumns(1, tableHeaders.length);
    
    // Применяем фильтры
    const dataRange = sheet.getDataRange();
    sheet.getRange(1, 1, dataRange.getNumRows(), dataRange.getNumColumns()).createFilter();
    
    console.log(`📄 Отчет создан: ${reportSheetName}`);
    return tempSpreadsheet.getUrl();
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
    
    // Более точный поиск с приоритетом точных совпадений
    for (const month of months) {
      // Точные совпадения (высший приоритет)
      const exactMatches = [
        month.name.toLowerCase(),
        month.short.toLowerCase(),
        `${month.short}25`,
        `${month.name}25`,
        `${month.short}2025`,
        `${month.name}2025`
      ];
      
      // Проверяем точные совпадения
      for (const exactMatch of exactMatches) {
        if (lowerText === exactMatch || lowerText.includes(exactMatch)) {
          return {
            name: month.name,
            short: month.short,
            number: month.number,
            year: 2025
          };
        }
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

  // Восстановленный фиксированный маппинг колонок для всех месяцев
  getColumnMapping() {
    return {
      platform: 1,
      theme: 3,
      text: 4,
      date: 6,
      author: 7,
      views: 11,
      engagement: 12,
      postType: 13,
      link: 2
    };
  }

  /**
   * ИСПРАВЛЕНИЕ: Поиск границ разделов (УЛУЧШЕННАЯ ВЕРСИЯ)
   */
  findSectionBoundaries(data) {
    const sections = [];
    let currentSection = null;
    let sectionStart = -1;
    
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // ✅ ИСПРАВЛЕНИЕ: Пропускаем строки статистики
      if (this.isStatisticsRow(row)) {
        continue;
      }
      
      // Определяем тип раздела
      let sectionType = null;
      let sectionName = '';
      
      // ✅ ИСПРАВЛЕНИЕ: Более строгие условия для определения заголовков разделов
      if (firstCell === 'отзывы' || (firstCell.includes('отзывы') && !firstCell.includes('топ-20') && !firstCell.includes('обсуждения') && !firstCell.includes('количество'))) {
        sectionType = 'reviews';
        sectionName = 'Отзывы';
      } else if (firstCell.includes('комментарии топ-20') || firstCell.includes('топ-20 выдачи')) {
        sectionType = 'commentsTop20';
        sectionName = 'Комментарии Топ-20';
      } else if (firstCell.includes('активные обсуждения') || firstCell.includes('мониторинг')) {
        sectionType = 'activeDiscussions';
        sectionName = 'Активные обсуждения';
      }
      
      // Если найден новый раздел
      if (sectionType && sectionType !== currentSection) {
        // Закрываем предыдущий раздел
        if (currentSection && sectionStart !== -1) {
          // ✅ ИСПРАВЛЕНИЕ: Улучшенное определение конца раздела
          let endRow = i - 1;
          
          // Ищем последнюю строку данных (исключаем статистику и пустые строки)
          for (let j = i - 1; j >= sectionStart; j--) {
            const checkRow = data[j];
            if (!this.isStatisticsRow(checkRow) && !this.isEmptyRow(checkRow)) {
              endRow = j;
              break;
            }
          }
          
          sections.push({
            type: currentSection,
            name: this.getSectionName(currentSection),
            startRow: sectionStart,
            endRow: endRow
          });
        }
        
        // Начинаем новый раздел
        currentSection = sectionType;
        sectionStart = i + 1; // ✅ ИСПРАВЛЕНО: исключаем заголовок секции из данных
        console.log(`📂 Найден раздел "${sectionName}" в строке ${i + 1}`);
      }
    }
    
    // Закрываем последний раздел
    if (currentSection && sectionStart !== -1) {
      // ✅ ИСПРАВЛЕНИЕ: Улучшенное определение конца последнего раздела
      let endRow = data.length - 1;
      
      // Ищем последнюю строку данных (исключаем статистику)
      for (let j = data.length - 1; j >= sectionStart; j--) {
        const checkRow = data[j];
        if (!this.isStatisticsRow(checkRow) && !this.isEmptyRow(checkRow)) {
          endRow = j;
          break;
        }
      }
      
      sections.push({
        type: currentSection,
        name: this.getSectionName(currentSection),
        startRow: sectionStart,
        endRow: endRow
      });
    }
    
    return sections;
  }

  /**
   * Получение названия раздела
   */
  getSectionName(sectionType) {
    const names = {
      'reviews': 'Отзывы',
      'commentsTop20': 'Комментарии Топ-20',
      'activeDiscussions': 'Активные обсуждения'
    };
    return names[sectionType] || sectionType;
  }

  /**
   * Проверка на заголовок раздела
   */
  isSectionHeader(row) {
    if (!row || row.length === 0) return false;
    
    const firstCell = String(row[0] || '').toLowerCase().trim();
    return firstCell.includes('отзывы') || 
           firstCell.includes('комментарии') || 
           firstCell.includes('обсуждения') ||
           firstCell.includes('топ-20') ||
           firstCell.includes('мониторинг');
  }

  /**
   * Проверка на строку статистики
   */
  isStatisticsRow(row) {
    if (!row || row.length === 0) return false;
    
    const firstCell = String(row[0] || '').toLowerCase().trim();
    return firstCell.includes('суммарное количество просмотров') || 
           firstCell.includes('количество карточек товара') ||
           firstCell.includes('количество обсуждений') ||
           firstCell.includes('доля обсуждений') ||
           firstCell.includes('площадки со статистикой') ||
           firstCell.includes('количество прочтений увеличивается');
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