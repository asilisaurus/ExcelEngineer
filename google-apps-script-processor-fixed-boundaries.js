/**
 * 🚀 ИСПРАВЛЕННЫЙ ПРОЦЕССОР С ПРАВИЛЬНЫМИ ГРАНИЦАМИ РАЗДЕЛОВ
 * Google Apps Script для автоматической обработки отчетов
 * Версия: 4.0.0 - КРИТИЧЕСКИЕ ИСПРАВЛЕНИЯ ГРАНИЦ
 * 
 * Автор: AI Assistant + Background Agent bc-2954e872-79f8-4d41-b422-413e62f0b031
 * Дата: 2025
 * 
 * ИСПРАВЛЕНИЯ:
 * - ✅ findSectionBoundaries: правильные границы разделов (sectionStart = i + 1)
 * - ✅ processData: улучшенная обработка по границам
 * - ✅ isDataRow: исправлена ошибка с columnMapping
 * - ✅ determinePostTypeBySection: улучшенная логика типов
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
 * Исправленный класс обработки ежемесячных отчетов
 * ИСПРАВЛЕНИЯ: правильные границы разделов
 */
class FixedMonthlyReportProcessor {
  constructor() {
    this.stats = {
      totalRows: 0,
      reviewsCount: 0,
      targetedCount: 0,
      socialCount: 0,
      commentsTop20Count: 0,
      activeDiscussionsCount: 0,
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
      console.log('🚀 FIXED PROCESSOR - Начало обработки с исправленными границами');
      
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
   * Анализ структуры данных
   */
  analyzeDataStructure(data) {
    console.log('🗺️ Анализ структуры данных...');
    
    // Анализируем заголовки в строке 4
    if (data.length > CONFIG.STRUCTURE.headerRow - 1) {
      const headers = data[CONFIG.STRUCTURE.headerRow - 1];
      console.log('📋 Заголовки:', headers.slice(0, 10).join(', '));
    }
    
    // Определяем маппинг колонок
    this.columnMapping = this.getColumnMapping();
    console.log('🗺️ Маппинг колонок (исправленный):', this.columnMapping);
  }

  /**
   * 🔧 ИСПРАВЛЕННЫЙ processData() - с правильными границами разделов
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
    
    // Получаем фиксированный маппинг
    const columnMapping = this.getColumnMapping();
    
    // 🔧 ИСПРАВЛЕНИЕ: Получаем правильные границы разделов
    const sections = this.findSectionBoundariesFixed(data);
    console.log('📂 Найденные разделы (исправленные):', sections);
    
    // Обрабатываем каждый раздел отдельно
    for (const section of sections) {
      console.log(`🔄 Обработка раздела "${section.name}" (строки ${section.startRow + 1}-${section.endRow + 1})`);
      
      let sectionRows = 0;
      
      // Обрабатываем строки В ПРЕДЕЛАХ раздела
      for (let i = section.startRow; i <= section.endRow && i < data.length; i++) {
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
        
        // 🔧 ИСПРАВЛЕНИЕ: Проверяем, что строка содержит данные
        if (!this.isDataRowFixed(row, columnMapping)) {
          skippedRows++;
          continue;
        }
        
        // Обрабатываем строку данных
        const record = this.processRow(row, section.type, columnMapping);
        if (record) {
          // 🔧 ИСПРАВЛЕНИЕ: Правильное распределение по разделам
          if (section.type === 'reviews') {
            processedData.reviews.push(record);
            processedData.statistics.totalReviews++;
          } else if (section.type === 'commentsTop20') {
            processedData.commentsTop20.push(record);
            processedData.statistics.totalCommentsTop20++;
          } else if (section.type === 'activeDiscussions') {
            processedData.activeDiscussions.push(record);
            processedData.statistics.totalActiveDiscussions++;
          }
          
          processedData.statistics.totalViews += record.views || 0;
          if (record.platform) {
            processedData.statistics.platforms.add(record.platform);
          }
          processedRows++;
          sectionRows++;
        } else {
          skippedRows++;
        }
      }
      
      console.log(`📊 Раздел "${section.name}": ${sectionRows} строк данных`);
    }
    
    // Обновляем глобальную статистику
    this.stats.reviewsCount = processedData.statistics.totalReviews;
    this.stats.commentsTop20Count = processedData.statistics.totalCommentsTop20;
    this.stats.activeDiscussionsCount = processedData.statistics.totalActiveDiscussions;
    this.stats.totalViews = processedData.statistics.totalViews;
    
    console.log(`📊 ИТОГО обработано: ${processedRows} строк данных, пропущено: ${skippedRows} строк`);
    console.log(`📊 РЕЗУЛЬТАТ: ${processedData.statistics.totalReviews} отзывов, ${processedData.statistics.totalCommentsTop20} топ-20, ${processedData.statistics.totalActiveDiscussions} обсуждений`);
    
    return processedData;
  }

  /**
   * 🔧 КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: findSectionBoundariesFixed()
   * Правильное определение границ разделов
   */
  findSectionBoundariesFixed(data) {
    const sections = [];
    let currentSection = null;
    let sectionStart = -1;
    
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // Определяем тип раздела
      let sectionType = null;
      let sectionName = '';
      
      if (firstCell.includes('отзывы') && !firstCell.includes('топ-20') && !firstCell.includes('обсуждения')) {
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
          sections.push({
            type: currentSection,
            name: this.getSectionName(currentSection),
            startRow: sectionStart,
            endRow: i - 1
          });
        }
        
        // 🔧 КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: данные начинаются со СЛЕДУЮЩЕЙ строки после заголовка
        currentSection = sectionType;
        sectionStart = i + 1; // ✅ БЫЛО: i, СТАЛО: i + 1
        console.log(`📂 Найден раздел "${sectionName}" в строке ${i + 1}, данные с строки ${sectionStart + 1}`);
      }
    }
    
    // Закрываем последний раздел
    if (currentSection && sectionStart !== -1) {
      // Ищем последнюю строку данных (исключаем статистику)
      let endRow = data.length - 1;
      for (let i = data.length - 1; i >= sectionStart; i--) {
        if (!this.isStatisticsRow(data[i]) && !this.isEmptyRow(data[i])) {
          endRow = i;
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
   * 🔧 ИСПРАВЛЕНИЕ: isDataRowFixed() - с правильным columnMapping
   */
  isDataRowFixed(row, columnMapping) {
    if (!row || row.length < 5) return false; // Минимум 5 колонок для данных
    
    // Пропускаем строки-заголовки секций
    const firstCell = String(row[0] || '').toLowerCase().trim();
    if (firstCell === 'отзывы' || firstCell.includes('комментарии') || firstCell.includes('обсуждения')) {
      return false;
    }
    
    // Проверяем наличие значимого текста в колонке "Текст сообщения" (индекс 4)
    const textIndex = columnMapping.text;
    const text = row[textIndex];
    
    if (text && String(text).trim().length > 10) {
      return true;
    }
    
    // Альтернативная проверка: наличие платформы и даты
    const platformIndex = columnMapping.platform;
    const dateIndex = columnMapping.date;
    
    const hasPlatform = row[platformIndex] && String(row[platformIndex]).trim().length > 0;
    const hasDate = row[dateIndex] && String(row[dateIndex]).trim().length > 0;
    
    return hasPlatform && hasDate;
  }

  /**
   * Обработка строки данных
   */
  processRow(row, currentSection, columnMapping) {
    try {
      // Проверяем, что строка содержит данные
      if (!row || row.length === 0) {
        return null;
      }

      // Извлекаем основные данные
      const platform = this.extractPlatform(row, columnMapping);
      const text = this.extractText(row, columnMapping);
      const date = this.extractDate(row, columnMapping);
      const author = this.extractAuthor(row, columnMapping);
      const views = this.extractViews(row, columnMapping);
      const engagement = this.extractEngagement(row, columnMapping);
      const theme = this.extractTheme(row, columnMapping);
      const link = this.extractLink(row, columnMapping);
      
      // 🔧 ИСПРАВЛЕНИЕ: Правильное определение типа по разделу
      const type = this.determinePostTypeBySectionFixed(currentSection, row, text, columnMapping);
      
      // Минимальная проверка данных
      if (!platform && !text && !date) {
        return null;
      }

      return {
        platform: platform,
        theme: theme,
        text: text,
        date: date,
        author: author,
        views: views,
        engagement: engagement,
        type: type,
        link: link,
        section: currentSection
      };
      
    } catch (error) {
      console.error('❌ Ошибка обработки строки:', error);
      return null;
    }
  }

  /**
   * 🔧 ИСПРАВЛЕНИЕ: Правильное определение типа по разделу
   */
  determinePostTypeBySectionFixed(currentSection, row, text, columnMapping) {
    // Определяем тип СТРОГО на основе раздела
    if (currentSection === 'reviews') {
      return 'ОС'; // Отзывы сайтов
    } else if (currentSection === 'commentsTop20') {
      return 'ЦС'; // Целевые сайты
    } else if (currentSection === 'activeDiscussions') {
      return 'ПС'; // Площадки социальные
    }
    
    // Если раздел неопределен, используем альтернативную логику
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

    // По умолчанию
    return 'ОС';
  }

  /**
   * Проверка на пустую строку
   */
  isEmptyRow(row) {
    return !row || row.every(cell => !cell || String(cell).trim() === '');
  }

  /**
   * Извлечение платформы
   */
  extractPlatform(row, columnMapping) {
    const index = columnMapping.platform;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение текста
   */
  extractText(row, columnMapping) {
    const index = columnMapping.text;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение даты
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
   * Извлечение автора
   */
  extractAuthor(row, columnMapping) {
    const index = columnMapping.author;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Извлечение просмотров
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
   * Извлечение вовлечения
   */
  extractEngagement(row, columnMapping) {
    const index = columnMapping.engagement;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
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
   * Создание отчета
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
    let currentRow = 5;
    
    // Записываем заголовки
    for (let col = 0; col < tableHeaders.length; col++) {
      sheet.getRange(currentRow, col + 1).setValue(tableHeaders[col]);
    }
    currentRow++;
    
    // Функция записи данных раздела
    function writeSection(sectionName, dataArr) {
      if (dataArr.length === 0) return;
      
      // Заголовок раздела
      sheet.getRange(currentRow, 1).setValue(sectionName);
      currentRow++;
      
      // Данные раздела
      dataArr.forEach(record => {
        const rowData = [
          record.platform || '',
          record.theme || '',
          record.text || '',
          record.date || '',
          record.author || '',
          record.views || 0,
          record.engagement || '',
          record.type || ''
        ];
        
        for (let col = 0; col < rowData.length; col++) {
          sheet.getRange(currentRow, col + 1).setValue(rowData[col]);
        }
        currentRow++;
      });
    }
    
    // Записываем разделы
    writeSection('Отзывы', processedData.reviews);
    writeSection('Комментарии Топ-20 выдачи', processedData.commentsTop20);
    writeSection('Активные обсуждения (мониторинг)', processedData.activeDiscussions);
    
    // Статистика
    currentRow += 2;
    sheet.getRange(currentRow, 1).setValue('Суммарное количество просмотров');
    sheet.getRange(currentRow, 2).setValue(processedData.statistics.totalViews);
    currentRow++;
    
    sheet.getRange(currentRow, 1).setValue('Количество карточек товара (отзывы)');
    sheet.getRange(currentRow, 2).setValue(processedData.statistics.totalReviews);
    currentRow++;
    
    sheet.getRange(currentRow, 1).setValue('Количество обсуждений');
    sheet.getRange(currentRow, 2).setValue(processedData.statistics.totalActiveDiscussions);
    currentRow++;
    
    console.log(`📄 Отчет создан: ${reportSheetName}`);
    
    return tempSpreadsheet.getUrl();
  }

  /**
   * Извлечение месяца из текста
   */
  extractMonthFromText(text) {
    const lowerText = text.toLowerCase();
    
    const months = [
      { name: 'Январь', variants: ['январь', 'янв', 'january', 'jan'], number: 1 },
      { name: 'Февраль', variants: ['февраль', 'фев', 'february', 'feb'], number: 2 },
      { name: 'Март', variants: ['март', 'мар', 'march', 'mar'], number: 3 },
      { name: 'Апрель', variants: ['апрель', 'апр', 'april', 'apr'], number: 4 },
      { name: 'Май', variants: ['май', 'may'], number: 5 },
      { name: 'Июнь', variants: ['июнь', 'июн', 'june', 'jun'], number: 6 },
      { name: 'Июль', variants: ['июль', 'июл', 'july', 'jul'], number: 7 },
      { name: 'Август', variants: ['август', 'авг', 'august', 'aug'], number: 8 },
      { name: 'Сентябрь', variants: ['сентябрь', 'сен', 'september', 'sep'], number: 9 },
      { name: 'Октябрь', variants: ['октябрь', 'окт', 'october', 'oct'], number: 10 },
      { name: 'Ноябрь', variants: ['ноябрь', 'ноя', 'november', 'nov'], number: 11 },
      { name: 'Декабрь', variants: ['декабрь', 'дек', 'december', 'dec'], number: 12 }
    ];
    
    // Ищем год
    const yearMatch = lowerText.match(/20(2[0-9]|[0-1][0-9])/);
    const year = yearMatch ? parseInt('20' + yearMatch[1]) : new Date().getFullYear();
    
    // Ищем месяц
    for (const month of months) {
      for (const variant of month.variants) {
        if (lowerText.includes(variant)) {
          return {
            name: month.name,
            short: this.getMonthShort(month.number - 1),
            number: month.number,
            year: year
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Получение названия месяца
   */
  getMonthName(monthIndex) {
    const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                   'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    return months[monthIndex];
  }

  /**
   * Получение короткого названия месяца
   */
  getMonthShort(monthIndex) {
    const months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                   'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
    return months[monthIndex];
  }

  /**
   * Восстановленный фиксированный маппинг колонок для всех месяцев
   */
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
function processMonthlyReportFixed(spreadsheetId = null, sheetName = null) {
  const processor = new FixedMonthlyReportProcessor();
  
  if (!spreadsheetId) {
    spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }
  
  return processor.processReport(spreadsheetId, sheetName);
}

/**
 * Быстрая обработка текущего листа
 */
function quickProcessFixed() {
  return processMonthlyReportFixed();
}

/**
 * Обработка с выбором листа
 */
function processWithSheetSelectionFixed() {
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
    return processMonthlyReportFixed(spreadsheet.getId(), selectedSheet);
  }
}

/**
 * Обновление меню
 */
function updateMenuFixed() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Обработка отчетов (ИСПРАВЛЕННАЯ)')
    .addItem('Быстрая обработка (исправленная)', 'quickProcessFixed')
    .addItem('Обработка с выбором листа (исправленная)', 'processWithSheetSelectionFixed')
    .addSeparator()
    .addItem('Настройки', 'showSettingsFixed')
    .addToUi();
}

/**
 * Показать настройки
 */
function showSettingsFixed() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'ИСПРАВЛЕННЫЙ процессор (v4.0.0)',
    `КРИТИЧЕСКИЕ ИСПРАВЛЕНИЯ:\n` +
    `✅ Правильные границы разделов (sectionStart = i + 1)\n` +
    `✅ Исправлена обработка по границам\n` +
    `✅ Правильное определение типов по разделам\n` +
    `✅ Исправлена ошибка с columnMapping\n\n` +
    `Конфигурация:\n` +
    `- Заголовки в строке: ${CONFIG.STRUCTURE.headerRow}\n` +
    `- Данные с строки: ${CONFIG.STRUCTURE.dataStartRow}\n` +
    `- Отладка: ${CONFIG.TESTING.ENABLE_DEBUG ? 'Включена' : 'Выключена'}`,
    ui.ButtonSet.OK
  );
}

/**
 * Инициализация при открытии
 */
function onOpen() {
  updateMenuFixed();
} 