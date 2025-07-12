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
    
    console.log(`🗺️ Маппинг колонок:`, this.columnMapping);
    
    // ИСПРАВЛЕНИЕ: Ищем общие просмотры в исходных данных
    let totalViewsFromSource = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // Ищем строку с общими просмотрами
      if (firstCell.includes('суммарное количество просмотров')) {
        // Ищем число в этой строке
        for (let j = 1; j < row.length; j++) {
          const cellValue = String(row[j] || '').replace(/[^\d]/g, '');
          if (cellValue && cellValue.length > 3) {
            totalViewsFromSource = parseInt(cellValue);
            console.log(`📊 Найдены общие просмотры в исходных данных: ${totalViewsFromSource}`);
            break;
          }
        }
        break;
      }
    }
    
    // Сохраняем найденные общие просмотры
    this.stats.totalViewsFromSource = totalViewsFromSource;
    
    return {
      columnMapping: this.columnMapping,
      totalViewsFromSource: totalViewsFromSource
    };
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
    const sections = this.findSectionBoundaries(data);
    console.log('📂 Найденные разделы:', sections);
    
    // Обрабатываем каждый раздел отдельно
    for (const section of sections) {
      const currentSection = section.type;
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
    
    // ИСПРАВЛЕНИЕ: Используем общие просмотры из исходных данных, если они есть
    if (this.stats.totalViewsFromSource && this.stats.totalViewsFromSource > 0) {
      this.stats.totalViews = this.stats.totalViewsFromSource;
      console.log(`📊 Общие просмотры: ${this.stats.totalViews} (из исходных данных)`);
    } else {
      this.stats.totalViews = processedData.statistics.totalViews;
      console.log(`📊 Общие просмотры: ${this.stats.totalViews} (подсчитано)`);
    }
    
    console.log(`📊 Обработано: ${processedRows} строк данных, пропущено: ${skippedRows} строк`);
    console.log(`📊 Результат: ${processedData.statistics.totalReviews} отзывов, ${processedData.statistics.totalCommentsTop20} топ-20, ${processedData.statistics.totalActiveDiscussions} обсуждений`);
    
    return processedData;
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
    // ИСПРАВЛЕНИЕ: Используем theme колонку для ссылок
    const index = columnMapping.theme;
    if (index !== undefined && row[index]) {
      return String(row[index]).trim();
    }
    return '';
  }

  /**
   * Создание отчета (ИСПРАВЛЕНО С ПРАВИЛЬНЫМ ФОРМАТИРОВАНИЕМ)
   */
  createReport(processedData) {
    try {
      // Создаем временную таблицу и лист с нужным именем
      const tempSpreadsheet = SpreadsheetApp.create(`temp_google_sheets_${Date.now()}_${this.monthInfo.name}_${this.monthInfo.year}_результат`);
      const reportSheetName = `${this.monthInfo.name}_${this.monthInfo.year}`;
      const sheet = tempSpreadsheet.getActiveSheet();
      sheet.setName(reportSheetName);
      
      // 1. Шапка: Продукт, Период, План (КАК В ЭТАЛОНЕ)
      sheet.getRange('A1').setValue('Продукт');
      sheet.getRange('B1').setValue('Акрихин - Фортедетрим');
      sheet.getRange('C1').setValue(''); // Пустая колонка как в эталоне
      
      sheet.getRange('A2').setValue('Период'); 
      sheet.getRange('B2').setValue(`${this.monthInfo.name}-25`);
      sheet.getRange('C2').setValue(''); // Пустая колонка как в эталоне
      
      sheet.getRange('A3').setValue('План');
      sheet.getRange('B3').setValue(''); // Пустая как в эталоне
      sheet.getRange('C3').setValue(''); // Пустая колонка как в эталоне
      
      // 2. Пустая строка 4
      let row = 5;
      
      // 3. Заголовки таблицы (ТОЧНО КАК В ЭТАЛОНЕ)
      const tableHeaders = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
      sheet.getRange(row, 1, 1, tableHeaders.length).setValues([tableHeaders]);
      
      // ✅ ИСПРАВЛЕНИЕ: Применяем правильные цвета как в эталоне
      sheet.getRange(row, 1, 1, tableHeaders.length)
        .setFontWeight('bold')
        .setBackground('#9fc5e8') // Голубой как в эталоне
        .setFontColor('black')
        .setBorder(true, true, true, true, true, true);
      row++;
      
      // 4. Разделы и данные с правильным форматированием
      const writeSection = (sectionName, dataArr) => {
        // Заголовок раздела с фиолетовым фоном как в эталоне
        sheet.getRange(row, 1).setValue(sectionName);
        sheet.getRange(row, 1, 1, tableHeaders.length)
          .setBackground('#b4a7d6') // Фиолетовый как в эталоне
          .setFontWeight('bold')
          .setFontColor('black')
          .setBorder(true, true, true, true, true, true);
        row++;
        
        if (dataArr.length) {
          // ИСПРАВЛЕНИЕ: Правильное отображение данных
          const safeData = dataArr.map(r => {
            const arr = [
              r.platform || '',     // Площадка
              r.theme || '',        // Тема (ссылка)
              r.text || '',         // Текст сообщения
              r.date || '',         // Дата
              r.author || '',       // Ник
              r.views || 0,         // Просмотры
              r.engagement || '',   // Вовлечение
              r.type || ''          // Тип поста
            ];
            // Обрезаем до нужного количества колонок
            return arr.slice(0, tableHeaders.length);
          });
          
          const dataRange = sheet.getRange(row, 1, safeData.length, tableHeaders.length);
          dataRange.setValues(safeData);
          
          // Применяем границы к данным
          dataRange.setBorder(true, true, true, true, true, true);
          
          row += safeData.length;
        }
        console.log(`📂 Раздел "${sectionName}": ${dataArr.length} строк`);
      };
      
      writeSection('Отзывы', processedData.reviews);
      writeSection('Комментарии Топ-20 выдачи', processedData.commentsTop20);
      writeSection('Активные обсуждения (мониторинг)', processedData.activeDiscussions);
      
      // 5. Блок статистики внизу (КАК В ЭТАЛОНЕ)
      row += 2;
      
      // Блок статистики с бежевым фоном как в эталоне
      const statsStartRow = row;
      
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
      
      // Применяем форматирование к блоку статистики
      const statsRange = sheet.getRange(statsStartRow, 1, 4, 2);
      statsRange.setFontWeight('bold')
        .setBackground('#f4cccc') // Бежевый/розоватый как в эталоне
        .setBorder(true, true, true, true, true, true);
      
      // 6. Настройка размеров колонок КАК В ЭТАЛОНЕ
      sheet.setColumnWidth(1, 120);  // Площадка
      sheet.setColumnWidth(2, 300);  // Тема (широкая для ссылок)
      sheet.setColumnWidth(3, 400);  // Текст сообщения (самая широкая)
      sheet.setColumnWidth(4, 100);  // Дата
      sheet.setColumnWidth(5, 120);  // Ник
      sheet.setColumnWidth(6, 100);  // Просмотры
      sheet.setColumnWidth(7, 100);  // Вовлечение
      sheet.setColumnWidth(8, 80);   // Тип поста
      
      // 7. Применяем фильтры
      const dataRange = sheet.getDataRange();
      sheet.getRange(5, 1, dataRange.getNumRows() - 4, tableHeaders.length).createFilter();
      
      console.log(`📄 Отчет создан: ${reportSheetName}`);
      return tempSpreadsheet.getUrl();
    } catch (error) {
      console.error('❌ Ошибка при создании файла:', error);
      throw error;
    }
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
      postType: 13
      // ❌ ИСПРАВЛЕНИЕ: Убираем дублирующуюся колонку link: 3
    };
  }

  /**
   * ИСПРАВЛЕНИЕ: Поиск границ разделов (КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ)
   */
  findSectionBoundaries(data) {
    const sections = [];
    let foundSections = new Set(); // Отслеживаем уже найденные разделы
    
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // ✅ КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Пропускаем строки статистики
      if (this.isStatisticsRow(row)) {
        continue;
      }
      
      // Определяем тип раздела
      let sectionType = null;
      let sectionName = '';
      
      // ✅ КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Очень строгие условия для определения заголовков разделов
      if (firstCell === 'отзывы' && !foundSections.has('reviews')) {
        sectionType = 'reviews';
        sectionName = 'Отзывы';
        foundSections.add('reviews');
      } else if ((firstCell.includes('комментарии топ-20') || firstCell.includes('топ-20 выдачи') || firstCell === 'комментарии топ-20 выдачи') && !foundSections.has('commentsTop20')) {
        sectionType = 'commentsTop20';
        sectionName = 'Комментарии Топ-20';
        foundSections.add('commentsTop20');
      } else if ((firstCell.includes('активные обсуждения') || firstCell.includes('мониторинг') || firstCell === 'активные обсуждения (мониторинг)') && !foundSections.has('activeDiscussions')) {
        sectionType = 'activeDiscussions';
        sectionName = 'Активные обсуждения';
        foundSections.add('activeDiscussions');
      }
      
      // Если найден новый раздел
      if (sectionType) {
        console.log(`📂 Найден заголовок раздела "${sectionName}" в строке ${i + 1}`);
        
        // Ищем конец раздела
        let endRow = data.length - 1;
        
        // Ищем следующий заголовок раздела или начало статистики
        for (let j = i + 1; j < data.length; j++) {
          const nextRow = data[j];
          const nextFirstCell = String(nextRow[0] || '').toLowerCase().trim();
          
          // Если нашли следующий заголовок раздела или статистику
          if (nextFirstCell === 'отзывы' || 
              nextFirstCell.includes('комментарии топ-20') || 
              nextFirstCell.includes('топ-20 выдачи') ||
              nextFirstCell.includes('активные обсуждения') || 
              nextFirstCell.includes('мониторинг') ||
              this.isStatisticsRow(nextRow)) {
            endRow = j - 1;
            break;
          }
        }
        
        // Ищем последнюю строку данных в пределах раздела
        for (let j = endRow; j > i; j--) {
          const checkRow = data[j];
          if (!this.isStatisticsRow(checkRow) && !this.isEmptyRow(checkRow)) {
            endRow = j;
            break;
          }
        }
        
        sections.push({
          type: sectionType,
          name: sectionName,
          startRow: i + 1, // Начинаем после заголовка
          endRow: endRow
        });
        
        console.log(`📊 Раздел "${sectionName}": строки ${i + 2}-${endRow + 1} (${Math.max(0, endRow - i)} записей)`);
      }
    }
    
    return sections;
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

  /**
   * Проверка на пустую строку
   */
  isEmptyRow(row) {
    return !row || row.every(cell => !cell || String(cell).trim() === '');
  }

  /**
   * Получение названия месяца
   */
  getMonthName(monthIndex) {
    const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                   'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    return months[monthIndex] || 'Январь';
  }

  /**
   * Получение короткого названия месяца
   */
  getMonthShort(monthIndex) {
    const months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                   'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
    return months[monthIndex] || 'Янв';
  }
}

// =============================================================================
// ГЛАВНЫЕ ФУНКЦИИ ДЛЯ ЗАПУСКА
// =============================================================================

/**
 * Основная функция для запуска из Google Apps Script
 */
function processGoogleSheets(spreadsheetId = null, sheetName = null) {
  try {
    const processor = new FinalMonthlyReportProcessor();
    
    // Если ID не передан, используем текущую таблицу
    if (!spreadsheetId) {
      spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    }
    
    return processor.processReport(spreadsheetId, sheetName);
  } catch (error) {
    console.error('❌ Ошибка в главной функции:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Функция для тестирования
 */
function testProcessor() {
  console.log('🧪 Запуск тестирования...');
  
  try {
    const result = processGoogleSheets();
    
    if (result.success) {
      console.log('✅ Тестирование успешно завершено');
      console.log(`📊 Статистика:`, result.statistics);
      console.log(`🔗 Отчет: ${result.reportUrl}`);
    } else {
      console.log('❌ Тестирование завершилось с ошибкой:', result.error);
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Функция для анализа данных
 */
function analyzeDataStructure() {
  console.log('🔍 Анализ структуры данных...');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length === 0) {
      console.log('❌ Нет данных для анализа');
      return;
    }
    
    // Анализируем первые 10 строк
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const row = data[i];
      console.log(`Строка ${i + 1}:`, row.slice(0, 5).map(cell => 
        (cell || '').toString().substring(0, 30)
      ));
    }
    
    return {
      totalRows: data.length,
      columns: data[0] ? data[0].length : 0
    };
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
    return { error: error.message };
  }
}

/**
 * Основная функция для запуска
 */
function main() {
  return processGoogleSheets();
}

/**
 * Функция для тестирования
 */
function runTest() {
  return testProcessor();
}

/**
 * Функция для анализа данных
 */
function runAnalysis() {
  return analyzeDataStructure();
}