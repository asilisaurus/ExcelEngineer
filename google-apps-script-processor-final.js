/**
 * 🚀 ИСПРАВЛЕННЫЙ ФИНАЛЬНЫЙ ГИБКИЙ ОБРАБОТЧИК
 * Google Apps Script для автоматической обработки отчетов
 * Версия: 3.2.0 - FIXED
 * 
 * Автор: AI Assistant + Background Agent
 * Дата: 2025
 */

// ==================== КОНФИГУРАЦИЯ НА ОСНОВЕ АНАЛИЗА ====================

const CONFIG = {
  // Настройки структуры данных
  STRUCTURE: {
    headerRow: 4,        // Заголовки ВСЕГДА в строке 4
    dataStartRow: 5,     // Данные начинаются с строки 5
    infoRows: [1, 2, 3], // Мета-информация в строках 1-3
    maxRows: 10000
  },
  
  // Классификация контента
  CONTENT_TYPES: {
    REVIEWS: ['ОС', 'Отзывы Сайтов', 'ос', 'отзывы сайтов'],
    TARGETED: ['ЦС', 'Целевые Сайты', 'цс', 'целевые сайты'],
    SOCIAL: ['ПС', 'Площадки Социальные', 'пс', 'площадки социальные']
  },
  
  // Настройки колонок
  COLUMNS: {
    PLATFORM: ['площадка', 'platform', 'site'],
    TEXT: ['текст сообщения', 'текст', 'message'],
    DATE: ['дата', 'date', 'created'],
    AUTHOR: ['автор', 'ник', 'author'],
    VIEWS: ['просмотры', 'просмотров получено', 'views'],
    ENGAGEMENT: ['вовлечение', 'engagement'],
    POST_TYPE: ['тип поста', 'тип размещения', 'post_type'],
    THEME: ['тема', 'theme', 'subject'],
    LINK: ['ссылка', 'link', 'url']
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

class FinalMonthlyReportProcessor {
  constructor() {
    this.stats = {
      totalRows: 0,
      reviewsCount: 0,
      targetedCount: 0,
      socialCount: 0,
      totalViews: 0,
      processingTime: 0,
      errors: [],
      commentsTop20Count: 0,
      activeDiscussionsCount: 0,
      engagementShare: 0
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
      console.log('🚀 FINAL PROCESSOR - Начало обработки');
      
      // 1. Получение данных
      const sourceData = this.getSourceData(spreadsheetId, sheetName);
      
      // 2. Определение месяца
      this.monthInfo = this.detectMonth(sourceData, sheetName);
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
    console.log(`📋 Структура: заголовки в строке ${CONFIG.STRUCTURE.headerRow}, данные с строки ${CONFIG.STRUCTURE.dataStartRow}`);
    
    return data;
  }

  /**
   * Определение месяца из мета-информации
   */
  detectMonth(data, sheetName = null) {
    // 1. Пробуем определить месяц из sheetName
    if (sheetName) {
      const monthFromSheet = this.extractMonthFromText(sheetName);
      if (monthFromSheet) {
        console.log(`📅 Месяц определен по имени листа: ${monthFromSheet.name} ${monthFromSheet.year}`);
        return { ...monthFromSheet, detectedFrom: 'sheet' };
      }
    }
    
    // 2. Ищем в мета-информации (строки 1-3)
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      const monthFromMeta = this.extractMonthFromText(rowText);
      if (monthFromMeta) {
        console.log(`📅 Месяц определен по мета-информации: ${monthFromMeta.name} ${monthFromMeta.year}`);
        return { ...monthFromMeta, detectedFrom: 'meta' };
      }
    }
    
    // 3. По умолчанию — текущий месяц
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
    this.columnMapping = this.getColumnMapping();
    console.log('🗺️ Маппинг колонок:', this.columnMapping);
  }

  /**
   * Обработка данных
   */
  processData(data) {
    // Сначала извлекаем статистику из исходных данных
    const sourceStats = this.extractStatisticsFromSourceData(data);
    
    const processedData = {
      reviews: [],
      commentsTop20: [],
      activeDiscussions: [],
      statistics: {
        totalReviews: 0,
        totalCommentsTop20: 0,
        totalActiveDiscussions: 0,
        totalViews: sourceStats.totalViews || 0,
        platforms: new Set()
      }
    };
    
    let processedRows = 0;
    let skippedRows = 0;
    
    // Получаем фиксированный маппинг
    const columnMapping = this.getColumnMapping();
    
    // Находим все разделы и их границы
    const sections = this.findSectionBoundaries(data);
    console.log('📂 Найденные разделы:', sections);
    
    // Обрабатываем каждый раздел отдельно
    for (const section of sections) {
      const currentSection = section.type;
      console.log(`🔄 Обработка раздела "${section.name}" (строки ${section.startRow + 1}-${section.endRow + 1})`);
      
      // Пропускаем пустые разделы
      if (section.endRow < section.startRow) {
        console.log(`⏭️ Раздел "${section.name}" пустой, пропускаем`);
        continue;
      }
      
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
          
          // Если общие просмотры не заданы из источника, суммируем из записей
          if (!sourceStats.totalViews && record.views) {
            processedData.statistics.totalViews += record.views;
          }
          
          if (record.platform) {
            processedData.statistics.platforms.add(record.platform);
          }
          processedRows++;
        } else {
          skippedRows++;
        }
      }
    }
    
    // Обновляем глобальную статистику
    this.stats.reviewsCount = processedData.statistics.totalReviews;
    this.stats.commentsTop20Count = processedData.statistics.totalCommentsTop20;
    this.stats.activeDiscussionsCount = processedData.statistics.totalActiveDiscussions;
    this.stats.totalViews = processedData.statistics.totalViews;
    this.stats.engagementShare = sourceStats.engagementShare;
    
    console.log(`📊 Обработано: ${processedRows} строк данных, пропущено: ${skippedRows} строк`);
    console.log(`📊 Результат: ${processedData.statistics.totalReviews} отзывов, ${processedData.statistics.totalCommentsTop20} топ-20, ${processedData.statistics.totalActiveDiscussions} обсуждений`);
    console.log(`📊 Общие просмотры: ${processedData.statistics.totalViews}`);
    
    return processedData;
  }

  /**
   * Поиск границ разделов
   */
  findSectionBoundaries(data) {
    const sections = [];
    
    // Сначала находим все заголовки разделов
    const sectionHeaders = [];
    
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // Пропускаем строки статистики
      if (this.isStatisticsRow(row)) {
        break;
      }
      
      // Определяем тип раздела
      let sectionType = null;
      let sectionName = '';
      
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
      
      if (sectionType) {
        sectionHeaders.push({
          type: sectionType,
          name: sectionName,
          headerRow: i
        });
        console.log(`📂 Найден заголовок раздела "${sectionName}" в строке ${i + 1}`);
      }
    }
    
    // Теперь определяем границы каждого раздела
    for (let i = 0; i < sectionHeaders.length; i++) {
      const currentHeader = sectionHeaders[i];
      const nextHeader = sectionHeaders[i + 1];
      
      const startRow = currentHeader.headerRow + 1;
      let endRow;
      
      if (nextHeader) {
        endRow = nextHeader.headerRow - 1;
      } else {
        endRow = data.length - 1;
        
        // Ищем начало статистики
        for (let j = startRow; j < data.length; j++) {
          if (this.isStatisticsRow(data[j]) || this.isEmptyRow(data[j])) {
            let hasDataAfter = false;
            for (let k = j + 1; k < Math.min(j + 5, data.length); k++) {
              if (!this.isEmptyRow(data[k]) && !this.isStatisticsRow(data[k])) {
                hasDataAfter = true;
                break;
              }
            }
            
            if (!hasDataAfter) {
              endRow = j - 1;
              break;
            }
          }
        }
      }
      
      // Корректируем endRow, если раздел пустой
      if (endRow < startRow) {
        endRow = startRow - 1;
      }
      
      sections.push({
        type: currentHeader.type,
        name: currentHeader.name,
        startRow: startRow,
        endRow: endRow
      });
      
      console.log(`📊 Раздел "${currentHeader.name}": строки ${startRow + 1}-${endRow + 1} (${Math.max(0, endRow - startRow + 1)} записей)`);
    }
    
    return sections;
  }

  /**
   * Обработка строки данных
   */
  processRow(row, currentSection, columnMapping) {
    try {
      if (!row || row.length === 0) {
        return null;
      }

      const textIndex = columnMapping.text;
      const platformIndex = columnMapping.platform;
      
      const text = row[textIndex] ? String(row[textIndex]).trim() : '';
      const platform = row[platformIndex] ? String(row[platformIndex]).trim() : '';
      
      // Пропускаем строки без текста и платформы
      if (!text && !platform) {
        return null;
      }
      
      // Пропускаем заголовки и служебные строки
      const firstCell = String(row[0] || '').toLowerCase().trim();
      if (firstCell === 'отзывы' || firstCell === 'комментарии топ-20 выдачи' || 
          firstCell === 'активные обсуждения (мониторинг)' || firstCell === 'площадка') {
        return null;
      }

      // Извлекаем данные
      const extractedPlatform = this.extractPlatform(row, columnMapping);
      const theme = this.extractTheme(row, columnMapping);
      const textContent = this.extractText(row, columnMapping);
      const extractedDate = this.extractDate(row, columnMapping);
      const author = this.extractAuthor(row, columnMapping);
      const views = this.extractViews(row, columnMapping);
      const engagement = this.extractEngagement(row, columnMapping);
      const link = this.extractLink(row, columnMapping);
      
      // Определяем тип поста
      let type = 'ОС';
      
      const postTypeIndex = columnMapping.postType;
      if (postTypeIndex !== undefined && row.length > postTypeIndex && row[postTypeIndex]) {
        const postTypeValue = String(row[postTypeIndex]).trim().toUpperCase();
        if (postTypeValue === 'ОС' || postTypeValue === 'О.С.' || postTypeValue === 'OC') {
          type = 'ОС';
        } else if (postTypeValue === 'ЦС' || postTypeValue === 'Ц.С.' || postTypeValue === 'TC') {
          type = 'ЦС';
        } else if (postTypeValue === 'ПС' || postTypeValue === 'П.С.' || postTypeValue === 'PC') {
          type = 'ПС';
        }
      }

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
      console.warn(`⚠️ Ошибка обработки строки: ${error.message}`);
      return null;
    }
  }

  /**
   * Создание отчета
   */
  createReport(processedData) {
    const tempSpreadsheet = SpreadsheetApp.create(`temp_google_sheets_${Date.now()}_${this.monthInfo.name}_${this.monthInfo.year}_результат`);
    const reportSheetName = `${this.monthInfo.name}_${this.monthInfo.year}`;
    const sheet = tempSpreadsheet.getActiveSheet();
    sheet.setName(reportSheetName);
    
    // Шапка
    sheet.getRange('A1').setValue('Продукт');
    sheet.getRange('B1').setValue('Акрихин - Фортедетрим');
    sheet.getRange('A2').setValue('Период');
    sheet.getRange('B2').setValue(`${this.monthInfo.name}-25`);
    sheet.getRange('A3').setValue('План');
    
    // Заголовки таблицы
    const tableHeaders = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
    let row = 5;
    sheet.getRange(row, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(row, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#3f2355').setFontColor('white');
    row++;
    
    // Разделы и данные
    const writeSection = (sectionName, dataArr) => {
      sheet.getRange(row, 1).setValue(sectionName);
      sheet.getRange(row, 1, 1, tableHeaders.length).setBackground('#b7a6c9').setFontWeight('bold');
      row++;
      if (dataArr.length) {
        const safeData = dataArr.map(r => {
          const arr = [r.platform, r.theme, r.text, r.date, r.author, r.views, r.engagement, r.type];
          while (arr.length < tableHeaders.length) arr.push('');
          return arr.slice(0, tableHeaders.length);
        });
        sheet.getRange(row, 1, safeData.length, tableHeaders.length).setValues(safeData);
        row += safeData.length;
      }
      console.log(`📂 Раздел "${sectionName}": ${dataArr.length} строк`);
    };
    
    writeSection('Отзывы', processedData.reviews);
    writeSection('Комментарии Топ-20 выдачи', processedData.commentsTop20);
    writeSection('Активные обсуждения (мониторинг)', processedData.activeDiscussions);
    
    // Блок статистики
    row += 2;
    sheet.getRange(row, 1).setValue('Суммарное количество просмотров');
    sheet.getRange(row, 2).setValue(this.stats.totalViews || 0);
    row++;
    sheet.getRange(row, 1).setValue('Количество карточек товара (отзывы)');
    sheet.getRange(row, 2).setValue(processedData.reviews.length);
    row++;
    sheet.getRange(row, 1).setValue('Количество обсуждений (форумы, сообщества, комментарии к статьям)');
    sheet.getRange(row, 2).setValue(processedData.activeDiscussions.length);
    row++;
    sheet.getRange(row, 1).setValue('Доля обсуждений с вовлечением в диалог');
    const engagementValue = this.stats.engagementShare || 0;
    sheet.getRange(row, 2).setValue(engagementValue > 0 ? engagementValue + '%' : '0%');
    sheet.getRange(row - 3, 1, 4, 2).setFontWeight('bold');
    
    // Форматирование
    sheet.autoResizeColumns(1, tableHeaders.length);
    
    console.log(`📄 Отчет создан: ${reportSheetName}`);
    return tempSpreadsheet.getUrl();
  }

  // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================

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
      if (lowerText.includes(month.name.toLowerCase()) || 
          lowerText.includes(month.short.toLowerCase())) {
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
   * Получение имени месяца по индексу
   */
  getMonthName(index) {
    const names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                   'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    return names[index];
  }

  /**
   * Получение короткого имени месяца
   */
  getMonthShort(index) {
    const shorts = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                    'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
    return shorts[index];
  }

  /**
   * Фиксированный маппинг колонок
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
   * Извлечение статистики из исходных данных
   */
  extractStatisticsFromSourceData(data) {
    const stats = {
      totalViews: 0,
      totalCards: 0,
      totalDiscussions: 0,
      engagementShare: 0
    };
    
    // Ищем блок статистики в конце файла
    for (let i = data.length - 1; i >= Math.max(0, data.length - 20); i--) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // Ищем строку с общими просмотрами
      if (firstCell.includes('суммарное количество просмотров')) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            const value = parseFloat(String(row[j]).replace(/[^\d]/g, ''));
            if (!isNaN(value) && value > 0) {
              stats.totalViews = value;
              console.log(`📊 Найдены общие просмотры в исходных данных: ${value}`);
              break;
            }
          }
        }
      }
      
      // Ищем долю вовлечения
      if (firstCell.includes('доля обсуждений с вовлечением')) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            let value = String(row[j]).trim();
            if (value.includes('%')) {
              value = value.replace('%', '');
            }
            const floatValue = parseFloat(value);
            if (!isNaN(floatValue)) {
              stats.engagementShare = floatValue > 1 ? floatValue : floatValue * 100;
              break;
            }
          }
        }
      }
    }
    
    return stats;
  }

  /**
   * Проверки типов строк
   */
  isEmptyRow(row) {
    return !row || row.every(cell => !cell || String(cell).trim() === '');
  }

  isSectionHeader(row) {
    if (!row || row.length === 0) return false;
    const firstCell = String(row[0] || '').toLowerCase().trim();
    return firstCell.includes('отзывы') || 
           firstCell.includes('комментарии') || 
           firstCell.includes('обсуждения') ||
           firstCell.includes('топ-20') ||
           firstCell.includes('мониторинг');
  }

  isStatisticsRow(row) {
    if (!row || row.length === 0) return false;
    const firstCell = String(row[0] || '').toLowerCase().trim();
    return firstCell.includes('суммарное количество просмотров') || 
           firstCell.includes('количество карточек товара') ||
           firstCell.includes('количество обсуждений') ||
           firstCell.includes('доля обсуждений');
  }

  /**
   * Методы извлечения данных
   */
  extractPlatform(row, columnMapping) {
    const index = columnMapping.platform;
    return index !== undefined && row[index] ? String(row[index]).trim() : '';
  }

  extractText(row, columnMapping) {
    const index = columnMapping.text;
    return index !== undefined && row[index] ? String(row[index]).trim() : '';
  }

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

  extractAuthor(row, columnMapping) {
    const index = columnMapping.author;
    return index !== undefined && row[index] ? String(row[index]).trim() : '';
  }

  extractViews(row, columnMapping) {
    const index = columnMapping.views;
    if (index !== undefined && row[index]) {
      const viewsValue = row[index];
      
      if (typeof viewsValue === 'number' && !isNaN(viewsValue)) {
        return Math.max(0, Math.floor(viewsValue));
      }
      
      const viewsStr = String(viewsValue).trim();
      const cleanStr = viewsStr.replace(/[^\d.,]/g, '');
      const normalizedStr = cleanStr.replace(',', '.');
      const parsed = parseFloat(normalizedStr);
      
      if (!isNaN(parsed)) {
        return Math.max(0, Math.floor(parsed));
      }
    }
    return 0;
  }

  extractEngagement(row, columnMapping) {
    const index = columnMapping.engagement;
    return index !== undefined && row[index] ? String(row[index]).trim() : '';
  }

  extractTheme(row, columnMapping) {
    const index = columnMapping.theme;
    return index !== undefined && row[index] ? String(row[index]).trim() : '';
  }

  extractLink(row, columnMapping) {
    const index = columnMapping.link;
    return index !== undefined && row[index] ? String(row[index]).trim() : '';
  }
}

// ==================== ФУНКЦИИ ДЛЯ GOOGLE APPS SCRIPT ====================

/**
 * Основная функция для запуска из меню
 */
function processMonthlyReport() {
  const processor = new FinalMonthlyReportProcessor();
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  
  const result = processor.processReport(spreadsheetId, sheetName);
  
  if (result.success) {
    SpreadsheetApp.getUi().alert(
      'Обработка завершена',
      `Отчет создан успешно!\n\nСсылка: ${result.reportUrl}\n\n` +
      `Обработано:\n- Отзывов: ${result.statistics.reviewsCount}\n` +
      `- Комментариев топ-20: ${result.statistics.commentsTop20Count}\n` +
      `- Активных обсуждений: ${result.statistics.activeDiscussionsCount}\n` +
      `- Общие просмотры: ${result.statistics.totalViews}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'Ошибка обработки',
      `Произошла ошибка: ${result.error}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Создание меню при открытии таблицы
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Обработка отчетов')
    .addItem('🚀 Обработать текущий месяц', 'processMonthlyReport')
    .addToUi();
}