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
   * Обработка данных (ИСПРАВЛЕНО - версия 3 с типами постов)
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
        engagementShare: sourceStats.engagementShare || 0,
        platforms: new Set()
      }
    };
    
    let processedRows = 0;
    let skippedRows = 0;
    
    // Получаем фиксированный маппинг
    const columnMapping = this.getColumnMapping();
    
    // Определяем границы разделов
    const sections = this.findSectionBoundaries(data);
    
    if (sections.length === 0) {
      console.error('❌ Не удалось определить разделы в данных');
      return processedData;
    }
    
    // Обрабатываем каждый раздел
    sections.forEach(section => {
      console.log(`🔄 Обработка раздела "${section.name}" (строки ${section.startRow + 1}-${section.endRow + 1})`);
      
      for (let i = section.startRow; i <= section.endRow; i++) {
        const row = data[i];
        
        // Пропускаем пустые строки
        if (this.isEmptyRow(row)) {
          skippedRows++;
          continue;
        }
        
        // Пропускаем заголовки
        if (this.isSectionHeader(row)) {
          skippedRows++;
          continue;
        }
        
        // Обрабатываем строку
        const processedRow = this.processRow(row, section.type, columnMapping);
        
        if (processedRow) {
          processedRows++;
          
          // Добавляем платформу в статистику
          if (processedRow.platform) {
            processedData.statistics.platforms.add(processedRow.platform);
          }
          
          // Распределяем по разделам
          if (section.type === 'reviews') {
            processedData.reviews.push(processedRow);
            processedData.statistics.totalReviews++;
          } else if (section.type === 'commentsTop20') {
            processedData.commentsTop20.push(processedRow);
            processedData.statistics.totalCommentsTop20++;
          } else if (section.type === 'activeDiscussions') {
            processedData.activeDiscussions.push(processedRow);
            processedData.statistics.totalActiveDiscussions++;
          }
        } else {
          skippedRows++;
          
          // Логируем пропущенную строку для отладки
          if (row && row[0]) {
            console.log(`[SKIP] processRow вернул null для строки ${i + 1}: ${row.slice(0, 5).map(cell => String(cell || '').substring(0, 20))}`);
          }
        }
      }
    });
    
    // Если просмотры не были извлечены из статистики, считаем из данных
    if (processedData.statistics.totalViews === 0) {
      let totalViews = 0;
      
      // Суммируем просмотры из всех разделов
      [...processedData.reviews, ...processedData.commentsTop20, ...processedData.activeDiscussions]
        .forEach(item => {
          if (item.views && item.views > 0) {
            totalViews += item.views;
          }
        });
      
      if (totalViews > 0) {
        processedData.statistics.totalViews = totalViews;
        console.log(`📊 Просмотры подсчитаны из записей: ${totalViews}`);
      }
    }
    
    // Рассчитываем долю вовлечения если не была извлечена
    if (processedData.statistics.engagementShare === 0 && processedData.statistics.totalActiveDiscussions > 0) {
      // Считаем записи с вовлечением (где есть значение в колонке engagement)
      let engagedCount = 0;
      processedData.activeDiscussions.forEach(item => {
        if (item.engagement && item.engagement.trim() !== '' && item.engagement !== '0') {
          engagedCount++;
        }
      });
      
      if (engagedCount > 0) {
        processedData.statistics.engagementShare = engagedCount / processedData.statistics.totalActiveDiscussions;
        console.log(`📊 Доля вовлечения рассчитана: ${(processedData.statistics.engagementShare * 100).toFixed(0)}%`);
      }
    }
    
    console.log(`📊 Обработано: ${processedRows} строк данных, пропущено: ${skippedRows} строк`);
    console.log(`📊 Результат: ${processedData.statistics.totalReviews} отзывов, ${processedData.statistics.totalCommentsTop20} топ-20, ${processedData.statistics.totalActiveDiscussions} обсуждений`);
    
    return processedData;
  }

  /**
   * Поиск границ разделов (ИСПРАВЛЕНО - точное определение по заголовкам)
   */
  findSectionBoundaries(data) {
    const sections = [];
    
    console.log('🔍 Анализ структуры данных для определения разделов...');
    
    // Ищем заголовки разделов
    let reviewsHeaderRow = -1;
    let commentsHeaderRow = -1;
    let discussionsHeaderRow = -1;
    let statsStartRow = -1;
    
    // Проходим по всем строкам и ищем заголовки
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // Ищем заголовок "Отзывы"
      if (reviewsHeaderRow === -1 && firstCell === 'отзывы') {
        reviewsHeaderRow = i;
        console.log(`📂 Найден заголовок "Отзывы" в строке ${i + 1}`);
      }
      // Ищем заголовок "Комментарии Топ-20"
      else if (commentsHeaderRow === -1 && 
               (firstCell.includes('комментарии топ-20') || 
                firstCell.includes('комментарии топ 20') ||
                firstCell === 'комментарии топ-20 выдачи')) {
        commentsHeaderRow = i;
        console.log(`📂 Найден заголовок "Комментарии Топ-20" в строке ${i + 1}`);
      }
      // Ищем заголовок "Активные обсуждения"
      else if (discussionsHeaderRow === -1 && 
               (firstCell.includes('активные обсуждения') || 
                firstCell === 'активные обсуждения (мониторинг)')) {
        discussionsHeaderRow = i;
        console.log(`📂 Найден заголовок "Активные обсуждения" в строке ${i + 1}`);
      }
      // Ищем начало статистики
      else if (this.isStatisticsRow(row)) {
        statsStartRow = i;
        console.log(`📊 Найдена статистика в строке ${i + 1}`);
        break;
      }
    }
    
    // Создаем разделы на основе найденных заголовков
    // Раздел "Отзывы"
    if (reviewsHeaderRow !== -1) {
      const startRow = reviewsHeaderRow + 1; // Данные начинаются после заголовка
      let endRow = startRow;
      
      // Ищем конец раздела (следующий заголовок или пустые строки)
      if (commentsHeaderRow > reviewsHeaderRow) {
        endRow = commentsHeaderRow - 1;
      } else if (discussionsHeaderRow > reviewsHeaderRow) {
        endRow = discussionsHeaderRow - 1;
      } else if (statsStartRow > reviewsHeaderRow) {
        endRow = statsStartRow - 1;
      }
      
      // Убираем пустые строки в конце
      while (endRow > startRow && this.isEmptyRow(data[endRow])) {
        endRow--;
      }
      
      if (endRow >= startRow) {
        sections.push({
          type: 'reviews',
          name: 'Отзывы',
          startRow: startRow,
          endRow: endRow
        });
      }
    }
    
    // Раздел "Комментарии Топ-20"
    if (commentsHeaderRow !== -1) {
      const startRow = commentsHeaderRow + 1;
      let endRow = startRow;
      
      if (discussionsHeaderRow > commentsHeaderRow) {
        endRow = discussionsHeaderRow - 1;
      } else if (statsStartRow > commentsHeaderRow) {
        endRow = statsStartRow - 1;
      }
      
      // Убираем пустые строки в конце
      while (endRow > startRow && this.isEmptyRow(data[endRow])) {
        endRow--;
      }
      
      if (endRow >= startRow) {
        sections.push({
          type: 'commentsTop20',
          name: 'Комментарии Топ-20',
          startRow: startRow,
          endRow: endRow
        });
      }
    }
    
    // Раздел "Активные обсуждения"
    if (discussionsHeaderRow !== -1) {
      const startRow = discussionsHeaderRow + 1;
      let endRow = startRow;
      
      if (statsStartRow > discussionsHeaderRow) {
        endRow = statsStartRow - 1;
      } else {
        // Если статистика не найдена, идем до конца данных
        endRow = data.length - 1;
      }
      
      // Убираем пустые строки и статистику в конце
      while (endRow > startRow && 
             (this.isEmptyRow(data[endRow]) || this.isStatisticsRow(data[endRow]))) {
        endRow--;
      }
      
      if (endRow >= startRow) {
        sections.push({
          type: 'activeDiscussions',
          name: 'Активные обсуждения',
          startRow: startRow,
          endRow: endRow
        });
      }
    }
    
    // Логируем результаты
    console.log('📂 Найденные разделы:');
    sections.forEach(section => {
      const count = section.endRow - section.startRow + 1;
      console.log(`   - ${section.name}: строки ${section.startRow + 1}-${section.endRow + 1} (${count} записей)`);
    });
    
    return sections;
  }

  /**
   * Обработка строки данных (ИСПРАВЛЕНО - версия 2)
   */
  processRow(row, currentSection, columnMapping) {
    try {
      // Проверяем, что строка содержит данные
      if (!row || row.length === 0) {
        return null;
      }

      // Пропускаем пустые строки
      if (this.isEmptyRow(row)) {
        return null;
      }
      
      // Пропускаем заголовки разделов
      const firstCell = String(row[0] || '').toLowerCase().trim();
      if (firstCell.includes('отзывы') || 
          firstCell.includes('комментарии') || 
          firstCell.includes('обсуждения') ||
          firstCell.includes('топ-20')) {
        return null;
      }

      // Проверяем наличие значимых данных
      const textIndex = columnMapping.text || 4;
      const platformIndex = columnMapping.platform || 1;
      const linkIndex = columnMapping.link || 2;
      
      const text = row[textIndex] ? String(row[textIndex]).trim() : '';
      const platform = row[platformIndex] ? String(row[platformIndex]).trim() : '';
      const link = row[linkIndex] ? String(row[linkIndex]).trim() : '';
      
      // Пропускаем строки без текста или платформы
      if (!text && !platform && !link) {
        return null;
      }
      
      // Определяем тип записи
      let recordType = currentSection;
      
      // Проверяем колонку "Тип поста"
      const postTypeIndex = columnMapping.postType || 13;
      if (row[postTypeIndex]) {
        const postType = String(row[postTypeIndex]).trim().toUpperCase();
        
        if (postType === 'ОС' || postType === 'О.С.') {
          recordType = 'reviews';
        } else if (postType === 'ЦС' || postType === 'Ц.С.') {
          recordType = 'commentsTop20';
        } else if (postType === 'ПС' || postType === 'П.С.') {
          recordType = 'activeDiscussions';
        }
      }

      // Извлекаем данные из строки
      const processedRow = {
        platform: platform,
        theme: row[columnMapping.theme || 3] ? String(row[columnMapping.theme || 3]).trim() : '',
        link: link,
        text: text,
        date: this.extractDate(row, columnMapping),
        author: row[columnMapping.author || 7] ? String(row[columnMapping.author || 7]).trim() : '',
        views: this.extractViews(row, columnMapping),
        engagement: row[columnMapping.engagement || 12] ? String(row[columnMapping.engagement || 12]).trim() : '',
        postType: row[postTypeIndex] ? String(row[postTypeIndex]).trim() : '',
        recordType: recordType
      };

      // Дополнительная валидация
      if (!processedRow.text && !processedRow.platform) {
        return null;
      }

      return processedRow;

    } catch (e) {
      console.error(`❌ Ошибка обработки строки: ${e.message}`);
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
    
    // 4. Блок статистики внизу (ИСПРАВЛЕНО)
    row += 2;
    sheet.getRange(row, 1).setValue('Суммарное количество просмотров');
    sheet.getRange(row, 2).setValue(processedData.statistics.totalViews || 0);
    row++;
    sheet.getRange(row, 1).setValue('Количество карточек товара (отзывы)');
    sheet.getRange(row, 2).setValue(processedData.statistics.totalReviews || 0);
    row++;
    sheet.getRange(row, 1).setValue('Количество обсуждений (форумы, сообщества, комментарии к статьям)');
    const totalDiscussions = (processedData.statistics.totalActiveDiscussions || 0) + 
                           (processedData.statistics.totalCommentsTop20 || 0);
    sheet.getRange(row, 2).setValue(totalDiscussions);
    row++;
    sheet.getRange(row, 1).setValue('Доля обсуждений с вовлечением в диалог');
    const engagementValue = processedData.statistics.engagementShare || 0;
    // Форматируем как проценты
    if (engagementValue > 0) {
      sheet.getRange(row, 2).setValue(engagementValue);
      sheet.getRange(row, 2).setNumberFormat("0%");
    } else {
      sheet.getRange(row, 2).setValue(0);
    }
    
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
   * Извлечение статистики из исходных данных (УЛУЧШЕНО)
   */
  extractStatisticsFromSourceData(data) {
    const stats = {
      totalViews: 0,
      totalCards: 0,
      totalDiscussions: 0,
      engagementShare: 0
    };
    
    console.log('📊 Извлечение статистики из исходных данных...');
    
    // Ищем блок статистики в последних 20 строках файла
    const startSearch = Math.max(0, data.length - 20);
    
    for (let i = startSearch; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const firstCell = String(row[0] || '').toLowerCase().trim();
      const secondCell = row[1] ? String(row[1]).trim() : '';
      
      // Ищем строку с общими просмотрами
      if (firstCell.includes('суммарное количество просмотров')) {
        // Ищем число в строке
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            const value = parseFloat(String(row[j]).replace(/[^\d]/g, ''));
            if (!isNaN(value) && value > 0) {
              stats.totalViews = Math.floor(value);
              console.log(`✅ Найдены общие просмотры: ${stats.totalViews}`);
              break;
            }
          }
        }
      }
      
      // Ищем количество карточек товара
      if (firstCell.includes('количество карточек товара')) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            const value = parseFloat(String(row[j]).replace(/[^\d]/g, ''));
            if (!isNaN(value) && value >= 0) {
              stats.totalCards = Math.floor(value);
              console.log(`✅ Найдено карточек товара: ${stats.totalCards}`);
              break;
            }
          }
        }
      }
      
      // Ищем количество обсуждений
      if (firstCell.includes('количество обсуждений')) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            const value = parseFloat(String(row[j]).replace(/[^\d]/g, ''));
            if (!isNaN(value) && value >= 0) {
              stats.totalDiscussions = Math.floor(value);
              console.log(`✅ Найдено обсуждений: ${stats.totalDiscussions}`);
              break;
            }
          }
        }
      }
      
      // Ищем долю вовлечения
      if (firstCell.includes('доля обсуждений с вовлечением')) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            const cellValue = String(row[j]).trim();
            let value = 0;
            
            // Проверяем разные форматы
            if (cellValue.includes('%')) {
              // Формат с процентом: "20%"
              value = parseFloat(cellValue.replace('%', '')) / 100;
            } else if (cellValue.includes('.')) {
              // Десятичный формат: "0.20"
              value = parseFloat(cellValue);
            } else {
              // Целое число: "20" (предполагаем проценты)
              const num = parseFloat(cellValue);
              if (!isNaN(num)) {
                value = num > 1 ? num / 100 : num;
              }
            }
            
            if (!isNaN(value) && value >= 0) {
              stats.engagementShare = value;
              console.log(`✅ Найдена доля вовлечения: ${(value * 100).toFixed(0)}%`);
              break;
            }
          }
        }
      }
    }
    
    // Альтернативный поиск если основной не сработал
    if (stats.totalViews === 0) {
      // Суммируем просмотры из колонки просмотров
      let sumViews = 0;
      const viewsIndex = 11; // колонка просмотров
      
      for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < Math.min(data.length - 10, 650); i++) {
        const row = data[i];
        if (row && row[viewsIndex]) {
          const views = this.extractViews(row, this.getColumnMapping());
          if (views > 0) {
            sumViews += views;
          }
        }
      }
      
      if (sumViews > 0) {
        stats.totalViews = sumViews;
        console.log(`📊 Просмотры подсчитаны из данных: ${sumViews}`);
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