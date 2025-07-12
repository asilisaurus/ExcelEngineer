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
    
    console.log(`📊 Загружено ${data.length} строк из листа ${SpreadsheetApp.getActiveSheet().getName()}`);
    
  } catch (error) {
    console.error('❌ Ошибка при получении данных:', error);
    throw error;
  }
}

/**
 * Поиск листа с текущим месяцем
 */
function findCurrentMonthSheet(sheets) {
  const currentMonth = new Date().getMonth();
  const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 
                     'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
  
  // Сначала ищем точное совпадение
  for (let sheet of sheets) {
    const sheetName = sheet.getName().toLowerCase();
    if (sheetName.includes(monthNames[currentMonth])) {
      return sheet;
    }
  }
  
  // Если не найдено, берем первый лист
  return sheets[0];
}

/**
 * Основная обработка данных
 */
function processData(rawData) {
  if (!rawData || rawData.length === 0) {
    throw new Error('Нет данных для обработки');
  }
  
  // 1. Поиск заголовков
  const headerInfo = findHeaders(rawData);
  console.log(`🔍 Найдена строка заголовков: ${headerInfo.row}`);
  
  // 2. Извлечение данных
  const dataRows = rawData.slice(headerInfo.row + 1);
  
  // 3. Обработка строк
  const processedRows = [];
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    
    // Пропускаем пустые строки
    if (isEmptyRow(row)) continue;
    
    // Пропускаем строки-заголовки
    if (isHeaderRow(row)) {
      console.log(`⏭️ Пропущена строка-заголовок: ${row[0]}`);
      continue;
    }
    
    // Определяем тип контента
    const contentType = determineContentType(row);
    
    if (contentType === 'review' || contentType === 'comment') {
      const processedRow = extractRowData(row, contentType, headerInfo.mapping);
      if (processedRow) {
        processedRows.push(processedRow);
      }
    }
  }
  
  // 4. Группировка по типам
  const reviews = processedRows.filter(row => row.type === 'review');
  const comments = processedRows.filter(row => row.type === 'comment');
  
  console.log(`📊 Найдено отзывов: ${reviews.length}, комментариев: ${comments.length}`);
  
  return processedRows;
}

/**
 * Поиск строки заголовков
 */
function findHeaders(data) {
  const columnMapping = {};
  
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    const rowStr = row.map(cell => (cell || '').toString().toLowerCase()).join(' ');
    
    if (rowStr.includes('тип размещения') || 
        rowStr.includes('площадка') || 
        rowStr.includes('текст сообщения')) {
      
      // Создаем маппинг колонок
      row.forEach((cell, index) => {
        const cleanHeader = (cell || '').toString().toLowerCase().trim();
        if (cleanHeader) {
          columnMapping[cleanHeader] = index;
        }
      });
      
      return {
        row: i,
        mapping: columnMapping
      };
    }
  }
  
  // Если заголовки не найдены, используем стандартный маппинг
  return {
    row: 0,
    mapping: getDefaultMapping()
  };
}

/**
 * Стандартный маппинг колонок для Google Sheets
 */
function getDefaultMapping() {
  return {
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

/**
 * Определение типа контента
 */
function determineContentType(row) {
  // Проверяем последнюю колонку (тип поста)
  const lastColIndex = row.length - 1;
  const lastColValue = (row[lastColIndex] || '').toString().toLowerCase().trim();
  
  if (lastColValue === 'ос' || lastColValue === 'основное сообщение') {
    return 'review';
  }
  
  if (lastColValue === 'цс' || lastColValue === 'целевое сообщение') {
    return 'comment';
  }
  
  // Дополнительная проверка по первой колонке
  const firstColValue = (row[0] || '').toString().toLowerCase().trim();
  if (firstColValue.includes('отзыв') || firstColValue.includes('ос')) {
    return 'review';
  }
  
  if (firstColValue.includes('комментарий') || firstColValue.includes('цс')) {
    return 'comment';
  }
  
  return 'unknown';
}

/**
 * Проверка, является ли строка заголовком
 */
function isHeaderRow(row) {
  const firstCell = (row[0] || '').toString().toLowerCase().trim();
  const headerPatterns = ['отзывы', 'комментарии', 'обсуждения'];
  
  return headerPatterns.includes(firstCell);
}

/**
 * Проверка, является ли строка пустой
 */
function isEmptyRow(row) {
  return !row || row.every(cell => !cell || cell.toString().trim() === '');
}

/**
 * Извлечение данных из строки
 */
function extractRowData(row, type, mapping) {
  try {
    // Индексы колонок (с fallback на стандартные позиции)
    const siteIndex = mapping['площадка'] || 1;
    const linkIndex = mapping['ссылка на сообщение'] || 3;
    const textIndex = mapping['текст сообщения'] || 4;
    const dateIndex = mapping['дата'] || 6;
    const authorIndex = mapping['автор'] || 8;
    const viewsIndex = mapping['просмотров получено'] || 11;
    const engagementIndex = mapping['вовлечение'] || 12;
    
    // Извлекаем данные
    const site = cleanValue(row[siteIndex]);
    const link = cleanValue(row[linkIndex]);
    const text = cleanValue(row[textIndex]);
    const date = cleanValue(row[dateIndex]);
    const author = cleanValue(row[authorIndex]);
    const views = extractViews(row[viewsIndex]);
    const engagement = extractEngagement(row[engagementIndex]);
    
    // Валидация обязательных полей
    if (!site || !text) {
      return null;
    }
    
    return {
      type: type,
      site: site,
      link: link,
      text: text,
      date: date,
      author: author,
      views: views,
      engagement: engagement
    };
    
  } catch (error) {
    console.error('❌ Ошибка при извлечении данных строки:', error);
    return null;
  }
}

/**
 * Очистка значения ячейки
 */
function cleanValue(value) {
  if (!value) return '';
  return value.toString().trim();
}

/**
 * Извлечение просмотров
 */
function extractViews(value) {
  if (!value) return 0;
  
  const str = value.toString().replace(/\D/g, '');
  const num = parseInt(str);
  
  return isNaN(num) ? 0 : num;
}

/**
 * Извлечение вовлечения
 */
function extractEngagement(value) {
  if (!value) return 0;
  
  const str = value.toString().replace(/[^\d.]/g, '');
  const num = parseFloat(str);
  
  return isNaN(num) ? 0 : num;
}

/**
 * Создание результирующего файла
 */
function createResultFile(processedData) {
  try {
    // Создаем новый спредшит
    const currentDate = new Date();
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    
    const fileName = `Результат_${monthNames[currentDate.getMonth()]}_${currentDate.getFullYear()}`;
    const spreadsheet = SpreadsheetApp.create(fileName);
    
    // Получаем активный лист
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName('Результаты');
    
    // Заголовки
    const headers = ['Тип', 'Площадка', 'Ссылка', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Вовлечение'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Данные
    const dataForSheet = processedData.map(row => [
      row.type === 'review' ? 'Отзыв' : 'Комментарий',
      row.site,
      row.link,
      row.text,
      row.date,
      row.author,
      row.views,
      row.engagement
    ]);
    
    if (dataForSheet.length > 0) {
      sheet.getRange(2, 1, dataForSheet.length, headers.length).setValues(dataForSheet);
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
    const index = columnMapping.views; // колонка 11
    if (index !== undefined && row[index]) {
      const viewsValue = row[index];
      
      // Если это уже число, возвращаем как есть
      if (typeof viewsValue === 'number' && !isNaN(viewsValue)) {
        return Math.max(0, Math.floor(viewsValue));
      }
      
      // Преобразуем строку в число
      const viewsStr = String(viewsValue).trim();
      
      // Удаляем все нечисловые символы кроме точки и запятой
      const cleanStr = viewsStr.replace(/[^\d.,]/g, '');
      
      // Заменяем запятую на точку для корректного парсинга
      const normalizedStr = cleanStr.replace(',', '.');
      
      const parsed = parseFloat(normalizedStr);
      
      if (!isNaN(parsed)) {
        return Math.max(0, Math.floor(parsed));
      }
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
    sheet.getRange(row, 2).setValue(this.stats.totalViews || processedData.statistics.totalViews || 0);
    row++;
    sheet.getRange(row, 1).setValue('Количество карточек товара (отзывы)');
    sheet.getRange(row, 2).setValue(processedData.reviews.length);
    row++;
    sheet.getRange(row, 1).setValue('Количество обсуждений (форумы, сообщества, комментарии к статьям)');
    sheet.getRange(row, 2).setValue(processedData.activeDiscussions.length);
    row++;
    sheet.getRange(row, 1).setValue('Доля обсуждений с вовлечением в диалог');
    const engagementValue = this.stats.engagementShare || 0;
    // Форматируем как проценты
    if (engagementValue > 0) {
      sheet.getRange(row, 2).setValue(engagementValue + '%');
    } else {
      sheet.getRange(row, 2).setValue('0%');
    }
    sheet.getRange(row - 3, 1, 4, 2).setFontWeight('bold');
    
    // Форматирование
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
    // Оформление
    sheet.autoResizeColumns(1, tableHeaders.length);
    
    // Применяем фильтры
    const dataRange = sheet.getDataRange();
    sheet.getRange(1, 1, dataRange.getNumRows(), dataRange.getNumColumns()).createFilter();
    
    console.log(`💾 Создан файл: ${fileName}`);
    console.log(`🔗 ID файла: ${spreadsheet.getId()}`);
    
    return spreadsheet.getId();
    
  } catch (error) {
    console.error('❌ Ошибка при создании файла:', error);
    throw error;
  }
}

// =============================================================================
// ТЕСТОВЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Тестирование процессора
 */
function testProcessor() {
  console.log('🧪 Запуск тестирования...');
  
  try {
    const result = processGoogleSheets();
    
    if (result.success) {
      console.log('✅ Тестирование успешно завершено');
      console.log(`📊 Обработано строк: ${result.processedRows}`);
      console.log(`🔗 ID результата: ${result.resultFileId}`);
    } else {
      console.log('❌ Тестирование завершилось с ошибкой:', result.error);
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
    
    return result;
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    return { success: false, error: error.message };
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
    
    // Сначала находим все заголовки разделов
    const sectionHeaders = [];
    
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // Пропускаем строки статистики
      if (this.isStatisticsRow(row)) {
        break; // Статистика идет после всех данных
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
      
      const startRow = currentHeader.headerRow + 1; // Данные начинаются после заголовка
      let endRow;
      
      if (nextHeader) {
        // Если есть следующий раздел, данные заканчиваются перед его заголовком
        endRow = nextHeader.headerRow - 1;
      } else {
        // Если это последний раздел, ищем конец данных (до статистики или конца файла)
        endRow = data.length - 1;
        
        // Ищем начало статистики
        for (let j = startRow; j < data.length; j++) {
          if (this.isStatisticsRow(data[j]) || this.isEmptyRow(data[j])) {
            // Проверяем, есть ли еще данные после пустой строки
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
      
      // Корректируем endRow, если раздел пустой или содержит только заголовки
      if (endRow < startRow) {
        endRow = startRow - 1; // Пустой раздел
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
        // Обычно значение находится в следующей колонке
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
      
      // Ищем количество карточек товара
      if (firstCell.includes('количество карточек товара')) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            const value = parseInt(String(row[j]).replace(/[^\d]/g, ''));
            if (!isNaN(value)) {
              stats.totalCards = value;
              break;
            }
          }
        }
      }
      
      // Ищем количество обсуждений
      if (firstCell.includes('количество обсуждений')) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) {
            const value = parseInt(String(row[j]).replace(/[^\d]/g, ''));
            if (!isNaN(value)) {
              stats.totalDiscussions = value;
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
            // Если это процент (например, "20%")
            if (value.includes('%')) {
              value = value.replace('%', '');
            }
            // Если это десятичная дробь (например, "0.20")
            const floatValue = parseFloat(value);
            if (!isNaN(floatValue)) {
              // Если значение больше 1, считаем что это проценты
              stats.engagementShare = floatValue > 1 ? floatValue : floatValue * 100;
              break;
            }
          }
        }
      }
    }
    
    return stats;
  }
}

/**
 * Анализ структуры данных
 */
function analyzeDataStructure() {
  console.log('🔍 Анализ структуры данных...');
  
  try {
    const sourceData = getSourceData();
    
    if (sourceData.length === 0) {
      console.log('❌ Нет данных для анализа');
      return;
    }
    
    // Анализируем первые 10 строк
    for (let i = 0; i < Math.min(10, sourceData.length); i++) {
      const row = sourceData[i];
      console.log(`Строка ${i + 1}:`, row.slice(0, 5).map(cell => 
        (cell || '').toString().substring(0, 30)
      ));
    }
    
    // Поиск заголовков
    const headerInfo = findHeaders(sourceData);
    console.log('📋 Найденные заголовки:', Object.keys(headerInfo.mapping));
    
    return {
      totalRows: sourceData.length,
      headerRow: headerInfo.row,
      columns: Object.keys(headerInfo.mapping)
    };
    
  } catch (error) {
    console.error('❌ Ошибка при анализе:', error);
    return { error: error.message };
  }
}

// =============================================================================
// ГЛАВНЫЕ ФУНКЦИИ ДЛЯ ЗАПУСКА
// =============================================================================

/**
 * Основная функция для запуска из Google Apps Script
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