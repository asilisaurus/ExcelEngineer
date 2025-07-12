/**
 * ============================================================================
 * GOOGLE APPS SCRIPT PROCESSOR - FINAL VERSION
 * ============================================================================
 * Исправленная версия на основе реального анализа данных
 * Основные улучшения:
 * - Определение типа контента по колонке "тип поста"
 * - Фиксированный маппинг колонок для Google Sheets
 * - Фильтрация заголовков
 * - Упрощенные методы извлечения
 * ============================================================================
 */

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Главная функция обработки Google Sheets
 */
function processGoogleSheets() {
  try {
    console.log('🚀 Начинаю обработку Google Sheets...');
    
    // 1. Получаем исходные данные
    const sourceData = getSourceData();
    console.log(`📊 Загружено ${sourceData.length} строк из источника`);
    
    // 2. Обрабатываем данные
    const processedData = processData(sourceData);
    console.log(`✅ Обработано записей: отзывов ${processedData.statistics.reviewsCount}, целевых ${processedData.statistics.targetedCount}, социальных ${processedData.statistics.socialCount}`);
    
    // 3. Создаем результирующий файл
    const resultFileId = createResultFile(processedData);
    console.log(`💾 Создан файл результата: ${resultFileId}`);
    
    return {
      success: true,
      sourceRows: sourceData.length,
      processedRows: processedData.statistics.totalRows,
      reviewsCount: processedData.statistics.reviewsCount,
      targetedCount: processedData.statistics.targetedCount,
      socialCount: processedData.statistics.socialCount,
      totalViews: processedData.statistics.totalViews,
      totalEngagement: processedData.statistics.totalEngagement,
      resultFileId: resultFileId,
      statistics: processedData.statistics
    };
    
  } catch (error) {
    console.error('❌ Ошибка при обработке:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Получение исходных данных из Google Sheets
 */
function getSourceData() {
  // ID реальной таблицы для тестирования
  const SHEET_ID = '1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI';
  
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    // Автоматический поиск листа с текущим месяцем
    const targetSheet = findCurrentMonthSheet(sheets);
    
    if (!targetSheet) {
      throw new Error('Не найден лист для текущего месяца');
    }
    
    console.log(`📋 Используется лист: ${targetSheet.getName()}`);
    
    // Получаем все данные листа
    const range = targetSheet.getDataRange();
    const values = range.getValues();
    
    console.log(`📊 Загружено ${values.length} строк из листа ${targetSheet.getName()}`);
    
    return values;
    
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
  
  // 2. Обработка данных по разделам
  const processedData = {
    reviews: [],
    targeted: [],
    social: [],
    statistics: {
      totalRows: 0,
      reviewsCount: 0,
      targetedCount: 0,
      socialCount: 0,
      totalViews: 0,
      totalEngagement: 0
    }
  };
  
  let currentSection = null;
  
  // 3. Проходим по всем строкам
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    
    // Пропускаем пустые строки
    if (isEmptyRow(row)) continue;
    
    // Определяем секцию по первой ячейке
    const firstCell = (row[0] || '').toString().toLowerCase().trim();
    
    // Проверяем, является ли строка заголовком раздела
    if (isSectionHeader(firstCell)) {
      currentSection = determineSectionType(firstCell);
      console.log(`📋 Найден раздел: ${currentSection} (строка ${i + 1})`);
      continue;
    }
    
    // Пропускаем заголовки таблицы
    if (isTableHeader(row)) {
      console.log(`📊 Пропуск заголовка таблицы: строка ${i + 1}`);
      continue;
    }
    
    // Обрабатываем строки данных
    if (currentSection && isDataRow(row, headerInfo.mapping)) {
      const processedRow = extractRowData(row, currentSection, headerInfo.mapping);
      
      if (processedRow) {
        processedData[currentSection].push(processedRow);
        processedData.statistics.totalRows++;
        
        // Обновляем статистику
        switch (currentSection) {
          case 'reviews':
            processedData.statistics.reviewsCount++;
            break;
          case 'targeted':
            processedData.statistics.targetedCount++;
            break;
          case 'social':
            processedData.statistics.socialCount++;
            break;
        }
        
        processedData.statistics.totalViews += processedRow.views || 0;
        processedData.statistics.totalEngagement += processedRow.engagement || 0;
        
        console.log(`✅ Обработана строка ${i + 1}: ${processedRow.site} - ${processedRow.views} просмотров`);
      }
    }
  }
  
  console.log(`📊 Итого: отзывов ${processedData.statistics.reviewsCount}, целевых ${processedData.statistics.targetedCount}, социальных ${processedData.statistics.socialCount}`);
  console.log(`📈 Общие просмотры: ${processedData.statistics.totalViews}, вовлечение: ${processedData.statistics.totalEngagement}`);
  
  return processedData;
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
 * Проверка, является ли строка заголовком раздела
 */
function isSectionHeader(text) {
  const sectionPatterns = [
    'отзывы сайтов', 'ос',
    'целевые сайты', 'цс', 
    'площадки социальные', 'пс',
    'отзыв', 'целевые', 'площадки'
  ];
  
  return sectionPatterns.some(pattern => text.includes(pattern));
}

/**
 * Определение типа раздела
 */
function determineSectionType(text) {
  if (text.includes('отзыв') || text.includes('ос')) {
    return 'reviews';
  }
  if (text.includes('целевые') || text.includes('цс')) {
    return 'targeted';
  }
  if (text.includes('площадки') || text.includes('пс')) {
    return 'social';
  }
  return 'other';
}

/**
 * Проверка, является ли строка заголовком таблицы
 */
function isTableHeader(row) {
  if (!row || row.length === 0) return false;
  
  const rowText = row.join(' ').toLowerCase();
  const headerPatterns = ['тип размещения', 'площадка', 'текст сообщения'];
  
  return headerPatterns.some(pattern => rowText.includes(pattern));
}

/**
 * Проверка, является ли строка данными
 */
function isDataRow(row, mapping) {
  if (!row || row.length < 3) return false;
  
  // Проверяем наличие основных данных
  const platformIndex = mapping['площадка'] || 1;
  const textIndex = mapping['текст сообщения'] || 4;
  
  const platform = row[platformIndex];
  const text = row[textIndex];
  
  return platform && text && 
         platform.toString().trim().length > 0 && 
         text.toString().trim().length > 10;
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
function extractRowData(row, sectionType, mapping) {
  try {
    // Индексы колонок (с fallback на стандартные позиции)
    const placementTypeIndex = mapping['тип размещения'] || 0;
    const siteIndex = mapping['площадка'] || 1;
    const productIndex = mapping['продукт'] || 2;
    const linkIndex = mapping['ссылка на сообщение'] || 3;
    const textIndex = mapping['текст сообщения'] || 4;
    const approvalIndex = mapping['согласование/комментарии'] || 5;
    const dateIndex = mapping['дата'] || 6;
    const nicknameIndex = mapping['ник'] || 7;
    const authorIndex = mapping['автор'] || 8;
    const startViewsIndex = mapping['просмотры темы на старте'] || 9;
    const endViewsIndex = mapping['просмотры в конце месяца'] || 10;
    const viewsIndex = mapping['просмотров получено'] || 11;
    const engagementIndex = mapping['вовлечение'] || 12;
    const postTypeIndex = mapping['тип поста'] || 13;
    
    // Извлекаем данные
    const placementType = cleanValue(row[placementTypeIndex]);
    const site = cleanValue(row[siteIndex]);
    const product = cleanValue(row[productIndex]);
    const link = cleanValue(row[linkIndex]);
    const text = cleanValue(row[textIndex]);
    const approval = cleanValue(row[approvalIndex]);
    const date = cleanValue(row[dateIndex]);
    const nickname = cleanValue(row[nicknameIndex]);
    const author = cleanValue(row[authorIndex]);
    const startViews = extractViews(row[startViewsIndex]);
    const endViews = extractViews(row[endViewsIndex]);
    const views = extractViews(row[viewsIndex]);
    const engagement = extractEngagement(row[engagementIndex]);
    const postType = cleanValue(row[postTypeIndex]);
    
    // Валидация обязательных полей
    if (!site || !text) {
      return null;
    }
    
    return {
      section: sectionType,
      placementType: placementType,
      site: site,
      product: product,
      link: link,
      text: text,
      approval: approval,
      date: date,
      nickname: nickname,
      author: author,
      startViews: startViews,
      endViews: endViews,
      views: views,
      engagement: engagement,
      postType: postType
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
    const headers = [
      'Раздел', 'Тип размещения', 'Площадка', 'Продукт', 'Ссылка', 
      'Текст сообщения', 'Дата', 'Автор', 'Просмотры получено', 'Вовлечение'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Форматируем заголовки
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    let currentRow = 2;
    
    // Добавляем данные по разделам
    const sections = [
      { key: 'reviews', name: 'ОТЗЫВЫ САЙТОВ (ОС)' },
      { key: 'targeted', name: 'ЦЕЛЕВЫЕ САЙТЫ (ЦС)' },
      { key: 'social', name: 'ПЛОЩАДКИ СОЦИАЛЬНЫЕ (ПС)' }
    ];
    
    sections.forEach(section => {
      const sectionData = processedData[section.key] || [];
      
      if (sectionData.length > 0) {
        // Добавляем заголовок раздела
        sheet.getRange(currentRow, 1).setValue(section.name);
        sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#e8f0fe');
        sheet.getRange(currentRow, 1, 1, headers.length).setFontWeight('bold');
        currentRow++;
        
        // Добавляем данные раздела
        sectionData.forEach(row => {
          const rowData = [
            section.key,
            row.placementType || '',
            row.site || '',
            row.product || '',
            row.link || '',
            (row.text || '').substring(0, 100) + (row.text && row.text.length > 100 ? '...' : ''),
            row.date || '',
            row.author || '',
            row.views || 0,
            row.engagement || 0
          ];
          
          sheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
          currentRow++;
        });
        
        currentRow++; // Пустая строка между разделами
      }
    });
    
    // Добавляем итоговую строку
    const stats = processedData.statistics || {};
    const totalData = [
      'ИТОГО',
      '',
      `Всего площадок: ${(processedData.reviews?.length || 0) + (processedData.targeted?.length || 0) + (processedData.social?.length || 0)}`,
      '',
      '',
      `Отзывов: ${stats.reviewsCount || 0}, Целевых: ${stats.targetedCount || 0}, Социальных: ${stats.socialCount || 0}`,
      '',
      '',
      stats.totalViews || 0,
      stats.totalEngagement || 0
    ];
    
    // Добавляем пустую строку
    sheet.getRange(currentRow, 1, 1, headers.length).setValues([Array(headers.length).fill('')]);
    currentRow++;
    
    // Добавляем итоговую строку
    const totalRange = sheet.getRange(currentRow, 1, 1, totalData.length);
    totalRange.setValues([totalData]);
    totalRange.setFontWeight('bold');
    totalRange.setBackground('#e8f0fe');
    
    // Автоматическое изменение размера колонок
    sheet.autoResizeColumns(1, headers.length);
    
    console.log(`💾 Создан файл: ${fileName}`);
    console.log(`🔗 ID файла: ${spreadsheet.getId()}`);
    console.log(`📊 Статистика: отзывов ${stats.reviewsCount}, целевых ${stats.targetedCount}, социальных ${stats.socialCount}`);
    console.log(`📈 Просмотры: ${stats.totalViews}, вовлечение: ${stats.totalEngagement}`);
    
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
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    return { success: false, error: error.message };
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