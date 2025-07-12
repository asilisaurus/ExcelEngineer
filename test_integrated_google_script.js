/**
 * ============================================================================
 * ИНТЕГРИРОВАННЫЙ ТЕСТОВЫЙ ФАЙЛ ДЛЯ GOOGLE APPS SCRIPT ПРОЦЕССОРА
 * ============================================================================
 * Включает все функции из основного процессора для тестирования в Node.js
 * ============================================================================
 */

// =============================================================================
// ФУНКЦИИ ИЗ GOOGLE APPS SCRIPT ПРОЦЕССОРА
// =============================================================================

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

// =============================================================================
// МОКОВЫЕ ДАННЫЕ ДЛЯ ТЕСТИРОВАНИЯ
// =============================================================================

const mockGoogleSheetsData = [
  // Заголовки (строка 0)
  ['тип размещения', 'площадка', 'продукт', 'ссылка на сообщение', 'текст сообщения', 
   'согласование/комментарии', 'дата', 'ник', 'автор', 'просмотры темы на старте', 
   'просмотры в конце месяца', 'просмотров получено', 'вовлечение', 'тип поста'],
  
  // Пустая строка (строка 1)
  ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  
  // Заголовок секции (строка 2)
  ['Отзывы', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  
  // Отзыв 1 (строка 3)
  ['основное', 'Яндекс.Маркет', 'Товар А', 'https://market.yandex.ru/review/1', 
   'Отличный товар, рекомендую всем!', 'согласовано', '2024-12-01', 'user123', 
   'Иван Иванов', '150', '280', '130', '4.2', 'ос'],
  
  // Отзыв 2 (строка 4)
  ['основное', 'Ozon', 'Товар Б', 'https://ozon.ru/review/2', 
   'Качество хорошее, доставка быстрая', 'согласовано', '2024-12-02', 'user456', 
   'Петр Петров', '200', '350', '150', '3.8', 'ос'],
  
  // Заголовок секции (строка 5)
  ['Комментарии', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  
  // Комментарий 1 (строка 6)
  ['целевое', 'Яндекс.Маркет', 'Товар А', 'https://market.yandex.ru/comment/1', 
   'Согласен с отзывом, товар действительно хороший', 'согласовано', '2024-12-03', 
   'user789', 'Сидор Сидоров', '50', '80', '30', '2.1', 'цс'],
  
  // Комментарий 2 (строка 7)
  ['целевое', 'Wildberries', 'Товар В', 'https://wildberries.ru/comment/2', 
   'Использую уже месяц, очень доволен', 'согласовано', '2024-12-04', 
   'user101', 'Анна Анненкова', '75', '120', '45', '3.0', 'цс'],
  
  // Заголовок секции (строка 8)
  ['Обсуждения', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  
  // Обсуждение (строка 9) - должно быть пропущено
  ['обсуждение', 'Форум', 'Товар Г', 'https://forum.ru/topic/1', 
   'Интересная тема для обсуждения', 'на модерации', '2024-12-05', 
   'user202', 'Мария Петрова', '100', '150', '50', '1.5', 'пс'],
  
  // Пустая строка (строка 10)
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '']
];

// =============================================================================
// ТЕСТОВЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Главная функция тестирования
 */
function runAllTests() {
  console.log('🧪 Запуск полного тестирования Google Apps Script процессора...');
  
  const testResults = {
    headerDetection: testHeaderDetection(),
    contentTypeDetection: testContentTypeDetection(),
    dataExtraction: testDataExtraction(),
    filtering: testFiltering(),
    fullProcessing: testFullProcessing()
  };
  
  // Подводим итоги
  const passedTests = Object.values(testResults).filter(result => result.passed).length;
  const totalTests = Object.keys(testResults).length;
  
  console.log(`\n📊 Результаты тестирования:`);
  console.log(`✅ Пройдено тестов: ${passedTests}/${totalTests}`);
  console.log(`❌ Провалено тестов: ${totalTests - passedTests}/${totalTests}`);
  
  if (passedTests === totalTests) {
    console.log('🎉 Все тесты пройдены успешно!');
  } else {
    console.log('⚠️ Некоторые тесты не пройдены, требуется доработка');
  }
  
  return testResults;
}

/**
 * Тест определения заголовков
 */
function testHeaderDetection() {
  console.log('\n🔍 Тест: Определение заголовков');
  
  try {
    const headerInfo = findHeaders(mockGoogleSheetsData);
    
    // Проверяем, что заголовки найдены
    const expectedHeaders = ['тип размещения', 'площадка', 'текст сообщения', 'тип поста'];
    const foundHeaders = Object.keys(headerInfo.mapping);
    
    const hasRequiredHeaders = expectedHeaders.every(header => 
      foundHeaders.includes(header)
    );
    
    if (hasRequiredHeaders && headerInfo.row === 0) {
      console.log('✅ Заголовки определены корректно');
      console.log(`   Найдено заголовков: ${foundHeaders.length}`);
      console.log(`   Строка заголовков: ${headerInfo.row}`);
      return { passed: true, message: 'Заголовки найдены' };
    } else {
      console.log('❌ Ошибка определения заголовков');
      console.log(`   Найденные заголовки: ${foundHeaders.join(', ')}`);
      return { passed: false, message: 'Заголовки не найдены' };
    }
    
  } catch (error) {
    console.error('❌ Ошибка в тесте заголовков:', error);
    return { passed: false, message: error.message };
  }
}

/**
 * Тест определения типов контента
 */
function testContentTypeDetection() {
  console.log('\n🔍 Тест: Определение типов контента');
  
  try {
    const testCases = [
      { row: mockGoogleSheetsData[3], expected: 'review', description: 'Отзыв (ос)' },
      { row: mockGoogleSheetsData[4], expected: 'review', description: 'Отзыв (ос)' },
      { row: mockGoogleSheetsData[6], expected: 'comment', description: 'Комментарий (цс)' },
      { row: mockGoogleSheetsData[7], expected: 'comment', description: 'Комментарий (цс)' },
      { row: mockGoogleSheetsData[9], expected: 'unknown', description: 'Обсуждение (пс)' }
    ];
    
    let passed = 0;
    
    for (const testCase of testCases) {
      const result = determineContentType(testCase.row);
      
      if (result === testCase.expected) {
        console.log(`✅ ${testCase.description}: ${result}`);
        passed++;
      } else {
        console.log(`❌ ${testCase.description}: ожидалось ${testCase.expected}, получено ${result}`);
      }
    }
    
    if (passed === testCases.length) {
      return { passed: true, message: 'Все типы определены корректно' };
    } else {
      return { passed: false, message: `Пройдено ${passed}/${testCases.length}` };
    }
    
  } catch (error) {
    console.error('❌ Ошибка в тесте типов контента:', error);
    return { passed: false, message: error.message };
  }
}

/**
 * Тест извлечения данных
 */
function testDataExtraction() {
  console.log('\n🔍 Тест: Извлечение данных');
  
  try {
    const headerInfo = findHeaders(mockGoogleSheetsData);
    const testRow = mockGoogleSheetsData[3]; // Первый отзыв
    
    const extractedData = extractRowData(testRow, 'review', headerInfo.mapping);
    
    if (extractedData && 
        extractedData.site === 'Яндекс.Маркет' &&
        extractedData.text === 'Отличный товар, рекомендую всем!' &&
        extractedData.type === 'review' &&
        extractedData.views === 130 &&
        extractedData.engagement === 4.2) {
      
      console.log('✅ Данные извлечены корректно');
      console.log(`   Площадка: ${extractedData.site}`);
      console.log(`   Текст: ${extractedData.text.substring(0, 30)}...`);
      console.log(`   Просмотры: ${extractedData.views}`);
      console.log(`   Вовлечение: ${extractedData.engagement}`);
      
      return { passed: true, message: 'Данные извлечены корректно' };
    } else {
      console.log('❌ Ошибка извлечения данных');
      console.log(`   Ожидалось: site=Яндекс.Маркет, views=130, engagement=4.2`);
      console.log(`   Получено: site=${extractedData?.site}, views=${extractedData?.views}, engagement=${extractedData?.engagement}`);
      return { passed: false, message: 'Данные не соответствуют ожиданиям' };
    }
    
  } catch (error) {
    console.error('❌ Ошибка в тесте извлечения данных:', error);
    return { passed: false, message: error.message };
  }
}

/**
 * Тест фильтрации строк
 */
function testFiltering() {
  console.log('\n🔍 Тест: Фильтрация строк');
  
  try {
    const testCases = [
      { row: mockGoogleSheetsData[0], isHeader: false, isEmpty: false, description: 'Строка заголовков' },
      { row: mockGoogleSheetsData[1], isHeader: false, isEmpty: true, description: 'Пустая строка' },
      { row: mockGoogleSheetsData[2], isHeader: true, isEmpty: false, description: 'Заголовок секции "Отзывы"' },
      { row: mockGoogleSheetsData[5], isHeader: true, isEmpty: false, description: 'Заголовок секции "Комментарии"' },
      { row: mockGoogleSheetsData[8], isHeader: true, isEmpty: false, description: 'Заголовок секции "Обсуждения"' },
      { row: mockGoogleSheetsData[3], isHeader: false, isEmpty: false, description: 'Строка с данными' }
    ];
    
    let passed = 0;
    
    for (const testCase of testCases) {
      const isEmpty = isEmptyRow(testCase.row);
      const isHeader = isHeaderRow(testCase.row);
      
      if (isEmpty === testCase.isEmpty && isHeader === testCase.isHeader) {
        console.log(`✅ ${testCase.description}: пустая=${isEmpty}, заголовок=${isHeader}`);
        passed++;
      } else {
        console.log(`❌ ${testCase.description}: ожидалось пустая=${testCase.isEmpty}, заголовок=${testCase.isHeader}, получено пустая=${isEmpty}, заголовок=${isHeader}`);
      }
    }
    
    if (passed === testCases.length) {
      return { passed: true, message: 'Фильтрация работает корректно' };
    } else {
      return { passed: false, message: `Пройдено ${passed}/${testCases.length}` };
    }
    
  } catch (error) {
    console.error('❌ Ошибка в тесте фильтрации:', error);
    return { passed: false, message: error.message };
  }
}

/**
 * Тест полной обработки данных
 */
function testFullProcessing() {
  console.log('\n🔍 Тест: Полная обработка данных');
  
  try {
    const processedData = processData(mockGoogleSheetsData);
    
    // Ожидаем: 2 отзыва + 2 комментария = 4 записи
    const expectedTotal = 4;
    const expectedReviews = 2;
    const expectedComments = 2;
    
    const actualTotal = processedData.length;
    const actualReviews = processedData.filter(row => row.type === 'review').length;
    const actualComments = processedData.filter(row => row.type === 'comment').length;
    
    console.log(`📊 Обработано записей: ${actualTotal}`);
    console.log(`   Отзывы: ${actualReviews}`);
    console.log(`   Комментарии: ${actualComments}`);
    
    if (actualTotal === expectedTotal && 
        actualReviews === expectedReviews && 
        actualComments === expectedComments) {
      
      console.log('✅ Полная обработка прошла успешно');
      
      // Показываем примеры обработанных данных
      console.log('\n📋 Примеры обработанных данных:');
      processedData.forEach((row, index) => {
        console.log(`   ${index + 1}. ${row.type === 'review' ? 'Отзыв' : 'Комментарий'}: ${row.site} - ${row.text.substring(0, 30)}...`);
      });
      
      return { passed: true, message: 'Обработка корректна' };
    } else {
      console.log('❌ Ошибка полной обработки');
      console.log(`   Ожидалось: всего=${expectedTotal}, отзывы=${expectedReviews}, комментарии=${expectedComments}`);
      console.log(`   Получено: всего=${actualTotal}, отзывы=${actualReviews}, комментарии=${actualComments}`);
      return { passed: false, message: 'Неверное количество обработанных записей' };
    }
    
  } catch (error) {
    console.error('❌ Ошибка в тесте полной обработки:', error);
    return { passed: false, message: error.message };
  }
}

// =============================================================================
// ЗАПУСК ТЕСТОВ
// =============================================================================

// Автоматический запуск при загрузке
console.log('🚀 Google Apps Script Process Integrated Testing');
console.log('='.repeat(60));
runAllTests();