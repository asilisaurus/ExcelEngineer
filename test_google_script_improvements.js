/**
 * ============================================================================
 * ТЕСТОВЫЙ СКРИПТ ДЛЯ ВАЛИДАЦИИ GOOGLE APPS SCRIPT ПРОЦЕССОРА
 * ============================================================================
 * Этот скрипт предназначен для тестирования исправлений в Google Apps Script
 * ============================================================================
 */

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
      return { passed: true, message: 'Заголовки найдены' };
    } else {
      console.log('❌ Ошибка определения заголовков');
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
// ДОПОЛНИТЕЛЬНЫЕ ТЕСТОВЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Тест производительности
 */
function performanceTest() {
  console.log('\n⏱️ Тест производительности');
  
  try {
    const start = Date.now();
    
    // Создаем большой массив данных для тестирования
    const largeDataset = [];
    for (let i = 0; i < 1000; i++) {
      largeDataset.push(...mockGoogleSheetsData);
    }
    
    const result = processData(largeDataset);
    
    const end = Date.now();
    const processingTime = end - start;
    
    console.log(`⏱️ Время обработки ${largeDataset.length} строк: ${processingTime}мс`);
    console.log(`📊 Обработано записей: ${result.length}`);
    console.log(`🚀 Производительность: ${Math.round(largeDataset.length / processingTime * 1000)} строк/сек`);
    
    return { time: processingTime, processed: result.length };
    
  } catch (error) {
    console.error('❌ Ошибка в тесте производительности:', error);
    return { error: error.message };
  }
}

/**
 * Тест граничных случаев
 */
function edgeCaseTests() {
  console.log('\n🔬 Тест граничных случаев');
  
  const testCases = [
    {
      name: 'Пустой массив',
      data: [],
      shouldFail: true
    },
    {
      name: 'Только заголовки',
      data: [mockGoogleSheetsData[0]],
      shouldFail: false
    },
    {
      name: 'Данные без заголовков',
      data: [mockGoogleSheetsData[3], mockGoogleSheetsData[4]],
      shouldFail: false
    },
    {
      name: 'Только пустые строки',
      data: [['', '', '', ''], ['', '', '', '']],
      shouldFail: false
    }
  ];
  
  let passed = 0;
  
  for (const testCase of testCases) {
    try {
      const result = processData(testCase.data);
      
      if (testCase.shouldFail) {
        console.log(`❌ ${testCase.name}: ожидалась ошибка, но получен результат`);
      } else {
        console.log(`✅ ${testCase.name}: обработано ${result.length} записей`);
        passed++;
      }
      
    } catch (error) {
      if (testCase.shouldFail) {
        console.log(`✅ ${testCase.name}: ожидаемая ошибка - ${error.message}`);
        passed++;
      } else {
        console.log(`❌ ${testCase.name}: неожиданная ошибка - ${error.message}`);
      }
    }
  }
  
  console.log(`📊 Граничные случаи: ${passed}/${testCases.length} пройдено`);
  return { passed: passed, total: testCases.length };
}

// =============================================================================
// ЗАПУСК ТЕСТОВ
// =============================================================================

// Для запуска в Google Apps Script
function runGoogleAppsScriptTests() {
  console.log('🚀 Запуск тестов в Google Apps Script...');
  return runAllTests();
}

// Для запуска в Node.js
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    runAllTests,
    performanceTest,
    edgeCaseTests
  };
}

// Автоматический запуск при загрузке
if (typeof window === 'undefined' && typeof global !== 'undefined') {
  // Node.js environment
  console.log('🧪 Автоматический запуск тестов...');
  runAllTests();
}