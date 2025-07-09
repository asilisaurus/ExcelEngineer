/**
 * FORTEDETRIM ORM PROCESSOR для Google Sheets
 * Автоматическая обработка ежемесячных отчетов
 */

const CONFIG = {
  MONTHS: {
    'январь': 'Январь', 'февраль': 'Февраль', 'март': 'Март',
    'апрель': 'Апрель', 'май': 'Май', 'июнь': 'Июнь',
    'июль': 'Июль', 'август': 'Август', 'сентябрь': 'Сентябрь',
    'октябрь': 'Октябрь', 'ноябрь': 'Ноябрь', 'декабрь': 'Декабрь'
  },
  QUALITY: {
    MIN_TEXT_LENGTH: 20,
    MIN_PLATFORM_LENGTH: 3,
    HIGH_QUALITY_THRESHOLD: 85
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Fortedetrim ORM')
    .addItem('📊 Создать отчет', 'createReport')
    .addItem('🔍 Анализ качества', 'analyzeQuality')
    .addItem('ℹ️ Справка', 'showHelp')
    .addToUi();
}

function createReport() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const monthResponse = ui.prompt(
      'Создание отчета',
      'Введите месяц (например: июнь):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (monthResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const targetMonth = monthResponse.getResponseText().toLowerCase().trim();
    
    const sheetResponse = ui.prompt(
      'Выбор данных',
      'Введите название листа с данными:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (sheetResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const sourceSheetName = sheetResponse.getResponseText().trim();
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName);
    
    if (!sourceSheet) {
      ui.alert('Ошибка', `Лист "${sourceSheetName}" не найден`, ui.ButtonSet.OK);
      return;
    }
    
    const processedData = processData(sourceSheet, targetMonth);
    const reportSheet = createReportSheet(processedData, targetMonth);
    const quality = analyzeDataQuality(processedData);
    
    showResults(reportSheet, quality, processedData);
    
  } catch (error) {
    ui.alert('Ошибка', `Не удалось создать отчет: ${error.message}`, ui.ButtonSet.OK);
  }
}

function processData(sourceSheet, targetMonth) {
  const data = sourceSheet.getDataRange().getValues();
  const headers = findHeaders(data);
  
  if (headers.rowIndex === -1) {
    throw new Error('Не найдены заголовки');
  }
  
  const columnMapping = mapColumns(headers.columns);
  const extractedData = extractData(data, headers.rowIndex, columnMapping);
  const categorizedData = categorizeData(extractedData);
  
  return {
    month: CONFIG.MONTHS[targetMonth] || targetMonth,
    reviews: categorizedData.reviews,
    comments: categorizedData.comments,
    discussions: categorizedData.discussions,
    total: extractedData.length
  };
}

function findHeaders(data) {
  for (let i = 0; i < Math.min(15, data.length); i++) {
    const rowText = data[i].join(' ').toLowerCase();
    if (rowText.includes('площадка') || rowText.includes('текст') || rowText.includes('ссылка')) {
      return { rowIndex: i, columns: data[i] };
    }
  }
  return { rowIndex: -1, columns: [] };
}

function mapColumns(headers) {
  const mapping = {};
  headers.forEach((header, index) => {
    const h = header.toString().toLowerCase();
    if (h.includes('площадка')) mapping.platform = index;
    if (h.includes('ссылка')) mapping.link = index;
    if (h.includes('текст')) mapping.text = index;
    if (h.includes('дата')) mapping.date = index;
    if (h.includes('ник') || h.includes('автор')) mapping.author = index;
    if (h.includes('просмотр')) mapping.views = index;
    if (h.includes('вовлечение')) mapping.engagement = index;
  });
  return mapping;
}

function extractData(data, startRow, mapping) {
  const results = [];
  
  for (let i = startRow + 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => !cell || cell.toString().trim() === '')) continue;
    
    const item = {
      platform: getValue(row[mapping.platform]),
      link: getValue(row[mapping.link]),
      text: getValue(row[mapping.text]),
      date: getValue(row[mapping.date]),
      author: getValue(row[mapping.author]),
      views: getValue(row[mapping.views]),
      engagement: getValue(row[mapping.engagement])
    };
    
    if (item.platform && item.text && item.text.length > 10) {
      results.push(item);
    }
  }
  
  return results;
}

function getValue(cell) {
  if (cell === null || cell === undefined) return '';
  return cell.toString().trim();
}

function categorizeData(data) {
  const reviews = [];
  const comments = [];
  const discussions = [];
  
  data.forEach(item => {
    const text = item.text.toLowerCase();
    
    if (text.includes('отзыв') || text.includes('рекомендую') || text.includes('покупал')) {
      reviews.push(item);
    } else if (text.includes('комментарий') || text.includes('вопрос') || text.includes('спрашива')) {
      comments.push(item);
    } else {
      discussions.push(item);
    }
  });
  
  return { reviews, comments, discussions };
}

function createReportSheet(data, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Fortedetrim ORM report ${month} 2025 result`;
  
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  
  const sheet = ss.insertSheet(sheetName);
  fillSheet(sheet, data);
  
  return sheet;
}

function fillSheet(sheet, data) {
  let row = 1;
  
  // Заголовки
  const headers = ['Площадка', 'Тема', 'Текст', 'Дата', 'Ник', 'Просмотры', 'Вовлечение'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  row++;
  
  // Отзывы
  if (data.reviews.length > 0) {
    sheet.getRange(row, 1).setValue('Отзывы').setFontWeight('bold').setBackground('#E8F4FD');
    row++;
    
    data.reviews.forEach(item => {
      const rowData = [item.platform, item.link, item.text, item.date, item.author, item.views || 'Нет данных', item.engagement || 'Нет данных'];
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
    row++;
  }
  
  // Комментарии
  if (data.comments.length > 0) {
    sheet.getRange(row, 1).setValue('Комментарии').setFontWeight('bold').setBackground('#FFF2CC');
    row++;
    
    data.comments.forEach(item => {
      const rowData = [item.platform, item.link, item.text, item.date, item.author, item.views || 'Нет данных', item.engagement || 'Нет данных'];
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
    row++;
  }
  
  // Обсуждения
  if (data.discussions.length > 0) {
    sheet.getRange(row, 1).setValue('Активные обсуждения').setFontWeight('bold').setBackground('#FCE5CD');
    row++;
    
    data.discussions.forEach(item => {
      const rowData = [item.platform, item.link, item.text, item.date, item.author, item.views || 'Нет данных', item.engagement || 'Нет данных'];
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
  }
  
  // Форматирование
  sheet.autoResizeColumns(1, 7);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 3, sheet.getLastRow(), 1).setWrap(true);
}

function analyzeDataQuality(data) {
  const allItems = [...data.reviews, ...data.comments, ...data.discussions];
  let totalScore = 0;
  let issues = 0;
  
  allItems.forEach(item => {
    let score = 100;
    
    if (!item.platform || item.platform.length < CONFIG.QUALITY.MIN_PLATFORM_LENGTH) {
      score -= 25;
      issues++;
    }
    
    if (!item.text || item.text.length < CONFIG.QUALITY.MIN_TEXT_LENGTH) {
      score -= 30;
      issues++;
    }
    
    if (!item.link || !item.link.includes('http')) {
      score -= 20;
      issues++;
    }
    
    if (!item.date) {
      score -= 15;
      issues++;
    }
    
    totalScore += Math.max(score, 0);
  });
  
  const avgScore = allItems.length > 0 ? totalScore / allItems.length : 0;
  
  return {
    score: Math.round(avgScore),
    issues: issues,
    total: allItems.length,
    grade: getGrade(avgScore)
  };
}

function getGrade(score) {
  if (score >= 95) return 'Отличное качество';
  if (score >= 85) return 'Высокое качество';
  if (score >= 70) return 'Хорошее качество';
  if (score >= 50) return 'Удовлетворительное качество';
  return 'Требует улучшения';
}

function showResults(sheet, quality, data) {
  const ui = SpreadsheetApp.getUi();
  
  const message = `
📊 Отчет создан: ${sheet.getName()}
📈 Всего записей: ${data.total}
📝 Отзывы: ${data.reviews.length}
💬 Комментарии: ${data.comments.length}
🗣️ Обсуждения: ${data.discussions.length}
🎯 Качество: ${quality.grade} (${quality.score}%)
⚠️ Проблемы: ${quality.issues}
  `;
  
  ui.alert('Результат', message, ui.ButtonSet.OK);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
}

function analyzeQuality() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const stats = {
    totalRows: data.length,
    emptyRows: 0,
    incompleteRows: 0
  };
  
  data.forEach((row, index) => {
    if (index === 0) return;
    
    const isEmpty = row.every(cell => !cell || cell.toString().trim() === '');
    const isIncomplete = row.some(cell => !cell || cell.toString().trim() === '');
    
    if (isEmpty) stats.emptyRows++;
    if (isIncomplete && !isEmpty) stats.incompleteRows++;
  });
  
  const completeness = Math.round(((stats.totalRows - stats.emptyRows - stats.incompleteRows) / stats.totalRows) * 100);
  
  ui.alert(
    'Анализ качества',
    `Всего строк: ${stats.totalRows}\nПустые: ${stats.emptyRows}\nНеполные: ${stats.incompleteRows}\nПолнота: ${completeness}%`,
    ui.ButtonSet.OK
  );
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Справка - Fortedetrim ORM Processor',
    `
🎯 Создание качественных ежемесячных отчетов

📊 Функции:
• Автоматическая обработка данных
• Категоризация контента
• Контроль качества
• Создание готового отчета

🔧 Использование:
1. Подготовьте данные с заголовками
2. Запустите "Создать отчет"
3. Укажите месяц и лист
4. Получите готовый отчет

📋 Требования:
• Площадка (мин. 3 символа)
• Текст (мин. 20 символов)
• Ссылка (с http)
• Дата в формате дд.мм.гггг
    `,
    ui.ButtonSet.OK
  );
}
