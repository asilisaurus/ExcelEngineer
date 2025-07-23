/**
 * GOOGLE APPS SCRIPT PROCESSOR V6
 * Auto-selects correct data sheet and handles UI context
 */

// ==================== CONFIGURATION ====================

var CONFIG = {
  STRUCTURE: {
    headerRow: 4,
    dataStartRow: 5,
    infoRows: [1, 2, 3]
  },
  FORMATTING: {
    DATE_FORMAT: 'dd.mm.yyyy',
    NUMBER_FORMAT: '#,##0'
  }
};

// ==================== MAIN FUNCTIONS ====================

/**
 * Main processing function
 */
function processMonthlyReport() {
  try {
    console.log('Starting V6 processor...');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Find the correct data sheet
    var sheet = findDataSheet(spreadsheet);
    if (!sheet) {
      console.error('No data sheet found!');
      showMessage('Error', 'No data sheet found! Please make sure you have a sheet with month data (e.g., Февраль25, Март25, etc.)');
      return;
    }
    
    var sheetName = sheet.getName();
    console.log('Processing sheet: ' + sheetName);
    
    // Get all data
    var data = sheet.getDataRange().getValues();
    console.log('Loaded ' + data.length + ' rows');
    
    // Detect month
    var monthInfo = detectMonth(sheetName);
    console.log('Month: ' + monthInfo.name + ' ' + monthInfo.year);
    
    // Process data
    var result = processDataFlexible(data);
    
    // Create report
    var reportUrl = createReport(result, monthInfo);
    
    // Show success message
    var message = 'Report created: ' + reportUrl + '\n\n' +
                  'Processed:\n' +
                  '- Reviews: ' + result.stats.reviews + '\n' +
                  '- Comments Top-20: ' + result.stats.comments + '\n' +
                  '- Active Discussions: ' + result.stats.discussions;
    
    console.log(message);
    showMessage('Processing completed', message);
    
  } catch (error) {
    console.error('Error: ' + error.toString());
    console.error(error.stack);
    showMessage('Error', 'Error: ' + error.toString());
  }
}

/**
 * Find data sheet (most recent month)
 */
function findDataSheet(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var monthSheets = [];
  
  var monthPatterns = ['январ', 'феврал', 'март', 'апрел', 'май', 'мая', 'июн', 'июл', 'август', 'сентябр', 'октябр', 'ноябр', 'декабр'];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName().toLowerCase();
    
    // Check if this is a month sheet
    for (var j = 0; j < monthPatterns.length; j++) {
      if (sheetName.indexOf(monthPatterns[j]) !== -1) {
        // Extract month number
        var monthNum = j + 1;
        if (monthPatterns[j] === 'май' || monthPatterns[j] === 'мая') monthNum = 5;
        
        monthSheets.push({
          sheet: sheets[i],
          name: sheets[i].getName(),
          monthNumber: monthNum
        });
        break;
      }
    }
  }
  
  // Sort by month number descending and return the most recent
  if (monthSheets.length > 0) {
    monthSheets.sort(function(a, b) { return b.monthNumber - a.monthNumber; });
    console.log('Found month sheets: ' + monthSheets.map(function(s) { return s.name; }).join(', '));
    console.log('Selected most recent: ' + monthSheets[0].name);
    return monthSheets[0].sheet;
  }
  
  // If no month sheets found, try to find a sheet with data
  for (var k = 0; k < sheets.length; k++) {
    var name = sheets[k].getName().toLowerCase();
    if (name.indexOf('инструкция') === -1 && name.indexOf('instruction') === -1) {
      var testData = sheets[k].getDataRange().getValues();
      if (testData.length > 10) {  // Has substantial data
        console.log('Using sheet with data: ' + sheets[k].getName());
        return sheets[k];
      }
    }
  }
  
  return null;
}

/**
 * Show message (handles UI context)
 */
function showMessage(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // If UI not available, just log
    console.log(title + ': ' + message);
  }
}

/**
 * Process data flexibly
 */
function processDataFlexible(data) {
  var allRecords = [];
  var columnMapping = detectColumns(data);
  
  if (!columnMapping) {
    throw new Error('Could not detect column structure');
  }
  
  console.log('Column mapping detected: ' + JSON.stringify(columnMapping));
  
  // Process all data rows
  for (var i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
    var row = data[i];
    
    // Skip empty rows
    if (isEmptyRow(row)) continue;
    
    // Stop at statistics
    if (isStatisticsRow(row)) break;
    
    // Extract data
    var platform = row[columnMapping.platform] ? String(row[columnMapping.platform]).trim() : '';
    var theme = row[columnMapping.theme] ? String(row[columnMapping.theme]).trim() : '';
    var text = row[columnMapping.text] ? String(row[columnMapping.text]).trim() : '';
    var date = formatDate(row[columnMapping.date]);
    var author = row[columnMapping.author] ? String(row[columnMapping.author]).trim() : '';
    var views = parseViews(row[columnMapping.views]);
    var engagement = row[columnMapping.engagement] ? String(row[columnMapping.engagement]).trim() : '';
    var postType = row[columnMapping.postType] ? String(row[columnMapping.postType]).trim().toUpperCase() : '';
    
    // Skip rows without type or content
    if (!postType || (!text && !platform)) continue;
    
    var record = {
      platform: platform,
      theme: theme,
      text: text,
      date: date,
      author: author,
      views: views,
      engagement: engagement,
      postType: postType
    };
    
    allRecords.push(record);
  }
  
  console.log('Total records found: ' + allRecords.length);
  
  // Separate by type
  var reviews = [];
  var targetSiteRecords = [];
  
  for (var j = 0; j < allRecords.length; j++) {
    var rec = allRecords[j];
    if (rec.postType === 'ОС' || rec.postType.indexOf('ОТЗЫВ') !== -1) {
      reviews.push(rec);
    } else if (rec.postType === 'ЦС' || rec.postType.indexOf('КОММЕНТАРИЙ') !== -1 || rec.postType.indexOf('ОБСУЖДЕНИЕ') !== -1) {
      targetSiteRecords.push(rec);
    }
  }
  
  // Sort ЦС records by views
  targetSiteRecords.sort(function(a, b) { return b.views - a.views; });
  
  // Split into Top-20 and discussions
  var comments = [];
  var discussions = [];
  
  for (var k = 0; k < targetSiteRecords.length; k++) {
    if (k < 20) {
      comments.push(targetSiteRecords[k]);
    } else {
      discussions.push(targetSiteRecords[k]);
    }
  }
  
  console.log('Processed: ' + reviews.length + ' reviews, ' + comments.length + ' comments, ' + discussions.length + ' discussions');
  
  return {
    reviews: reviews,
    comments: comments,
    discussions: discussions,
    stats: {
      reviews: reviews.length,
      comments: comments.length,
      discussions: discussions.length,
      totalViews: calculateTotalViews(allRecords)
    }
  };
}

/**
 * Detect columns dynamically
 */
function detectColumns(data) {
  // Try to find header row
  for (var i = 0; i < Math.min(10, data.length); i++) {
    var row = data[i];
    if (!row) continue;
    
    var mapping = {};
    var found = false;
    
    for (var j = 0; j < row.length; j++) {
      var header = String(row[j]).toLowerCase();
      
      if (header.indexOf('площадка') !== -1 || header.indexOf('платформа') !== -1) {
        mapping.platform = j;
        found = true;
      } else if (header.indexOf('тема') !== -1) {
        mapping.theme = j;
      } else if (header.indexOf('текст') !== -1 || header.indexOf('сообщение') !== -1) {
        mapping.text = j;
      } else if (header.indexOf('дата') !== -1) {
        mapping.date = j;
      } else if (header.indexOf('ник') !== -1 || header.indexOf('автор') !== -1) {
        mapping.author = j;
      } else if (header.indexOf('просмотр') !== -1) {
        mapping.views = j;
      } else if (header.indexOf('вовлечение') !== -1 || header.indexOf('ответ') !== -1) {
        mapping.engagement = j;
      } else if (header.indexOf('тип') !== -1 && header.indexOf('поста') !== -1) {
        mapping.postType = j;
      }
    }
    
    if (found && mapping.postType !== undefined) {
      console.log('Found headers at row ' + (i + 1));
      CONFIG.STRUCTURE.headerRow = i + 1;
      CONFIG.STRUCTURE.dataStartRow = i + 2;
      return mapping;
    }
  }
  
  // Fallback to default mapping
  console.log('Using default column mapping');
  return {
    platform: 1,      // B
    theme: 3,         // D
    text: 4,          // E
    date: 6,          // G
    author: 7,        // H
    views: 11,        // L
    engagement: 12,   // M
    postType: 13      // N
  };
}

/**
 * Create report
 */
function createReport(data, monthInfo) {
  var reportName = 'Report_' + monthInfo.name + '_' + monthInfo.year + '_' + new Date().getTime();
  var newSpreadsheet = SpreadsheetApp.create(reportName);
  var sheet = newSpreadsheet.getActiveSheet();
  sheet.setName(monthInfo.name + '_' + monthInfo.year);
  
  // Header
  sheet.getRange('A1').setValue('Продукт');
  sheet.getRange('B1').setValue('Акрихин - Фортедетрим');
  sheet.getRange('A2').setValue('Период');
  sheet.getRange('B2').setValue(monthInfo.name + '-25');
  sheet.getRange('A3').setValue('План');
  
  // Table headers
  var headers = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
  var row = 5;
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#3f2355').setFontColor('white');
  row++;
  
  // Write sections
  row = writeSection(sheet, row, 'Отзывы', data.reviews, headers.length);
  row = writeSection(sheet, row, 'Комментарии Топ-20 выдачи', data.comments, headers.length);
  row = writeSection(sheet, row, 'Активные обсуждения (мониторинг)', data.discussions, headers.length);
  
  // Statistics
  row += 2;
  sheet.getRange(row, 1).setValue('Суммарное количество просмотров');
  sheet.getRange(row, 2).setValue(data.stats.totalViews);
  row++;
  sheet.getRange(row, 1).setValue('Количество карточек товара (отзывы)');
  sheet.getRange(row, 2).setValue(data.stats.reviews);
  row++;
  sheet.getRange(row, 1).setValue('Количество обсуждений (форумы, сообщества, комментарии к статьям)');
  sheet.getRange(row, 2).setValue(data.stats.comments + data.stats.discussions);
  row++;
  sheet.getRange(row, 1).setValue('Доля обсуждений с вовлечением в диалог');
  var engagementRate = calculateEngagementRate(data.comments.concat(data.discussions));
  sheet.getRange(row, 2).setValue(engagementRate);
  sheet.getRange(row, 2).setNumberFormat("0%");
  
  // Format columns
  sheet.autoResizeColumns(1, headers.length);
  
  return newSpreadsheet.getUrl();
}

/**
 * Write section to sheet
 */
function writeSection(sheet, startRow, sectionName, dataArr, colCount) {
  sheet.getRange(startRow, 1).setValue(sectionName);
  sheet.getRange(startRow, 1, 1, colCount).setBackground('#b7a6c9').setFontWeight('bold');
  startRow++;
  
  if (dataArr.length > 0) {
    var rows = [];
    for (var i = 0; i < dataArr.length; i++) {
      var r = dataArr[i];
      rows.push([
        r.platform,
        r.theme,
        r.text,
        r.date,
        r.author,
        r.views,
        r.engagement,
        r.postType
      ]);
    }
    sheet.getRange(startRow, 1, rows.length, colCount).setValues(rows);
    startRow += rows.length;
  }
  
  console.log('Section "' + sectionName + '": ' + dataArr.length + ' rows');
  return startRow;
}

// ==================== HELPER FUNCTIONS ====================

function detectMonth(sheetName) {
  var monthsMap = {
    'янв': { name: 'Январь', number: 1 },
    'фев': { name: 'Февраль', number: 2 },
    'мар': { name: 'Март', number: 3 },
    'апр': { name: 'Апрель', number: 4 },
    'май': { name: 'Май', number: 5 },
    'мая': { name: 'Май', number: 5 },
    'июн': { name: 'Июнь', number: 6 },
    'июл': { name: 'Июль', number: 7 },
    'авг': { name: 'Август', number: 8 },
    'сен': { name: 'Сентябрь', number: 9 },
    'окт': { name: 'Октябрь', number: 10 },
    'ноя': { name: 'Ноябрь', number: 11 },
    'дек': { name: 'Декабрь', number: 12 }
  };
  
  var lowerName = sheetName.toLowerCase();
  for (var key in monthsMap) {
    if (lowerName.indexOf(key) !== -1) {
      return {
        name: monthsMap[key].name,
        number: monthsMap[key].number,
        year: 2025
      };
    }
  }
  
  // Default
  var now = new Date();
  var monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  return {
    name: monthNames[now.getMonth()],
    number: now.getMonth() + 1,
    year: now.getFullYear()
  };
}

function isEmptyRow(row) {
  if (!row) return true;
  for (var i = 0; i < row.length; i++) {
    if (row[i] && String(row[i]).trim() !== '') return false;
  }
  return true;
}

function isStatisticsRow(row) {
  if (!row || row.length === 0) return false;
  var firstCell = String(row[0] || '').toLowerCase();
  return firstCell.indexOf('суммарное количество просмотров') !== -1 || 
         firstCell.indexOf('количество карточек товара') !== -1 ||
         firstCell.indexOf('количество обсуждений') !== -1;
}

function formatDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), CONFIG.FORMATTING.DATE_FORMAT);
  }
  return String(value);
}

function parseViews(value) {
  if (!value) return 0;
  if (typeof value === 'number') return Math.floor(value);
  
  var str = String(value).replace(/[^\d]/g, '');
  var num = parseInt(str);
  return isNaN(num) ? 0 : num;
}

function calculateTotalViews(records) {
  var total = 0;
  for (var i = 0; i < records.length; i++) {
    total += records[i].views;
  }
  return total;
}

function calculateEngagementRate(discussions) {
  if (discussions.length === 0) return 0;
  var withEngagement = 0;
  for (var i = 0; i < discussions.length; i++) {
    var eng = discussions[i].engagement;
    if (eng && eng !== '0' && eng !== '') {
      withEngagement++;
    }
  }
  return withEngagement / discussions.length;
}

/**
 * Create menu
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Processing V6')
    .addItem('🚀 Process month (auto-select)', 'processMonthlyReport')
    .addItem('📋 Show available sheets', 'showSheets')
    .addToUi();
}

/**
 * Show available sheets
 */
function showSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var message = 'Available sheets:\n\n';
  
  for (var i = 0; i < sheets.length; i++) {
    message += (i + 1) + '. ' + sheets[i].getName() + '\n';
  }
  
  showMessage('Sheets', message);
}