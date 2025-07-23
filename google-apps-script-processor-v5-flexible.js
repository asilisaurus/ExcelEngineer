/**
 * 🚀 GOOGLE APPS SCRIPT PROCESSOR V5 - FLEXIBLE
 * Handles ANY amount of data for future months
 * 
 * Author: AI Assistant
 * Date: 2025
 */

// ==================== CONFIGURATION ====================

const CONFIG = {
  // Data structure
  STRUCTURE: {
    headerRow: 4,        // Headers in row 4
    dataStartRow: 5,     // Data starts from row 5
    infoRows: [1, 2, 3], // Meta-information in rows 1-3
  },
  
  // Formatting
  FORMATTING: {
    DATE_FORMAT: 'dd.mm.yyyy',
    NUMBER_FORMAT: '#,##0'
  }
};

// ==================== MAIN PROCESSOR CLASS ====================

class ProcessorV5 {
  constructor() {
    this.stats = {
      totalRows: 0,
      reviewsCount: 0,
      commentsCount: 0,
      discussionsCount: 0,
      totalViews: 0,
      processingTime: 0,
      errors: []
    };
    
    this.monthInfo = null;
  }

  /**
   * Main processing method
   */
  processReport(spreadsheetId, sheetName = null) {
    const startTime = Date.now();
    
    try {
      console.log('🚀 PROCESSOR V5 - Starting flexible processing');
      
      // 1. Get data
      const sourceData = this.getSourceData(spreadsheetId, sheetName);
      
      // 2. Detect month
      this.monthInfo = this.detectMonth(sourceData, sheetName);
      console.log(`📅 Month: ${this.monthInfo.name} ${this.monthInfo.year}`);
      
      // 3. Process data dynamically
      const processedData = this.processDataFlexible(sourceData);
      
      // 4. Create report
      const reportUrl = this.createReport(processedData);
      
      // 5. Update statistics
      this.stats.processingTime = Date.now() - startTime;
      
      console.log('✅ Processing completed successfully');
      console.log(`📊 Results: ${this.stats.reviewsCount} reviews, ${this.stats.commentsCount} comments, ${this.stats.discussionsCount} discussions`);
      
      return {
        success: true,
        reportUrl: reportUrl,
        statistics: this.stats,
        monthInfo: this.monthInfo
      };
      
    } catch (error) {
      console.error('❌ Processing error:', error);
      this.stats.errors.push(error.toString());
      
      return {
        success: false,
        error: error.toString(),
        statistics: this.stats
      };
    }
  }

  /**
   * Get source data
   */
  getSourceData(spreadsheetId, sheetName) {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    this.stats.totalRows = data.length;
    
    console.log(`📊 Loaded ${data.length} rows of data`);
    
    return data;
  }

  /**
   * Detect month from sheet name or data
   */
  detectMonth(data, sheetName = null) {
    if (sheetName) {
      const monthMatch = sheetName.match(/(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)/i);
      if (monthMatch) {
        const monthsMap = {
          'янв': { name: 'Январь', number: 1 },
          'фев': { name: 'Февраль', number: 2 },
          'мар': { name: 'Март', number: 3 },
          'апр': { name: 'Апрель', number: 4 },
          'май': { name: 'Май', number: 5 },
          'июн': { name: 'Июнь', number: 6 },
          'июл': { name: 'Июль', number: 7 },
          'авг': { name: 'Август', number: 8 },
          'сен': { name: 'Сентябрь', number: 9 },
          'окт': { name: 'Октябрь', number: 10 },
          'ноя': { name: 'Ноябрь', number: 11 },
          'дек': { name: 'Декабрь', number: 12 }
        };
        
        const monthKey = monthMatch[1].toLowerCase();
        const month = monthsMap[monthKey];
        
        return {
          name: month.name,
          number: month.number,
          year: 2025
        };
      }
    }
    
    // Default to current month
    const now = new Date();
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    
    return {
      name: monthNames[now.getMonth()],
      number: now.getMonth() + 1,
      year: now.getFullYear()
    };
  }

  /**
   * Process data flexibly without fixed limits
   */
  processDataFlexible(data) {
    // Arrays to hold different types of records
    const allRecords = [];
    
    // Get column mapping
    const columnMapping = this.getColumnMapping();
    
    // Process all data rows
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row || row.every(cell => !cell || String(cell).trim() === '')) continue;
      
      // Stop at statistics
      if (this.isStatisticsRow(row)) break;
      
      // Extract data
      const platform = row[columnMapping.platform] ? String(row[columnMapping.platform]).trim() : '';
      const theme = row[columnMapping.theme] ? String(row[columnMapping.theme]).trim() : '';
      const text = row[columnMapping.text] ? String(row[columnMapping.text]).trim() : '';
      const date = this.formatDate(row[columnMapping.date]);
      const author = row[columnMapping.author] ? String(row[columnMapping.author]).trim() : '';
      const views = this.parseViews(row[columnMapping.views]);
      const engagement = row[columnMapping.engagement] ? String(row[columnMapping.engagement]).trim() : '';
      const postType = row[columnMapping.postType] ? String(row[columnMapping.postType]).trim().toUpperCase() : '';
      
      // Skip rows without type or content
      if (!postType || (!text && !platform)) continue;
      
      const record = {
        platform,
        theme,
        text,
        date,
        author,
        views,
        engagement,
        postType,
        originalType: postType
      };
      
      allRecords.push(record);
    }
    
    console.log(`📊 Total records found: ${allRecords.length}`);
    
    // Separate records by type
    const reviews = [];
    const targetSiteRecords = []; // All ЦС records
    
    // First pass: separate reviews (ОС) and target site records (ЦС)
    for (const record of allRecords) {
      if (record.postType === 'ОС' || record.postType.includes('ОТЗЫВ')) {
        reviews.push(record);
      } else if (record.postType === 'ЦС' || 
                 record.postType.includes('КОММЕНТАРИЙ') || 
                 record.postType.includes('ОБСУЖДЕНИЕ')) {
        targetSiteRecords.push(record);
      }
    }
    
    // Sort ЦС records by views (descending) to identify Top-20
    targetSiteRecords.sort((a, b) => b.views - a.views);
    
    // Split ЦС records:
    // - Top-20: the 20 most viewed ЦС records
    // - Active discussions: all remaining ЦС records
    const comments = [];
    const discussions = [];
    
    for (let i = 0; i < targetSiteRecords.length; i++) {
      if (i < 20) {
        // First 20 most viewed records are Top-20 comments
        comments.push(targetSiteRecords[i]);
      } else {
        // Rest are active discussions
        discussions.push(targetSiteRecords[i]);
      }
    }
    
    // Update statistics
    this.stats.reviewsCount = reviews.length;
    this.stats.commentsCount = comments.length;
    this.stats.discussionsCount = discussions.length;
    
    // Extract total views from statistics section
    const stats = this.extractStatistics(data);
    this.stats.totalViews = stats.totalViews;
    
    // If no total views found in statistics, calculate from data
    if (!this.stats.totalViews) {
      this.stats.totalViews = allRecords.reduce((sum, record) => sum + record.views, 0);
    }
    
    console.log(`📊 Processed: ${reviews.length} reviews, ${comments.length} comments, ${discussions.length} discussions`);
    console.log(`👁️ Top comment views: ${comments[0]?.views || 0}, Last comment views: ${comments[comments.length-1]?.views || 0}`);
    
    return {
      reviews,
      commentsTop20: comments,
      activeDiscussions: discussions,
      statistics: {
        totalViews: this.stats.totalViews,
        totalReviews: reviews.length,
        totalComments: comments.length,
        totalDiscussions: discussions.length
      }
    };
  }

  /**
   * Column mapping
   */
  getColumnMapping() {
    return {
      platform: 1,      // B - Platform
      theme: 3,         // D - Theme
      text: 4,          // E - Text
      date: 6,          // G - Date
      author: 7,        // H - Nick
      views: 11,        // L - Views
      engagement: 12,   // M - Engagement
      postType: 13      // N - Post Type
    };
  }

  /**
   * Extract statistics from data
   */
  extractStatistics(data) {
    let totalViews = 0;
    
    // Search for statistics in last rows
    for (let i = data.length - 20; i < data.length; i++) {
      if (i < 0) continue;
      
      const row = data[i];
      if (!row) continue;
      
      const firstCell = String(row[0] || '').toLowerCase();
      
      if (firstCell.includes('суммарное количество просмотров')) {
        for (let j = 1; j < row.length; j++) {
          const value = parseFloat(String(row[j] || '').replace(/[^\d]/g, ''));
          if (!isNaN(value) && value > 0) {
            totalViews = Math.floor(value);
            console.log(`📊 Found total views in statistics: ${totalViews}`);
            break;
          }
        }
        break;
      }
    }
    
    return { totalViews };
  }

  /**
   * Create report
   */
  createReport(processedData) {
    const tempSpreadsheet = SpreadsheetApp.create(`temp_google_sheets_${Date.now()}_${this.monthInfo.name}_${this.monthInfo.year}_результат`);
    const reportSheetName = `${this.monthInfo.name}_${this.monthInfo.year}`;
    const sheet = tempSpreadsheet.getActiveSheet();
    sheet.setName(reportSheetName);
    
    // Header
    sheet.getRange('A1').setValue('Продукт');
    sheet.getRange('B1').setValue('Акрихин - Фортедетрим');
    sheet.getRange('A2').setValue('Период');
    sheet.getRange('B2').setValue(`${this.monthInfo.name}-25`);
    sheet.getRange('A3').setValue('План');
    
    // Table headers
    const tableHeaders = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
    let row = 5;
    sheet.getRange(row, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(row, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#3f2355').setFontColor('white');
    row++;
    
    // Function to write section
    const writeSection = (sectionName, dataArr) => {
      sheet.getRange(row, 1).setValue(sectionName);
      sheet.getRange(row, 1, 1, tableHeaders.length).setBackground('#b7a6c9').setFontWeight('bold');
      row++;
      
      if (dataArr.length > 0) {
        const rows = dataArr.map(r => [
          r.platform,
          r.theme,
          r.text,
          r.date,
          r.author,
          r.views,
          r.engagement,
          r.originalType || r.postType
        ]);
        sheet.getRange(row, 1, rows.length, tableHeaders.length).setValues(rows);
        row += rows.length;
      }
      
      console.log(`📂 Section "${sectionName}": ${dataArr.length} rows`);
    };
    
    // Write data sections
    writeSection('Отзывы', processedData.reviews);
    writeSection('Комментарии Топ-20 выдачи', processedData.commentsTop20);
    writeSection('Активные обсуждения (мониторинг)', processedData.activeDiscussions);
    
    // Statistics
    row += 2;
    sheet.getRange(row, 1).setValue('Суммарное количество просмотров');
    sheet.getRange(row, 2).setValue(processedData.statistics.totalViews || 0);
    row++;
    sheet.getRange(row, 1).setValue('Количество карточек товара (отзывы)');
    sheet.getRange(row, 2).setValue(processedData.statistics.totalReviews || 0);
    row++;
    sheet.getRange(row, 1).setValue('Количество обсуждений (форумы, сообщества, комментарии к статьям)');
    sheet.getRange(row, 2).setValue((processedData.statistics.totalComments || 0) + (processedData.statistics.totalDiscussions || 0));
    row++;
    sheet.getRange(row, 1).setValue('Доля обсуждений с вовлечением в диалог');
    
    // Calculate engagement percentage for all discussions (comments + active)
    let engagementPercentage = 0;
    const allDiscussions = [...processedData.commentsTop20, ...processedData.activeDiscussions];
    if (allDiscussions.length > 0) {
      const withEngagement = allDiscussions.filter(d => d.engagement && d.engagement !== '0' && d.engagement !== '').length;
      engagementPercentage = withEngagement / allDiscussions.length;
    }
    sheet.getRange(row, 2).setValue(engagementPercentage);
    sheet.getRange(row, 2).setNumberFormat("0%");
    
    // Format columns
    sheet.autoResizeColumns(1, tableHeaders.length);
    
    console.log(`📄 Report created: ${reportSheetName}`);
    return tempSpreadsheet.getUrl();
  }

  // ==================== HELPER METHODS ====================

  isStatisticsRow(row) {
    if (!row || row.length === 0) return false;
    const firstCell = String(row[0] || '').toLowerCase();
    return firstCell.includes('суммарное количество просмотров') || 
           firstCell.includes('количество карточек товара') ||
           firstCell.includes('количество обсуждений');
  }

  formatDate(value) {
    if (!value) return '';
    if (value instanceof Date) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), CONFIG.FORMATTING.DATE_FORMAT);
    }
    return String(value);
  }

  parseViews(value) {
    if (!value) return 0;
    if (typeof value === 'number') return Math.floor(value);
    
    const str = String(value).replace(/[^\d]/g, '');
    const num = parseInt(str);
    return isNaN(num) ? 0 : num;
  }
}

// ==================== GOOGLE APPS SCRIPT FUNCTIONS ====================

/**
 * Main function to run
 */
function processMonthlyReport() {
  const processor = new ProcessorV5();
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  
  const result = processor.processReport(spreadsheetId, sheetName);
  
  if (result.success) {
    SpreadsheetApp.getUi().alert(
      'Processing completed',
      `Report created: ${result.reportUrl}\n\n` +
      `Processed:\n` +
      `- Reviews: ${result.statistics.reviewsCount}\n` +
      `- Comments Top-20: ${result.statistics.commentsCount}\n` +
      `- Active Discussions: ${result.statistics.discussionsCount}\n` +
      `- Total Views: ${result.statistics.totalViews}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert('Error', result.error, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create menu on open
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Processing V5')
    .addItem('🚀 Process month', 'processMonthlyReport')
    .addToUi();
}