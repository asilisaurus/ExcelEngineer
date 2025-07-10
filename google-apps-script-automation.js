/**
 * 🤖 АВТОМАТИЗАЦИЯ И ДОПОЛНИТЕЛЬНЫЕ ФУНКЦИИ
 * Расширенные возможности для Google Apps Script
 * 
 * Автор: AI Assistant + Background Agent
 * Версия: 1.0.0
 * Дата: 2025
 */

// ==================== АВТОМАТИЗАЦИЯ ====================

/**
 * Класс для автоматической обработки
 */
class AutomatedProcessor {
  constructor() {
    this.triggers = [];
    this.settings = this.loadSettings();
  }

  /**
   * Настройка автоматической обработки
   */
  setupAutomation() {
    console.log('🤖 Настройка автоматизации...');
    
    // Удаляем старые триггеры
    this.removeAllTriggers();
    
    // Создаем новые триггеры
    this.createTriggers();
    
    console.log('✅ Автоматизация настроена');
  }

  /**
   * Создание триггеров
   */
  createTriggers() {
    // Триггер при открытии таблицы
    ScriptApp.newTrigger('onOpen')
      .onOpen()
      .create();
    
    // Триггер по времени (ежедневно в 9:00)
    ScriptApp.newTrigger('dailyProcessing')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    
    // Триггер при изменении данных
    ScriptApp.newTrigger('onDataChange')
      .onEdit()
      .create();
    
    console.log('📅 Триггеры созданы');
  }

  /**
   * Удаление всех триггеров
   */
  removeAllTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });
    console.log('🗑️ Старые триггеры удалены');
  }

  /**
   * Ежедневная обработка
   */
  dailyProcessing() {
    console.log('📅 Запуск ежедневной обработки...');
    
    try {
      // Проверяем, есть ли новые данные
      const hasNewData = this.checkForNewData();
      
      if (hasNewData) {
        console.log('🆕 Обнаружены новые данные, запускаем обработку...');
        processMonthlyReport();
      } else {
        console.log('ℹ️ Новых данных не обнаружено');
      }
    } catch (error) {
      console.error('❌ Ошибка ежедневной обработки:', error);
      this.sendErrorNotification(error);
    }
  }

  /**
   * Обработка при изменении данных
   */
  onDataChange(e) {
    console.log('📝 Обнаружено изменение данных...');
    
    // Проверяем, что изменение в нужном диапазоне
    const range = e.range;
    const sheet = range.getSheet();
    
    // Если изменение в исходных данных
    if (sheet.getName().includes('Исходные данные') || sheet.getName().includes('Source')) {
      console.log('🔄 Изменение в исходных данных, запускаем обработку...');
      
      // Небольшая задержка для завершения изменений
      Utilities.sleep(2000);
      processMonthlyReport();
    }
  }

  /**
   * Проверка новых данных
   */
  checkForNewData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    for (const sheet of sheets) {
      const lastRow = sheet.getLastRow();
      const lastModified = sheet.getRange(lastRow, 1).getLastModified();
      const now = new Date();
      
      // Если данные были изменены в последние 24 часа
      if (now.getTime() - lastModified.getTime() < 24 * 60 * 60 * 1000) {
        return true;
      }
    }
    
    return false;
  }

  /**
   * Отправка уведомления об ошибке
   */
  sendErrorNotification(error) {
    try {
      const email = Session.getActiveUser().getEmail();
      const subject = '❌ Ошибка автоматической обработки отчетов';
      const body = `
        Произошла ошибка при автоматической обработке отчетов:
        
        Ошибка: ${error.toString()}
        Время: ${new Date().toLocaleString('ru-RU')}
        Пользователь: ${Session.getActiveUser().getEmail()}
        
        Проверьте логи в Google Apps Script для подробностей.
      `;
      
      MailApp.sendEmail(email, subject, body);
      console.log('📧 Уведомление об ошибке отправлено');
    } catch (emailError) {
      console.error('❌ Не удалось отправить уведомление:', emailError);
    }
  }

  /**
   * Загрузка настроек
   */
  loadSettings() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const settings = properties.getProperty('processor_settings');
      return settings ? JSON.parse(settings) : this.getDefaultSettings();
    } catch (error) {
      console.warn('⚠️ Ошибка загрузки настроек, используем по умолчанию:', error);
      return this.getDefaultSettings();
    }
  }

  /**
   * Сохранение настроек
   */
  saveSettings(settings) {
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty('processor_settings', JSON.stringify(settings));
      console.log('💾 Настройки сохранены');
    } catch (error) {
      console.error('❌ Ошибка сохранения настроек:', error);
    }
  }

  /**
   * Настройки по умолчанию
   */
  getDefaultSettings() {
    return {
      autoProcessing: true,
      dailyCheck: true,
      emailNotifications: true,
      maxRows: 10000,
      qualityThreshold: 80,
      platforms: ['VK', 'Telegram', 'Instagram', 'YouTube']
    };
  }
}

// ==================== РАСШИРЕННЫЕ ФУНКЦИИ ====================

/**
 * Класс для работы с данными
 */
class DataAnalyzer {
  constructor() {
    this.cache = {};
  }

  /**
   * Анализ качества данных
   */
  analyzeDataQuality(data) {
    console.log('🔍 Анализ качества данных...');
    
    const quality = {
      totalRows: data.length,
      emptyRows: 0,
      incompleteRows: 0,
      duplicateRows: 0,
      qualityScore: 0,
      issues: []
    };

    const seenRows = new Set();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Проверка пустых строк
      if (this.isEmptyRow(row)) {
        quality.emptyRows++;
        quality.issues.push(`Строка ${i + 1}: пустая строка`);
        continue;
      }
      
      // Проверка неполных строк
      const requiredColumns = ['PLATFORM', 'TEXT', 'DATE'];
      const missingColumns = requiredColumns.filter(col => !this.hasValidData(row, col));
      
      if (missingColumns.length > 0) {
        quality.incompleteRows++;
        quality.issues.push(`Строка ${i + 1}: отсутствуют ${missingColumns.join(', ')}`);
      }
      
      // Проверка дубликатов
      const rowHash = this.createRowHash(row);
      if (seenRows.has(rowHash)) {
        quality.duplicateRows++;
        quality.issues.push(`Строка ${i + 1}: дубликат`);
      } else {
        seenRows.add(rowHash);
      }
    }
    
    // Расчет качества
    const validRows = quality.totalRows - quality.emptyRows - quality.incompleteRows - quality.duplicateRows;
    quality.qualityScore = Math.round((validRows / quality.totalRows) * 100);
    
    console.log('📊 Результаты анализа качества:', quality);
    return quality;
  }

  /**
   * Проверка пустой строки
   */
  isEmptyRow(row) {
    return !row.some(cell => cell && String(cell).trim().length > 0);
  }

  /**
   * Проверка наличия валидных данных
   */
  hasValidData(row, columnType) {
    const columnIndex = this.getColumnIndex(columnType);
    if (columnIndex === undefined || columnIndex >= row.length) {
      return false;
    }
    const value = row[columnIndex];
    return value && String(value).trim().length > 0;
  }

  /**
   * Создание хеша строки
   */
  createRowHash(row) {
    return row.map(cell => String(cell || '')).join('|');
  }

  /**
   * Получение индекса колонки
   */
  getColumnIndex(columnType) {
    // Здесь должна быть логика получения индекса колонки
    // Для упрощения возвращаем примерные значения
    const columnMap = {
      'PLATFORM': 1,
      'TEXT': 4,
      'DATE': 6,
      'AUTHOR': 7
    };
    return columnMap[columnType];
  }

  /**
   * Анализ трендов
   */
  analyzeTrends(data) {
    console.log('📈 Анализ трендов...');
    
    const trends = {
      platforms: {},
      engagement: {},
      views: {},
      timeDistribution: {}
    };
    
    // Анализ по площадкам
    data.forEach(row => {
      const platform = this.getColumnValue(row, 'PLATFORM');
      if (platform) {
        trends.platforms[platform] = (trends.platforms[platform] || 0) + 1;
      }
    });
    
    // Анализ вовлечения
    data.forEach(row => {
      const engagement = this.getColumnValue(row, 'ENGAGEMENT');
      if (engagement) {
        const engagementNum = this.parseEngagement(engagement);
        if (engagementNum > 0) {
          trends.engagement[engagementNum] = (trends.engagement[engagementNum] || 0) + 1;
        }
      }
    });
    
    console.log('📊 Результаты анализа трендов:', trends);
    return trends;
  }

  /**
   * Получение значения из колонки
   */
  getColumnValue(row, columnType) {
    const columnIndex = this.getColumnIndex(columnType);
    if (columnIndex === undefined || columnIndex >= row.length) {
      return '';
    }
    return String(row[columnIndex] || '').trim();
  }

  /**
   * Парсинг вовлечения
   */
  parseEngagement(engagement) {
    if (!engagement) return 0;
    
    const engagementStr = String(engagement).replace(/[^\d]/g, '');
    const engagementNum = parseInt(engagementStr);
    
    return isNaN(engagementNum) ? 0 : engagementNum;
  }
}

// ==================== ЭКСПОРТ И ИМПОРТ ====================

/**
 * Класс для экспорта данных
 */
class DataExporter {
  constructor() {
    this.formats = ['xlsx', 'csv', 'json'];
  }

  /**
   * Экспорт в Excel
   */
  exportToExcel(data, filename) {
    console.log('📊 Экспорт в Excel...');
    
    try {
      const spreadsheet = SpreadsheetApp.create(filename);
      const sheet = spreadsheet.getActiveSheet();
      
      // Добавляем заголовки
      const headers = ['Площадка', 'Тема', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Вовлечение', 'Тип'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Добавляем данные
      if (data.length > 0) {
        sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      }
      
      // Форматирование
      sheet.autoResizeColumns(1, headers.length);
      
      console.log(`✅ Экспорт завершен: ${spreadsheet.getUrl()}`);
      return spreadsheet.getUrl();
      
    } catch (error) {
      console.error('❌ Ошибка экспорта в Excel:', error);
      throw error;
    }
  }

  /**
   * Экспорт в CSV
   */
  exportToCSV(data, filename) {
    console.log('📄 Экспорт в CSV...');
    
    try {
      const headers = ['Площадка', 'Тема', 'Текст', 'Дата', 'Автор', 'Просмотры', 'Вовлечение', 'Тип'];
      let csvContent = headers.join(',') + '\n';
      
      data.forEach(row => {
        const csvRow = row.map(cell => `"${String(cell || '').replace(/"/g, '""')}"`).join(',');
        csvContent += csvRow + '\n';
      });
      
      const blob = Utilities.newBlob(csvContent, 'text/csv', filename);
      const file = DriveApp.createFile(blob);
      
      console.log(`✅ Экспорт завершен: ${file.getUrl()}`);
      return file.getUrl();
      
    } catch (error) {
      console.error('❌ Ошибка экспорта в CSV:', error);
      throw error;
    }
  }

  /**
   * Экспорт в JSON
   */
  exportToJSON(data, filename) {
    console.log('📋 Экспорт в JSON...');
    
    try {
      const jsonData = {
        metadata: {
          exportDate: new Date().toISOString(),
          totalRecords: data.length,
          version: '1.0.0'
        },
        data: data
      };
      
      const jsonContent = JSON.stringify(jsonData, null, 2);
      const blob = Utilities.newBlob(jsonContent, 'application/json', filename);
      const file = DriveApp.createFile(blob);
      
      console.log(`✅ Экспорт завершен: ${file.getUrl()}`);
      return file.getUrl();
      
    } catch (error) {
      console.error('❌ Ошибка экспорта в JSON:', error);
      throw error;
    }
  }
}

// ==================== ГЛОБАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Настройка автоматизации
 */
function setupAutomation() {
  const automation = new AutomatedProcessor();
  automation.setupAutomation();
  
  SpreadsheetApp.getUi().alert(
    'Автоматизация настроена!',
    'Автоматическая обработка отчетов настроена:\n\n- Ежедневная проверка в 9:00\n- Обработка при изменении данных\n- Уведомления об ошибках',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Анализ качества данных
 */
function analyzeDataQuality() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const analyzer = new DataAnalyzer();
  const quality = analyzer.analyzeDataQuality(data);
  
  // Создаем отчет о качестве
  const qualitySheet = spreadsheet.insertSheet('Качество данных');
  
  const qualityData = [
    ['АНАЛИЗ КАЧЕСТВА ДАННЫХ'],
    [''],
    [`Всего строк: ${quality.totalRows}`],
    [`Пустых строк: ${quality.emptyRows}`],
    [`Неполных строк: ${quality.incompleteRows}`],
    [`Дубликатов: ${quality.duplicateRows}`],
    [`Качество: ${quality.qualityScore}%`],
    [''],
    ['ПРОБЛЕМЫ:'],
    ...quality.issues.map(issue => [issue])
  ];
  
  qualitySheet.getRange(1, 1, qualityData.length, 1).setValues(qualityData);
  qualitySheet.autoResizeColumn(1);
  
  SpreadsheetApp.getUi().alert(
    'Анализ качества завершен',
    `Качество данных: ${quality.qualityScore}%\n\nПроблем: ${quality.issues.length}\n\nПодробности в листе "Качество данных"`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Экспорт данных
 */
function exportData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const exporter = new DataExporter();
  const filename = `Экспорт_${new Date().toISOString().split('T')[0]}`;
  
  try {
    const excelUrl = exporter.exportToExcel(data, `${filename}.xlsx`);
    const csvUrl = exporter.exportToCSV(data, `${filename}.csv`);
    const jsonUrl = exporter.exportToJSON(data, `${filename}.json`);
    
    SpreadsheetApp.getUi().alert(
      'Экспорт завершен!',
      `Данные экспортированы в форматах:\n\nExcel: ${excelUrl}\nCSV: ${csvUrl}\nJSON: ${jsonUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Ошибка экспорта',
      `Произошла ошибка при экспорте:\n${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Обновление меню с новыми функциями
 */
function updateMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Обработчик отчетов')
    .addItem('Обработать ежемесячный отчет', 'processMonthlyReport')
    .addSeparator()
    .addItem('Анализ качества данных', 'analyzeDataQuality')
    .addItem('Экспорт данных', 'exportData')
    .addSeparator()
    .addItem('Настройка автоматизации', 'setupAutomation')
    .addItem('Настройки', 'showSettings')
    .addToUi();
}

/**
 * Тестирование расширенных функций
 */
function testAdvancedFeatures() {
  console.log('🧪 Тестирование расширенных функций...');
  
  // Тест анализатора данных
  const analyzer = new DataAnalyzer();
  console.log('✅ Анализатор данных инициализирован');
  
  // Тест экспортера
  const exporter = new DataExporter();
  console.log('✅ Экспортер инициализирован');
  
  // Тест автоматизации
  const automation = new AutomatedProcessor();
  console.log('✅ Автоматизация инициализирована');
  
  console.log('✅ Тестирование расширенных функций завершено');
} 