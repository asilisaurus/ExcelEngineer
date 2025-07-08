import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

// Интерфейс статистики для unified процессора
interface ProcessingStats {
  totalRows: number;
  reviewsCount: number;
  commentsCount: number;
  activeDiscussionsCount: number;
  totalViews: number;
  engagementRate: number;
  platformsWithData: number;
  processingTime: number;
  qualityScore: number;
}

interface DataRow {
  площадка: string;
  тема: string;
  текст: string;
  дата: string;
  ник: string;
  просмотры: number | string;
  вовлечение: string;
  типПоста: string;
  qualityScore: number;
  originalRow: any[];
}

interface ProcessedData {
  reviews: DataRow[];
  topComments: DataRow[];
  activeDiscussions: DataRow[];
  statistics: ProcessingStats;
  monthName: string;
  fileName: string;
}

interface ProcessingConfig {
  maxReviews: number;
  maxTopComments: number;
  minTextLength: number;
  enableQualityFiltering: boolean;
  enableSmartFiltering: boolean;
}

export class ExcelProcessorUnified {
  // Платформы из smart processor
  private reviewPlatforms = [
    'otzovik', 'irecommend', 'otzyvru', 'pravogolosa', 'medum', 
    'vseotzyvy', 'otzyvy.pro', 'market.yandex', 'dialog.ru',
    'goodapteka', 'megapteka', 'uteka', 'nfapteka', 'piluli.ru',
    'eapteka.ru', 'pharmspravka.ru', 'gde.ru', 'ozon.ru'
  ];

  private commentPlatforms = [
    'dzen.ru', 'woman.ru', 'forum.baby.ru', 'vk.com', 't.me',
    'ok.ru', 'otvet.mail.ru', 'babyblog.ru', 'mom.life', 
    'youtube.com', 'pikabu.ru', 'livejournal.com', 'facebook.com'
  ];

  private defaultConfig: ProcessingConfig = {
    maxReviews: 22,
    maxTopComments: 20,
    minTextLength: 10,
    enableQualityFiltering: true,
    enableSmartFiltering: true
  };

  // Гибкий вход как из smart processor
  async processExcelFile(
    input: string | Buffer, 
    fileName?: string,
    config?: Partial<ProcessingConfig>
  ): Promise<{ outputPath: string; statistics: ProcessingStats }> {
    const startTime = Date.now();
    const processingConfig = { ...this.defaultConfig, ...config };
    
    try {
      console.log('🔄 UNIFIED PROCESSOR - Начало обработки файла:', fileName || 'unknown');
      
      let workbook: XLSX.WorkBook;
      let originalFileName: string;
      
      // Поддержка как строки так и буфера (из smart)
      if (typeof input === 'string') {
        // Если передан путь к файлу
        if (!fs.existsSync(input)) {
          throw new Error(`Файл не найден: ${input}`);
        }
        const buffer = fs.readFileSync(input);
        workbook = XLSX.read(buffer, { 
          type: 'buffer',
          cellDates: false,
          cellNF: false,
          cellText: false,
          raw: true
        });
        originalFileName = fileName || path.basename(input);
      } else {
        // Если передан буфер
        workbook = XLSX.read(input, { 
          type: 'buffer',
          cellDates: false,
          cellNF: false,
          cellText: false,
          raw: true
        });
        originalFileName = fileName || 'unknown.xlsx';
      }
      
      console.log('📋 Найденные листы:', workbook.SheetNames);
      
      // Поиск листа с данными месяца (из smart)
      const months = ["Янв25", "Фев25", "Мар25", "Март25", "Апр25", "Май25", "Июн25", 
                     "Июл25", "Авг25", "Сен25", "Окт25", "Ноя25", "Дек25"];
      
      const worksheet = workbook.Sheets[workbook.SheetNames.find((name: string) => 
        months.some(month => name.includes(month))
      ) || workbook.SheetNames[0]];
      
      console.log('📊 Обрабатываем лист:', workbook.SheetNames[0]);
      
      // Конвертируем в массив данных (из perfect)
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1, 
        defval: '',
        raw: true
      }) as any[][];
      
      console.log('📈 Извлечено строк из файла:', jsonData.length);
      
      // Извлекаем данные с лучшей логикой
      const processedData = this.extractDataPerfectly(jsonData, originalFileName, processingConfig);
      
      // Создаем форматированный отчет
      const outputPath = await this.createPerfectWorkbook(processedData, originalFileName);
      
      const processingTime = Date.now() - startTime;
      processedData.statistics.processingTime = processingTime;
      
      console.log('✅ Файл успешно обработан за', processingTime, 'мс');
      
      return {
        outputPath,
        statistics: processedData.statistics
      };
      
    } catch (error) {
      console.error('❌ Ошибка обработки файла:', error);
      throw error;
    }
  }

  private extractDataPerfectly(
    jsonData: any[][],
    originalFileName: string,
    config: ProcessingConfig
  ): ProcessedData {
    console.log('🔍 UNIFIED EXTRACTION - Анализ данных...');
    
    const reviews: DataRow[] = [];
    const topComments: DataRow[] = [];
    const activeDiscussions: DataRow[] = [];
    
    let totalViews = 0;
    let engagementCount = 0;
    let commentsWithViews = 0;
    let totalQualityScore = 0;
    
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || !Array.isArray(row) || row.length === 0) continue;
      
      const rowType = this.analyzeRowType(row);
      
      if (rowType === 'review') {
        const reviewData = this.extractReviewData(row, i + 1);
        if (reviewData && this.passesQualityFilter(reviewData, config)) {
          reviews.push(reviewData);
          totalQualityScore += reviewData.qualityScore;
          
          if (typeof reviewData.просмотры === 'number') {
            totalViews += reviewData.просмотры;
          }
        }
      } else if (rowType === 'comment') {
        const commentData = this.extractCommentData(row, i + 1);
        if (commentData && this.passesQualityFilter(commentData, config)) {
          // Разделяем на Топ-20 и Активные обсуждения (из ultimate)
          if (topComments.length < config.maxTopComments) {
            topComments.push(commentData);
          } else {
            activeDiscussions.push(commentData);
          }
          
          totalQualityScore += commentData.qualityScore;
          
          // Подсчет статистики
          if (typeof commentData.просмотры === 'number') {
            totalViews += commentData.просмотры;
            commentsWithViews++;
          }
          
          if (commentData.вовлечение && 
              (commentData.вовлечение.toLowerCase().includes('есть') || 
               commentData.вовлечение.toLowerCase().includes('да') ||
               commentData.вовлечение.toLowerCase().includes('+'))) {
            engagementCount++;
          }
        }
      }
    }
    
    // Применяем умную фильтрацию (из smart)
    const filteredReviews = this.applySmartFiltering(reviews, config.maxReviews);
    const filteredTopComments = this.applySmartFiltering(topComments, config.maxTopComments);
    
    console.log(`📊 Извлечено: ${filteredReviews.length} отзывов, ${filteredTopComments.length} топ-комментариев, ${activeDiscussions.length} активных обсуждений`);
    
    const totalRows = filteredReviews.length + filteredTopComments.length + activeDiscussions.length;
    const discussionCount = filteredTopComments.length + activeDiscussions.length;
    const engagementRate = discussionCount > 0 ? Math.round((engagementCount / discussionCount) * 100) : 0;
    const platformsWithData = commentsWithViews > 0 ? Math.round((commentsWithViews / discussionCount) * 100) : 0;
    const averageQualityScore = totalRows > 0 ? Math.round(totalQualityScore / totalRows) : 0;
    
    const statistics: ProcessingStats = {
      totalRows,
      reviewsCount: filteredReviews.length,
      commentsCount: filteredTopComments.length,
      activeDiscussionsCount: activeDiscussions.length,
      totalViews,
      engagementRate,
      platformsWithData,
      processingTime: 0, // Будет установлено позже
      qualityScore: averageQualityScore
    };
    
    // Определяем название месяца и файла
    const monthName = this.extractMonthName(originalFileName);
    const fileName = this.generateFileName(originalFileName, monthName);
    
    return {
      reviews: filteredReviews,
      topComments: filteredTopComments,
      activeDiscussions,
      statistics,
      monthName,
      fileName
    };
  }

  private analyzeRowType(row: any[]): string {
    if (!row || row.length === 0) return 'empty';
    
    const colA = this.getCleanValue(row[0]).toLowerCase();
    const colB = this.getCleanValue(row[1]).toLowerCase();
    const colD = this.getCleanValue(row[3]).toLowerCase();
    const colE = this.getCleanValue(row[4]).toLowerCase();
    const colN = this.getCleanValue(row[13]).toLowerCase();
    
    // Заголовки и служебные строки
    if (colA.includes('тип размещения') || colA.includes('площадка') || 
        colB.includes('площадка') || colE.includes('текст сообщения') ||
        colA.includes('план') || colA.includes('итого')) {
      return 'header';
    }
    
    // Точная логика определения типов (из ultimate)
    if (colA.includes('отзыв')) {
      return 'review';
    }
    
    if (colA.includes('комментарии')) {
      return 'comment';
    }
    
    // Анализ по URL и платформам (из smart)
    const urlText = colB + ' ' + colD;
    const isReviewPlatform = this.reviewPlatforms.some(platform => 
      urlText.toLowerCase().includes(platform)
    );
    const isCommentPlatform = this.commentPlatforms.some(platform => 
      urlText.toLowerCase().includes(platform)
    );
    
    // Анализ типа поста в колонке N
    const postType = colN;
    
    // Определяем тип по платформе и типу поста
    if ((colB || colD || colE) && (isReviewPlatform || postType === 'ос')) {
      return 'review';
    }
    
    if ((colB || colD || colE) && (isCommentPlatform || postType === 'цс')) {
      return 'comment';
    }
    
    // Если есть контент, но тип неясен
    if (colB || colD || colE) {
      return 'content';
    }
    
    return 'empty';
  }

  private extractReviewData(row: any[], rowIndex: number): DataRow | null {
    try {
      const площадка = this.getCleanValue(row[1]);
      const тема = this.extractThemeFromText(this.getCleanValue(row[4]));
      const текст = this.getCleanValue(row[4]);
      const дата = this.convertExcelDateToString(row[6]);
      const ник = this.getCleanValue(row[7]);
      const просмотры = this.extractViewsFromRow(row, 'review');
      const qualityScore = this.calculateQualityScore(row);
      
      if (!площадка && !текст) return null;
      
      return {
        площадка,
        тема,
        текст,
        дата,
        ник,
        просмотры,
        вовлечение: 'Нет данных',
        типПоста: 'ОС',
        qualityScore,
        originalRow: row
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения отзыва в строке ${rowIndex}:`, error);
      return null;
    }
  }

  private extractCommentData(row: any[], rowIndex: number): DataRow | null {
    try {
      const площадка = this.getCleanValue(row[1]);
      const тема = this.extractThemeFromText(this.getCleanValue(row[4]));
      const текст = this.getCleanValue(row[4]);
      const дата = this.convertExcelDateToString(row[6]);
      const ник = this.getCleanValue(row[7]);
      const просмотры = this.extractViewsFromRow(row, 'comment');
      const вовлечение = this.getCleanValue(row[12]) || 'Нет данных';
      const qualityScore = this.calculateQualityScore(row);
      
      if (!площадка && !текст) return null;
      
      return {
        площадка,
        тема,
        текст,
        дата,
        ник,
        просмотры,
        вовлечение,
        типПоста: 'ЦС',
        qualityScore,
        originalRow: row
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения комментария в строке ${rowIndex}:`, error);
      return null;
    }
  }

  // Правильная обработка дат Excel serial numbers (из perfect)
  private convertExcelDateToString(dateValue: any): string {
    if (!dateValue) return '';
    
    // Если это Excel serial number (число больше 40000)
    if (typeof dateValue === 'number' && dateValue > 40000) {
      try {
        const date = new Date((dateValue - 25569) * 86400 * 1000);
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        return `${day}.${month}.${year}`;
      } catch (error) {
        console.warn('Ошибка конверсии даты:', dateValue, error);
        return dateValue.toString();
      }
    }
    
    // Если это строка с датой в формате MM/DD/YYYY
    if (typeof dateValue === 'string' && dateValue.includes('/')) {
      const parts = dateValue.split('/');
      if (parts.length === 3) {
        const day = parts[1].padStart(2, '0');
        const month = parts[0].padStart(2, '0');
        const year = parts[2];
        return `${day}.${month}.${year}`;
      }
    }
    
    return dateValue.toString();
  }

  private extractViewsFromRow(row: any[], type: 'review' | 'comment'): number | string {
    if (type === 'review') {
      // Для отзывов просмотры НЕ в колонке G (там дата)
      // Проверяем другие колонки
      const possibleViews = [row[10], row[11], row[12]];
      
      for (const value of possibleViews) {
        if (typeof value === 'number' && value > 0 && value < 10000000) {
          return Math.round(value);
        }
      }
    } else {
      // Для комментариев просмотры в колонках K, L, M (10, 11, 12)
      const possibleViews = [row[10], row[11], row[12]];
      
      for (const value of possibleViews) {
        if (typeof value === 'number' && value > 0 && value < 10000000) {
          return Math.round(value);
        }
      }
    }
    
    return 'Нет данных';
  }

  // Извлечение темы из текста (из ultimate)
  private extractThemeFromText(text: string): string {
    if (!text) return '';
    
    const cleanText = text.replace(/Название:\s*/i, '').trim();
    const firstSentence = cleanText.split('.')[0];
    
    if (firstSentence.length > 50) {
      return firstSentence.substring(0, 47) + '...';
    }
    
    return firstSentence;
  }

  // Система качества из smart processor
  private calculateQualityScore(row: any[]): number {
    let score = 100;
    
    const colB = this.getCleanValue(row[1]);
    const colD = this.getCleanValue(row[3]);
    const colE = this.getCleanValue(row[4]);
    const colG = row[6];
    const colH = this.getCleanValue(row[7]);
    const colN = this.getCleanValue(row[13]);
    
    // Штрафы за отсутствие важных данных
    if (!colE || colE.length < 20) score -= 30;
    if (!colD || !colD.includes('http')) score -= 25;
    if (!colG || typeof colG !== 'number') score -= 20;
    if (!colH || colH.length < 3) score -= 15;
    if (!colN || (colN !== 'ос' && colN !== 'цс')) score -= 10;
    
    // Дополнительные штрафы
    if (colE.length < 50) score -= 10;
    if (colD.includes('deleted') || colD.includes('removed')) score -= 50;
    
    return Math.max(0, score);
  }

  private passesQualityFilter(row: DataRow, config: ProcessingConfig): boolean {
    if (!config.enableQualityFiltering) return true;
    
    // Минимальная длина текста
    if (row.текст.length < config.minTextLength) return false;
    
    // Минимальный балл качества
    if (row.qualityScore < 30) return false;
    
    // Исключаем удаленные записи
    if (row.текст.includes('удален') || row.текст.includes('removed')) return false;
    
    return true;
  }

  // Умная фильтрация (из smart)
  private applySmartFiltering(rows: DataRow[], maxCount: number): DataRow[] {
    if (!rows.length) return rows;
    
    console.log(`🔍 Применяем умную фильтрацию: ${rows.length} → макс. ${maxCount}`);
    
    // Сортируем по качеству
    const sortedRows = [...rows].sort((a, b) => b.qualityScore - a.qualityScore);
    
    // Берем лучшие
    const filtered = sortedRows.slice(0, maxCount);
    
    console.log(`✅ После умной фильтрации: ${filtered.length} записей`);
    
    return filtered;
  }

  private getCleanValue(value: any): string {
    if (value === null || value === undefined) return '';
    return value.toString().trim();
  }

  private extractMonthName(fileName: string): string {
    const monthsMap: { [key: string]: string } = {
      'янв': 'Январь', 'фев': 'Февраль', 'мар': 'Март', 'март': 'Март',
      'апр': 'Апрель', 'май': 'Май', 'июн': 'Июнь',
      'июл': 'Июль', 'авг': 'Август', 'сен': 'Сентябрь',
      'окт': 'Октябрь', 'ноя': 'Ноябрь', 'дек': 'Декабрь'
    };
    
    const lowerFileName = fileName.toLowerCase();
    for (const [key, value] of Object.entries(monthsMap)) {
      if (lowerFileName.includes(key)) {
        return value;
      }
    }
    
    return 'Март';
  }

  private generateFileName(originalFileName: string, monthName: string): string {
    const baseName = originalFileName.replace(/\.[^/.]+$/, '');
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    return `${baseName}_${monthName}_2025_результат_${timestamp}.xlsx`;
  }

  // Создание идеального workbook (из perfect)
  private async createPerfectWorkbook(data: ProcessedData, originalFileName: string): Promise<string> {
    console.log('📝 Создание идеального форматированного отчета...');
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${data.monthName} 2025`);

    // Настройка ширины колонок
    worksheet.columns = [
      { width: 60 }, { width: 30 }, { width: 80 }, { width: 12 },
      { width: 20 }, { width: 12 }, { width: 15 }, { width: 10 },
      { width: 8 }, { width: 12 }, { width: 15 }, { width: 8 }
    ];

    // Создание шапки отчета
    this.createPerfectHeader(worksheet, data);

    // Создание заголовков таблицы
    this.createPerfectTableHeaders(worksheet, data);

    let currentRow = 5;

    // Добавление секций
    currentRow = this.addPerfectSection(worksheet, 'Отзывы', data.reviews, currentRow);
    currentRow = this.addPerfectSection(worksheet, 'Комментарии Топ-20 выдачи', data.topComments, currentRow);
    
    if (data.activeDiscussions.length > 0) {
      currentRow = this.addPerfectSection(worksheet, 'Активные обсуждения (мониторинг)', data.activeDiscussions, currentRow);
    }

    // Добавление итоговой статистики
    this.addPerfectSummary(worksheet, data.statistics, currentRow + 2);

    // Генерация пути к файлу
    const outputPath = path.join(process.cwd(), 'uploads', data.fileName);
    await workbook.xlsx.writeFile(outputPath);

    console.log('✅ Идеальный отчет создан:', outputPath);
    return outputPath;
  }

  private createPerfectHeader(worksheet: ExcelJS.Worksheet, data: ProcessedData): void {
    const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF2D1B69' } };
    const headerFont = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    const centerAlign = { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: true };
    
    // Строка 1: Продукт
    worksheet.mergeCells('A1:B1');
    worksheet.getCell('A1').value = 'Продукт';
    worksheet.mergeCells('C1:H1');
    worksheet.getCell('C1').value = 'Акрихин - Фортедетрим';
    
    // Строка 2: Период
    worksheet.mergeCells('A2:B2');
    worksheet.getCell('A2').value = 'Период';
    worksheet.mergeCells('C2:H2');
    worksheet.getCell('C2').value = `${data.monthName} 2025`;
    
    // Строка 3: План
    worksheet.mergeCells('A3:B3');
    worksheet.getCell('A3').value = 'План';
    worksheet.mergeCells('C3:H3');
    worksheet.getCell('C3').value = `Отзывы - ${data.statistics.reviewsCount}, Комментарии - ${data.statistics.commentsCount}`;
    
    // Дополнительные колонки
    worksheet.getCell('I3').value = 'Отзыв';
    worksheet.getCell('J3').value = 'Упоминание';
    worksheet.getCell('K3').value = 'Поддерживающее';
    worksheet.getCell('L3').value = 'Всего';

    // Применение стилей
    for (let row = 1; row <= 3; row++) {
      for (let col = 1; col <= 12; col++) {
        const cell = worksheet.getCell(row, col);
        cell.fill = headerFill;
        cell.font = headerFont;
        cell.alignment = centerAlign;
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      }
    }
  }

  private createPerfectTableHeaders(worksheet: ExcelJS.Worksheet, data: ProcessedData): void {
    const headers = [
      'Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 
      'Просмотры', 'Вовлечение', 'Тип поста', 
      data.statistics.reviewsCount.toString(),
      data.statistics.commentsCount.toString(),
      '', 
      data.statistics.totalRows.toString()
    ];
    
    const headerRow = worksheet.getRow(4);
    headerRow.values = headers;
    
    const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF2D1B69' } };
    const headerFont = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    const centerAlign = { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: true };

    headers.forEach((_, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.alignment = centerAlign;
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    });
  }

  private addPerfectSection(worksheet: ExcelJS.Worksheet, sectionName: string, data: DataRow[], startRow: number): number {
    let currentRow = startRow;

    // Заголовок секции
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    const sectionCell = worksheet.getCell(`A${currentRow}`);
    sectionCell.value = sectionName;
    sectionCell.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF000000' } };
    sectionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
    sectionCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    sectionCell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' }
    };
    worksheet.getRow(currentRow).height = 12;
    currentRow++;

    // Данные секции
    data.forEach(row => {
      const dataRow = worksheet.getRow(currentRow);
      dataRow.values = [
        row.площадка, row.тема, row.текст, row.дата,
        row.ник, row.просмотры, row.вовлечение, row.типПоста
      ];
      
      // Форматирование ячеек
      dataRow.eachCell((cell: any, colNumber: number) => {
        cell.font = { name: 'Arial', size: 9 };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        
        if (colNumber === 4 || colNumber === 6) {
          cell.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        }
      });
      
      dataRow.height = 12;
      currentRow++;
    });

    return currentRow;
  }

  private addPerfectSummary(worksheet: ExcelJS.Worksheet, statistics: ProcessingStats, startRow: number): void {
    const summaryData = [
      ['', '', '', '', 'Суммарное количество просмотров*', statistics.totalViews],
      ['', '', '', '', 'Количество карточек товара (отзывы)', statistics.reviewsCount],
      ['', '', '', '', 'Количество обсуждений (форумы, сообщества, комментарии к статьям)', statistics.commentsCount],
      ['', '', '', '', 'Доля обсуждений с вовлечением в диалог', `${statistics.engagementRate}%`],
      ['', '', '', '', 'Средний балл качества записей', `${statistics.qualityScore}%`],
      [],
      ['', '', '*Без учета площадок с закрытой статистикой прочтений'],
      ['', '', 'Площадки со статистикой просмотров', '', '', `${statistics.platformsWithData}%`],
      ['', '', 'Количество прочтений увеличивается в среднем на 30% в течение 3 месяцев, следующих за публикацией.']
    ];

    summaryData.forEach((rowData, index) => {
      const row = worksheet.getRow(startRow + index);
      row.values = rowData;
      row.font = { name: 'Arial', size: 9 };
    });
  }
}

// Создаем экземпляр объединенного процессора
export const unifiedProcessor = new ExcelProcessorUnified();