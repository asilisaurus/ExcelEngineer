import * as XLSX from 'xlsx';
const ExcelJS = require('exceljs');
import { ProcessingStats } from '@shared/schema';

interface ProcessedData {
  reviews: any[];
  comments: any[];
  activeDiscussions: any[];
  statistics: ProcessingStats;
}

interface CleanedRow {
  площадка: string;
  тема: string;
  текст: string;
  дата: Date | string;
  ник: string;
  просмотры: number | string;
  вовлечение: string;
  типПоста: string;
}

export class ExcelProcessor {
  private cleanViews(value: any): number | string {
    if (!value || value === null || value === undefined || String(value).trim() === '' || String(value).trim() === 'Нет данных') {
      return 'Нет данных';
    }
    
    try {
      const cleaned = String(value).replace(/\s/g, '').replace(/,/g, '').replace(/'/g, '');
      const num = parseFloat(cleaned);
      return !isNaN(num) && num > 0 ? num : 'Нет данных';
    } catch {
      return 'Нет данных';
    }
  }

  private calculateEngagementRate(comments: number, likes: number, reposts: number, views: number): number {
    if (views === 0 || isNaN(views)) return 0;
    return ((comments + likes + reposts) / views) * 100;
  }

  async processExcelFile(buffer: Buffer, originalFileName: string): Promise<{ workbook: any; statistics: ProcessingStats }> {
    // Read the source file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    
    // Find the month sheet
    const months = ["Янв25", "Фев25", "Мар25", "Март25", "Апр25", "Май25", "Июн25", 
                   "Июл25", "Авг25", "Сен25", "Окт25", "Ноя25", "Дек25"];
    
    const sheetName = workbook.SheetNames.find(name => 
      months.some(month => name.includes(month))
    );
    
    if (!sheetName) {
      throw new Error(`Лист с данными месяца не найден. Доступные листы: ${workbook.SheetNames.join(', ')}`);
    }

    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

    console.log(`Processing file with ${jsonData.length} total rows`);

    // Find sections by content analysis
    const { reviews, commentsTop20, activeDiscussions } = this.extractSections(jsonData);
    
    // Clean and process the data
    const processedReviews = this.cleanData(reviews);
    const processedComments = this.cleanData(commentsTop20);
    const processedActiveDiscussions = this.cleanData(activeDiscussions);

    console.log(`Processed: ${processedReviews.length} reviews, ${processedComments.length} comments, ${processedActiveDiscussions.length} active discussions`);

    // Calculate statistics
    const allData = [...processedReviews, ...processedComments, ...processedActiveDiscussions];
    const statistics = this.calculateStatistics(allData, processedReviews.length, processedComments.length, processedActiveDiscussions.length);

    // Create the formatted output workbook
    const outputWorkbook = await this.createFormattedWorkbook(
      { reviews: processedReviews, comments: processedComments, topPosts: processedActiveDiscussions, activeDiscussions: [] },
      statistics,
      sheetName
    );

    return { workbook: outputWorkbook, statistics };
  }

  private extractSections(jsonData: any[][]): { reviews: any[][]; commentsTop20: any[][]; activeDiscussions: any[][] } {
    const reviews: any[][] = [];
    const commentsTop20: any[][] = [];
    const activeDiscussions: any[][] = [];
    
    // Process all data rows, ignoring section headers
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || row.length === 0) continue;
      
      const firstCell = String(row[0] || '').toLowerCase().trim();
      const secondCell = String(row[1] || '').toLowerCase().trim();
      
      // Skip header rows and section titles
      if (firstCell.includes('тип размещения') || 
          firstCell.includes('отзыв') || 
          firstCell.includes('комментари') || 
          firstCell.includes('активн') ||
          secondCell.includes('площадка') ||
          !row[1] || // Skip empty rows
          row.length < 10) { // Skip rows with insufficient columns
        continue;
      }
      
      // Extract data from any valid data row
      // Check if this looks like a data row (has platform name and other data)
      if (row[1] && String(row[1]).trim() !== '') {
        const extractedRow = [
          row[1] || '',   // Площадка
          row[2] || '',   // Продукт/Тема  
          row[4] || '',   // Текст сообщения
          row[6] || '',   // Дата
          row[7] || '',   // Ник
          row[11] || '',  // Просмотров получено
          row[12] || '',  // Вовлечение
          row[13] || ''   // Тип поста
        ];
        
        // Determine section based on content or type
        const text = String(row[4] || '').toLowerCase();
        const postType = String(row[13] || '').toLowerCase();
        
        // Simple categorization - all data goes to reviews for now
        // since we want to process ALL data
        reviews.push(extractedRow);
      }
    }
    
    return { reviews, commentsTop20, activeDiscussions };
  }

  private cleanData(rawData: any[][]): CleanedRow[] {
    return rawData.map(row => ({
      площадка: String(row[0] || '').trim(),
      тема: String(row[1] || '').trim(),
      текст: String(row[2] || '').trim(),
      дата: row[3] || '',
      ник: String(row[4] || '').trim(),
      просмотры: this.cleanViews(row[5]),
      вовлечение: String(row[6] || '').trim(),
      типПоста: String(row[7] || '').trim()
    })).filter(row => row.площадка || row.тема || row.текст); // Filter out completely empty rows
  }

  private calculateStatistics(allData: CleanedRow[], reviewsCount: number, commentsCount: number, activeCount: number): ProcessingStats {
    const totalViews = allData.reduce((sum, row) => {
      const views = typeof row.просмотры === 'number' ? row.просмотры : 0;
      return sum + views;
    }, 0);

    const platformsWithViews = allData.filter(row => 
      typeof row.просмотры === 'number' && row.просмотры > 0
    ).length;

    const engagementRate = allData.filter(row => 
      row.вовлечение && row.вовлечение.toLowerCase().includes('есть')
    ).length / allData.length * 100;

    const platformsWithDataPercentage = allData.length > 0 ? (platformsWithViews / allData.length) * 100 : 0;

    return {
      totalRows: allData.length,
      reviewsCount,
      commentsCount,
      activeDiscussionsCount: activeCount,
      totalViews,
      engagementRate: Math.round(engagementRate),
      platformsWithData: Math.round(platformsWithDataPercentage),
    };
  }

  private async createFormattedWorkbook(
    data: { reviews: CleanedRow[]; comments: CleanedRow[]; topPosts: CleanedRow[]; activeDiscussions: CleanedRow[] },
    statistics: ProcessingStats,
    sheetName: string
  ): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    
    // Determine month name for sheet title
    const monthMap: { [key: string]: string } = {
      "Янв": "Январь", "Фев": "Февраль", "Мар": "Март", "Март": "Март",
      "Апр": "Апрель", "Май": "Май", "Июн": "Июнь", "Июл": "Июль",
      "Авг": "Август", "Сен": "Сентябрь", "Окт": "Октябрь", "Ноя": "Ноябрь", "Дек": "Декабрь"
    };
    
    const sheetPrefix = sheetName.startsWith("Март") ? sheetName.slice(0, 4) : sheetName.slice(0, 3);
    const monthName = monthMap[sheetPrefix] || "Отчет";
    const year = `20${sheetName.slice(-2)}`;
    
    const worksheet = workbook.addWorksheet(`${monthName} ${year}`);

    // Define styles
    const headerStyle = {
      font: { bold: true, size: 12, color: { argb: 'FFFFFF' } },
      fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: '2D1341' } },
      alignment: { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: true }
    };

    const sectionStyle = {
      font: { bold: true, size: 9 },
      fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'C5D9F1' } },
      alignment: { horizontal: 'center' as const, vertical: 'middle' as const }
    };

    const dataStyle = {
      font: { name: 'Arial', size: 9 },
      alignment: { horizontal: 'left' as const, vertical: 'top' as const, wrapText: true }
    };

    // Header setup
    const monthIndex = sheetName.includes('Март') ? 3 : Math.max(1, ["Янв25", "Фев25", "Мар25", "Апр25", "Май25", "Июн25", "Июл25", "Авг25", "Сен25", "Окт25", "Ноя25", "Дек25"].indexOf(sheetName) + 1);
    const dateStr = `01.${monthIndex.toString().padStart(2, '0')}.2025`;

    // Create header rows
    worksheet.mergeCells('A1:B1');
    worksheet.getCell('A1').value = 'Продукт';
    worksheet.mergeCells('C1:G1');
    worksheet.getCell('C1').value = 'Акрихин - Фортедетрим';

    worksheet.mergeCells('A2:B2');
    worksheet.getCell('A2').value = 'Период';
    worksheet.mergeCells('C2:G2');
    worksheet.getCell('C2').value = dateStr;

    worksheet.mergeCells('A3:B3');
    worksheet.getCell('A3').value = 'План';
    worksheet.mergeCells('C3:G3');
    worksheet.getCell('C3').value = 'Отзывы - 22, Комментарии - 650';

    // Apply header styling
    for (let row = 1; row <= 3; row++) {
      for (let col = 1; col <= 8; col++) {
        const cell = worksheet.getCell(row, col);
        cell.style = headerStyle;
      }
    }

    // Column headers
    const columnHeaders = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
    for (let i = 0; i < columnHeaders.length; i++) {
      const cell = worksheet.getCell(4, i + 1);
      cell.value = columnHeaders[i];
      cell.style = headerStyle;
    }

    // Set column widths
    worksheet.getColumn(1).width = 25; // Площадка
    worksheet.getColumn(2).width = 20; // Тема
    worksheet.getColumn(3).width = 50; // Текст сообщения
    worksheet.getColumn(4).width = 12; // Дата
    worksheet.getColumn(5).width = 15; // Ник
    worksheet.getColumn(6).width = 12; // Просмотры
    worksheet.getColumn(7).width = 12; // Вовлечение
    worksheet.getColumn(8).width = 12; // Тип поста

    let currentRow = 5;

    // Add sections with data
    const sections = [
      { title: 'Отзывы', data: data.reviews },
      { title: 'Комментарии Топ-20 выдачи', data: data.topPosts },
      { title: 'Активные обсуждения (мониторинг)', data: data.activeDiscussions }
    ];

    for (const section of sections) {
      // Add section header
      worksheet.mergeCells(currentRow, 1, currentRow, 8);
      const sectionCell = worksheet.getCell(currentRow, 1);
      sectionCell.value = section.title;
      sectionCell.style = sectionStyle;
      worksheet.getRow(currentRow).height = 12;
      currentRow++;

      // Add section data
      for (const row of section.data) {
        const worksheetRow = worksheet.getRow(currentRow);
        worksheetRow.getCell(1).value = row.площадка;
        worksheetRow.getCell(2).value = row.тема;
        worksheetRow.getCell(3).value = row.текст;
        worksheetRow.getCell(4).value = row.дата;
        worksheetRow.getCell(5).value = row.ник;
        worksheetRow.getCell(6).value = row.просмотры;
        worksheetRow.getCell(7).value = row.вовлечение;
        worksheetRow.getCell(8).value = row.типПоста;

        // Apply data styling
        for (let col = 1; col <= 8; col++) {
          const cell = worksheetRow.getCell(col);
          cell.style = dataStyle;
          if (col === 4 && row.дата) {
            cell.numFmt = 'dd.mm.yyyy';
          }
          if (col === 6) {
            cell.alignment = { horizontal: 'center', vertical: 'top' };
          }
        }

        worksheetRow.height = 12;
        currentRow++;
      }
    }

    // Add statistics at the bottom
    currentRow += 2;
    const statsStartRow = currentRow;
    
    worksheet.mergeCells(currentRow, 6, currentRow, 8);
    worksheet.getCell(currentRow, 6).value = `Суммарное количество просмотров: ${statistics.totalViews}`;
    currentRow++;
    
    worksheet.mergeCells(currentRow, 6, currentRow, 8);
    worksheet.getCell(currentRow, 6).value = `Количество карточек товара (отзывы): ${statistics.reviewsCount}`;
    currentRow++;
    
    worksheet.mergeCells(currentRow, 6, currentRow, 8);
    worksheet.getCell(currentRow, 6).value = `Количество обсуждений (форумы, соцсети, комментарии и статьи): ${statistics.commentsCount}`;
    currentRow++;
    
    worksheet.mergeCells(currentRow, 6, currentRow, 8);
    worksheet.getCell(currentRow, 6).value = `Доля обсуждений с вовлечением в диалог: ${statistics.engagementRate}%`;
    currentRow++;
    
    worksheet.mergeCells(currentRow, 6, currentRow, 8);
    worksheet.getCell(currentRow, 6).value = `*Без учета площадок с закрытой статистикой прочтений`;
    currentRow++;
    
    worksheet.mergeCells(currentRow, 6, currentRow, 8);
    worksheet.getCell(currentRow, 6).value = `Площадки со статистикой просмотров: ${statistics.platformsWithData}%`;
    currentRow++;
    
    worksheet.mergeCells(currentRow, 6, currentRow, 8);
    worksheet.getCell(currentRow, 6).value = `Количество прочтений увеличивается в среднем на 50% в течение 3 месяцев, следующих за публикацией`;

    // Style statistics section
    for (let row = statsStartRow; row <= currentRow; row++) {
      for (let col = 6; col <= 8; col++) {
        const cell = worksheet.getCell(row, col);
        cell.style = {
          font: { size: 9, bold: true },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCE4D6' } },
          alignment: { horizontal: 'left', vertical: 'top', wrapText: true }
        };
      }
    }

    return workbook;
  }
}
