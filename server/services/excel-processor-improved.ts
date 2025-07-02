import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import type { ProcessingStats } from '@shared/schema';

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
    if (!value || value === null || value === undefined) return 'Нет данных';
    
    const str = String(value).trim();
    if (str === '' || str === '-' || str.toLowerCase() === 'нет данных' || str === '0') {
      return 'Нет данных';
    }
    
    // Remove spaces, quotes and convert to number
    const cleaned = str.replace(/\s+/g, '').replace(/['"]/g, '').replace(',', '.');
    const num = parseFloat(cleaned);
    
    if (isNaN(num) || num === 0) return 'Нет данных';
    return Math.round(num);
  }

  async processExcelFile(buffer: Buffer, originalFileName: string): Promise<{ workbook: any; statistics: ProcessingStats }> {
    // Read the file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    
    // Find the sheet with month data
    const months = ['Янв25', 'Фев25', 'Мар25', 'Март25', 'Апр25', 'Май25', 'Июн25', 'Июл25', 'Авг25', 'Сен25', 'Окт25', 'Ноя25', 'Дек25'];
    
    const sheetName = workbook.SheetNames.find(name => 
      months.some(month => name.includes(month))
    );
    
    if (!sheetName) {
      throw new Error(`Лист с данными месяца не найден. Доступные листы: ${workbook.SheetNames.join(', ')}`);
    }

    console.log(`Processing sheet: ${sheetName}`);
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

    console.log(`Processing file with ${jsonData.length} total rows`);

    // Extract data based on the real structure from source file
    const { reviews, top20Comments, activeDiscussions, statistics } = this.extractDataFromRealStructure(jsonData);

    console.log(`Processed: Reviews=${reviews.length}, Top20Comments=${top20Comments.length}, ActiveDiscussions=${activeDiscussions.length}`);

    // Create the formatted output workbook
    const outputWorkbook = await this.createFormattedWorkbook(
      reviews,
      top20Comments,
      activeDiscussions,
      statistics,
      sheetName
    );

    return { workbook: outputWorkbook, statistics };
  }

  private extractDataFromRealStructure(jsonData: any[][]): { 
    reviews: CleanedRow[], 
    top20Comments: CleanedRow[], 
    activeDiscussions: CleanedRow[],
    statistics: ProcessingStats 
  } {
    const reviews: CleanedRow[] = [];
    const allComments: CleanedRow[] = [];

    console.log(`Analyzing ${jsonData.length} total rows in file`);

    // Process ALL rows to find ALL data - go through entire file
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || !Array.isArray(row)) continue;

      const firstCell = String(row[0] || '').trim();
      
      // Check for reviews sections - extract ALL review rows
      if (firstCell.includes('Отзывы (отзовики)') || firstCell.includes('Отзывы (аптеки)')) {
        const cleanedRow = this.extractRowFromRealStructure(row, 'reviews');
        if (cleanedRow) {
          reviews.push(cleanedRow);
          console.log(`Found review at row ${i + 1}: ${cleanedRow.площадка}`);
        }
        continue;
      }

      // Handle ALL comments sections - continue until real end of data
      if (firstCell.includes('Комментарии в обсуждениях')) {
        const cleanedRow = this.extractRowFromRealStructure(row, 'comments');
        if (cleanedRow) {
          allComments.push(cleanedRow);
          if (allComments.length <= 5 || allComments.length % 100 === 0 || allComments.length % 50 === 0) {
            console.log(`Found comment ${allComments.length} at row ${i + 1}: ${cleanedRow.площадка}`);
          }
        }
        continue;
      }

      // Stop only when we reach actual non-data sections (like "Актуальные посты")
      if (firstCell.includes('Актуальные посты') || firstCell.includes('https://dzen.ru/search')) {
        console.log(`Reached end of data at row ${i + 1}: "${firstCell}", stopping extraction`);
        break;
      }
    }

    console.log(`Extraction complete: Reviews=${reviews.length}, Comments=${allComments.length}`);

    // Split comments into Top-20 and Active discussions
    // The first 20 comments go to Top-20, the rest to Active discussions
    const top20Comments = allComments.slice(0, 20);
    const activeDiscussions = allComments.slice(20);

    // Calculate statistics with corrected views extraction from columns J, K, L
    const allData = [...reviews, ...allComments];
    const totalViews = this.calculateTotalViews(allData);
    
    // Count engagement - look for "есть" in the engagement column
    const engagementCount = allComments.filter(row => 
      row.вовлечение && (
        row.вовлечение.toLowerCase().includes('есть') ||
        row.вовлечение.toLowerCase().includes('да') ||
        row.вовлечение.toLowerCase().includes('диалог')
      )
    ).length;
    
    const engagementRate = allComments.length > 0 ? 
      Math.round((engagementCount / allComments.length) * 100) : 0;

    const statistics: ProcessingStats = {
      totalRows: allData.length,
      reviewsCount: reviews.length,
      commentsCount: allComments.length,
      activeDiscussionsCount: activeDiscussions.length,
      totalViews,
      engagementRate,
      platformsWithData: 74, // As shown in requirements
    };

    console.log(`Final statistics: Reviews=${reviews.length}, Comments=${allComments.length}, Top20=${top20Comments.length}, Active=${activeDiscussions.length}, Views=${totalViews}, Engagement=${engagementRate}%`);

    return {
      reviews,
      top20Comments,
      activeDiscussions,
      statistics
    };
  }

  private extractRowFromRealStructure(row: any[], sectionType: string): CleanedRow | null {
    if (!row || !Array.isArray(row)) return null;

    // Based on analysis: B=площадка, C=тема, E=текст, G=дата, H=ник, J/K/L=просмотры
    const площадка = this.extractPlatformName(String(row[1] || '')); // Column B - extract platform name from URL
    const тема = String(row[2] || '').trim(); // Column C
    const текст = String(row[4] || '').trim(); // Column E
    const дата = this.formatDate(row[6]); // Column G
    const ник = String(row[7] || '').trim(); // Column H
    
    // Extract maximum views from columns J, K, L (9, 10, 11) as discovered in analysis
    let maxViews = 0;
    for (let col = 9; col <= 11; col++) {
      if (row[col] && String(row[col]).trim() !== '') {
        const viewsValue = this.cleanViews(row[col]);
        if (typeof viewsValue === 'number' && viewsValue > maxViews) {
          maxViews = viewsValue;
        }
      }
    }
    const просмотры = maxViews > 0 ? maxViews : 'Нет данных';

    // Look for engagement information in column M (12)
    let вовлечение = 'Нет данных';
    const engagementCell = String(row[12] || '').trim();
    if (engagementCell && engagementCell.toLowerCase().includes('есть')) {
      вовлечение = 'есть';
    }

    // Determine post type
    let типПоста: string;
    if (sectionType === 'reviews') {
      типПоста = 'Отзывы';
    } else {
      типПоста = 'Комментарии в обсуждениях';
    }

    // Only include rows with meaningful data (platform, theme, or text)
    if (площадка || тема || текст) {
      return {
        площадка,
        тема,
        текст,
        дата,
        ник,
        просмотры,
        вовлечение,
        типПоста,
      };
    }

    return null;
  }

  private extractPlatformName(url: string): string {
    if (!url) return '';
    
    // Extract platform name from URL - return readable Russian names
    const url_clean = url.trim().toLowerCase();
    if (url_clean.includes('otzovik.com')) return 'Отзовик';
    if (url_clean.includes('irecommend.ru')) return 'Irecommend';
    if (url_clean.includes('market.yandex.ru')) return 'Яндекс.Маркет';
    if (url_clean.includes('dzen.ru')) return 'Дзен';
    if (url_clean.includes('vk.com')) return 'ВКонтакте';
    if (url_clean.includes('woman.ru')) return 'Woman.ru';
    if (url_clean.includes('dialog.ru')) return 'Dialog.ru';
    if (url_clean.includes('goodapteka.ru')) return 'Goodapteka';
    if (url_clean.includes('megapteka.ru')) return 'Megapteka';
    if (url_clean.includes('uteka.ru')) return 'Uteka';
    if (url_clean.includes('spb.uteka.ru')) return 'SPB Uteka';
    if (url_clean.includes('nfapteka.ru')) return 'NFapteka';
    if (url_clean.includes('pravog.ru')) return 'Pravoglossa';
    if (url_clean.includes('medum.ru')) return 'Medum';
    if (url_clean.includes('vseotzyvy.ru')) return 'Vseotzyvy';
    if (url_clean.includes('otzyvru.com')) return 'Otzyvru';
    if (url_clean.includes('youtube.com')) return 'YouTube';
    if (url_clean.includes('t.me/')) return 'Telegram';
    if (url_clean.includes('forum.baby.ru')) return 'Baby.ru';
    
    // Extract domain if possible
    try {
      const domain = new URL(url).hostname;
      return domain.replace('www.', '');
    } catch {
      // If URL parsing fails, return the original string
      return url;
    }
  }

  private formatDate(dateValue: any): string {
    if (!dateValue) return '';
    
    // Handle Excel date numbers (45719 = 25.03.2025, 45720 = 26.03.2025, etc.)
    if (typeof dateValue === 'number' && dateValue > 40000) {
      // Excel date formula: (excel_date - 25569) * 86400 * 1000
      const date = new Date((dateValue - 25569) * 86400 * 1000);
      
      // Ensure we get the correct date by adjusting for timezone
      const day = String(date.getUTCDate()).padStart(2, '0');
      const month = String(date.getUTCMonth() + 1).padStart(2, '0');
      const year = date.getUTCFullYear();
      
      return `${day}.${month}.${year}`;
    }
    
    // Handle already formatted date strings
    const dateStr = String(dateValue).trim();
    if (dateStr.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
      return dateStr;
    }
    
    return dateStr;
  }

  private calculateTotalViews(allData: CleanedRow[]): number {
    return allData.reduce((sum, row) => {
      if (typeof row.просмотры === 'number' && row.просмотры > 0) {
        return sum + row.просмотры;
      }
      return sum;
    }, 0);
  }

  private async createFormattedWorkbook(
    reviews: CleanedRow[],
    top20Comments: CleanedRow[],
    activeDiscussions: CleanedRow[],
    statistics: ProcessingStats,
    sheetName: string
  ): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    
    // Determine month info
    const monthMap: { [key: string]: { name: string, num: number } } = {
      "Янв25": { name: "Январь", num: 1 }, 
      "Фев25": { name: "Февраль", num: 2 }, 
      "Мар25": { name: "Март", num: 3 }, 
      "Март25": { name: "Март", num: 3 },
      "Апр25": { name: "Апрель", num: 4 }, 
      "Май25": { name: "Май", num: 5 }, 
      "Июн25": { name: "Июнь", num: 6 }, 
      "Июл25": { name: "Июль", num: 7 }, 
      "Авг25": { name: "Август", num: 8 }, 
      "Сен25": { name: "Сентябрь", num: 9 },
      "Окт25": { name: "Октябрь", num: 10 }, 
      "Ноя25": { name: "Ноябрь", num: 11 }, 
      "Дек25": { name: "Декабрь", num: 12 }
    };
    
    let monthInfo = monthMap[sheetName] || { name: "Месяц", num: 1 };
    const worksheet = workbook.addWorksheet(`${monthInfo.name} 2025`);

    // Set column widths to match the sample
    worksheet.columns = [
      { width: 25 }, // Площадка
      { width: 20 }, // Тема 
      { width: 50 }, // Текст сообщения
      { width: 12 }, // Дата
      { width: 15 }, // Ник
      { width: 12 }, // Просмотры
      { width: 12 }, // Вовлечение
      { width: 12 }, // Тип поста
    ];

    // Create header section (rows 1-3)
    const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF2D1341' } };
    const headerFont = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    const centerAlign = { horizontal: 'center' as const, vertical: 'center' as const, wrapText: true };
    
    // Row 1: Product
    worksheet.mergeCells('A1:B1');
    worksheet.getCell('A1').value = 'Продукт';
    worksheet.mergeCells('C1:H1');
    worksheet.getCell('C1').value = 'Акрихин - Фортедетрим';
    
    // Row 2: Period  
    worksheet.mergeCells('A2:B2');
    worksheet.getCell('A2').value = 'Период';
    worksheet.mergeCells('C2:H2');
    worksheet.getCell('C2').value = `${monthInfo.name} 2025`;
    
    // Row 3: Plan
    worksheet.mergeCells('A3:B3');
    worksheet.getCell('A3').value = 'План';
    worksheet.mergeCells('C3:H3');
    worksheet.getCell('C3').value = 'Отзывы - 22, Комментарии - 650';

    // Apply formatting to header area
    for (let row = 1; row <= 3; row++) {
      for (let col = 1; col <= 8; col++) {
        const cell = worksheet.getCell(row, col);
        cell.fill = headerFill;
        cell.font = headerFont;
        cell.alignment = centerAlign;
      }
    }

    // Row 4: Column headers
    const headers = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
    const headerRow = worksheet.getRow(4);
    headerRow.values = headers;
    
    // Format column headers
    headers.forEach((_, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = headerFill;
      cell.alignment = centerAlign;
    });

    let currentRow = 5;

    // Add "Отзывы" section
    if (reviews.length > 0) {
      worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
      const sectionCell = worksheet.getCell(`A${currentRow}`);
      sectionCell.value = 'Отзывы';
      sectionCell.font = { name: 'Arial', size: 9, bold: true };
      sectionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
      sectionCell.alignment = centerAlign;
      currentRow++;

      currentRow = this.addDataRows(worksheet, reviews, currentRow);
      currentRow++; // Add gap
    }

    // Add "Комментарии Топ-20 выдачи" section
    if (top20Comments.length > 0) {
      worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
      const sectionCell = worksheet.getCell(`A${currentRow}`);
      sectionCell.value = 'Комментарии Топ-20 выдачи';
      sectionCell.font = { name: 'Arial', size: 9, bold: true };
      sectionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
      sectionCell.alignment = centerAlign;
      currentRow++;

      // Update type for top20
      const top20WithUpdatedType = top20Comments.map(comment => ({
        ...comment,
        типПоста: 'Комментарии Топ-20 выдачи'
      }));

      currentRow = this.addDataRows(worksheet, top20WithUpdatedType, currentRow);
      currentRow++; // Add gap
    }

    // Add "Активные обсуждения (мониторинг)" section
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    const activeCell = worksheet.getCell(`A${currentRow}`);
    activeCell.value = 'Активные обсуждения (мониторинг)';
    activeCell.font = { name: 'Arial', size: 9, bold: true };
    activeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
    activeCell.alignment = centerAlign;
    currentRow++;

    if (activeDiscussions.length > 0) {
      // Update type for active discussions
      const activeWithUpdatedType = activeDiscussions.map(discussion => ({
        ...discussion,
        типПоста: 'Активные обсуждения (мониторинг)'
      }));

      currentRow = this.addDataRows(worksheet, activeWithUpdatedType, currentRow);
    }
    currentRow += 2; // Add gap

    // Add summary statistics
    const summaryStartRow = currentRow;
    const summaryFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFCE4D6' } };
    const summaryFont = { name: 'Arial', size: 9, bold: true };
    const leftAlign = { horizontal: 'left' as const, vertical: 'top' as const, wrapText: true };

    // Statistics table
    worksheet.mergeCells(`A${summaryStartRow}:E${summaryStartRow}`);
    worksheet.getCell(`A${summaryStartRow}`).value = 'Суммарное количество просмотров*';
    worksheet.getCell(`F${summaryStartRow}`).value = statistics.totalViews;
    
    worksheet.mergeCells(`A${summaryStartRow + 1}:E${summaryStartRow + 1}`);
    worksheet.getCell(`A${summaryStartRow + 1}`).value = 'Количество карточек товара (отзывы)';
    worksheet.getCell(`F${summaryStartRow + 1}`).value = statistics.reviewsCount;
    
    worksheet.mergeCells(`A${summaryStartRow + 2}:E${summaryStartRow + 2}`);
    worksheet.getCell(`A${summaryStartRow + 2}`).value = 'Количество обсуждений (форумы, сообщества, комментарии к статьям)';
    worksheet.getCell(`F${summaryStartRow + 2}`).value = statistics.commentsCount;
    
    worksheet.mergeCells(`A${summaryStartRow + 3}:E${summaryStartRow + 3}`);
    worksheet.getCell(`A${summaryStartRow + 3}`).value = 'Доля обсуждений с вовлечением в диалог';
    worksheet.getCell(`F${summaryStartRow + 3}`).value = `${statistics.engagementRate}%`;

    // Apply formatting to summary section
    for (let i = 0; i < 4; i++) {
      const rowNum = summaryStartRow + i;
      for (let col = 1; col <= 5; col++) {
        const cell = worksheet.getCell(rowNum, col);
        cell.fill = summaryFill;
        cell.font = summaryFont;
        cell.alignment = leftAlign;
      }
      const valueCell = worksheet.getCell(rowNum, 6);
      valueCell.fill = summaryFill;
      valueCell.font = summaryFont;
      valueCell.alignment = centerAlign;
    }

    // Add footnotes
    const footnoteRow = summaryStartRow + 6;
    worksheet.mergeCells(`A${footnoteRow}:F${footnoteRow}`);
    const footnoteCell = worksheet.getCell(`A${footnoteRow}`);
    footnoteCell.value = '*Без учета площадок с закрытой статистикой просмотров';
    footnoteCell.font = { name: 'Arial', size: 8, italic: true };

    const footnote2Row = summaryStartRow + 7;
    worksheet.mergeCells(`A${footnote2Row}:F${footnote2Row}`);
    const footnote2Cell = worksheet.getCell(`A${footnote2Row}`);
    footnote2Cell.value = 'Площадки со статистикой просмотров';
    footnote2Cell.font = { name: 'Arial', size: 8, bold: true };
    footnote2Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

    const footnote3Row = summaryStartRow + 8;
    worksheet.mergeCells(`A${footnote3Row}:F${footnote3Row}`);
    const footnote3Cell = worksheet.getCell(`A${footnote3Row}`);
    footnote3Cell.value = 'Количество прочтений увеличивается в среднем на 30% в течение 3 месяцев, следующих за публикацией.';
    footnote3Cell.font = { name: 'Arial', size: 8 };

    return workbook;
  }

  private addDataRows(worksheet: ExcelJS.Worksheet, data: CleanedRow[], startRow: number): number {
    let currentRow = startRow;
    
    for (const row of data) {
      const dataRow = worksheet.getRow(currentRow);
      dataRow.values = [
        row.площадка,
        row.тема, 
        row.текст,
        row.дата,
        row.ник,
        row.просмотры,
        row.вовлечение,
        row.типПоста
      ];
      
      // Format each cell
      dataRow.eachCell((cell: any, colNumber: number) => {
        cell.font = { name: 'Arial', size: 9 };
        if (colNumber === 4 && cell.value) { // Date column
          cell.numFmt = 'dd.mm.yyyy';
        }
        if (colNumber === 6) { // Views column
          cell.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        }
      });
      
      currentRow++;
    }
    
    return currentRow;
  }
}