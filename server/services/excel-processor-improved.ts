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
    if (!value) return 0;
    
    const str = String(value).trim();
    if (str === '' || str === '-' || str.toLowerCase() === 'нет данных') {
      return 0;
    }
    
    // Remove spaces and convert to number
    const cleaned = str.replace(/\s+/g, '').replace(',', '.');
    const num = parseFloat(cleaned);
    
    return isNaN(num) ? 0 : Math.round(num);
  }

  private calculateEngagementRate(comments: number, likes: number, reposts: number, views: number): number {
    if (views === 0) return 0;
    return ((comments + likes + reposts) / views) * 100;
  }

  async processExcelFile(buffer: Buffer, originalFileName: string): Promise<{ workbook: any; statistics: ProcessingStats }> {
    // Read the file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    
    // Find the sheet with month data
    const months = ['Янв', 'Фев', 'Мар', 'Март', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
    
    const sheetName = workbook.SheetNames.find(name => 
      months.some(month => name.includes(month))
    );
    
    if (!sheetName) {
      throw new Error(`Лист с данными месяца не найден. Доступные листы: ${workbook.SheetNames.join(', ')}`);
    }

    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

    console.log(`Processing file with ${jsonData.length} total rows`);

    // Extract ALL data rows (improved logic)
    const allDataRows = this.extractAllDataRows(jsonData);
    
    // Clean and process the data
    const processedData = this.cleanData(allDataRows);

    console.log(`Processed: ${processedData.length} total data rows`);

    // Calculate statistics
    const statistics = this.calculateStatistics(processedData);

    // Create the formatted output workbook
    const outputWorkbook = await this.createFormattedWorkbook(
      processedData,
      statistics,
      sheetName
    );

    return { workbook: outputWorkbook, statistics };
  }

  /**
   * Найти строку с заголовками и вернуть список нужных данных
   * @param jsonData лист как массив массивов
   */
  private extractAllDataRows(jsonData: any[][]): any[][] {
    // требуемые колонны и возможные варианты их названия
    const headerAliases: Record<string, string[]> = {
      platform: ['площадка'],
      topic: ['продукт/тема', 'тема', 'продукт'],
      text: ['текст', 'текст сообщения'],
      date: ['дата'],
      author: ['ник', 'автор'],
      views: ['просмотров получено', 'просмотры получено', 'просмотры'],
      engagement: ['вовлечение'],
      postType: ['тип поста', 'тип'],
    };

    const normalize = (s: string) => s.toLowerCase().trim();

    // 1. ищем строку-заголовок
    let headerRowIdx = -1;
    let colIndex: Record<keyof typeof headerAliases, number> | null = null;

    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row) continue;
      const mapping: any = {};
      Object.entries(headerAliases).forEach(([key, aliases]) => {
        const idx = row.findIndex((cell: any) => aliases.includes(normalize(String(cell ?? ''))));
        if (idx !== -1) mapping[key as keyof typeof headerAliases] = idx;
      });

      // если нашли все колонки — это заголовок
      if (Object.keys(mapping).length === Object.keys(headerAliases).length) {
        headerRowIdx = i;
        colIndex = mapping as Record<keyof typeof headerAliases, number>;
        break;
      }
    }

    if (headerRowIdx === -1 || !colIndex) {
      throw new Error('Не удалось найти строку заголовков с необходимыми колонками');
    }

    // 2. собираем данные ниже заголовка
    const rows: any[][] = [];
    for (let i = headerRowIdx + 1; i < jsonData.length; i++) {
      const r = jsonData[i];
      if (!r) continue;

      const platform = String(r[colIndex.platform] ?? '').trim();
      const topic = String(r[colIndex.topic] ?? '').trim();

      if (platform === '' && topic === '') {
        // достигли пустой области — заканчиваем, но не раньше чем 3 пустые подряд
        continue;
      }

      // Формируем массив данных в фиксированном порядке
      rows.push([
        platform,
        topic,
        r[colIndex.text] ?? '',
        r[colIndex.date] ?? '',
        r[colIndex.author] ?? '',
        r[colIndex.views] ?? '',
        r[colIndex.engagement] ?? '',
        r[colIndex.postType] ?? '',
      ]);
    }

    return rows;
  }

  private cleanData(rawData: any[][]): CleanedRow[] {
    return rawData
      .map(row => {
        const rawType = String(row[7] || '').trim().toLowerCase();
        let category = 'Активные обсуждения (мониторинг)';
        if (rawType.includes('отзыв')) {
          category = 'Отзывы';
        } else if (rawType.includes('коммент')) {
          category = 'Комментарии Топ-20 выдачи';
        }

        return {
          площадка: String(row[0] || '').trim(),
          тема: String(row[1] || '').trim(),
          текст: String(row[2] || '').trim(),
          дата: row[3] || '',
          ник: String(row[4] || '').trim(),
          просмотры: this.cleanViews(row[5]),
          вовлечение: String(row[6] || '').trim(),
          типПоста: category,
        } as CleanedRow;
      })
      .filter(row => row.площадка !== '' || row.тема !== '' || row.текст !== '');
  }

  private calculateStatistics(allData: CleanedRow[]): ProcessingStats {
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

    // Categorize rows by type
    const reviews = allData.filter(r => r.типПоста === 'Отзывы');
    const comments = allData.filter(r => r.типПоста === 'Комментарии Топ-20 выдачи');
    const activePosts = allData.filter(r => r.типПоста === 'Активные обсуждения (мониторинг)');

    return {
      totalRows: allData.length,
      reviewsCount: reviews.length,
      commentsCount: comments.length,
      activeDiscussionsCount: activePosts.length,
      totalViews,
      engagementRate: Math.round(engagementRate),
      platformsWithData: Math.round(platformsWithDataPercentage),
    };
  }

  private async createFormattedWorkbook(
    allData: CleanedRow[],
    statistics: ProcessingStats,
    sheetName: string
  ): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    
    // Determine month name for sheet title
    const monthMap: { [key: string]: string } = {
      "Янв": "Январь", "Фев": "Февраль", "Мар": "Март", "Март": "Март",
      "Апр": "Апрель", "Май": "Май", "Июн": "Июнь", 
      "Июл": "Июль", "Авг": "Август", "Сен": "Сентябрь",
      "Окт": "Октябрь", "Ноя": "Ноябрь", "Дек": "Декабрь"
    };
    
    let monthName = "Месяц";
    for (const [abbr, full] of Object.entries(monthMap)) {
      if (sheetName.includes(abbr)) {
        monthName = full;
        break;
      }
    }
    
    const worksheet = workbook.addWorksheet(`${monthName} 2025`);

    // Set column widths
    worksheet.columns = [
      { width: 20 }, // Площадка
      { width: 25 }, // Продукт/Тема
      { width: 50 }, // Текст сообщения
      { width: 15 }, // Дата
      { width: 20 }, // Ник
      { width: 15 }, // Просмотры
      { width: 15 }, // Вовлечение
      { width: 15 }, // Тип поста
    ];

    // Add title
    worksheet.mergeCells('A1:H1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `Отчет ORM - ${monthName} 2025`;
    titleCell.font = { bold: true, size: 16 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' }
    };

    // Add statistics summary
    worksheet.mergeCells('A3:H3');
    const summaryCell = worksheet.getCell('A3');
    summaryCell.value = `Всего записей: ${statistics.totalRows} | Просмотров: ${statistics.totalViews.toLocaleString('ru-RU')} | Вовлечение: ${statistics.engagementRate}%`;
    summaryCell.font = { bold: true };
    summaryCell.alignment = { horizontal: 'center' };
    summaryCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE7E6E6' }
    };

    // Add headers
    const headers = ['Площадка', 'Продукт/Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
    const headerRow = worksheet.getRow(5);
    headerRow.values = headers;
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9E2F3' }
    };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

    // Add data
    let currentRow = 6;
    
    // Group data by type
    const dataByType: { [key: string]: CleanedRow[] } = {};
    allData.forEach(row => {
      const type = row.типПоста || 'Другое';
      if (!dataByType[type]) {
        dataByType[type] = [];
      }
      dataByType[type].push(row);
    });

    const categoriesOrder = ['Отзывы', 'Комментарии Топ-20 выдачи', 'Активные обсуждения (мониторинг)'];

    for (const type of categoriesOrder) {
      const rows = dataByType[type] || [];
      if (rows.length === 0) continue;

      // Add section header
      worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
      const sectionCell = worksheet.getCell(`A${currentRow}`);
      sectionCell.value = `${type} (${rows.length} записей)`;
      sectionCell.font = { bold: true };
      sectionCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF2F2F2' }
      };
      currentRow++;

      // Add rows
      rows.forEach(row => {
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
        
        // Format views column
        if (typeof row.просмотры === 'number' && row.просмотры > 0) {
          dataRow.getCell(6).numFmt = '#,##0';
        }
        
        currentRow++;
      });
    }

    // Add borders to all cells with data
    for (let i = 1; i < currentRow; i++) {
      const row = worksheet.getRow(i);
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    }

    return workbook;
  }
}