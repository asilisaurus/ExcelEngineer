import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import type { ProcessingStats } from '@shared/schema';

interface DataRow {
  площадка: string;
  тема: string;
  текст: string;
  дата: string | Date;
  ник: string;
  просмотры: number | string;
  вовлечение: string;
  типПоста: string;
}

export class ExcelProcessor {
  private getCleanValue(value: any): string {
    if (value === null || value === undefined) return '';
    return value.toString().trim();
  }

  private cleanViews(value: any): number | string {
    if (typeof value === 'number' && value > 0) {
      return Math.round(value);
    }
    
    if (typeof value === 'string') {
      // Проверяем, если это дата в формате "3/7/25"
      const dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
      if (dateMatch) {
        const month = parseInt(dateMatch[1]);
        const day = parseInt(dateMatch[2]);
        let year = parseInt(dateMatch[3]);
        
        // Если год двузначный, добавляем 2000
        if (year < 100) {
          year += 2000;
        }
        
        // Создаем дату и конвертируем в Excel serial number
        const date = new Date(year, month - 1, day);
        const excelEpoch = new Date(1900, 0, 1);
        const daysDiff = Math.floor((date.getTime() - excelEpoch.getTime()) / (1000 * 60 * 60 * 24));
        
        // Excel считает 1900 год високосным (хотя это не так), поэтому добавляем 2
        return daysDiff + 2;
      }
      
      // Пытаемся парсить как обычное число
      const str = value.toString().replace(/[\s,'"]/g, '');
      const num = parseFloat(str);
      if (!isNaN(num) && num > 0) {
        return Math.round(num);
      }
    }
    
    return 'Нет данных';
  }

  async processExcelFile(buffer: Buffer, originalFileName: string): Promise<{ workbook: any; statistics: ProcessingStats }> {
    try {
      console.log('🔄 Начало обработки файла:', originalFileName);
      
      // Читаем Excel файл
      const workbook = XLSX.read(buffer, { 
        type: 'buffer',
        cellDates: true,
        cellNF: false,
        cellText: false
      });
      
      console.log('📋 Найденные листы:', workbook.SheetNames);
      
      // Находим нужный лист
      const monthSheet = this.findMonthSheet(workbook.SheetNames);
      if (!monthSheet) {
        throw new Error('Не найден лист с данными месяца');
      }
      
      console.log('📊 Обрабатываем лист:', monthSheet);
      
      // Конвертируем в JSON с сохранением структуры
      const worksheet = workbook.Sheets[monthSheet];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1, 
        defval: '',
        raw: false // Важно для корректного извлечения чисел
      }) as any[][];
      
      console.log('📈 Извлечено строк из файла:', jsonData.length);
      
      // Извлекаем данные с правильной логикой
      console.log('🔍 Начинаем извлечение данных...');
      const { reviews, comments, active, statistics } = this.extractDataCorrectly(jsonData);
      
      console.log('📊 Статистика извлечения:', {
        reviews: reviews.length,
        comments: comments.length,
        active: active.length,
        total: statistics.totalRows,
        totalViews: statistics.totalViews
      });
      
      // Создаем форматированный отчет
      console.log('📝 Создание отчета...');
      const formattedWorkbook = await this.createFormattedReport(
        { отзывы: reviews, комментарии: comments, активные: active },
        statistics,
        monthSheet
      );
      
      console.log('✅ Файл успешно обработан');
      
      return {
        workbook: formattedWorkbook,
        statistics
      };
      
    } catch (error) {
      console.error('❌ Ошибка обработки файла:', error);
      throw error;
    }
  }

  private findMonthSheet(sheetNames: string[]): string | null {
    const monthPatterns = ['март', 'мар', 'march', 'mar'];
    return sheetNames.find(name => 
      monthPatterns.some(pattern => name.toLowerCase().includes(pattern))
    ) || sheetNames[0];
  }

  private extractDataCorrectly(jsonData: any[][]): { 
    reviews: DataRow[], 
    comments: DataRow[], 
    active: DataRow[], 
    statistics: ProcessingStats 
  } {
    const reviews: DataRow[] = [];
    const comments: DataRow[] = [];
    const active: DataRow[] = [];
    
    console.log('🔍 Анализ структуры данных...');
    
    // Находим секции по ключевым словам в первой колонке
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || !Array.isArray(row) || row.length === 0) continue;
      
      const firstCell = this.getCleanValue(row[0]).toLowerCase();
      
      // Отзывы - строки с "отзыв" в первой колонке
      if (firstCell.includes('отзыв') && !firstCell.includes('карточек')) {
        const rowData = this.extractReviewRow(row, i + 1);
        if (rowData) {
          reviews.push(rowData);
        }
      }
      
      // Комментарии - строки с "комментарии" в первой колонке  
      else if (firstCell.includes('комментарии') && firstCell.includes('обсуждениях')) {
        const rowData = this.extractCommentRow(row, i + 1);
        if (rowData) {
          comments.push(rowData);
        }
      }
    }
    
    console.log(`📊 Извлечено: Отзывы=${reviews.length}, Комментарии=${comments.length}`);
    
    // Рассчитываем статистику
    const statistics = this.calculateStatistics(reviews, comments, active);
    
    return { reviews, comments, active, statistics };
  }

  private extractReviewRow(row: any[], rowIndex: number): DataRow | null {
    try {
      // Структура отзывов: [Тип, Площадка, Продукт, Ссылка, Текст, Согласование, Дата, Ник]
      const площадка = this.getCleanValue(row[1]);
      const продукт = this.getCleanValue(row[2]);
      const текст = this.getCleanValue(row[4]);
      const дата = row[6] || '';
      const ник = this.getCleanValue(row[7]);
      
      // Для отзывов просмотры - это дата в колонке 6, конвертируем в Excel serial number
      const просмотры = this.cleanViews(row[6]);
      
      if (!площадка && !текст) return null;
      
      return {
        площадка,
        тема: продукт,
        текст,
        дата: дата,
        ник,
        просмотры,
        вовлечение: 'Нет данных',
        типПоста: 'Отзывы'
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения отзыва в строке ${rowIndex}:`, error);
      return null;
    }
  }

  private extractCommentRow(row: any[], rowIndex: number): DataRow | null {
    try {
      // Структура комментариев: [Тип, Площадка, Продукт, Ссылка, Текст, _, Дата, Ник, Автор, Старт, Конец, Получено, Вовлечение]
      const площадка = this.getCleanValue(row[1]);
      const продукт = this.getCleanValue(row[2]);
      const текст = this.getCleanValue(row[4]);
      const дата = row[6] || '';
      const ник = this.getCleanValue(row[7]);
      
      // Для комментариев просмотры также в колонке 6
      const просмотры = this.cleanViews(row[6]);
      const вовлечение = this.getCleanValue(row[12]) || 'Нет данных';
      
      if (!площадка && !текст) return null;
      
      return {
        площадка,
        тема: продукт,
        текст,
        дата: дата,
        ник,
        просмотры,
        вовлечение,
        типПоста: 'Комментарии Топ-20 выдачи'
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения комментария в строке ${rowIndex}:`, error);
      return null;
    }
  }

  private calculateStatistics(reviews: DataRow[], comments: DataRow[], active: DataRow[]): ProcessingStats {
    const allData = [...reviews, ...comments, ...active];
    
    // Суммируем просмотры
    const totalViews = allData.reduce((sum, row) => {
      if (typeof row.просмотры === 'number') {
        return sum + row.просмотры;
      }
      return sum;
    }, 0);
    
    // Считаем записи с просмотрами
    const recordsWithViews = allData.filter(row => 
      typeof row.просмотры === 'number' && row.просмотры > 0
    ).length;
    
    // Процент площадок с данными
    const platformsWithData = allData.length > 0 ? 
      Math.round((recordsWithViews / allData.length) * 100) : 0;
    
    // Считаем вовлечение для комментариев
    const discussionData = [...comments, ...active];
    const engagedDiscussions = discussionData.filter(row => 
      row.вовлечение && 
      row.вовлечение !== 'Нет данных' && 
      (row.вовлечение.toLowerCase().includes('есть') || 
       row.вовлечение.toLowerCase().includes('да') ||
       row.вовлечение.toLowerCase().includes('+'))
    ).length;
    
    const engagementRate = discussionData.length > 0 ? 
      Math.round((engagedDiscussions / discussionData.length) * 100) : 0;
    
    return {
      totalRows: allData.length,
      reviewsCount: reviews.length,
      commentsCount: comments.length,
      activeDiscussionsCount: active.length,
      totalViews,
      engagementRate,
      platformsWithData
    };
  }

  private async createFormattedReport(
    groupedData: { отзывы: DataRow[]; комментарии: DataRow[]; активные: DataRow[] },
    statistics: ProcessingStats,
    sheetName: string
  ): Promise<any> {
    console.log('📝 Создание форматированного отчета...');
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Март 2025');

    // Настройка ширины колонок
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

    console.log('📋 Создание шапки отчета...');
    // Создание шапки отчета
    this.createReportHeader(worksheet, statistics);

    console.log('📋 Создание заголовков таблицы...');
    // Создание заголовков таблицы
    this.createTableHeaders(worksheet);

    let currentRow = 5;

    console.log('📋 Добавление отзывов...');
    // Добавление разделов данных
    currentRow = this.addDataSection(worksheet, 'Отзывы', groupedData.отзывы, currentRow);
    
    console.log('📋 Добавление комментариев...');
    currentRow = this.addDataSection(worksheet, 'Комментарии Топ-20 выдачи', groupedData.комментарии, currentRow);
    
    // Активные обсуждения (если есть)
    if (groupedData.активные.length > 0) {
      console.log('📋 Добавление активных обсуждений...');
      currentRow = this.addDataSection(worksheet, 'Активные обсуждения (мониторинг)', groupedData.активные, currentRow);
    }

    console.log('📋 Добавление итоговых показателей...');
    // Добавление итоговых показателей
    this.addSummaryMetrics(worksheet, statistics, currentRow + 2);

    console.log('✅ Форматированный отчет создан');
    return workbook;
  }

  private createReportHeader(worksheet: ExcelJS.Worksheet, statistics: ProcessingStats): void {
    const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF2D1341' } };
    const headerFont = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
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
    worksheet.getCell('C2').value = 'Март 2025';
    
    // Строка 3: План
    worksheet.mergeCells('A3:B3');
    worksheet.getCell('A3').value = 'План';
    worksheet.mergeCells('C3:H3');
    worksheet.getCell('C3').value = `Отзывы - ${statistics.reviewsCount}, Комментарии - ${statistics.commentsCount}`;

    // Применение форматирования к шапке
    for (let row = 1; row <= 3; row++) {
      for (let col = 1; col <= 8; col++) {
        const cell = worksheet.getCell(row, col);
        cell.fill = headerFill;
        cell.font = headerFont;
        cell.alignment = centerAlign;
      }
    }
  }

  private createTableHeaders(worksheet: ExcelJS.Worksheet): void {
    const headers = ['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста'];
    const headerRow = worksheet.getRow(4);
    headerRow.values = headers;
    
    const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF2D1341' } };
    const headerFont = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    const centerAlign = { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: true };

    headers.forEach((_, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.alignment = centerAlign;
    });
  }

  private addDataSection(worksheet: ExcelJS.Worksheet, sectionName: string, data: DataRow[], startRow: number): number {
    let currentRow = startRow;

    // Добавляем заголовок секции
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    const sectionCell = worksheet.getCell(`A${currentRow}`);
    sectionCell.value = sectionName;
    sectionCell.font = { name: 'Arial', size: 9, bold: true };
    sectionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
    sectionCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    worksheet.getRow(currentRow).height = 12;
    currentRow++;

    // Добавляем данные секции
    data.forEach(row => {
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
      
      // Форматирование ячеек данных
      dataRow.eachCell((cell: any, colNumber: number) => {
        cell.font = { name: 'Arial', size: 9 };
        if (colNumber === 4 && cell.value) { // Дата
          cell.numFmt = 'dd.mm.yyyy';
        }
        if (colNumber === 6) { // Просмотры - центрирование
          cell.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        }
      });
      
      dataRow.height = 12;
      currentRow++;
    });
    
    return currentRow + 1; // Добавляем пробел после секции
  }

  private addSummaryMetrics(worksheet: ExcelJS.Worksheet, statistics: ProcessingStats, startRow: number): void {
    const summaryData = [
      ['', '', '', '', '', '', '', ''],
      ['Суммарное количество просмотров', '', '', '', '', statistics.totalViews, '', ''],
      ['Количество карточек товара (отзывы)', '', '', '', '', statistics.reviewsCount, '', ''],
      ['Количество обсуждений (форумы, сообщества, комментарии к статьям)', '', '', '', '', statistics.commentsCount, '', ''],
      ['Доля обсуждений с вовлечением в диалог', '', '', '', '', `${statistics.engagementRate}%`, '', ''],
      ['', '', '', '', '', '', '', ''],
      ['*Без учета площадок с закрытой статистикой просмотров', '', '', '', '', '', '', ''],
      ['Площадки со статистикой просмотров', '', '', '', '', `${statistics.platformsWithData}%`, '', ''],
      ['Количество прочтений увеличивается в среднем на 30% в течение 3 месяцев, следующих за публикацией.', '', '', '', '', '', '', '']
    ];

    summaryData.forEach((rowData, index) => {
      const row = worksheet.getRow(startRow + index);
      row.values = rowData;
      
      // Специальное форматирование для строки с процентами
      if (index === 7) { // Строка "Площадки со статистикой просмотров"
        row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Желтый фон
        for (let col = 1; col <= 8; col++) {
          row.getCell(col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        }
      }
      
      row.eachCell((cell: any) => {
        cell.font = { name: 'Arial', size: 9 };
        cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      });
    });
  }
}