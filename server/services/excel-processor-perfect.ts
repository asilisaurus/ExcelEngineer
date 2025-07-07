import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import type { ProcessingStats } from '@shared/schema';

interface DataRow {
  площадка: string;
  тема: string;
  текст: string;
  дата: string;
  ник: string;
  просмотры: number | string;
  вовлечение: string;
  типПоста: string;
}

interface ProcessedData {
  reviews: DataRow[];
  comments: DataRow[];
  activeDiscussions: DataRow[];
  statistics: ProcessingStats;
  monthName: string;
  fileName: string;
}

export class ExcelProcessorPerfect {
  
  async processExcelFile(buffer: Buffer, originalFileName: string): Promise<{ workbook: any; statistics: ProcessingStats }> {
    console.log('🔄 PERFECT PROCESSOR - Начало обработки файла:', originalFileName);
    
    try {
      // Читаем Excel файл
      const workbook = XLSX.read(buffer, { 
        type: 'buffer',
        cellDates: false, // Важно: читаем как числа, чтобы правильно обработать даты
        cellNF: false,
        cellText: false
      });
      
      console.log('📋 Найденные листы:', workbook.SheetNames);
      
      // Находим лист с данными
      const sheetName = workbook.SheetNames[0]; // Берем первый лист
      console.log('📊 Обрабатываем лист:', sheetName);
      
      // Конвертируем в JSON
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1, 
        defval: '',
        raw: true // Важно: сохраняем исходные значения
      }) as any[][];
      
      console.log('📈 Извлечено строк из файла:', jsonData.length);
      
      // Извлекаем данные с правильной логикой
      const processedData = this.extractDataPerfectly(jsonData, sheetName);
      
      console.log('📊 Результат обработки:', {
        reviews: processedData.reviews.length,
        comments: processedData.comments.length,
        active: processedData.activeDiscussions.length,
        totalViews: processedData.statistics.totalViews,
        fileName: processedData.fileName
      });
      
      // Создаем форматированный отчет
      const formattedWorkbook = await this.createPerfectWorkbook(processedData);
      
      console.log('✅ Файл успешно обработан');
      
      return {
        workbook: formattedWorkbook,
        statistics: processedData.statistics
      };
      
    } catch (error) {
      console.error('❌ Ошибка обработки файла:', error);
      throw error;
    }
  }

  private extractDataPerfectly(jsonData: any[][], sheetName: string): ProcessedData {
    console.log('🔍 PERFECT EXTRACTION - Анализ данных...');
    
    const reviews: DataRow[] = [];
    const comments: DataRow[] = [];
    const activeDiscussions: DataRow[] = [];
    
    let totalViews = 0;
    let engagementCount = 0;
    let commentsWithViews = 0;
    
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || !Array.isArray(row) || row.length === 0) continue;
      
      const firstCell = this.getCleanValue(row[0]).toLowerCase();
      
      // Отзывы - строки с "отзыв" в первой колонке
      if (firstCell.includes('отзыв')) {
        const reviewData = this.extractReviewData(row, i + 1);
        if (reviewData) {
          reviews.push(reviewData);
        }
      }
      
      // Комментарии - строки с "комментарии" в первой колонке  
      else if (firstCell.includes('комментарии')) {
        const commentData = this.extractCommentData(row, i + 1);
        if (commentData) {
          // Первые 20 комментариев идут в "Комментарии Топ-20 выдачи"
          if (comments.length < 20) {
            comments.push(commentData);
          } else {
            // Остальные в "Активные обсуждения"
            activeDiscussions.push(commentData);
          }
          
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
    
    console.log(`📊 Извлечено: ${reviews.length} отзывов, ${comments.length} комментариев, ${activeDiscussions.length} активных обсуждений`);
    console.log(`📊 Статистика: ${totalViews} просмотров, ${engagementCount} вовлечений`);
    
    const totalRows = reviews.length + comments.length + activeDiscussions.length;
    const discussionCount = comments.length + activeDiscussions.length;
    const engagementRate = discussionCount > 0 ? Math.round((engagementCount / discussionCount) * 100) : 0;
    const platformsWithData = commentsWithViews > 0 ? Math.round((commentsWithViews / discussionCount) * 100) : 0;
    
    const statistics: ProcessingStats = {
      totalRows,
      reviewsCount: reviews.length,
      commentsCount: discussionCount,
      activeDiscussionsCount: activeDiscussions.length,
      totalViews,
      engagementRate,
      platformsWithData
    };
    
    // Определяем название месяца и файла
    const monthName = this.extractMonthName(sheetName);
    const fileName = this.generateFileName('source_file', monthName);
    
    return {
      reviews,
      comments,
      activeDiscussions,
      statistics,
      monthName,
      fileName
    };
  }

  private extractReviewData(row: any[], rowIndex: number): DataRow | null {
    try {
      const площадка = this.getCleanValue(row[1]);
      const тема = this.getCleanValue(row[2]);
      const текст = this.getCleanValue(row[4]);
      const дата = this.convertExcelDateToString(row[6]);
      const ник = this.getCleanValue(row[7]);
      
      if (!площадка && !текст) return null;
      
      return {
        площадка,
        тема,
        текст,
        дата,
        ник,
        просмотры: 'Нет данных',
        вовлечение: 'Нет данных',
        типПоста: 'ОС'
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения отзыва в строке ${rowIndex}:`, error);
      return null;
    }
  }

  private extractCommentData(row: any[], rowIndex: number): DataRow | null {
    try {
      const площадка = this.getCleanValue(row[1]);
      const тема = this.getCleanValue(row[2]);
      const текст = this.getCleanValue(row[4]);
      
      // Колонка G (6) ВСЕГДА содержит дату для комментариев
      const дата = this.convertExcelDateToString(row[6]);
      const ник = this.getCleanValue(row[7]);
      
      // Для комментариев просмотры НЕ в колонке G (там дата)!
      const просмотры = this.extractViews(row);
      const вовлечение = this.getCleanValue(row[12]) || 'Нет данных';
      
      if (!площадка && !текст) return null;
      
      return {
        площадка,
        тема,
        текст,
        дата,
        ник,
        просмотры,
        вовлечение,
        типПоста: 'ЦС'
      };
    } catch (error) {
      console.warn(`⚠️ Ошибка извлечения комментария в строке ${rowIndex}:`, error);
      return null;
    }
  }

  private convertExcelDateToString(dateValue: any): string {
    if (!dateValue) return '';
    
    // Если это Excel serial number (число больше 40000)
    if (typeof dateValue === 'number' && dateValue > 40000) {
      try {
        // Конвертируем Excel serial number в дату
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
    
    // Если это уже строка с датой
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

  private extractViews(row: any[]): number | string {
    // ВАЖНО: Колонка G (6) содержит ДАТЫ в формате Excel serial number, НЕ просмотры!
    // Проверяем только правильные колонки для просмотров
    
    // Проверяем колонку K (10) - там чаще всего просмотры
    if (row[10] && typeof row[10] === 'number' && row[10] > 100 && row[10] < 1000000) {
      return Math.round(row[10]);
    }
    
    // Проверяем колонку L (11) 
    if (row[11] && typeof row[11] === 'number' && row[11] > 100 && row[11] < 1000000) {
      return Math.round(row[11]);
    }
    
    // Проверяем колонку M (12)
    if (row[12] && typeof row[12] === 'number' && row[12] > 100 && row[12] < 1000000) {
      return Math.round(row[12]);
    }
    
    // НЕ проверяем колонку G (6) - там даты!
    
    return 'Нет данных';
  }

  private getCleanValue(value: any): string {
    if (value === null || value === undefined) return '';
    return value.toString().trim();
  }

  private extractMonthName(sheetName: string): string {
    const monthsMap: { [key: string]: string } = {
      'март': 'Март',
      'мар': 'Март',
      'march': 'Март',
      'mar': 'Март',
      'янв': 'Январь',
      'фев': 'Февраль',
      'апр': 'Апрель',
      'май': 'Май',
      'июн': 'Июнь',
      'июл': 'Июль',
      'авг': 'Август',
      'сен': 'Сентябрь',
      'окт': 'Октябрь',
      'ноя': 'Ноябрь',
      'дек': 'Декабрь'
    };
    
    const lowerSheetName = sheetName.toLowerCase();
    for (const [key, value] of Object.entries(monthsMap)) {
      if (lowerSheetName.includes(key)) {
        return value;
      }
    }
    
    return 'Март'; // По умолчанию
  }

  private generateFileName(originalFileName: string, monthName: string): string {
    const baseName = originalFileName.replace(/\.[^/.]+$/, ''); // Убираем расширение
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    return `${baseName}_${monthName}_2025_результат_${timestamp}.xlsx`;
  }

  private async createPerfectWorkbook(data: ProcessedData): Promise<any> {
    console.log('📝 Создание идеального форматированного отчета...');
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`${data.monthName} 2025`);

    // Настройка ширины колонок согласно образцу
    worksheet.columns = [
      { width: 60 }, // Площадка - широкая колонка
      { width: 30 }, // Тема
      { width: 80 }, // Текст сообщения - самая широкая
      { width: 12 }, // Дата
      { width: 20 }, // Ник
      { width: 12 }, // Просмотры
      { width: 15 }, // Вовлечение
      { width: 10 }, // Тип поста
      { width: 8 },  // Отзыв
      { width: 12 }, // Упоминание
      { width: 15 }, // Поддерживающее
      { width: 8 }   // Всего
    ];

    // Создание шапки отчета с правильными цветами
    this.createPerfectHeader(worksheet, data);

    // Создание заголовков таблицы
    this.createPerfectTableHeaders(worksheet, data);

    let currentRow = 5;

    // Добавление отзывов
    currentRow = this.addPerfectSection(worksheet, 'Отзывы', data.reviews, currentRow);
    
    // Добавление комментариев топ-20
    currentRow = this.addPerfectSection(worksheet, 'Комментарии Топ-20 выдачи', data.comments, currentRow);
    
    // Добавление активных обсуждений (если есть)
    if (data.activeDiscussions.length > 0) {
      currentRow = this.addPerfectSection(worksheet, 'Активные обсуждения (мониторинг)', data.activeDiscussions, currentRow);
    }

    // Добавление итоговой статистики
    this.addPerfectSummary(worksheet, data.statistics, currentRow + 2);

    console.log('✅ Идеальный отчет создан');
    return workbook;
  }

  private createPerfectHeader(worksheet: ExcelJS.Worksheet, data: ProcessedData): void {
    // Стили для шапки - фиолетовый фон как в образце
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
    
    // Строка 3: План с дополнительными колонками
    worksheet.mergeCells('A3:B3');
    worksheet.getCell('A3').value = 'План';
    worksheet.mergeCells('C3:H3');
    worksheet.getCell('C3').value = `Отзывы - ${data.statistics.reviewsCount}, Комментарии - ${data.statistics.commentsCount}`;
    
    // Дополнительные колонки в строке 3
    worksheet.getCell('I3').value = 'Отзыв';
    worksheet.getCell('J3').value = 'Упоминание';
    worksheet.getCell('K3').value = 'Поддерживающее';
    worksheet.getCell('L3').value = 'Всего';

    // Применение стилей к шапке
    for (let row = 1; row <= 3; row++) {
      for (let col = 1; col <= 12; col++) {
        const cell = worksheet.getCell(row, col);
        cell.fill = headerFill;
        cell.font = headerFont;
        cell.alignment = centerAlign;
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
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
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  }

  private addPerfectSection(worksheet: ExcelJS.Worksheet, sectionName: string, data: DataRow[], startRow: number): number {
    let currentRow = startRow;

    // Добавляем заголовок секции с голубым фоном
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    const sectionCell = worksheet.getCell(`A${currentRow}`);
    sectionCell.value = sectionName;
    sectionCell.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF000000' } };
    sectionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
    sectionCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    sectionCell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
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
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        
        if (colNumber === 4 && cell.value) { // Дата
          cell.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
        } else if (colNumber === 6) { // Просмотры
          cell.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        }
      });
      
      dataRow.height = 12;
      currentRow++;
    });
    
    return currentRow + 1;
  }

  private addPerfectSummary(worksheet: ExcelJS.Worksheet, statistics: ProcessingStats, startRow: number): void {
    const summaryData = [
      ['', '', '', '', '', `Суммарное количество просмотров*`, statistics.totalViews.toString()],
      ['', '', '', '', '', `Количество карточек товара (отзывы)`, statistics.reviewsCount.toString()],
      ['', '', '', '', '', `Количество обсуждений (форумы, сообщества, комментарии к статьям)`, statistics.commentsCount.toString()],
      ['', '', '', '', '', `Доля обсуждений с вовлечением в диалог`, `${statistics.engagementRate}%`],
      ['', '', '', '', '', '', ''],
      ['', '', '*Без учета площадок с закрытой статистикой прочтений', '', '', '', ''],
      ['', '', `Площадки со статистикой просмотров`, '', '', '', `${statistics.platformsWithData}%`],
      ['', '', 'Количество прочтений увеличивается в среднем на 30% в течение 3 месяцев, следующих за публикацией.', '', '', '', '']
    ];

    summaryData.forEach((rowData, index) => {
      const row = worksheet.getRow(startRow + index);
      row.values = rowData;
      
      // Специальное форматирование для строки с процентами (желтый фон)
      if (index === 6) {
        for (let col = 1; col <= 8; col++) {
          const cell = row.getCell(col);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
          cell.font = { name: 'Arial', size: 9, bold: true };
        }
      }
      
      row.eachCell((cell: any) => {
        if (index !== 6) {
          cell.font = { name: 'Arial', size: 9 };
        }
        cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      });
    });
  }
}

export const perfectProcessor = new ExcelProcessorPerfect(); 