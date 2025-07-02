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
    if (!value || value === null || value === undefined) return '';
    return String(value).trim();
  }

  private cleanViews(value: any): number | string {
    if (!value || value === null || value === undefined) return 'Нет данных';
    
    // Если уже число, просто возвращаем его
    if (typeof value === 'number') {
      return value > 0 ? Math.round(value) : 'Нет данных';
    }
    
    const str = String(value).trim();
    if (str === '' || str === '-' || str.toLowerCase() === 'нет данных' || str === '0') {
      return 'Нет данных';
    }

    // Пытаемся преобразовать в число напрямую
    const num = parseFloat(str.replace(/[\s,'"]/g, ''));
    if (!isNaN(num) && num > 0) {
      return Math.round(num);
    }

    // Если не получилось, возвращаем как есть (возможно текстовое значение)
    return str;
  }

  async processExcelFile(buffer: Buffer, originalFileName: string): Promise<{ workbook: any; statistics: ProcessingStats }> {
    try {
      console.log('🔄 Начинаем обработку файла...');
      
      // 1. Читаем файл
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      console.log('📁 Файл прочитан, листы:', workbook.SheetNames);
      
      // 2. Находим лист с данными месяца
      const sheetName = this.findMonthSheet(workbook.SheetNames);
      if (!sheetName) {
        throw new Error(`Лист с данными месяца не найден. Доступные листы: ${workbook.SheetNames.join(', ')}`);
      }
      console.log('📊 Найден лист:', sheetName);
      
      // 3. Извлекаем данные из листа
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      console.log('📋 Извлечено строк:', jsonData.length);
      
      // 4. Используем фиксированные диапазоны для надежности
      const { reviews, comments, active, statistics } = this.extractDataByFixedRanges(jsonData);
      console.log('📂 Извлечено:', {
        отзывы: reviews.length,
        комментарии: comments.length,
        активные: active.length
      });
      
      // 5. Создаем итоговый отчет
      const outputWorkbook = await this.createFormattedReport(
        { отзывы: reviews, комментарии: comments, активные: active }, 
        statistics, 
        sheetName
      );
      console.log('📄 Отчет создан успешно!');
      
      return { workbook: outputWorkbook, statistics };
      
    } catch (error) {
      console.error('❌ Ошибка обработки файла:', error);
      throw new Error(`Ошибка обработки файла: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`);
    }
  }

  private findMonthSheet(sheetNames: string[]): string | null {
    const months = ['Янв25', 'Фев25', 'Мар25', 'Март25', 'Апр25', 'Май25', 'Июн25', 'Июл25', 'Авг25', 'Сен25', 'Окт25', 'Ноя25', 'Дек25'];
    return sheetNames.find(name => months.some(month => name.includes(month))) || null;
  }

  private extractDataByFixedRanges(jsonData: any[][]): { 
    reviews: DataRow[], 
    comments: DataRow[], 
    active: DataRow[], 
    statistics: ProcessingStats 
  } {
    // Ищем все данные динамически после заголовков
    const allData = this.extractAllDataFromFile(jsonData);
    
    // Группируем по уже определенным типам
    const reviews = allData.filter(row => row.типПоста === 'Отзывы');
    const comments = allData.filter(row => row.типПоста === 'Комментарии Топ-20 выдачи');
    const active = allData.filter(row => row.типПоста === 'Активные обсуждения (мониторинг)');

    console.log(`📊 Распределение данных: Отзывы=${reviews.length}, Комментарии=${comments.length}, Активные=${active.length}, Всего=${allData.length}`);
    
    // Рассчитываем статистику
    const totalViews = allData.reduce((sum, row) => {
      return sum + (typeof row.просмотры === 'number' ? row.просмотры : 0);
    }, 0);

    const recordsWithViews = allData.filter(row => 
      typeof row.просмотры === 'number' && row.просмотры > 0
    ).length;
    
    const platformsWithData = allData.length > 0 ? Math.round((recordsWithViews / allData.length) * 100) : 0;

    const discussionData = [...comments, ...active];
    const engagedDiscussions = discussionData.filter(row => 
      row.вовлечение && (
        row.вовлечение.toLowerCase().includes('есть') ||
        row.вовлечение.toLowerCase().includes('да') ||
        row.вовлечение.toLowerCase().includes('+') ||
        (row.вовлечение.trim() !== '' && row.вовлечение !== 'Нет данных')
      )
    ).length;
    const engagementRate = discussionData.length > 0 ? Math.round((engagedDiscussions / discussionData.length) * 100) : 0;

    const statistics: ProcessingStats = {
      totalRows: allData.length,
      reviewsCount: reviews.length,
      commentsCount: comments.length,
      activeDiscussionsCount: active.length,
      totalViews,
      engagementRate,
      platformsWithData
    };

    return { reviews, comments, active, statistics };
  }

  private extractAllDataFromFile(jsonData: any[][]): DataRow[] {
    const allData: DataRow[] = [];
    let currentSection = '';
    let dataStarted = false;
    
    // Ищем данные по всему файлу
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || !Array.isArray(row)) continue;

      // Проверяем, не является ли это заголовком секции
      const rowText = this.getCleanValue(row[0]).toLowerCase();
      if (rowText.includes('отзывы') && !rowText.includes('карточек')) {
        currentSection = 'Отзывы';
        dataStarted = true;
        continue;
      } else if (rowText.includes('комментарии') && rowText.includes('топ')) {
        currentSection = 'Комментарии Топ-20 выдачи';
        dataStarted = true;
        continue;
      } else if (rowText.includes('активные') && rowText.includes('обсуждения')) {
        currentSection = 'Активные обсуждения (мониторинг)';
        dataStarted = true;
        continue;
      }

      // Пропускаем заголовки и служебные строки
      if (this.isHeaderRow(row) || this.isSummaryRow(row)) {
        continue;
      }

      // Извлекаем данные из строки
      const площадка = this.getCleanValue(row[0]);
      const тема = this.getCleanValue(row[1]);
      const текст = this.getCleanValue(row[2]);

      // Если нашли строку с данными
      if ((площадка || тема || текст) && dataStarted) {
        const дата = row[3] || '';
        const ник = this.getCleanValue(row[4]);
        const вовлечение = this.getCleanValue(row[6]);
        const типПоста = currentSection || this.determinePostType(площадка, тема);

        allData.push({
          площадка,
          тема,
          текст,
          дата,
          ник,
          просмотры: this.cleanViews(row[5]),
          вовлечение,
          типПоста
        });
      }
    }
    
    console.log(`📋 Извлечено ${allData.length} строк данных из файла`);
    return allData;
  }

  private isHeaderRow(row: any[]): boolean {
    const rowText = row.join('').toLowerCase();
    return rowText.includes('площадка') || 
           rowText.includes('тема') || 
           rowText.includes('просмотры') ||
           rowText.includes('вовлечение') ||
           rowText.includes('продукт') ||
           rowText.includes('период') ||
           rowText.includes('план');
  }

  private isSummaryRow(row: any[]): boolean {
    const rowText = row.join('').toLowerCase();
    return rowText.includes('суммарное') || 
           rowText.includes('количество карточек') ||
           rowText.includes('количество обсуждений') ||
           rowText.includes('доля обсуждений') ||
           rowText.includes('без учета') ||
           rowText.includes('площадки со статистикой') ||
           rowText.includes('количество прочтений');
  }

  private determinePostType(площадка: string, тема: string): string {
    const area = площадка.toLowerCase();
    const topic = тема.toLowerCase();
    
    if (area.includes('otzovik') || area.includes('irecommend') || 
        area.includes('market.yandex') || area.includes('ozon') ||
        area.includes('goodapteka') || area.includes('megapteka') ||
        area.includes('uteka') || topic.includes('отзыв')) {
      return 'Отзывы';
    } else if (area.includes('dzen.ru') || area.includes('woman.ru') ||
               area.includes('forum.baby.ru') || area.includes('vk.com') ||
               topic.includes('коммент')) {
      return 'Комментарии Топ-20 выдачи';
    } else {
      return 'Активные обсуждения (мониторинг)';
    }
  }

  private async createFormattedReport(
    groupedData: { отзывы: DataRow[]; комментарии: DataRow[]; активные: DataRow[] },
    statistics: ProcessingStats,
    sheetName: string
  ): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    
    // Определяем название месяца
    const monthName = this.getMonthName(sheetName);
    const worksheet = workbook.addWorksheet(`${monthName} 2025`);

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

    // Создание шапки отчета
    this.createReportHeader(worksheet, monthName);

    // Создание заголовков таблицы
    this.createTableHeaders(worksheet);

    let currentRow = 5;

    // Добавление разделов данных
    currentRow = this.addDataSection(worksheet, 'Отзывы', groupedData.отзывы, currentRow);
    currentRow = this.addDataSection(worksheet, 'Комментарии Топ-20 выдачи', groupedData.комментарии, currentRow);
    currentRow = this.addDataSection(worksheet, 'Активные обсуждения (мониторинг)', groupedData.активные, currentRow);

    // Добавление итоговых показателей
    this.addSummaryMetrics(worksheet, statistics, currentRow + 2);

    return workbook;
  }

  private getMonthName(sheetName: string): string {
    const monthMap: { [key: string]: string } = {
      'Янв': 'Январь', 'Фев': 'Февраль', 'Мар': 'Март', 'Март': 'Март',
      'Апр': 'Апрель', 'Май': 'Май', 'Июн': 'Июнь', 'Июл': 'Июль',
      'Авг': 'Август', 'Сен': 'Сентябрь', 'Окт': 'Октябрь',
      'Ноя': 'Ноябрь', 'Дек': 'Декабрь'
    };

    for (const [abbr, full] of Object.entries(monthMap)) {
      if (sheetName.includes(abbr)) return full;
    }
    return 'Месяц';
  }

  private createReportHeader(worksheet: ExcelJS.Worksheet, monthName: string): void {
    const headerFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF2D1341' } };
    const headerFont = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    const centerAlign = { horizontal: 'center' as const, vertical: 'center' as const, wrapText: true };

    // Строка 1: Продукт
    worksheet.mergeCells('A1:B1');
    worksheet.getCell('A1').value = 'Продукт';
    worksheet.mergeCells('C1:H1');
    worksheet.getCell('C1').value = 'Акрихин - Фортедетрим';

    // Строка 2: Период
    worksheet.mergeCells('A2:B2');
    worksheet.getCell('A2').value = 'Период';
    worksheet.mergeCells('C2:H2');
    worksheet.getCell('C2').value = `${monthName} 2025`;

    // Строка 3: План
    worksheet.mergeCells('A3:B3');
    worksheet.getCell('A3').value = 'План';
    worksheet.mergeCells('C3:H3');
    worksheet.getCell('C3').value = 'Отзывы - 22, Комментарии - 650';

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
    const centerAlign = { horizontal: 'center' as const, vertical: 'center' as const, wrapText: true };

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
    sectionCell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
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
    const summaryFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFCE4D6' } };
    const summaryFont = { name: 'Arial', size: 9, bold: true };
    const leftAlign = { horizontal: 'left' as const, vertical: 'top' as const, wrapText: true };
    const centerAlign = { horizontal: 'center' as const, vertical: 'center' as const };

    const metrics = [
      ['Суммарное количество просмотров', statistics.totalViews],
      ['Количество карточек товара (отзывы)', statistics.reviewsCount],
      ['Количество обсуждений (форумы, сообщества, комментарии к статьям)', statistics.commentsCount + statistics.activeDiscussionsCount],
      ['Доля обсуждений с вовлечением в диалог', `${statistics.engagementRate}%`]
    ];

    metrics.forEach((metric, index) => {
      const rowNum = startRow + index;
      
      // Объединяем ячейки A-E для описания
      worksheet.mergeCells(`A${rowNum}:E${rowNum}`);
      worksheet.getCell(`A${rowNum}`).value = metric[0];
      worksheet.getCell(`F${rowNum}`).value = metric[1];

      // Применяем форматирование
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

      // Для процентов - специальное форматирование
      if (index === 3) {
        valueCell.value = statistics.engagementRate / 100;
        valueCell.numFmt = '0%';
      }

      worksheet.getRow(rowNum).height = 12;
    });

    // Добавляем сноски
    const footnoteRow = startRow + 6;
    worksheet.mergeCells(`A${footnoteRow}:F${footnoteRow}`);
    const footnoteCell = worksheet.getCell(`A${footnoteRow}`);
    footnoteCell.value = '*Без учета площадок с закрытой статистикой просмотров';
    footnoteCell.font = { name: 'Arial', size: 8, italic: true };
    footnoteCell.alignment = leftAlign;
    worksheet.getRow(footnoteRow).height = 12;

    const footnote2Row = startRow + 7;
    worksheet.mergeCells(`A${footnote2Row}:F${footnote2Row}`);
    const footnote2Cell = worksheet.getCell(`A${footnote2Row}`);
    footnote2Cell.value = `Площадки со статистикой просмотров                    ${statistics.platformsWithData}%`;
    footnote2Cell.font = { name: 'Arial', size: 8, bold: true };
    footnote2Cell.alignment = leftAlign;
    footnote2Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    worksheet.getRow(footnote2Row).height = 12;

    const footnote3Row = startRow + 8;
    worksheet.mergeCells(`A${footnote3Row}:F${footnote3Row}`);
    const footnote3Cell = worksheet.getCell(`A${footnote3Row}`);
    footnote3Cell.value = 'Количество прочтений увеличивается в среднем на 30% в течение 3 месяцев, следующих за публикацией.';
    footnote3Cell.font = { name: 'Arial', size: 8 };
    footnote3Cell.alignment = leftAlign;
    worksheet.getRow(footnote3Row).height = 12;
  }
}