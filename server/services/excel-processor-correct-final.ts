import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

interface ProcessedData {
  reviews: number;
  comments: number;
  totalViews: number;
  processedRecords: number;
  engagementRate: number;
}

export class ExcelProcessorCorrectFinal {
  private convertExcelDateToString(excelDate: number): string {
    if (!excelDate || typeof excelDate !== 'number') {
      return '';
    }
    
    // Excel date serial number to JavaScript Date
    const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
    
    // Format as DD.MM.YYYY
    const day = jsDate.getDate().toString().padStart(2, '0');
    const month = (jsDate.getMonth() + 1).toString().padStart(2, '0');
    const year = jsDate.getFullYear().toString();
    
    return `${day}.${month}.${year}`;
  }

  private extractTotalViews(headerRow: any[]): number {
    // Ищем в заголовке строку с просмотрами
    for (const cell of headerRow) {
      if (cell && typeof cell === 'string' && cell.includes('Просмотры:')) {
        const match = cell.match(/(\d+)/);
        if (match) {
          return parseInt(match[1]);
        }
      }
    }
    return 0;
  }

  private countDataInRange(data: any[][], startRow: number, endRow: number): number {
    let count = 0;
    for (let i = startRow; i <= endRow; i++) {
      if (data[i] && data[i].length > 0) {
        // Проверяем, есть ли данные в ключевых колонках B(1), D(3), E(4)
        const hasData = data[i][1] || data[i][3] || data[i][4];
        if (hasData) {
          count++;
        }
      }
    }
    return count;
  }

  private findCommentSections(data: any[][]): number {
    let commentCount = 0;
    
    // Ищем все строки с комментариями после строки 51
    for (let i = 51; i < data.length; i++) {
      if (data[i] && data[i].length > 0) {
        // Проверяем, есть ли данные в ключевых колонках
        const hasData = data[i][1] || data[i][3] || data[i][4];
        if (hasData) {
          commentCount++;
        }
      }
    }
    
    return commentCount;
  }

  async processExcelFile(filePath: string): Promise<string> {
    try {
      console.log('🔄 Начинаю обработку файла:', filePath);
      
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      
      console.log('📊 Листы в файле:', workbook.worksheets.map(ws => ws.name));
      
      // Находим лист с данными месяца
      const months = ["Янв25", "Фев25", "Мар25", "Март25", "Апр25", "Май25", "Июн25", 
                     "Июл25", "Авг25", "Сен25", "Окт25", "Ноя25", "Дек25"];
      
      const sourceSheet = workbook.worksheets.find(ws => 
        months.some(month => ws.name.includes(month))
      );
      
      if (!sourceSheet) {
        throw new Error('Не найден лист с данными месяца');
      }
      
      console.log('📋 Обрабатываю лист:', sourceSheet.name);
      
      // Конвертируем в массив данных
      const data: any[][] = [];
      sourceSheet.eachRow((row, rowNumber) => {
        const rowData: any[] = [];
        row.eachCell((cell, colNumber) => {
          rowData[colNumber - 1] = cell.value;
        });
        data[rowNumber - 1] = rowData;
      });
      
      console.log('📊 Всего строк данных:', data.length);
      
      // Извлекаем просмотры из заголовка
      const totalViews = this.extractTotalViews(data[0] || []);
      console.log('👀 Просмотры из заголовка:', totalViews.toLocaleString());
      
      // Подсчитываем отзывы по диапазонам
      const otzReviews = this.countDataInRange(data, 5, 14); // строки 6-15 (индексы 5-14)
      const aptReviews = this.countDataInRange(data, 14, 27); // строки 15-28 (индексы 14-27)
      const totalReviews = otzReviews + aptReviews;
      
      console.log('📝 Отзывы OTZ (строки 6-15):', otzReviews);
      console.log('📝 Отзывы APT (строки 15-28):', aptReviews);
      console.log('📝 Всего отзывов:', totalReviews);
      
      // Подсчитываем комментарии
      const top20Comments = this.countDataInRange(data, 30, 50); // строки 31-51 (индексы 30-50)
      const additionalComments = this.findCommentSections(data);
      const totalComments = top20Comments + additionalComments;
      
      console.log('💬 Комментарии ТОП-20 (строки 31-51):', top20Comments);
      console.log('💬 Дополнительные комментарии:', additionalComments);
      console.log('💬 Всего комментариев:', totalComments);
      
      // Создаем новый файл результата
      const resultWorkbook = new ExcelJS.Workbook();
      const resultSheet = resultWorkbook.addWorksheet('Март 2025');
      
      // Заголовок
      resultSheet.getCell('A1').value = 'Продукт';
      resultSheet.getCell('B1').value = 'Акрихин - Фортедетрим';
      
      // Период
      resultSheet.getCell('A2').value = 'Период';
      resultSheet.getCell('B2').value = 45717; // Excel date for March 2025
      
      // План
      resultSheet.getCell('A3').value = 'План';
      resultSheet.getCell('B3').value = `Отзывы - ${totalReviews}, Комментарии - ${totalComments}`;
      
      // Статистика
      resultSheet.getCell('I3').value = 'Отзыв';
      resultSheet.getCell('J3').value = 'Упоминание';
      resultSheet.getCell('K3').value = 'Поддерживающее';
      resultSheet.getCell('L3').value = 'Всего';
      
      // Заголовки таблицы
      resultSheet.getCell('A4').value = 'Площадка';
      resultSheet.getCell('B4').value = 'Тема';
      resultSheet.getCell('C4').value = 'Текст сообщения';
      resultSheet.getCell('D4').value = 'Дата';
      resultSheet.getCell('E4').value = 'Ник';
      resultSheet.getCell('F4').value = 'Просмотры';
      resultSheet.getCell('G4').value = 'Вовлечение';
      resultSheet.getCell('H4').value = 'Тип поста';
      resultSheet.getCell('I4').value = totalReviews;
      resultSheet.getCell('J4').value = totalComments;
      resultSheet.getCell('K4').value = 0; // Поддерживающее
      resultSheet.getCell('L4').value = totalReviews + totalComments;
      
      // Секция отзывов
      resultSheet.getCell('A5').value = 'Отзывы';
      let currentRow = 6;
      
      // Обрабатываем отзывы OTZ (строки 6-15)
      for (let i = 5; i < 15; i++) {
        if (data[i] && (data[i][1] || data[i][3] || data[i][4])) {
          const row = resultSheet.getRow(currentRow);
          
          row.getCell(1).value = data[i][1] || ''; // Площадка
          row.getCell(2).value = data[i][3] || ''; // Ссылка
          row.getCell(3).value = data[i][4] || ''; // Текст
          
          // Дата
          if (data[i][6] && typeof data[i][6] === 'number') {
            row.getCell(4).value = this.convertExcelDateToString(data[i][6]);
          }
          
          row.getCell(5).value = data[i][7] || ''; // Ник
          row.getCell(6).value = data[i][9] || 'Нет данных'; // Просмотры
          row.getCell(8).value = 'ОС'; // Тип поста
          
          currentRow++;
        }
      }
      
      // Обрабатываем отзывы APT (строки 15-28)
      for (let i = 14; i < 28; i++) {
        if (data[i] && (data[i][1] || data[i][3] || data[i][4])) {
          const row = resultSheet.getRow(currentRow);
          
          row.getCell(1).value = data[i][1] || ''; // Площадка
          row.getCell(2).value = data[i][3] || ''; // Ссылка
          row.getCell(3).value = data[i][4] || ''; // Текст
          
          // Дата
          if (data[i][6] && typeof data[i][6] === 'number') {
            row.getCell(4).value = this.convertExcelDateToString(data[i][6]);
          }
          
          row.getCell(5).value = data[i][7] || ''; // Ник
          row.getCell(6).value = data[i][9] || 'Нет данных'; // Просмотры
          row.getCell(8).value = 'ОС'; // Тип поста
          
          currentRow++;
        }
      }
      
      // Секция комментариев
      resultSheet.getCell(`A${currentRow}`).value = 'Комментарии';
      currentRow++;
      
      // Обрабатываем комментарии ТОП-20 (строки 31-51)
      for (let i = 30; i < 51; i++) {
        if (data[i] && (data[i][1] || data[i][3] || data[i][4])) {
          const row = resultSheet.getRow(currentRow);
          
          row.getCell(1).value = data[i][1] || ''; // Площадка
          row.getCell(2).value = data[i][3] || ''; // Ссылка
          row.getCell(3).value = data[i][4] || ''; // Текст
          
          // Дата
          if (data[i][6] && typeof data[i][6] === 'number') {
            row.getCell(4).value = this.convertExcelDateToString(data[i][6]);
          }
          
          row.getCell(5).value = data[i][7] || ''; // Ник
          row.getCell(6).value = data[i][9] || 'Нет данных'; // Просмотры
          row.getCell(8).value = 'ЦС'; // Тип поста
          
          currentRow++;
        }
      }
      
      // Обрабатываем дополнительные комментарии
      for (let i = 51; i < data.length; i++) {
        if (data[i] && (data[i][1] || data[i][3] || data[i][4])) {
          const row = resultSheet.getRow(currentRow);
          
          row.getCell(1).value = data[i][1] || ''; // Площадка
          row.getCell(2).value = data[i][3] || ''; // Ссылка
          row.getCell(3).value = data[i][4] || ''; // Текст
          
          // Дата
          if (data[i][6] && typeof data[i][6] === 'number') {
            row.getCell(4).value = this.convertExcelDateToString(data[i][6]);
          }
          
          row.getCell(5).value = data[i][7] || ''; // Ник
          row.getCell(6).value = data[i][9] || 'Нет данных'; // Просмотры
          row.getCell(8).value = 'ЦС'; // Тип поста
          
          currentRow++;
        }
      }
      
      // Форматирование
      this.applyFormatting(resultSheet, totalViews, totalReviews, totalComments);
      
      // Сохранение файла
      const originalFileName = path.basename(filePath, path.extname(filePath));
      const monthName = sourceSheet.name.includes('Март') ? 'Март' : 'Март';
      const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      const outputFileName = `${originalFileName}_${monthName}_2025_результат_${dateStr}.xlsx`;
      const outputPath = path.join(path.dirname(filePath), outputFileName);
      
      await resultWorkbook.xlsx.writeFile(outputPath);
      
      console.log('✅ Файл обработан успешно:', outputFileName);
      console.log('📊 Итоговая статистика:');
      console.log(`   Отзывы: ${totalReviews}`);
      console.log(`   Комментарии: ${totalComments}`);
      console.log(`   Просмотры: ${totalViews.toLocaleString()}`);
      console.log(`   Всего записей: ${totalReviews + totalComments}`);
      
      return outputPath;
      
    } catch (error) {
      console.error('❌ Ошибка при обработке файла:', error);
      throw error;
    }
  }

  private applyFormatting(worksheet: ExcelJS.Worksheet, totalViews: number, reviews: number, comments: number) {
    // Заголовки - фиолетовый фон
    const headerCells = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3', 'K3', 'L3'];
    headerCells.forEach(cellAddress => {
      const cell = worksheet.getCell(cellAddress);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF2D1B69' }
      };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    });
    
    // Заголовки таблицы - синий фон
    const tableCells = ['A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4', 'K4', 'L4'];
    tableCells.forEach(cellAddress => {
      const cell = worksheet.getCell(cellAddress);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC5D9F1' }
      };
      cell.font = { bold: true };
    });
    
    // Секционные заголовки - желтый фон
    const sectionHeaders = ['A5', 'A' + (reviews + 6)];
    sectionHeaders.forEach(cellAddress => {
      const cell = worksheet.getCell(cellAddress);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' }
      };
      cell.font = { bold: true };
    });
    
    // Итоговая статистика
    const statsRow = worksheet.getRow(4);
    statsRow.getCell(9).value = reviews;
    statsRow.getCell(10).value = comments;
    statsRow.getCell(11).value = 0;
    statsRow.getCell(12).value = reviews + comments;
    
    // Автоширина столбцов
    worksheet.columns.forEach(column => {
      column.width = 15;
    });
    
    // Особые ширины для текстовых столбцов
    worksheet.getColumn(3).width = 50; // Текст сообщения
    worksheet.getColumn(2).width = 30; // Ссылка
    
    // Границы для всех ячеек
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
  }
}

export const correctFinalProcessor = new ExcelProcessorCorrectFinal(); 