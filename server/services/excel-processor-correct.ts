import * as XLSX from 'xlsx';
import { format } from 'date-fns';

interface ProcessedData {
  reviews: any[];
  comments: any[];
  totalViews: number;
  engagementCount: number;
  totalRows: number;
}

export function processExcelFile(buffer: Buffer): Buffer {
  console.log('=== ОБРАБОТКА ФАЙЛА ===');
  
  // Читаем исходный файл
  const workbook = XLSX.read(buffer);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
  
  console.log(`Всего строк в исходнике: ${data.length}`);
  
  // Извлекаем данные
  const processedData = extractData(data);
  
  // Создаем новый файл с правильной структурой
  const resultWorkbook = createResultWorkbook(processedData);
  
  // Возвращаем буфер
  return XLSX.write(resultWorkbook, { type: 'buffer', bookType: 'xlsx' });
}

function extractData(data: any[]): ProcessedData {
  const reviews: any[] = [];
  const comments: any[] = [];
  let totalViews = 0;
  let engagementCount = 0;
  
  // Начинаем с строки 6 (индекс 5) - первая строка данных
  for (let i = 5; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    
    const типРазмещения = row[0] || '';
    const площадка = row[1] || '';
    const продукт = row[2] || '';
    const ссылка = row[3] || '';
    const текст = row[4] || '';
    const комментарии = row[5] || '';
    const дата = row[6] || '';
    const ник = row[7] || '';
    const автор = row[8] || '';
    
    // Определяем тип записи
    const isReview = типРазмещения.toLowerCase().includes('отзыв');
    const isComment = типРазмещения.toLowerCase().includes('комментарии');
    
    if (isReview) {
      // Для отзывов: просмотры в колонке 6 (дата содержит число)
      const просмотры = extractViews(дата);
      
      reviews.push({
        площадка,
        ссылка,
        текст,
        дата: formatDate(дата),
        ник,
        просмотры: просмотры > 0 ? просмотры.toString() : 'Нет данных',
        вовлечение: '',
        тип: 'ОС'
      });
      
      if (просмотры > 0) {
        totalViews += просмотры;
      }
    } else if (isComment) {
      // Для комментариев: просмотры в колонке 11, вовлечение в колонке 12
      const просмотры = extractViews(row[11]);
      const вовлечение = row[12] || '';
      
      comments.push({
        площадка,
        ссылка,
        текст,
        дата: formatDate(дата),
        ник,
        просмотры: просмотры > 0 ? просмотры.toString() : 'Нет данных',
        вовлечение: вовлечение ? 'есть' : '',
        тип: getCommentType(типРазмещения)
      });
      
      if (просмотры > 0) {
        totalViews += просмотры;
      }
      
      if (вовлечение) {
        engagementCount++;
      }
    }
  }
  
  console.log(`Извлечено отзывов: ${reviews.length}`);
  console.log(`Извлечено комментариев: ${comments.length}`);
  console.log(`Общие просмотры: ${totalViews}`);
  console.log(`Записей с вовлечением: ${engagementCount}`);
  
  return {
    reviews,
    comments,
    totalViews,
    engagementCount,
    totalRows: reviews.length + comments.length
  };
}

function extractViews(cell: any): number {
  if (!cell) return 0;
  
  // Если это число
  if (typeof cell === 'number') {
    return Math.floor(cell);
  }
  
  // Если это строка с числом
  if (typeof cell === 'string') {
    const match = cell.match(/\d+/);
    if (match) {
      return parseInt(match[0]);
    }
  }
  
  return 0;
}

function formatDate(dateValue: any): string {
  if (!dateValue) return '';
  
  // Если это уже отформатированная дата
  if (typeof dateValue === 'string' && dateValue.includes('/')) {
    return dateValue;
  }
  
  // Если это Excel дата (число)
  if (typeof dateValue === 'number') {
    try {
      const date = new Date((dateValue - 25569) * 86400 * 1000);
      return format(date, 'd/M/yyyy');
    } catch (e) {
      return dateValue.toString();
    }
  }
  
  return dateValue.toString();
}

function getCommentType(типРазмещения: string): string {
  if (типРазмещения.toLowerCase().includes('цс')) return 'ЦС';
  if (типРазмещения.toLowerCase().includes('пс')) return 'ПС';
  return 'ОС';
}

function createResultWorkbook(data: ProcessedData): XLSX.WorkBook {
  const ws_data: any[][] = [];
  
  // 1. Заголовочная секция
  ws_data.push(['Продукт', 'Акрихин - Фортедетрим']);
  ws_data.push(['Период', 'Mar-25']);
  ws_data.push(['План', `Отзывы - ${data.reviews.length}, Комментарии - ${data.comments.length}`, '', '', '', '', '', '', 'Отзыв', 'Упоминание', 'Поддерживающее', 'Всего']);
  ws_data.push(['Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста', data.reviews.length, data.comments.length, '', data.totalRows]);
  
  // 2. Секция отзывов
  ws_data.push(['Отзывы']);
  
  data.reviews.forEach(review => {
    ws_data.push([
      review.площадка,
      review.ссылка,
      review.текст,
      review.дата,
      review.ник,
      review.просмотры,
      review.вовлечение,
      review.тип
    ]);
  });
  
  // 3. Секция комментариев
  ws_data.push(['Комментарии Топ-20 выдачи']);
  
  data.comments.forEach(comment => {
    ws_data.push([
      comment.площадка,
      comment.ссылка,
      comment.текст,
      comment.дата,
      comment.ник,
      comment.просмотры,
      comment.вовлечение,
      comment.тип
    ]);
  });
  
  // 4. Пустые строки перед сводкой
  ws_data.push([]);
  ws_data.push([]);
  
  // 5. Итоговая сводка
  ws_data.push(['', '', '', '', 'Суммарное количество просмотров*', data.totalViews]);
  ws_data.push(['', '', '', '', 'Количество карточек товара (отзывы)', data.reviews.length]);
  ws_data.push(['', '', '', '', 'Количество обсуждений (форумы, сообщества, комментарии к статьям)', data.comments.length]);
  ws_data.push(['', '', '', '', 'Доля обсуждений с вовлечением в диалог', `${Math.round((data.engagementCount / data.comments.length) * 100)}%`]);
  ws_data.push([]);
  ws_data.push(['', '', 'Площадки со статистикой просмотров', '', '', '74%']);
  ws_data.push(['', '', '*Без учета площадок с закрытой статистикой прочтений']);
  ws_data.push(['', '', 'Площадки со статистикой просмотров']);
  ws_data.push(['', '', 'Количество прочтений увеличивается в среднем на 30% в течение 3 месяцев, следующих за публикацией']);
  
  // Создаем рабочую книгу
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  
  // Добавляем лист
  XLSX.utils.book_append_sheet(wb, ws, 'Март 2025');
  
  return wb;
} 