import * as XLSX from 'xlsx';
import { format } from 'date-fns';

interface ProcessedData {
  reviews: any[];
  topComments: any[];
  activeDiscussions: any[];
  totalViews: number;
  engagementCount: number;
  totalRows: number;
}

export function processExcelFile(buffer: Buffer): Buffer {
  console.log('=== ОБРАБОТКА ФАЙЛА (ФИНАЛЬНАЯ ВЕРСИЯ) ===');
  
  // Читаем исходный файл
  const workbook = XLSX.read(buffer);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
  
  console.log(`Всего строк в исходнике: ${data.length}`);
  
  // Извлекаем данные
  const processedData = extractDataCorrectly(data);
  
  // Создаем новый файл с правильной структурой
  const resultWorkbook = createResultWorkbook(processedData);
  
  // Возвращаем буфер
  return XLSX.write(resultWorkbook, { bookType: 'xlsx', type: 'buffer' });
}

function extractDataCorrectly(data: any[]): ProcessedData {
  console.log('=== ИЗВЛЕЧЕНИЕ ДАННЫХ ===');
  
  const reviews: any[] = [];
  const topComments: any[] = [];
  const activeDiscussions: any[] = [];
  let totalViews = 0;
  let engagementCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;
    
    const typeCell = row[0].toString().toLowerCase();
    
    // Отзывы (отзовики и аптеки)
    if (typeCell.includes('отзывы') && (typeCell.includes('отзовики') || typeCell.includes('аптеки'))) {
      const reviewData = {
        platform: extractPlatformFromUrl(row[1] || ''),
        topic: row[3] || '',
        text: row[4] || '',
        date: formatDate(row[6] || ''),
        username: row[7] || '',
        views: 'Нет данных',
        engagement: '',
        postType: 'ОС'
      };
      
      reviews.push(reviewData);
      console.log(`Отзыв добавлен: ${reviewData.username}`);
    }
    
    // Комментарии в обсуждениях (первые 20 для секции "Топ-20")
    else if (typeCell.includes('комментарии') && typeCell.includes('обсуждениях')) {
      const commentData = {
        platform: extractPlatformFromUrl(row[1] || ''),
        topic: row[3] || '',
        text: row[4] || '',
        date: formatDate(row[6] || ''),
        username: row[7] || '',
        views: row[11] || 'Нет данных',
        engagement: row[12] || '',
        postType: 'ЦС'
      };
      
      // Первые 20 идут в "Комментарии Топ-20 выдачи"
      if (topComments.length < 20) {
        topComments.push(commentData);
        console.log(`Топ-комментарий добавлен: ${commentData.username}`);
      } else {
        // Остальные в "Активные обсуждения"
        activeDiscussions.push(commentData);
      }
      
      // Подсчет просмотров
      if (commentData.views && commentData.views !== 'Нет данных') {
        const views = parseInt(commentData.views.toString());
        if (!isNaN(views)) {
          totalViews += views;
        }
      }
      
      // Подсчет вовлечения
      if (commentData.engagement && commentData.engagement.toString().toLowerCase().includes('есть')) {
        engagementCount++;
      }
    }
  }
  
  console.log(`Извлечено: ${reviews.length} отзывов, ${topComments.length} топ-комментариев, ${activeDiscussions.length} активных обсуждений`);
  console.log(`Всего просмотров: ${totalViews}, вовлечений: ${engagementCount}`);
  
  return {
    reviews,
    topComments,
    activeDiscussions,
    totalViews,
    engagementCount,
    totalRows: reviews.length + topComments.length + activeDiscussions.length
  };
}

function extractPlatformFromUrl(url: string): string {
  if (!url) return '';
  
  if (url.includes('otzovik.com')) return 'https://otzovik.com/reviews/fortedetrim/';
  if (url.includes('irecommend.ru')) return 'https://irecommend.ru/content/lekarstvennyi-preparat-ao-medana-farma-fortedetrim-vitamin-d';
  if (url.includes('market.yandex.ru')) return url;
  if (url.includes('dzen.ru')) return url;
  if (url.includes('vk.com')) return url;
  if (url.includes('ok.ru')) return url;
  if (url.includes('youtube.com')) return url;
  if (url.includes('t.me')) return url;
  if (url.includes('pikabu.ru')) return url;
  if (url.includes('woman.ru')) return url;
  if (url.includes('otvet.mail.ru')) return url;
  if (url.includes('mom.life')) return url;
  if (url.includes('babyblog.ru')) return url;
  if (url.includes('forum.baby.ru')) return url;
  
  return url;
}

function formatDate(dateValue: any): string {
  if (!dateValue) return '';
  
  const dateStr = dateValue.toString();
  
  // Если уже в формате даты
  if (dateStr.includes('/')) {
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      const day = parts[1];
      const month = parts[0];
      const year = parts[2];
      return `${day}/${month}/${year}`;
    }
  }
  
  return dateStr;
}

function createResultWorkbook(data: ProcessedData): XLSX.WorkBook {
  console.log('=== СОЗДАНИЕ РЕЗУЛЬТИРУЮЩЕГО ФАЙЛА ===');
  
  const wb = XLSX.utils.book_new();
  const ws_data: any[][] = [];
  
  // Заголовочная секция
  ws_data.push(['Продукт', 'Акрихин - Фортедетрим']);
  ws_data.push(['Период', 'Mar-25']);
  ws_data.push([
    'План',
    'Отзывы - 22, Комментарии - 650',
    '', '', '', '', '', '',
    'Отзыв', 'Упоминание', 'Поддерживающее', 'Всего'
  ]);
  ws_data.push([
    'Площадка', 'Тема', 'Текст сообщения', 'Дата', 'Ник', 'Просмотры', 'Вовлечение', 'Тип поста',
    data.reviews.length.toString(),
    (data.topComments.length + data.activeDiscussions.length).toString(),
    '', 
    data.totalRows.toString()
  ]);
  
  // Секция "Отзывы"
  ws_data.push(['Отзывы']);
  data.reviews.forEach(review => {
    ws_data.push([
      review.platform,
      review.topic,
      review.text,
      review.date,
      review.username,
      review.views,
      review.engagement,
      review.postType
    ]);
  });
  
  // Секция "Комментарии Топ-20 выдачи"
  ws_data.push(['Комментарии Топ-20 выдачи']);
  data.topComments.forEach(comment => {
    ws_data.push([
      comment.platform,
      comment.topic,
      comment.text,
      comment.date,
      comment.username,
      comment.views,
      comment.engagement,
      comment.postType
    ]);
  });
  
  // Секция "Активные обсуждения (мониторинг)"
  ws_data.push(['Активные обсуждения (мониторинг)']);
  data.activeDiscussions.forEach(discussion => {
    ws_data.push([
      discussion.platform,
      discussion.topic,
      discussion.text,
      discussion.date,
      discussion.username,
      discussion.views,
      discussion.engagement,
      discussion.postType
    ]);
  });
  
  // Пустые строки перед сводкой
  ws_data.push([]);
  ws_data.push([]);
  
  // Сводка
  ws_data.push(['', '', '', '', 'Суммарное количество просмотров*', data.totalViews.toString()]);
  ws_data.push(['', '', '', '', 'Количество карточек товара (отзывы)', data.reviews.length.toString()]);
  ws_data.push(['', '', '', '', 'Количество обсуждений (форумы, сообщества, комментарии к статьям)', 
    (data.topComments.length + data.activeDiscussions.length).toString()]);
  
  const engagementPercentage = data.totalRows > 0 ? 
    Math.round((data.engagementCount / data.totalRows) * 100) : 0;
  ws_data.push(['', '', '', '', 'Доля обсуждений с вовлечением в диалог', `${engagementPercentage}%`]);
  
  ws_data.push([]);
  ws_data.push(['', '', '*Без учета площадок с закрытой статистикой прочтений']);
  
  const platformsWithViews = data.topComments.concat(data.activeDiscussions)
    .filter(item => item.views && item.views !== 'Нет данных').length;
  const totalPlatforms = data.topComments.length + data.activeDiscussions.length;
  const platformsPercentage = totalPlatforms > 0 ? 
    Math.round((platformsWithViews / totalPlatforms) * 100) : 0;
  
  ws_data.push(['', '', 'Площадки со статистикой просмотров', '', '', `${platformsPercentage}%`]);
  ws_data.push(['', '', 'Количество прочтений увеличивается в среднем на 30% в течение 3 месяцев, следующих за публикацией.']);
  
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, 'Март 2025');
  
  console.log(`Создан файл с ${ws_data.length} строками`);
  
  return wb;
} 