import axios from 'axios';

interface ProgressCallback {
  (progress: number, message: string): void;
}

export async function importFromGoogleSheets(
  url: string, 
  progressCallback?: ProgressCallback
): Promise<Buffer> {
  console.log('=== ИМПОРТ ИЗ GOOGLE ТАБЛИЦ ===');
  console.log('URL:', url);
  
  if (progressCallback) {
    progressCallback(10, 'Подключение к Google Таблицам...');
  }
  
  // Извлекаем ID таблицы и GID листа из URL
  const { spreadsheetId, gid } = extractGoogleSheetsInfo(url);
  
  if (!spreadsheetId) {
    throw new Error('Неверный URL Google Таблиц');
  }
  
  if (progressCallback) {
    progressCallback(20, 'Формирование запроса к Google Таблицам...');
  }
  
  // Формируем URL для экспорта в формате Excel
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx${gid ? `&gid=${gid}` : ''}`;
  
  console.log('Export URL:', exportUrl);
  
  try {
    if (progressCallback) {
      progressCallback(30, 'Отправка запроса к Google Таблицам...');
    }
    
    // Загружаем данные
    const response = await axios.get(exportUrl, {
      responseType: 'arraybuffer',
      timeout: 30000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (progressCallback) {
      progressCallback(70, 'Данные загружены, обработка...');
    }
    
    console.log('Данные загружены, размер:', response.data.byteLength, 'байт');
    
    if (progressCallback) {
      progressCallback(90, 'Подготовка данных для обработки...');
    }
    
    // Возвращаем буфер
    return Buffer.from(response.data);
    
  } catch (error: any) {
    console.error('Ошибка при загрузке из Google Таблиц:', error.message);
    
    if (error.response?.status === 403) {
      throw new Error('Доступ к таблице запрещен. Убедитесь, что таблица доступна для просмотра.');
    } else if (error.response?.status === 404) {
      throw new Error('Таблица не найдена. Проверьте URL.');
    } else {
      throw new Error(`Ошибка загрузки: ${error.message}`);
    }
  }
}

function extractGoogleSheetsInfo(url: string): { spreadsheetId: string | null; gid: string | null } {
  try {
    // Различные форматы URL Google Таблиц
    const patterns = [
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)\/edit/,
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)\/edit#gid=(\d+)/
    ];
    
    let spreadsheetId: string | null = null;
    let gid: string | null = null;
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) {
        spreadsheetId = match[1];
        if (match[2]) {
          gid = match[2];
        }
        break;
      }
    }
    
    // Попытка извлечь GID из параметров URL
    if (!gid) {
      const gidMatch = url.match(/[#&]gid=(\d+)/);
      if (gidMatch) {
        gid = gidMatch[1];
      }
    }
    
    console.log('Извлечено:', { spreadsheetId, gid });
    
    return { spreadsheetId, gid };
    
  } catch (error) {
    console.error('Ошибка при разборе URL:', error);
    return { spreadsheetId: null, gid: null };
  }
}

export function validateGoogleSheetsUrl(url: string): boolean {
  return /docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9-_]+/.test(url);
} 