import { type ClassValue, clsx } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 Б';
  
  const k = 1024;
  const sizes = ['Б', 'КБ', 'МБ', 'ГБ'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

export function formatNumber(num: number | undefined | null): string {
  if (num === null || num === undefined || isNaN(num)) {
    return 'не число';
  }
  return new Intl.NumberFormat('ru-RU').format(num);
}

export function formatDate(date: string | Date): string {
  if (!date) return 'Нет данных';
  const d = new Date(date);
  if (isNaN(d.getTime())) return 'Нет данных';
  return d.toLocaleDateString('ru-RU', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

export function getTimeAgo(date: string | Date): string {
  const now = new Date();
  const past = new Date(date);
  const diffInSeconds = Math.floor((now.getTime() - past.getTime()) / 1000);
  
  if (diffInSeconds < 60) {
    return 'только что';
  } else if (diffInSeconds < 3600) {
    const minutes = Math.floor(diffInSeconds / 60);
    return `${minutes} мин назад`;
  } else if (diffInSeconds < 86400) {
    const hours = Math.floor(diffInSeconds / 3600);
    return `${hours} ч назад`;
  } else {
    const days = Math.floor(diffInSeconds / 86400);
    return `${days} дн назад`;
  }
}

export function getProcessingProgress(status: string, rowsProcessed?: number, totalRows?: number): number {
  switch (status) {
    case 'processing':
      if (rowsProcessed !== undefined) {
        // Map processing stages to progress percentages
        switch (rowsProcessed) {
          case 0:
            return 10; // Initial stage
          case 1:
            return 30; // Reading/downloading
          case 2:
            return 70; // Processing
          case 3:
            return 90; // Final formatting
          default:
            return Math.min((rowsProcessed / (totalRows || 4)) * 100, 95);
        }
      }
      // If no specific progress data, show minimal progress
      return 25;
    case 'completed':
      return 100;
    case 'error':
      return 0;
    default:
      return 0;
  }
}
