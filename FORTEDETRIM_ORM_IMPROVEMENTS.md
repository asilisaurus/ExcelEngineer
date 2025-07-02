# Fortedetrim ORM Report Analysis & Fix

## Обнаруженные проблемы

### 1. Неправильное извлечение данных о просмотрах
- **Проблема**: Данные о просмотрах находятся в колонках J (9), K (10), L (11), но скрипт искал их в других местах
- **Текущий результат**: 173,243 просмотра
- **Ожидаемый результат**: 3,398,560 просмотров

### 2. Структура данных в исходном файле
Анализ показал следующую структуру:
- **Колонка A (0)**: Тип записи ("Отзывы (отзовики)", "Отзывы (аптеки)", "Комментарии в обсуждениях")
- **Колонка B (1)**: URL площадки
- **Колонка C (2)**: Тема/заголовок (если есть)
- **Колонка E (4)**: Текст сообщения
- **Колонка G (6)**: Дата (Excel формат: 45719, 45720, etc.)
- **Колонка H (7)**: Ник пользователя
- **Колонки J-L (9-11)**: Различные типы просмотров
  - J (9): Просмотры темы на старте
  - K (10): Просмотры в конце месяца  
  - L (11): Просмотров получено
- **Колонка M (12)**: Вовлечение ("есть", пустые значения)
- **Колонка N (13)**: Тип записи ("ОС", "ЦС")

### 3. Примеры реальных данных
```
Row 50: Дата=45719, Просмотры=12000/50600/38600, Вовлечение="есть"
Row 51: Дата=45719, Просмотры=7452/50600/43148
Row 75: Дата=45719, Просмотры=60000/75300/15300
```

## Необходимые исправления

### 1. Исправить извлечение просмотров
```typescript
// В функции extractRowFromRealStructure
// СТАРЫЙ КОД:
for (let col = 8; col <= 15; col++) {
  if (row[col] && String(row[col]).trim() !== '') {
    const viewsValue = this.cleanViews(row[col]);
    if (typeof viewsValue === 'number') {
      просмотры = viewsValue;
      break;
    }
  }
}

// НОВЫЙ КОД:
// Извлекаем максимальное значение из колонок J, K, L (9, 10, 11)
let maxViews = 0;
for (let col = 9; col <= 11; col++) {
  if (row[col] && String(row[col]).trim() !== '') {
    const viewsValue = this.cleanViews(row[col]);
    if (typeof viewsValue === 'number' && viewsValue > maxViews) {
      maxViews = viewsValue;
    }
  }
}
просмотры = maxViews > 0 ? maxViews : 'Нет данных';
```

### 2. Исправить обработку дат
```typescript
private formatDate(dateValue: any): string {
  if (!dateValue) return '';
  
  // Обработка Excel дат (45719 = 25.03.2025)
  if (typeof dateValue === 'number' && dateValue > 40000) {
    const date = new Date((dateValue - 25569) * 86400 * 1000);
    return date.toLocaleDateString('ru-RU');
  }
  
  return String(dateValue);
}
```

### 3. Исправить извлечение названий площадок
```typescript
private extractPlatformName(url: string): string {
  if (!url) return '';
  
  const url_clean = url.trim().toLowerCase();
  if (url_clean.includes('otzovik.com')) return 'Отзовик';
  if (url_clean.includes('irecommend.ru')) return 'Irecommend';
  if (url_clean.includes('market.yandex.ru')) return 'Яндекс.Маркет';
  if (url_clean.includes('dzen.ru')) return 'Дзен';
  if (url_clean.includes('vk.com')) return 'ВКонтакте';
  if (url_clean.includes('woman.ru')) return 'Woman.ru';
  // ... и т.д.
  
  try {
    const domain = new URL(url).hostname;
    return domain.replace('www.', '');
  } catch {
    return url;
  }
}
```

### 4. Ожидаемые результаты после исправления
- **Общие просмотры**: 3,398,560
- **Отзывы**: 18-22
- **Комментарии**: 519-650  
- **Вовлечение**: 20%
- **Площадки со статистикой**: 74%

## Статус
- ✅ Проблема идентифицирована
- 🔄 Требуется обновление логики извлечения данных
- ⏳ Тестирование после исправлений

## Следующие шаги
1. Обновить функцию `extractRowFromRealStructure` для правильного извлечения просмотров
2. Протестировать на исходном файле
3. Проверить соответствие результата эталону
4. Провести окончательное тестирование