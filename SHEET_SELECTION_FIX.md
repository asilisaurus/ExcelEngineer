# 🔧 Исправление проблемы выбора листа

## 🚨 **Проблема**

Обнаружено несоответствие между определением месяца и выбором листа:

1. **detectMonthIntelligently** находит первый лист с месяцем (например, "Июнь25")
2. **findDataSheet** выбирает лист с наибольшим количеством строк (например, "Апр24")
3. **Результат**: файл называется "Июнь_2025", но содержит данные за апрель

## ✅ **Решение**

### 1. Исправлена логика выбора листа

```typescript
private findDataSheet(workbook: XLSX.WorkBook, monthInfo: MonthInfo, selectedSheet?: string): XLSX.WorkSheet {
  // Приоритет 0: Если явно выбран лист, используем его
  if (selectedSheet && sheetNames.includes(selectedSheet)) {
    console.log(`🎯 Используется выбранный лист: ${selectedSheet}`);
    return workbook.Sheets[selectedSheet];
  }
  
  // Приоритет 1: Точное совпадение с текущим месяцем
  const monthPatterns = [
    monthInfo.name.toLowerCase(),
    monthInfo.shortName.toLowerCase(),
    monthInfo.name.toLowerCase() + '25',
    monthInfo.shortName.toLowerCase() + '25'
  ];
  
  // Сначала ищем листы с совпадающим месяцем
  // Если не найден, выбираем лист с наибольшим количеством строк
}
```

### 2. Добавлен API для получения списка листов

```typescript
// POST /api/get-sheets
app.post("/api/get-sheets", (req, res) => {
  // Анализирует файл и возвращает список листов с информацией:
  // - name: название листа
  // - rowCount: количество строк
  // - hasData: есть ли данные
});
```

### 3. Обновлен процессор для поддержки выбранного листа

```typescript
async processExcelFile(
  input: string | Buffer, 
  fileName?: string,
  selectedSheet?: string  // Новый параметр
): Promise<{ outputPath: string; statistics: ProcessingStats }>
```

## 📋 **Логика выбора листа (по приоритету)**

1. **Выбранный пользователем лист** (если указан)
2. **Лист с совпадающим месяцем** (по шаблонам)
3. **Лист с наибольшим количеством строк** (fallback)

## 🔄 **Шаблоны поиска месяца**

```typescript
const monthPatterns = [
  monthInfo.name.toLowerCase(),     // "июнь"
  monthInfo.shortName.toLowerCase(), // "июн"
  monthInfo.name.toLowerCase() + '25', // "июнь25"
  monthInfo.shortName.toLowerCase() + '25', // "июн25"
  monthInfo.name.toLowerCase() + '24', // "июнь24"
  monthInfo.shortName.toLowerCase() + '24'  // "июн24"
];
```

## 📊 **Логирование**

Добавлено подробное логирование для отслеживания выбора листа:

```
📋 Найдены листы: Инструкция, Июнь25, Май25, Апр25, Март25...
🎯 Используется выбранный лист: Июнь25
📅 Найден лист с совпадающим месяцем: Июнь25 (156 строк)
📋 Выбран лист: Июнь25 (156 строк)
```

## 🎯 **Ожидаемый результат**

- ✅ Правильное соответствие месяца и данных
- ✅ Возможность выбора листа в интерфейсе
- ✅ Приоритет пользовательского выбора над автоматическим
- ✅ Подробное логирование для отладки