# Google Apps Script Processor Deployment Instructions

## Quick Start Guide

### 1. Access Your Google Sheet
Open the Google Sheet containing your source data:
https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?usp=sharing

### 2. Open Apps Script Editor
- Click on **Extensions** → **Apps Script**
- This will open the script editor in a new tab

### 3. Update the Processor Script
1. Delete any existing code in the editor
2. Copy the entire contents of `google-apps-script-processor-final.js`
3. Paste it into the Apps Script editor
4. Click the **Save** button (disk icon)

### 4. Run the Processor
1. In the Apps Script editor, select the `onOpen` function from the dropdown
2. Click **Run**
3. Grant necessary permissions when prompted
4. Return to your Google Sheet
5. You should see a new menu: **📊 Обработка отчетов**
6. Click on it and select **🚀 Обработать текущий месяц**

### 5. Process Specific Months
1. Navigate to the sheet for the month you want to process (e.g., "Май25")
2. Use the menu to run the processor
3. A new spreadsheet will be created with the processed results
4. The URL will appear in a popup

### 6. Verify Results
Compare the output with the reference sheets:
- Май 2025 (эталон)
- Апрель 2025 (эталон)  
- Март 2025 (эталон)
- Февраль 2025 (эталон)

## Expected Output Format

The processor will create a new spreadsheet with:
- Header information (Product, Period, Plan)
- 8 columns of data:
  1. Площадка
  2. Тема
  3. Текст сообщения
  4. Дата
  5. Ник
  6. Просмотры
  7. Вовлечение
  8. Тип поста
- Three sections:
  - Отзывы
  - Комментарии Топ-20 выдачи
  - Активные обсуждения (мониторинг)
- Statistics block at the bottom

## Troubleshooting

If the processor doesn't work correctly:
1. Check the **View** → **Logs** in Apps Script editor for errors
2. Ensure you're on the correct sheet before running
3. Make sure permissions are granted
4. Verify the source data structure matches expectations