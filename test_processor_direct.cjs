const path = require('path');
const fs = require('fs');

async function testProcessorDirectly() {
  console.log('🔍 ПРЯМОЙ ТЕСТ ПРОЦЕССОРА');
  console.log('='.repeat(50));
  
  try {
    // Импортируем напрямую из TypeScript
    const { simpleProcessor } = await import('./server/services/excel-processor-simple.ts');
    
    // Находим исходный файл
    const sourceFile = path.join('uploads', 'Fortedetrim ORM report source.xlsx');
    
    if (!fs.existsSync(sourceFile)) {
      console.log('❌ Исходный файл не найден:', sourceFile);
      return;
    }
    
    console.log('📂 Найден исходный файл:', sourceFile);
    
    // Обрабатываем файл напрямую
    console.log('🔄 Обрабатываем файл...');
    const outputPath = await simpleProcessor.processExcelFile(sourceFile);
    
    console.log('✅ Файл обработан:', outputPath);
    
    // Анализируем результат
    if (fs.existsSync(outputPath)) {
      const ExcelJS = require('exceljs');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(outputPath);
      const worksheet = workbook.getWorksheet(1);
      
      let reviewCount = 0;
      let commentCount = 0;
      let discussionCount = 0;
      let currentSection = '';
      
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 4) return;
        
        const cellA = row.getCell(1).value;
        if (cellA && typeof cellA === 'string') {
          const cellStr = cellA.toString().trim();
          
          if (cellStr === 'Отзывы') {
            currentSection = 'reviews';
            return;
          } else if (cellStr.includes('Комментарии')) {
            currentSection = 'comments';
            return;
          } else if (cellStr.includes('Активные обсуждения')) {
            currentSection = 'discussions';
            return;
          }
        }
        
        const hasData = cellA && cellA !== 'Площадка' && 
                       !cellA.toString().startsWith('Суммарное');
        
        if (hasData) {
          if (currentSection === 'reviews') reviewCount++;
          else if (currentSection === 'comments') commentCount++;
          else if (currentSection === 'discussions') discussionCount++;
        }
      });
      
      console.log('\n📊 РЕЗУЛЬТАТ ПРЯМОГО ТЕСТА:');
      console.log(`📝 Отзывы: ${reviewCount}`);
      console.log(`💬 Комментарии: ${commentCount}`);
      console.log(`🔥 Активные обсуждения: ${discussionCount}`);
      console.log(`📋 Всего: ${reviewCount + commentCount + discussionCount}`);
      
      if (reviewCount === 18 && commentCount === 519 && discussionCount > 0) {
        console.log('\n🎉 ПРЯМОЙ ТЕСТ ПРОЙДЕН!');
        console.log('✅ Процессор работает правильно');
      } else {
        console.log('\n❌ ПРЯМОЙ ТЕСТ НЕ ПРОЙДЕН');
        console.log('⚠️ Процессор дает неправильные результаты');
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка при прямом тесте:', error.message);
  }
}

testProcessorDirectly(); 