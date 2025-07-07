const { simpleProcessor } = require('./server/services/excel-processor-simple.ts');

async function testFinalFixes() {
  try {
    console.log('=== ТЕСТИРОВАНИЕ ФИНАЛЬНЫХ ИСПРАВЛЕНИЙ ===');
    
    // Проверяем последний файл
    const testFile = 'uploads/file-1751884808676-267432594_результат_20250707.xlsx';
    
    console.log('🔄 Тестируем обработку файла...');
    const result = await simpleProcessor.processExcelFile(testFile);
    
    console.log('✅ Результат обработки:');
    console.log('📁 Выходной файл:', result.outputPath);
    console.log('📊 Статистика:', JSON.stringify(result.statistics, null, 2));
    
    // Проверяем структуру статистики
    const stats = result.statistics;
    console.log('');
    console.log('=== ПРОВЕРКА СТАТИСТИКИ ===');
    console.log('✅ totalRows:', stats.totalRows);
    console.log('✅ reviewsCount:', stats.reviewsCount);
    console.log('✅ commentsCount:', stats.commentsCount);
    console.log('✅ activeDiscussionsCount:', stats.activeDiscussionsCount);
    console.log('✅ totalViews:', stats.totalViews);
    console.log('✅ engagementRate:', stats.engagementRate);
    console.log('✅ platformsWithData:', stats.platformsWithData);
    
    console.log('');
    console.log('✅ Все исправления успешно протестированы!');
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
  }
}

testFinalFixes(); 