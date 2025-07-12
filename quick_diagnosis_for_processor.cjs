/**
 * 🔍 БЫСТРАЯ ДИАГНОСТИКА ПРОБЛЕМ ПРОЦЕССОРА
 * Анализ кода процессора для выявления проблем
 * 
 * Автор: AI Assistant + Background Agent bc-2954e872-79f8-4d41-b422-413e62f0b031
 * Цель: Найти конкретные проблемы в логике без запуска Google Apps Script
 */

const fs = require('fs');

// ==================== АНАЛИЗ КОДА ПРОЦЕССОРА ====================

function analyzeProcessorCode() {
  console.log('🔍 АНАЛИЗ КОДА ПРОЦЕССОРА');
  console.log('========================');
  
  const processorCode = fs.readFileSync('google-apps-script-processor-final.js', 'utf8');
  
  const issues = [];
  
  // 1. Анализ метода findSectionBoundaries
  console.log('\n1️⃣ АНАЛИЗ findSectionBoundaries():');
  if (processorCode.includes('findSectionBoundaries')) {
    console.log('✅ Метод найден');
    
    // Ищем логику определения разделов
    const sectionRegex = /findSectionBoundaries\([\s\S]*?\n\s*\}/;
    const sectionMatch = processorCode.match(sectionRegex);
    
    if (sectionMatch) {
      const sectionCode = sectionMatch[0];
      console.log('📋 Логика определения разделов:');
      
      // Проверяем на критические проблемы
      if (!sectionCode.includes('отзывы') || !sectionCode.includes('комментарии') || !sectionCode.includes('обсуждения')) {
        issues.push({
          severity: 'КРИТИЧЕСКАЯ',
          problem: 'Отсутствуют проверки на все разделы',
          location: 'findSectionBoundaries()',
          solution: 'Добавить проверки на "отзывы", "комментарии топ-20", "активные обсуждения"'
        });
        console.log('❌ Не все разделы проверяются');
      } else {
        console.log('✅ Основные разделы проверяются');
      }
      
      if (!sectionCode.includes('startRow') || !sectionCode.includes('endRow')) {
        issues.push({
          severity: 'КРИТИЧЕСКАЯ', 
          problem: 'Нет определения границ разделов',
          location: 'findSectionBoundaries()',
          solution: 'Добавить startRow и endRow для каждого раздела'
        });
        console.log('❌ Нет определения границ разделов');
      } else {
        console.log('✅ Границы разделов определяются');
      }
    }
  } else {
    issues.push({
      severity: 'КРИТИЧЕСКАЯ',
      problem: 'Метод findSectionBoundaries отсутствует',
      location: 'Весь процессор',
      solution: 'Создать метод findSectionBoundaries для определения границ разделов'
    });
    console.log('❌ Метод отсутствует');
  }
  
  // 2. Анализ метода processData
  console.log('\n2️⃣ АНАЛИЗ processData():');
  if (processorCode.includes('processData')) {
    console.log('✅ Метод найден');
    
    const processRegex = /processData\([\s\S]*?\n\s*\}/;
    const processMatch = processorCode.match(processRegex);
    
    if (processMatch) {
      const processCode = processMatch[0];
      
      // Проверяем использование findSectionBoundaries
      if (!processCode.includes('findSectionBoundaries')) {
        issues.push({
          severity: 'КРИТИЧЕСКАЯ',
          problem: 'processData не использует findSectionBoundaries',
          location: 'processData()',
          solution: 'Вызвать findSectionBoundaries для правильного определения разделов'
        });
        console.log('❌ Не использует findSectionBoundaries');
      } else {
        console.log('✅ Использует findSectionBoundaries');
      }
      
      // Проверяем обработку разделов по границам
      if (processCode.includes('currentSection') && !processCode.includes('section.startRow')) {
        issues.push({
          severity: 'ВЫСОКАЯ',
          problem: 'Обработка не по границам разделов',
          location: 'processData()',
          solution: 'Обрабатывать данные строго в пределах boundaries каждого раздела'
        });
        console.log('❌ Обработка не по границам разделов');
      }
    }
  } else {
    issues.push({
      severity: 'КРИТИЧЕСКАЯ',
      problem: 'Метод processData отсутствует',
      location: 'Весь процессор',
      solution: 'Создать метод processData'
    });
    console.log('❌ Метод отсутствует');
  }
  
  // 3. Анализ determinePostTypeBySection
  console.log('\n3️⃣ АНАЛИЗ determinePostTypeBySection():');
  if (processorCode.includes('determinePostTypeBySection')) {
    console.log('✅ Метод найден');
    
    const typeRegex = /determinePostTypeBySection\([\s\S]*?\n\s*\}/;
    const typeMatch = processorCode.match(typeRegex);
    
    if (typeMatch) {
      const typeCode = typeMatch[0];
      
      // Проверяем маппинг разделов на типы
      const hasReviewsMapping = typeCode.includes('reviews') && typeCode.includes('ОС');
      const hasCommentsMapping = typeCode.includes('commentsTop20') && typeCode.includes('ЦС');
      const hasDiscussionsMapping = typeCode.includes('activeDiscussions') && typeCode.includes('ПС');
      
      if (!hasReviewsMapping || !hasCommentsMapping || !hasDiscussionsMapping) {
        issues.push({
          severity: 'ВЫСОКАЯ',
          problem: 'Неполный маппинг разделов на типы постов',
          location: 'determinePostTypeBySection()',
          solution: 'Добавить правильный маппинг: reviews->ОС, commentsTop20->ЦС, activeDiscussions->ПС'
        });
        console.log('❌ Неполный маппинг разделов на типы');
      } else {
        console.log('✅ Правильный маппинг разделов на типы');
      }
    }
  } else {
    issues.push({
      severity: 'ВЫСОКАЯ',
      problem: 'Метод determinePostTypeBySection отсутствует',
      location: 'Весь процессор',
      solution: 'Создать метод для определения типов постов по разделам'
    });
    console.log('❌ Метод отсутствует');
  }
  
  // 4. Анализ isSectionHeader и isStatisticsRow
  console.log('\n4️⃣ АНАЛИЗ ВСПОМОГАТЕЛЬНЫХ МЕТОДОВ:');
  const hasSectionHeader = processorCode.includes('isSectionHeader');
  const hasStatisticsRow = processorCode.includes('isStatisticsRow');
  
  if (!hasSectionHeader) {
    issues.push({
      severity: 'СРЕДНЯЯ',
      problem: 'Отсутствует метод isSectionHeader',
      location: 'Весь процессор',
      solution: 'Создать метод для проверки заголовков разделов'
    });
    console.log('❌ isSectionHeader отсутствует');
  } else {
    console.log('✅ isSectionHeader найден');
  }
  
  if (!hasStatisticsRow) {
    issues.push({
      severity: 'СРЕДНЯЯ',
      problem: 'Отсутствует метод isStatisticsRow',
      location: 'Весь процессор',
      solution: 'Создать метод для проверки строк статистики'
    });
    console.log('❌ isStatisticsRow отсутствует');
  } else {
    console.log('✅ isStatisticsRow найден');
  }
  
  return issues;
}

// ==================== СОЗДАНИЕ ПЛАНА ИСПРАВЛЕНИЙ ====================

function createFixPlan(issues) {
  console.log('\n🔧 ПЛАН ИСПРАВЛЕНИЙ');
  console.log('==================');
  
  // Группируем по приоритету
  const critical = issues.filter(i => i.severity === 'КРИТИЧЕСКАЯ');
  const high = issues.filter(i => i.severity === 'ВЫСОКАЯ');
  const medium = issues.filter(i => i.severity === 'СРЕДНЯЯ');
  
  console.log(`\n🔥 КРИТИЧЕСКИЕ (${critical.length}):`);
  critical.forEach((issue, index) => {
    console.log(`${index + 1}. ${issue.problem}`);
    console.log(`   📍 ${issue.location}`);
    console.log(`   🔧 ${issue.solution}`);
    console.log('');
  });
  
  console.log(`\n⚠️ ВЫСОКИЕ (${high.length}):`);
  high.forEach((issue, index) => {
    console.log(`${index + 1}. ${issue.problem}`);
    console.log(`   📍 ${issue.location}`);
    console.log(`   🔧 ${issue.solution}`);
    console.log('');
  });
  
  console.log(`\n💡 СРЕДНИЕ (${medium.length}):`);
  medium.forEach((issue, index) => {
    console.log(`${index + 1}. ${issue.problem}`);
    console.log(`   📍 ${issue.location}`);
    console.log(`   🔧 ${issue.solution}`);
    console.log('');
  });
  
  return { critical, high, medium };
}

// ==================== ГЕНЕРАЦИЯ ИСПРАВЛЕНИЙ ====================

function generateFixes(issues) {
  console.log('\n💻 КОНКРЕТНЫЕ ИСПРАВЛЕНИЯ КОДА');
  console.log('==============================');
  
  // Исправление 1: findSectionBoundaries
  const needsSectionBoundaries = issues.some(i => i.location.includes('findSectionBoundaries'));
  if (needsSectionBoundaries) {
    console.log('\n1️⃣ ИСПРАВЛЕНИЕ findSectionBoundaries():');
    console.log(`
  findSectionBoundaries(data) {
    const sections = [];
    let currentSection = null;
    let sectionStart = -1;
    
    for (let i = CONFIG.STRUCTURE.dataStartRow - 1; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').toLowerCase().trim();
      
      // Определяем тип раздела
      let sectionType = null;
      if (firstCell.includes('отзывы') && !firstCell.includes('топ-20') && !firstCell.includes('обсуждения')) {
        sectionType = 'reviews';
      } else if (firstCell.includes('комментарии топ-20') || firstCell.includes('топ-20 выдачи')) {
        sectionType = 'commentsTop20';
      } else if (firstCell.includes('активные обсуждения') || firstCell.includes('мониторинг')) {
        sectionType = 'activeDiscussions';
      }
      
      // Если найден новый раздел
      if (sectionType && sectionType !== currentSection) {
        // Закрываем предыдущий раздел
        if (currentSection && sectionStart !== -1) {
          sections.push({
            type: currentSection,
            startRow: sectionStart,
            endRow: i - 1
          });
        }
        
        // Начинаем новый раздел
        currentSection = sectionType;
        sectionStart = i + 1; // Данные начинаются со следующей строки
      }
    }
    
    // Закрываем последний раздел
    if (currentSection && sectionStart !== -1) {
      sections.push({
        type: currentSection,
        startRow: sectionStart,
        endRow: data.length - 1
      });
    }
    
    return sections;
  }`);
  }
  
  // Исправление 2: processData
  const needsProcessData = issues.some(i => i.location.includes('processData'));
  if (needsProcessData) {
    console.log('\n2️⃣ ИСПРАВЛЕНИЕ processData():');
    console.log(`
  processData(data) {
    const processedRecords = [];
    const columnMapping = this.getColumnMapping();
    
    // Получаем границы разделов
    const sections = this.findSectionBoundaries(data);
    
    // Обрабатываем каждый раздел отдельно
    for (const section of sections) {
      console.log(\`📂 Обработка раздела "\${section.type}" (строки \${section.startRow}-\${section.endRow})\`);
      
      for (let i = section.startRow; i <= section.endRow; i++) {
        if (i >= data.length) break;
        
        const row = data[i];
        
        // Пропускаем пустые строки и заголовки
        if (this.isEmptyRow(row) || this.isSectionHeader(row) || this.isStatisticsRow(row)) {
          continue;
        }
        
        // Обрабатываем строку в контексте раздела
        const processedRow = this.processRow(row, section.type, columnMapping);
        if (processedRow) {
          processedRecords.push(processedRow);
        }
      }
    }
    
    return processedRecords;
  }`);
  }
}

// ==================== ГЛАВНАЯ ФУНКЦИЯ ====================

function main() {
  console.log('🔍 БЫСТРАЯ ДИАГНОСТИКА ПРОЦЕССОРА');
  console.log('=================================');
  
  // 1. Анализ кода
  const issues = analyzeProcessorCode();
  
  // 2. План исправлений
  const plan = createFixPlan(issues);
  
  // 3. Генерация исправлений
  generateFixes(issues);
  
  // 4. Итоги
  console.log('\n🎯 ИТОГИ ДИАГНОСТИКИ:');
  console.log('====================');
  console.log(`🔥 Критических проблем: ${plan.critical.length}`);
  console.log(`⚠️ Высоких проблем: ${plan.high.length}`); 
  console.log(`💡 Средних проблем: ${plan.medium.length}`);
  console.log(`📊 Всего проблем: ${issues.length}`);
  
  const priority = plan.critical.length > 0 ? 'КРИТИЧЕСКИЙ' : 
                  plan.high.length > 0 ? 'ВЫСОКИЙ' : 'СРЕДНИЙ';
  console.log(`⚡ Приоритет исправлений: ${priority}`);
  
  return { issues, plan };
}

// Запуск
if (require.main === module) {
  main();
}

module.exports = { main }; 