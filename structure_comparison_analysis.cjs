const XLSX = require('xlsx');
const fs = require('fs');

function analyzeWorkbook(filename, description) {
    console.log(`\n${'='.repeat(80)}`);
    console.log(`АНАЛИЗ ${description.toUpperCase()}: ${filename}`);
    console.log(`${'='.repeat(80)}\n`);
    
    if (!fs.existsSync(filename)) {
        console.log(`❌ Файл ${filename} не найден!`);
        return null;
    }
    
    const workbook = XLSX.readFile(filename, {
        type: 'buffer',
        cellDates: true,
        cellNF: false,
        cellText: false,
        raw: false
    });
    
    const analysis = {
        filename,
        totalSheets: workbook.SheetNames.length,
        sheets: []
    };
    
    console.log(`📋 Общее количество листов: ${workbook.SheetNames.length}`);
    console.log(`📑 Названия листов: ${workbook.SheetNames.join(', ')}\n`);
    
    workbook.SheetNames.forEach((sheetName, index) => {
        console.log(`${'─'.repeat(60)}`);
        console.log(`📄 ЛИСТ ${index + 1}: "${sheetName}"`);
        console.log(`${'─'.repeat(60)}`);
        
        const worksheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        const maxRow = range.e.r + 1;
        const maxCol = range.e.c + 1;
        
        console.log(`📏 Размеры: ${maxRow} строк × ${maxCol} колонок`);
        console.log(`📍 Диапазон: ${worksheet['!ref'] || 'Пустой лист'}\n`);
        
        // Анализ первых 5 строк
        console.log(`🔍 ПЕРВЫЕ 5 СТРОК:`);
        for (let row = 0; row < Math.min(5, maxRow); row++) {
            const rowData = [];
            for (let col = 0; col < maxCol; col++) {
                const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
                const cell = worksheet[cellAddress];
                const value = cell ? (typeof cell.v === 'string' ? cell.v.trim() : cell.v) : '';
                rowData.push(value || '');
            }
            console.log(`   Строка ${row + 1}: [${rowData.slice(0, 10).map(v => `"${v}"`).join(', ')}${rowData.length > 10 ? '...' : ''}]`);
        }
        console.log();
        
        // Поиск ключевых разделов
        console.log(`🔍 ПОИСК КЛЮЧЕВЫХ РАЗДЕЛОВ:`);
        const sections = findSections(worksheet, maxRow, maxCol);
        sections.forEach(section => {
            console.log(`   ${section.type}: ${section.description} (строка ${section.row + 1})`);
        });
        console.log();
        
        // Анализ заголовков
        const headers = analyzeHeaders(worksheet, maxRow, maxCol);
        if (headers.found) {
            console.log(`📋 ЗАГОЛОВКИ НАЙДЕНЫ В СТРОКЕ ${headers.row + 1}:`);
            headers.columns.forEach((header, index) => {
                console.log(`   ${String.fromCharCode(65 + index)}: "${header}"`);
            });
            console.log();
        }
        
        // Поиск итоговых строк
        const totals = findTotalRows(worksheet, maxRow, maxCol);
        if (totals.length > 0) {
            console.log(`📊 ИТОГОВЫЕ СТРОКИ:`);
            totals.forEach(total => {
                console.log(`   Строка ${total.row + 1}: ${total.text}`);
            });
            console.log();
        }
        
        // Анализ данных
        const dataAnalysis = analyzeDataTypes(worksheet, headers, maxRow, maxCol);
        if (dataAnalysis) {
            console.log(`📈 АНАЛИЗ ДАННЫХ:`);
            Object.entries(dataAnalysis).forEach(([type, count]) => {
                console.log(`   ${type}: ${count}`);
            });
            console.log();
        }
        
        const sheetAnalysis = {
            name: sheetName,
            dimensions: {rows: maxRow, cols: maxCol},
            range: worksheet['!ref'],
            headers,
            sections,
            totals,
            dataAnalysis
        };
        
        analysis.sheets.push(sheetAnalysis);
    });
    
    return analysis;
}

function findSections(worksheet, maxRow, maxCol) {
    const sections = [];
    const sectionKeywords = [
        {keyword: 'отзыв', type: '📝 ОТЗЫВЫ'},
        {keyword: 'комментар', type: '💬 КОММЕНТАРИИ'},
        {keyword: 'обсужден', type: '🗣️ ОБСУЖДЕНИЯ'},
        {keyword: 'итого', type: '📊 ИТОГИ'},
        {keyword: 'total', type: '📊 TOTAL'},
        {keyword: 'сумма', type: '🧮 СУММЫ'},
        {keyword: 'инструкц', type: '📖 ИНСТРУКЦИИ'},
        {keyword: 'площадк', type: '🌐 ПЛОЩАДКИ'},
        {keyword: 'дата', type: '📅 ДАТЫ'},
        {keyword: 'статус', type: '📈 СТАТУСЫ'}
    ];
    
    for (let row = 0; row < maxRow; row++) {
        for (let col = 0; col < maxCol; col++) {
            const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
            const cell = worksheet[cellAddress];
            if (cell && cell.v && typeof cell.v === 'string') {
                const value = cell.v.toLowerCase().trim();
                
                sectionKeywords.forEach(({keyword, type}) => {
                    if (value.includes(keyword)) {
                        sections.push({
                            type,
                            row,
                            col,
                            description: cell.v.trim(),
                            keyword
                        });
                    }
                });
            }
        }
    }
    
    return sections;
}

function analyzeHeaders(worksheet, maxRow, maxCol) {
    // Ищем заголовки в первых 5 строках
    for (let headerRow = 0; headerRow < Math.min(5, maxRow); headerRow++) {
        const rowData = [];
        let hasHeaders = false;
        
        for (let col = 0; col < maxCol; col++) {
            const cellAddress = XLSX.utils.encode_cell({r: headerRow, c: col});
            const cell = worksheet[cellAddress];
            const value = cell ? (typeof cell.v === 'string' ? cell.v.trim() : cell.v) : '';
            rowData.push(value || '');
            
            // Проверяем, похоже ли на заголовок
            if (typeof value === 'string' && value.length > 1) {
                const headerKeywords = ['площадк', 'дата', 'текст', 'тип', 'статус', 'ссылк', 'примеч', 'просмотр', 'лайк'];
                if (headerKeywords.some(keyword => value.toLowerCase().includes(keyword))) {
                    hasHeaders = true;
                }
            }
        }
        
        if (hasHeaders) {
            return {
                found: true,
                row: headerRow,
                columns: rowData.filter(v => v !== '')
            };
        }
    }
    
    return {found: false};
}

function findTotalRows(worksheet, maxRow, maxCol) {
    const totals = [];
    const totalKeywords = ['итого', 'total', 'сумма', 'всего', 'количество'];
    
    for (let row = 0; row < maxRow; row++) {
        for (let col = 0; col < maxCol; col++) {
            const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
            const cell = worksheet[cellAddress];
            if (cell && cell.v && typeof cell.v === 'string') {
                const value = cell.v.toLowerCase().trim();
                
                if (totalKeywords.some(keyword => value.includes(keyword))) {
                    totals.push({
                        row,
                        col,
                        text: cell.v.trim()
                    });
                }
            }
        }
    }
    
    return totals;
}

function analyzeDataTypes(worksheet, headers, maxRow, maxCol) {
    if (!headers.found) return null;
    
    let reviews = 0, comments = 0, discussions = 0, dates = 0, links = 0;
    
    for (let row = headers.row + 1; row < maxRow; row++) {
        for (let col = 0; col < maxCol; col++) {
            const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
                const value = typeof cell.v === 'string' ? cell.v.toLowerCase().trim() : cell.v;
                
                if (typeof value === 'string') {
                    if (value.includes('ос') || value.includes('отзыв')) reviews++;
                    if (value.includes('цс') || value.includes('комментар')) comments++;
                    if (value.includes('пс') || value.includes('обсужден')) discussions++;
                    if (value.includes('http') || value.includes('www')) links++;
                }
                
                if (cell.v instanceof Date) dates++;
            }
        }
    }
    
    return {
        'Отзывы (ОС)': reviews,
        'Комментарии (ЦС)': comments,
        'Обсуждения (ПС)': discussions,
        'Даты': dates,
        'Ссылки': links
    };
}

function compareStructures(source, reference) {
    console.log(`\n${'='.repeat(80)}`);
    console.log(`СРАВНИТЕЛЬНЫЙ АНАЛИЗ СТРУКТУР`);
    console.log(`${'='.repeat(80)}\n`);
    
    if (!source || !reference) {
        console.log('❌ Невозможно провести сравнение - один из файлов не загружен');
        return;
    }
    
    console.log(`📊 ОБЩЕЕ СРАВНЕНИЕ:`);
    console.log(`   Исходник: ${source.totalSheets} листов`);
    console.log(`   Эталон: ${reference.totalSheets} листов\n`);
    
    // Находим наиболее подходящие листы для сравнения
    const sourceMainSheet = findMainDataSheet(source);
    const referenceMainSheet = findMainDataSheet(reference);
    
    if (sourceMainSheet && referenceMainSheet) {
        console.log(`🎯 СРАВНЕНИЕ ОСНОВНЫХ ЛИСТОВ:`);
        console.log(`   Исходник: "${sourceMainSheet.name}" (${sourceMainSheet.dimensions.rows}×${sourceMainSheet.dimensions.cols})`);
        console.log(`   Эталон: "${referenceMainSheet.name}" (${referenceMainSheet.dimensions.rows}×${referenceMainSheet.dimensions.cols})\n`);
        
        // Сравнение заголовков
        if (sourceMainSheet.headers.found && referenceMainSheet.headers.found) {
            console.log(`📋 СРАВНЕНИЕ ЗАГОЛОВКОВ:`);
            console.log(`   Исходник (строка ${sourceMainSheet.headers.row + 1}): ${sourceMainSheet.headers.columns.length} колонок`);
            console.log(`   Эталон (строка ${referenceMainSheet.headers.row + 1}): ${referenceMainSheet.headers.columns.length} колонок\n`);
            
            compareHeaders(sourceMainSheet.headers.columns, referenceMainSheet.headers.columns);
        }
        
        // Сравнение разделов
        console.log(`🔍 СРАВНЕНИЕ РАЗДЕЛОВ:`);
        console.log(`   Исходник: ${sourceMainSheet.sections.length} разделов`);
        console.log(`   Эталон: ${referenceMainSheet.sections.length} разделов\n`);
        
        // Сравнение итогов
        console.log(`📊 СРАВНЕНИЕ ИТОГОВ:`);
        console.log(`   Исходник: ${sourceMainSheet.totals.length} итоговых строк`);
        console.log(`   Эталон: ${referenceMainSheet.totals.length} итоговых строк\n`);
    }
    
    generateRecommendations(source, reference, sourceMainSheet, referenceMainSheet);
}

function findMainDataSheet(analysis) {
    // Ищем лист с наибольшим количеством данных и заголовков
    return analysis.sheets.reduce((best, sheet) => {
        const score = (sheet.headers.found ? 100 : 0) + 
                     sheet.sections.length * 10 + 
                     sheet.totals.length * 5 + 
                     sheet.dimensions.rows;
        
        if (!best || score > best.score) {
            return {...sheet, score};
        }
        return best;
    }, null);
}

function compareHeaders(sourceHeaders, referenceHeaders) {
    const common = [];
    const sourceOnly = [];
    const referenceOnly = [];
    
    sourceHeaders.forEach(header => {
        const normalizedSource = header.toLowerCase().trim();
        const found = referenceHeaders.some(refHeader => 
            refHeader.toLowerCase().trim().includes(normalizedSource) ||
            normalizedSource.includes(refHeader.toLowerCase().trim())
        );
        
        if (found) {
            common.push(header);
        } else {
            sourceOnly.push(header);
        }
    });
    
    referenceHeaders.forEach(header => {
        const normalizedRef = header.toLowerCase().trim();
        const found = sourceHeaders.some(srcHeader => 
            srcHeader.toLowerCase().trim().includes(normalizedRef) ||
            normalizedRef.includes(srcHeader.toLowerCase().trim())
        );
        
        if (!found) {
            referenceOnly.push(header);
        }
    });
    
    console.log(`   ✅ Общие колонки (${common.length}): ${common.join(', ')}`);
    console.log(`   🔸 Только в исходнике (${sourceOnly.length}): ${sourceOnly.join(', ')}`);
    console.log(`   🔹 Только в эталоне (${referenceOnly.length}): ${referenceOnly.join(', ')}\n`);
}

function generateRecommendations(source, reference, sourceMainSheet, referenceMainSheet) {
    console.log(`\n${'='.repeat(80)}`);
    console.log(`РЕКОМЕНДАЦИИ ДЛЯ НАСТРОЙКИ ПРОЦЕССОРА`);
    console.log(`${'='.repeat(80)}\n`);
    
    console.log(`🎯 КЛЮЧЕВЫЕ НАСТРОЙКИ:`);
    
    if (sourceMainSheet && sourceMainSheet.headers.found) {
        console.log(`   📍 Строка заголовков: ${sourceMainSheet.headers.row + 1}`);
        console.log(`   📏 Количество колонок: ${sourceMainSheet.headers.columns.length}`);
    }
    
    console.log(`\n🔑 КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПОИСКА:`);
    console.log(`   📝 Отзывы: "ОС", "отзыв"`);
    console.log(`   💬 Комментарии: "ЦС", "комментар"`);
    console.log(`   🗣️ Обсуждения: "ПС", "обсужден"`);
    console.log(`   📊 Итоги: "ИТОГО", "total", "сумма"`);
    
    console.log(`\n📋 МАППИНГ КОЛОНОК:`);
    if (sourceMainSheet && sourceMainSheet.headers.found) {
        sourceMainSheet.headers.columns.forEach((header, index) => {
            console.log(`   ${String.fromCharCode(65 + index)}: "${header}"`);
        });
    }
    
    console.log(`\n⚠️ ОСОБЕННОСТИ ОБРАБОТКИ:`);
    console.log(`   🔸 Использовать имена колонок вместо фиксированных позиций`);
    console.log(`   🔸 Учитывать строку ${sourceMainSheet?.headers.row + 1 || 1} как заголовки`);
    console.log(`   🔸 Автоматически определять разделы по ключевым словам`);
    console.log(`   🔸 Добавлять итоговые строки в конец обработки`);
    
    if (referenceMainSheet) {
        console.log(`\n📊 ЭТАЛОННЫЕ ПОКАЗАТЕЛИ:`);
        if (referenceMainSheet.dataAnalysis) {
            Object.entries(referenceMainSheet.dataAnalysis).forEach(([type, count]) => {
                console.log(`   ${type}: ~${count}`);
            });
        }
    }
}

// Основной запуск
async function main() {
    try {
        const sourceAnalysis = analyzeWorkbook('source_structure_analysis.xlsx', 'ИСХОДНОЙ ТАБЛИЦЫ');
        const referenceAnalysis = analyzeWorkbook('reference_structure_analysis.xlsx', 'ЭТАЛОННОЙ ТАБЛИЦЫ');
        
        if (sourceAnalysis && referenceAnalysis) {
            compareStructures(sourceAnalysis, referenceAnalysis);
        }
        
        // Сохраняем результаты в файл
        const results = {
            timestamp: new Date().toISOString(),
            source: sourceAnalysis,
            reference: referenceAnalysis
        };
        
        fs.writeFileSync('structure_analysis_results.json', JSON.stringify(results, null, 2));
        console.log(`\n✅ Результаты сохранены в structure_analysis_results.json`);
        
    } catch (error) {
        console.error('❌ Ошибка при анализе:', error.message);
    }
}

main();