const fs = require('fs');
const path = require('path');

// Помощник для фонового агента bc-be5df7e8-043a-4363-8fb5-572b1eb8c17c
class AgentHelper {
  constructor() {
    this.agentId = 'bc-be5df7e8-043a-4363-8fb5-572b1eb8c17c';
    this.requestsFile = path.join(__dirname, 'agent_requests.json');
    this.statusFile = path.join(__dirname, 'agent_status.json');
  }

  // Получить запросы от агента
  getRequests() {
    try {
      if (fs.existsSync(this.requestsFile)) {
        const data = fs.readFileSync(this.requestsFile, 'utf8');
        return JSON.parse(data);
      }
      return { requests: [], responses: [] };
    } catch (error) {
      console.error('Ошибка чтения запросов:', error);
      return { requests: [], responses: [] };
    }
  }

  // Ответить на запрос агента
  respondToRequest(requestId, response) {
    try {
      const data = this.getRequests();
      
      // Обновить статус запроса
      const request = data.requests.find(req => req.id === requestId);
      if (request) {
        request.status = 'completed';
      }
      
      // Обновить ответ
      const responseObj = data.responses.find(res => res.requestId === requestId);
      if (responseObj) {
        responseObj.response = response;
        responseObj.respondedAt = new Date().toISOString();
        responseObj.status = 'completed';
      }
      
      fs.writeFileSync(this.requestsFile, JSON.stringify(data, null, 2));
      console.log(`✅ Ответ на запрос ${requestId} предоставлен`);
      return true;
    } catch (error) {
      console.error('Ошибка при ответе на запрос:', error);
      return false;
    }
  }

  // Анализ Google Apps Script процессора
  analyzeGoogleAppsScript() {
    const processorPath = path.join(__dirname, 'google-apps-script-processor-enhanced.js');
    
    if (!fs.existsSync(processorPath)) {
      return 'Файл google-apps-script-processor-enhanced.js не найден';
    }

    const content = fs.readFileSync(processorPath, 'utf8');
    
    return {
      description: 'Google Apps Script процессор для обработки данных из таблиц',
      keyFeatures: [
        'Определение колонок через CONFIG.COLUMNS',
        'Разделение на отзывы и комментарии через CONFIG.POST_TYPES',
        'Автоматическое определение месяца через detectMonth()',
        'Разделение на секции через detectSection()',
        'Фильтрация и валидация данных'
      ],
      configColumns: this.extractConfigColumns(content),
      postTypes: this.extractPostTypes(content),
      mainFunctions: this.extractMainFunctions(content)
    };
  }

  // Извлечь конфигурацию колонок
  extractConfigColumns(content) {
    const match = content.match(/CONFIG\.COLUMNS\s*=\s*\{([^}]+)\}/);
    if (match) {
      return match[1].trim();
    }
    return 'Не найдено';
  }

  // Извлечь типы постов
  extractPostTypes(content) {
    const match = content.match(/CONFIG\.POST_TYPES\s*=\s*\{([^}]+)\}/);
    if (match) {
      return match[1].trim();
    }
    return 'Не найдено';
  }

  // Извлечь основные функции
  extractMainFunctions(content) {
    const functions = [];
    const funcRegex = /function\s+(\w+)\s*\(/g;
    let match;
    
    while ((match = funcRegex.exec(content)) !== null) {
      functions.push(match[1]);
    }
    
    return functions;
  }

  // Анализ тестового фреймворка
  analyzeTestingFramework() {
    const testingPath = path.join(__dirname, 'google-apps-script-testing.js');
    
    if (!fs.existsSync(testingPath)) {
      return 'Файл google-apps-script-testing.js не найден';
    }

    return {
      description: 'Тестовый фреймворк для проверки работы процессора',
      testFunctions: [
        'testSingleMonth() - тестирование одного месяца',
        'runFullTesting() - полное тестирование всех месяцев',
        'compareResults() - сравнение с эталонными данными',
        'validateStructure() - проверка структуры отчета'
      ],
      usage: 'Запустить testSingleMonth("Февраль") для тестирования февраля'
    };
  }

  // Информация о структуре данных
  getDataStructureInfo() {
    return {
      missionCritical: {
        spreadsheetId: '1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI',
        url: 'https://docs.google.com/spreadsheets/d/1RT8T5gnDPe0KMikTmVNdSvxqDal3aQUmelpEwItgxMI/edit?gid=1783011202#gid=1783011202',
        description: 'ГЛАВНАЯ ТАБЛИЦА с исходниками и эталонами',
        structure: {
          sourceSheets: [
            'Февраль 2025 - исходник',
            'Март 2025 - исходник', 
            'Апрель 2025 - исходник',
            'Май 2025 - исходник'
          ],
          referenceSheets: [
            'Февраль 2025 - эталон',
            'Март 2025 - эталон',
            'Апрель 2025 - эталон', 
            'Май 2025 - эталон'
          ]
        },
        requirements: {
          visualMatch: 'Внешний вид должен полностью соответствовать эталону',
          dataMatch: 'Данные в каждом столбце должны соответствовать эталону',
          statisticsMatch: 'Итоговая статистика должна соответствовать эталону',
          completeness: '4 полных соответствия по каждому месяцу'
        }
      },
      autonomousMode: true,
      maxResources: true,
      fullPermissions: true
    };
  }

  // Обработать все ожидающие запросы
  processAllRequests() {
    const data = this.getRequests();
    const pendingRequests = data.requests.filter(req => req.status === 'pending');
    
    console.log(`📋 Обработка ${pendingRequests.length} запросов...`);
    
    pendingRequests.forEach(request => {
      let response = '';
      
      switch (request.type) {
        case 'code_analysis':
          if (request.target === 'google-apps-script-processor-enhanced.js') {
            response = JSON.stringify(this.analyzeGoogleAppsScript(), null, 2);
          }
          break;
          
        case 'data_structure':
          response = JSON.stringify(this.getDataStructureInfo(), null, 2);
          break;
          
        case 'testing_framework':
          response = JSON.stringify(this.analyzeTestingFramework(), null, 2);
          break;
          
        default:
          response = `Тип запроса "${request.type}" не поддерживается`;
      }
      
      this.respondToRequest(request.id, response);
    });
    
    console.log('✅ Все запросы обработаны');
  }

  // Добавить новый запрос (для тестирования)
  addRequest(type, target, question, priority = 'medium') {
    const data = this.getRequests();
    const newId = `req-${Date.now()}`;
    
    const newRequest = {
      id: newId,
      type,
      target,
      question,
      status: 'pending',
      priority,
      createdAt: new Date().toISOString()
    };
    
    const newResponse = {
      requestId: newId,
      response: 'Будет заполнено при обработке запроса',
      respondedAt: null,
      status: 'pending'
    };
    
    data.requests.push(newRequest);
    data.responses.push(newResponse);
    
    fs.writeFileSync(this.requestsFile, JSON.stringify(data, null, 2));
    console.log(`📝 Добавлен новый запрос: ${newId}`);
    return newId;
  }

  // Показать справку
  showHelp() {
    console.log(`
🤖 ПОМОЩНИК ДЛЯ ФОНОВОГО АГЕНТА ${this.agentId}

📋 ДОСТУПНЫЕ КОМАНДЫ:
  processAllRequests() - Обработать все ожидающие запросы
  addRequest(type, target, question) - Добавить новый запрос
  analyzeGoogleAppsScript() - Анализ основного процессора
  analyzeTestingFramework() - Анализ тестового фреймворка
  getDataStructureInfo() - Информация о структуре данных

🔍 ТИПЫ ЗАПРОСОВ:
  - code_analysis: Анализ кода
  - data_structure: Структура данных
  - testing_framework: Тестовый фреймворк
  - error_debugging: Отладка ошибок
  - performance_optimization: Оптимизация производительности
  - configuration_help: Помощь с конфигурацией

📞 ПРИМЕР ИСПОЛЬЗОВАНИЯ:
  const helper = new AgentHelper();
  helper.processAllRequests();
`);
  }
}

// Если скрипт запущен напрямую
if (require.main === module) {
  const helper = new AgentHelper();
  
  // Показать справку
  helper.showHelp();
  
  // Обработать все запросы
  helper.processAllRequests();
}

module.exports = AgentHelper; 