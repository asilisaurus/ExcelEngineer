const fs = require('fs');
const path = require('path');

// Мониторинг состояния фонового агента bc-be5df7e8-043a-4363-8fb5-572b1eb8c17c
class AgentMonitor {
  constructor() {
    this.agentId = 'bc-be5df7e8-043a-4363-8fb5-572b1eb8c17c';
    this.statusFile = path.join(__dirname, 'agent_status.json');
    this.requestsFile = path.join(__dirname, 'agent_requests.json');
  }

  // Получить текущий статус агента
  getAgentStatus() {
    try {
      if (fs.existsSync(this.statusFile)) {
        const data = fs.readFileSync(this.statusFile, 'utf8');
        return JSON.parse(data);
      }
      return null;
    } catch (error) {
      console.error('Ошибка чтения статуса агента:', error);
      return null;
    }
  }

  // Обновить статус агента
  updateAgentStatus(updates) {
    try {
      const currentStatus = this.getAgentStatus() || {};
      const newStatus = {
        ...currentStatus,
        ...updates,
        lastUpdate: new Date().toISOString()
      };
      
      fs.writeFileSync(this.statusFile, JSON.stringify(newStatus, null, 2));
      console.log('Статус агента обновлен:', updates);
      return newStatus;
    } catch (error) {
      console.error('Ошибка обновления статуса:', error);
      return null;
    }
  }

  // Показать прогресс тестирования
  showProgress() {
    const status = this.getAgentStatus();
    if (!status) {
      console.log('❌ Статус агента недоступен');
      return;
    }

    console.log(`\n📊 ПРОГРЕСС ТЕСТИРОВАНИЯ АГЕНТА ${this.agentId}`);
    console.log(`📅 Последнее обновление: ${status.lastUpdate}`);
    console.log(`🎯 Текущий статус: ${status.status}`);
    console.log(`🔧 Текущая задача: ${status.currentTask || 'Нет'}`);
    
    console.log('\n📈 РЕЗУЛЬТАТЫ ПО МЕСЯЦАМ:');
    Object.entries(status.testResults).forEach(([month, results]) => {
      const statusIcon = results.status === 'completed' ? '✅' : 
                        results.status === 'in_progress' ? '🔄' : '⏳';
      console.log(`${statusIcon} ${month.toUpperCase()}: ${results.accuracy}% точность`);
      if (results.errors.length > 0) {
        console.log(`   ❌ Ошибок: ${results.errors.length}`);
      }
    });

    console.log('\n📊 ОБЩИЙ ПРОГРЕСС:');
    console.log(`✅ Завершено тестов: ${status.overallProgress.completedTests}/${status.overallProgress.totalTests}`);
    console.log(`📈 Средняя точность: ${status.overallProgress.averageAccuracy}%`);
    console.log(`🎯 Целевая точность: ${status.overallProgress.targetAccuracy}%`);
  }

  // Проверить запросы от агента
  checkRequests() {
    try {
      if (fs.existsSync(this.requestsFile)) {
        const data = fs.readFileSync(this.requestsFile, 'utf8');
        const requests = JSON.parse(data);
        
        const pendingRequests = requests.requests.filter(req => req.status === 'pending');
        
        if (pendingRequests.length > 0) {
          console.log(`\n📬 НОВЫЕ ЗАПРОСЫ ОТ АГЕНТА (${pendingRequests.length}):`);
          pendingRequests.forEach(req => {
            console.log(`🔍 ${req.id}: ${req.question}`);
            console.log(`   📁 Цель: ${req.target}`);
            console.log(`   ⚡ Приоритет: ${req.priority}`);
          });
        } else {
          console.log('\n📬 Новых запросов нет');
        }
      }
    } catch (error) {
      console.error('Ошибка проверки запросов:', error);
    }
  }

  // Добавить лог для агента
  addLog(message, type = 'info') {
    const status = this.getAgentStatus();
    if (status) {
      status.logs = status.logs || [];
      status.logs.push({
        timestamp: new Date().toISOString(),
        type,
        message
      });
      
      // Оставляем только последние 100 записей
      if (status.logs.length > 100) {
        status.logs = status.logs.slice(-100);
      }
      
      this.updateAgentStatus(status);
    }
  }

  // Запустить мониторинг
  startMonitoring() {
    console.log(`🔍 Запуск мониторинга агента ${this.agentId}`);
    
    // Показать текущий статус
    this.showProgress();
    
    // Проверить запросы
    this.checkRequests();
    
    // Добавить лог о начале мониторинга
    this.addLog('Мониторинг запущен', 'info');
  }
}

// Если скрипт запущен напрямую
if (require.main === module) {
  const monitor = new AgentMonitor();
  monitor.startMonitoring();
}

module.exports = AgentMonitor; 