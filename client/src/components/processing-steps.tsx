import React from 'react';
import { Check, Loader2, AlertCircle, Clock } from 'lucide-react';
import { cn } from '@/lib/utils';

interface ProcessingStep {
  id: string;
  title: string;
  description: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
}

interface ProcessingStepsProps {
  currentStep: string;
  error?: string | null;
  rowsProcessed?: number;
}

export function ProcessingSteps({ currentStep, error, rowsProcessed }: ProcessingStepsProps) {
  const isGoogleSheetsImport = !rowsProcessed || rowsProcessed < 10;
  
  // Determine step based on rowsProcessed value
  const getStepStatus = (stepIndex: number): ProcessingStep['status'] => {
    if (error) return 'error';
    if (currentStep === 'completed') return 'completed';
    if (currentStep === 'error') return 'error';
    
    if (rowsProcessed === undefined || rowsProcessed === 0) {
      return stepIndex === 0 ? 'processing' : 'pending';
    }
    
    if (rowsProcessed === 1) {
      return stepIndex <= 0 ? 'completed' : stepIndex === 1 ? 'processing' : 'pending';
    }
    
    if (rowsProcessed === 2) {
      return stepIndex <= 1 ? 'completed' : stepIndex === 2 ? 'processing' : 'pending';
    }
    
    if (rowsProcessed === 3) {
      return stepIndex <= 2 ? 'completed' : stepIndex === 3 ? 'processing' : 'pending';
    }
    
    return stepIndex <= rowsProcessed ? 'completed' : 'pending';
  };
  
  const steps: ProcessingStep[] = [
    {
      id: 'upload',
      title: isGoogleSheetsImport ? 'Загрузка данных из Google Таблиц' : 'Извлечение данных из листов',
      description: isGoogleSheetsImport 
        ? 'Подключение к Google Таблицам и загрузка данных'
        : 'Чтение структуры файла Excel',
      status: getStepStatus(0)
    },
    {
      id: 'cleaning',
      title: 'Удаление лишних колонок',
      description: 'Удалены: Согласование, Просмотры на старте, Просмотры в конце месяца',
      status: getStepStatus(1)
    },
    {
      id: 'calculating',
      title: 'Расчет показателей вовлечения',
      description: 'Применение формулы: (комментарии + лайки + репосты) / просмотры * 100',
      status: getStepStatus(2)
    },
    {
      id: 'formatting',
      title: 'Форматирование отчета',
      description: 'Создание групп и итоговых показателей',
      status: getStepStatus(3)
    }
  ];

  // Apply error state to current step if there's an error
  if (error) {
    const errorStepIndex = steps.findIndex(step => step.id === currentStep);
    if (errorStepIndex !== -1) {
      steps[errorStepIndex].status = 'error';
    }
  }

  const getStepIcon = (status: ProcessingStep['status']) => {
    switch (status) {
      case 'completed':
        return <Check className="w-4 h-4 text-white" />;
      case 'processing':
        return <Loader2 className="w-4 h-4 text-white animate-spin" />;
      case 'error':
        return <AlertCircle className="w-4 h-4 text-white" />;
      default:
        return <Clock className="w-4 h-4 text-gray-500" />;
    }
  };

  const getStepBadgeClass = (status: ProcessingStep['status']) => {
    switch (status) {
      case 'completed':
        return 'bg-green-500';
      case 'processing':
        return 'bg-yellow-500';
      case 'error':
        return 'bg-red-500';
      default:
        return 'bg-gray-200';
    }
  };

  const getStepTextClass = (status: ProcessingStep['status']) => {
    switch (status) {
      case 'completed':
        return 'text-green-600';
      case 'processing':
        return 'text-yellow-600';
      case 'error':
        return 'text-red-600';
      default:
        return 'text-gray-400';
    }
  };

  const getStatusText = (status: ProcessingStep['status']) => {
    switch (status) {
      case 'completed':
        return 'Завершено';
      case 'processing':
        return 'В процессе';
      case 'error':
        return 'Ошибка';
      default:
        return 'Ожидание';
    }
  };

  return (
    <div className="space-y-4">
      {steps.map((step, index) => (
        <div
          key={step.id}
          className={cn(
            "processing-step",
            step.status === 'completed' && "completed",
            step.status === 'processing' && "processing",
            step.status === 'pending' && "pending",
            step.status === 'error' && "error"
          )}
        >
          <div className="flex-shrink-0">
            <div className={cn(
              "w-6 h-6 rounded-full flex items-center justify-center",
              getStepBadgeClass(step.status)
            )}>
              {step.status === 'pending' ? (
                <span className="text-xs text-gray-500 font-medium">{index + 1}</span>
              ) : (
                getStepIcon(step.status)
              )}
            </div>
          </div>
          <div className="flex-1 min-w-0">
            <p className={cn(
              "text-sm font-medium",
              step.status === 'pending' ? 'text-gray-500' : 'text-gray-900'
            )}>
              {step.title}
            </p>
            <p className={cn(
              "text-xs",
              step.status === 'pending' ? 'text-gray-400' : 'text-gray-500'
            )}>
              {step.description}
            </p>
            {step.status === 'error' && error && (
              <p className="text-xs text-red-600 mt-1">{error}</p>
            )}
          </div>
          <span className={cn(
            "text-xs font-medium",
            getStepTextClass(step.status)
          )}>
            {getStatusText(step.status)}
          </span>
        </div>
      ))}
    </div>
  );
}
