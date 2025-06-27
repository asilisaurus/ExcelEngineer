import React, { useState, useEffect } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { FileSpreadsheet, Clock, Settings, Download, Eye, History, FileOutput, Info } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import { useToast } from '@/hooks/use-toast';
import { FileUpload } from '@/components/ui/file-upload';
import { ProgressBar } from '@/components/ui/progress-bar';
import { ProcessingSteps } from '@/components/processing-steps';
import { StatisticsPanel } from '@/components/statistics-panel';
import { apiRequest } from '@/lib/queryClient';
import { formatDate, formatFileSize, getProcessingProgress } from '@/lib/utils';
import type { ProcessedFile } from '@shared/schema';

export default function Home() {
  const [currentFileId, setCurrentFileId] = useState<number | null>(null);
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const { toast } = useToast();
  const queryClient = useQueryClient();

  // Query for current file status
  const { data: currentFile, isLoading: fileLoading } = useQuery({
    queryKey: ['/api/files', currentFileId],
    enabled: !!currentFileId,
    refetchInterval: (data) => {
      return data?.status === 'processing' ? 2000 : false;
    },
  });

  // Upload mutation
  const uploadMutation = useMutation({
    mutationFn: async (file: File) => {
      const formData = new FormData();
      formData.append('file', file);
      const response = await apiRequest('POST', '/api/upload', formData);
      return response.json();
    },
    onSuccess: (data: { fileId: number; file: ProcessedFile }) => {
      setCurrentFileId(data.fileId);
      setUploadedFile(null);
      toast({
        title: "Файл загружен",
        description: "Обработка файла началась. Прогресс будет обновляться автоматически.",
      });
      queryClient.invalidateQueries({ queryKey: ['/api/files'] });
    },
    onError: (error: Error) => {
      toast({
        title: "Ошибка загрузки",
        description: error.message,
        variant: "destructive",
      });
    },
  });

  // Download mutation
  const downloadMutation = useMutation({
    mutationFn: async (fileId: number) => {
      const response = await apiRequest('GET', `/api/files/${fileId}/download`);
      const blob = await response.blob();
      return { blob, filename: currentFile?.processedName || 'report.xlsx' };
    },
    onSuccess: ({ blob, filename }) => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      toast({
        title: "Файл скачан",
        description: "Обработанный отчет успешно скачан.",
      });
    },
    onError: (error: Error) => {
      toast({
        title: "Ошибка скачивания",
        description: error.message,
        variant: "destructive",
      });
    },
  });

  const handleFileUpload = (file: File) => {
    setUploadedFile(file);
    uploadMutation.mutate(file);
  };

  const handleRemoveFile = () => {
    setUploadedFile(null);
    setCurrentFileId(null);
  };

  const handleDownload = () => {
    if (currentFileId && currentFile?.status === 'completed') {
      downloadMutation.mutate(currentFileId);
    }
  };

  const getProcessingStep = (status: string) => {
    switch (status) {
      case 'processing':
        return 'calculating';
      case 'completed':
        return 'completed';
      case 'error':
        return 'upload';
      default:
        return 'upload';
    }
  };

  const progress = currentFile ? getProcessingProgress(currentFile.status, currentFile.rowsProcessed || 0) : 0;

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <div className="flex items-center space-x-4">
              <div className="flex-shrink-0">
                <div className="w-10 h-10 bg-primary rounded-lg flex items-center justify-center">
                  <FileSpreadsheet className="text-white text-lg" />
                </div>
              </div>
              <div>
                <h1 className="text-xl font-semibold text-gray-900">Фортедетрим XLS Процессор</h1>
                <p className="text-sm text-gray-500">Автоматизация обработки отчетов</p>
              </div>
            </div>
            <div className="flex items-center space-x-4">
              <div className="text-sm text-gray-500">
                <Clock className="inline w-4 h-4 mr-1" />
                Последняя обработка: <span className="font-medium">
                  {currentFile ? formatDate(currentFile.createdAt) : 'Нет данных'}
                </span>
              </div>
              <Button variant="ghost" size="sm">
                <Settings className="w-4 h-4" />
              </Button>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Process Overview */}
        <Card className="mb-8">
          <CardContent className="p-6">
            <h2 className="text-lg font-semibold text-gray-900 mb-4">Процесс обработки</h2>
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-8">
                <div className="flex items-center space-x-3">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center ${
                    uploadedFile || currentFile ? 'bg-primary' : 'bg-gray-200'
                  }`}>
                    <FileOutput className={`text-sm ${uploadedFile || currentFile ? 'text-white' : 'text-gray-500'}`} />
                  </div>
                  <div>
                    <p className="text-sm font-medium text-gray-900">Загрузка</p>
                    <p className="text-xs text-gray-500">Исходный XLS файл</p>
                  </div>
                </div>
                <div className="w-8 h-px bg-gray-300"></div>
                <div className="flex items-center space-x-3">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center ${
                    currentFile?.status === 'processing' ? 'bg-yellow-500' : 
                    currentFile?.status === 'completed' ? 'bg-green-500' : 'bg-gray-200'
                  }`}>
                    <Settings className={`text-sm ${
                      currentFile?.status ? 'text-white' : 'text-gray-500'
                    }`} />
                  </div>
                  <div>
                    <p className={`text-sm font-medium ${
                      currentFile?.status ? 'text-gray-900' : 'text-gray-500'
                    }`}>
                      Обработка
                    </p>
                    <p className={`text-xs ${
                      currentFile?.status ? 'text-gray-500' : 'text-gray-400'
                    }`}>
                      Очистка и расчеты
                    </p>
                  </div>
                </div>
                <div className="w-8 h-px bg-gray-300"></div>
                <div className="flex items-center space-x-3">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center ${
                    currentFile?.status === 'completed' ? 'bg-green-500' : 'bg-gray-200'
                  }`}>
                    <Download className={`text-sm ${
                      currentFile?.status === 'completed' ? 'text-white' : 'text-gray-500'
                    }`} />
                  </div>
                  <div>
                    <p className={`text-sm font-medium ${
                      currentFile?.status === 'completed' ? 'text-gray-900' : 'text-gray-500'
                    }`}>
                      Результат
                    </p>
                    <p className={`text-xs ${
                      currentFile?.status === 'completed' ? 'text-gray-500' : 'text-gray-400'
                    }`}>
                      Готовый отчет
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </CardContent>
        </Card>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Left Column: Upload & Processing */}
          <div className="lg:col-span-2 space-y-6">
            {/* File Upload Section */}
            <Card>
              <CardHeader>
                <div className="flex items-center justify-between">
                  <CardTitle>Загрузка исходного файла</CardTitle>
                  <Badge variant="secondary">Шаг 1</Badge>
                </div>
                <p className="text-sm text-gray-600">Загрузите исходный XLS файл с данными для обработки</p>
              </CardHeader>
              <CardContent>
                <FileUpload
                  onFileUpload={handleFileUpload}
                  isLoading={uploadMutation.isPending}
                  uploadedFile={uploadedFile ? {
                    name: uploadedFile.name,
                    size: uploadedFile.size,
                    uploadedAt: new Date().toISOString()
                  } : currentFile ? {
                    name: currentFile.originalName,
                    size: currentFile.fileSize,
                    uploadedAt: currentFile.createdAt
                  } : null}
                  onRemoveFile={handleRemoveFile}
                  error={uploadMutation.error?.message}
                />
              </CardContent>
            </Card>

            {/* Processing Section */}
            <Card>
              <CardHeader>
                <div className="flex items-center justify-between">
                  <CardTitle>Обработка данных</CardTitle>
                  <Badge variant={currentFile?.status === 'processing' ? 'default' : 'secondary'}>
                    Шаг 2
                  </Badge>
                </div>
                <p className="text-sm text-gray-600">Автоматическая очистка и форматирование данных</p>
              </CardHeader>
              <CardContent className="space-y-6">
                <ProcessingSteps
                  currentStep={getProcessingStep(currentFile?.status || 'pending')}
                  error={currentFile?.errorMessage}
                  rowsProcessed={currentFile?.rowsProcessed || undefined}
                />

                {currentFile && (
                  <ProgressBar
                    value={progress}
                    variant={currentFile.status === 'error' ? 'error' : 
                             currentFile.status === 'completed' ? 'success' : 'default'}
                  />
                )}
              </CardContent>
            </Card>

            {/* Results Section */}
            <Card>
              <CardHeader>
                <div className="flex items-center justify-between">
                  <CardTitle>Результат обработки</CardTitle>
                  <Badge variant={currentFile?.status === 'completed' ? 'default' : 'secondary'}>
                    Шаг 3
                  </Badge>
                </div>
                <p className="text-sm text-gray-600">Готовый отчет для скачивания</p>
              </CardHeader>
              <CardContent>
                {currentFile?.status === 'completed' ? (
                  <div className="space-y-4">
                    <div className="bg-gray-50 rounded-lg p-4">
                      <div className="flex items-center justify-between mb-3">
                        <h4 className="text-sm font-medium text-gray-900">{currentFile.processedName}</h4>
                        <span className="text-xs text-gray-500">{formatFileSize(currentFile.fileSize)}</span>
                      </div>
                      
                      {/* Data Preview Table */}
                      <div className="overflow-hidden border border-gray-200 rounded-lg mb-2">
                        <table className="min-w-full divide-y divide-gray-200">
                          <thead className="bg-gray-100">
                            <tr>
                              <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Площадка</th>
                              <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Тема</th>
                              <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Дата</th>
                              <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Просмотры</th>
                              <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Вовлечение</th>
                            </tr>
                          </thead>
                          <tbody className="bg-white divide-y divide-gray-200">
                            <tr>
                              <td className="px-3 py-2 text-xs text-gray-900">https://vk.com/wall</td>
                              <td className="px-3 py-2 text-xs text-gray-900">71 Рецепт против авитаминоза</td>
                              <td className="px-3 py-2 text-xs text-gray-900">28.03.2025</td>
                              <td className="px-3 py-2 text-xs text-gray-900">500</td>
                              <td className="px-3 py-2 text-xs text-gray-900">есть</td>
                            </tr>
                            <tr className="bg-gray-50">
                              <td className="px-3 py-2 text-xs text-gray-900">https://www.youtube.com/watch</td>
                              <td className="px-3 py-2 text-xs text-gray-900">Добрый день! Фортедетрим не БАД</td>
                              <td className="px-3 py-2 text-xs text-gray-900">28.03.2025</td>
                              <td className="px-3 py-2 text-xs text-gray-900">Нет данных</td>
                              <td className="px-3 py-2 text-xs text-gray-900">ЦС</td>
                            </tr>
                            <tr>
                              <td className="px-3 py-2 text-xs text-gray-900">https://vk.com/wall</td>
                              <td className="px-3 py-2 text-xs text-gray-900">для меня такой витамин</td>
                              <td className="px-3 py-2 text-xs text-gray-900">28.03.2025</td>
                              <td className="px-3 py-2 text-xs text-gray-900">3300</td>
                              <td className="px-3 py-2 text-xs text-gray-900">ЦС</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                      <p className="text-xs text-gray-500">Показано 3 из {currentFile.rowsProcessed} строк</p>
                    </div>

                    <div className="flex items-center justify-between">
                      <div className="flex items-center space-x-4">
                        <Button
                          onClick={handleDownload}
                          disabled={downloadMutation.isPending}
                          className="bg-green-600 hover:bg-green-700"
                        >
                          <Download className="w-4 h-4 mr-2" />
                          {downloadMutation.isPending ? 'Скачивание...' : 'Скачать отчет'}
                        </Button>
                        <Button variant="outline">
                          <Eye className="w-4 h-4 mr-2" />
                          Предварительный просмотр
                        </Button>
                      </div>
                      <div className="text-xs text-gray-500 flex items-center">
                        <span className="w-2 h-2 bg-green-500 rounded-full mr-2"></span>
                        Обработка завершена
                      </div>
                    </div>
                  </div>
                ) : currentFile?.status === 'error' ? (
                  <div className="text-center py-8">
                    <div className="w-12 h-12 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
                      <FileSpreadsheet className="w-6 h-6 text-red-600" />
                    </div>
                    <h3 className="text-lg font-medium text-gray-900 mb-2">Ошибка обработки</h3>
                    <p className="text-sm text-gray-600 mb-4">{currentFile.errorMessage}</p>
                    <Button onClick={handleRemoveFile} variant="outline">
                      Попробовать снова
                    </Button>
                  </div>
                ) : (
                  <div className="text-center py-8">
                    <div className="w-12 h-12 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                      <FileSpreadsheet className="w-6 h-6 text-gray-400" />
                    </div>
                    <h3 className="text-lg font-medium text-gray-900 mb-2">Ожидание обработки</h3>
                    <p className="text-sm text-gray-600">Загрузите файл для начала обработки</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>

          {/* Right Column: Statistics & Info */}
          <div className="space-y-6">
            {/* Processing Statistics */}
            <StatisticsPanel
              statistics={currentFile?.statistics || null}
              isLoading={fileLoading && currentFile?.status === 'processing'}
            />

            {/* Quick Actions */}
            <Card>
              <CardHeader>
                <CardTitle>Быстрые действия</CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                <Button variant="outline" className="w-full justify-start">
                  <History className="w-4 h-4 mr-2" />
                  История обработок
                </Button>
                <Button variant="outline" className="w-full justify-start">
                  <Settings className="w-4 h-4 mr-2" />
                  Настройки формата
                </Button>
                <Button variant="outline" className="w-full justify-start">
                  <FileOutput className="w-4 h-4 mr-2" />
                  Экспорт шаблона
                </Button>
              </CardContent>
            </Card>

            {/* System Info */}
            <Card className="bg-gradient-to-br from-blue-50 to-indigo-50 border-blue-200">
              <CardContent className="p-6">
                <div className="flex items-start space-x-3">
                  <div className="flex-shrink-0">
                    <Info className="text-primary text-lg" />
                  </div>
                  <div>
                    <h4 className="text-sm font-medium text-gray-900 mb-2">Информация о системе</h4>
                    <div className="text-xs text-gray-600 space-y-1">
                      <p>• Поддержка файлов до 10 МБ</p>
                      <p>• Автоматическое распознавание листов</p>
                      <p>• Сохранение истории обработок</p>
                      <p>• Проверка целостности данных</p>
                    </div>
                    <div className="mt-3 text-xs text-gray-500">
                      Версия: 2.1.0 • Обновлено: 15.01.2025
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-gray-200 mt-16">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
          <div className="flex items-center justify-between">
            <div className="text-sm text-gray-500">
              © 2025 Фортедетрим XLS Процессор. Все права защищены.
            </div>
            <div className="flex items-center space-x-6 text-sm text-gray-500">
              <a href="#" className="hover:text-gray-900">Документация</a>
              <a href="#" className="hover:text-gray-900">Поддержка</a>
              <a href="#" className="hover:text-gray-900">API</a>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
}
