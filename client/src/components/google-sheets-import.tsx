import React, { useState } from 'react';
import { Link, Upload } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { useMutation } from '@tanstack/react-query';

interface GoogleSheetsImportProps {
  onImportStart: (fileId: number) => void;
}

export function GoogleSheetsImport({ onImportStart }: GoogleSheetsImportProps) {
  const [url, setUrl] = useState('');

  const importMutation = useMutation({
    mutationFn: async (url: string) => {
      const response = await fetch('/api/import-google-sheets', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ url }),
      });

      if (!response.ok) {
        const error = await response.json();
        throw new Error(error.message || 'Ошибка при импорте');
      }

      return response.json();
    },
    onSuccess: (data) => {
      onImportStart(data.fileId);
      setUrl('');
    },
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (url.trim()) {
      importMutation.mutate(url.trim());
    }
  };

  const isValidGoogleSheetsUrl = (url: string): boolean => {
    return /docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9-_]+/.test(url);
  };

  return (
    <Card className="border-dashed border-2 border-blue-200 bg-blue-50/50">
      <CardHeader>
        <CardTitle className="flex items-center space-x-2 text-blue-700">
          <Link className="h-5 w-5" />
          <span>Импорт из Google Таблиц</span>
        </CardTitle>
      </CardHeader>
      <CardContent>
        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <Input
              type="url"
              placeholder="https://docs.google.com/spreadsheets/d/..."
              value={url}
              onChange={(e) => setUrl(e.target.value)}
              className="w-full"
              disabled={importMutation.isPending}
            />
            <p className="text-xs text-gray-500 mt-1">
              Вставьте ссылку на Google Таблицу (убедитесь, что доступ открыт для просмотра)
            </p>
          </div>

          {importMutation.error && (
            <div className="text-sm text-red-600 bg-red-50 p-2 rounded">
              {importMutation.error.message}
            </div>
          )}

          <div className="flex items-center space-x-2">
            <Button
              type="submit"
              disabled={!url.trim() || !isValidGoogleSheetsUrl(url) || importMutation.isPending}
              className="bg-blue-600 hover:bg-blue-700"
            >
              <Upload className="w-4 h-4 mr-2" />
              {importMutation.isPending ? 'Импорт...' : 'Импортировать'}
            </Button>
            
            {url.trim() && !isValidGoogleSheetsUrl(url) && (
              <span className="text-xs text-red-500">
                Неверный URL Google Таблиц
              </span>
            )}
          </div>
        </form>

        <div className="mt-4 text-xs text-gray-600">
          <p className="font-medium mb-1">Как получить ссылку:</p>
          <ol className="list-decimal list-inside space-y-1">
            <li>Откройте Google Таблицу</li>
            <li>Нажмите "Настройки доступа" → "Доступ по ссылке"</li>
            <li>Скопируйте ссылку и вставьте её выше</li>
          </ol>
        </div>
      </CardContent>
    </Card>
  );
} 