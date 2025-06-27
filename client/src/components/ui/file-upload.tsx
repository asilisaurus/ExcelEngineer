import React, { useCallback, useState } from 'react';
import { useDropzone } from 'react-dropzone';
import { Upload, File, X, CheckCircle, AlertCircle } from 'lucide-react';
import { Button } from './button';
import { cn, formatFileSize, getTimeAgo } from '@/lib/utils';

interface FileUploadProps {
  onFileUpload: (file: File) => void;
  isLoading?: boolean;
  uploadedFile?: {
    name: string;
    size: number;
    uploadedAt: string;
  } | null;
  onRemoveFile?: () => void;
  error?: string | null;
}

export function FileUpload({ 
  onFileUpload, 
  isLoading = false, 
  uploadedFile, 
  onRemoveFile,
  error 
}: FileUploadProps) {
  const [dragOver, setDragOver] = useState(false);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      onFileUpload(acceptedFiles[0]);
    }
    setDragOver(false);
  }, [onFileUpload]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.ms-excel': ['.xls'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
    },
    maxFiles: 1,
    maxSize: 10 * 1024 * 1024, // 10MB
    onDragEnter: () => setDragOver(true),
    onDragLeave: () => setDragOver(false)
  });

  if (uploadedFile) {
    return (
      <div className="p-4 bg-gray-50 rounded-lg">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="flex-shrink-0">
              <File className="h-8 w-8 text-green-500" />
            </div>
            <div>
              <p className="text-sm font-medium text-gray-900">{uploadedFile.name}</p>
              <p className="text-xs text-gray-500">
                {formatFileSize(uploadedFile.size)} • Загружен {getTimeAgo(uploadedFile.uploadedAt)}
              </p>
            </div>
          </div>
          {onRemoveFile && (
            <Button
              variant="ghost"
              size="sm"
              onClick={onRemoveFile}
              className="text-gray-400 hover:text-gray-600"
            >
              <X className="h-4 w-4" />
            </Button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div
        {...getRootProps()}
        className={cn(
          "file-upload-zone",
          isDragActive && "drag-over",
          isLoading && "opacity-50 cursor-not-allowed"
        )}
      >
        <input {...getInputProps()} disabled={isLoading} />
        <div className="mx-auto max-w-sm">
          <Upload className="mx-auto h-12 w-12 text-gray-400 mb-4" />
          <h4 className="text-lg font-medium text-gray-900 mb-2">
            {isDragActive ? "Отпустите файл здесь" : "Перетащите файл сюда"}
          </h4>
          <p className="text-sm text-gray-600 mb-4">или нажмите для выбора файла</p>
          <Button 
            type="button" 
            disabled={isLoading}
            className="mb-3"
          >
            <Upload className="w-4 h-4 mr-2" />
            {isLoading ? "Загрузка..." : "Выбрать файл"}
          </Button>
          <p className="text-xs text-gray-500">
            Поддерживаемые форматы: .xls, .xlsx (до 10 МБ)
          </p>
        </div>
      </div>

      {error && (
        <div className="flex items-center space-x-2 text-sm text-red-600 bg-red-50 p-3 rounded-lg">
          <AlertCircle className="h-4 w-4 flex-shrink-0" />
          <span>{error}</span>
        </div>
      )}
    </div>
  );
}
