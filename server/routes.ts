import type { Express, Request, Response } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { upload, cleanupFile, getOutputFileName } from "./services/file-handler";
import { unifiedProcessor } from "./services/excel-processor-unified";
import { importFromGoogleSheets, validateGoogleSheetsUrl } from "./services/google-sheets-importer";
import { insertProcessedFileSchema, processingStatsSchema } from "@shared/schema";
import fs from 'fs';
import path from 'path';

export async function registerRoutes(app: Express): Promise<Server> {
  // Upload and process Excel file
  app.post("/api/upload", (req: Request, res: Response) => {
    upload.single('file')(req, res, async (err: any) => {
      try {
        console.log('Upload request received:', {
          file: req.file,
          body: req.body,
          headers: req.headers['content-type'],
          error: err
        });
        
        if (err) {
          console.log('Multer error:', err);
          return res.status(400).json({ message: err.message || "Ошибка при загрузке файла" });
        }
        
        if (!req.file) {
          console.log('No file in request');
          return res.status(400).json({ message: "Файл не был загружен" });
        }

        // Multer (and many browsers) provide non-ASCII filenames in latin1.
        // Convert to UTF-8 so Cyrillic/Unicode characters display correctly.
        const originalNameUtf8 = Buffer.from(req.file.originalname, "latin1").toString("utf8");

        // Create initial record
        const processedFile = await storage.createProcessedFile({
          originalName: originalNameUtf8,
          processedName: getOutputFileName(originalNameUtf8),
          status: 'processing',
          fileSize: req.file.size,
          rowsProcessed: 0,
          statistics: null,
          errorMessage: null,
        });

        // Process file asynchronously with progress tracking
        setImmediate(async () => {
          try {
            console.log('🔄 Начинаем обработку файла:', req.file!.originalname);
            
            // Update to show reading stage
            console.log('📊 PROGRESS: Setting reading stage (30%)');
            await storage.updateProcessedFile(processedFile.id, {
              status: 'processing',
              rowsProcessed: 1,
              statistics: JSON.stringify({ stage: 'reading', message: 'Чтение файла...' }),
            });

            // Process the file using the unified processor
            console.log('🔄 PROGRESS: Starting file processing with unified processor...');
            const result = await unifiedProcessor.processExcelFile(req.file!.path);
            console.log('✅ PROGRESS: File processing completed, path:', result.outputPath);

            // Update to show processing stage
            console.log('📊 PROGRESS: Setting processing stage (70%)');
            await storage.updateProcessedFile(processedFile.id, {
              status: 'processing',
              rowsProcessed: 2,
              statistics: JSON.stringify({ stage: 'processing', message: 'Обработка данных...' }),
            });

            // Simulate final formatting stage
            console.log('📊 PROGRESS: Final formatting stage...');
            await new Promise(resolve => setTimeout(resolve, 500));

            // Update record with success and correct processed filename
            const actualFileName = path.basename(result.outputPath);
            console.log('✅ PROGRESS: Setting completed stage (100%), file:', actualFileName);
            await storage.updateProcessedFile(processedFile.id, {
              processedName: actualFileName,
              status: 'completed',
              rowsProcessed: result.statistics.totalRows,
              statistics: JSON.stringify(result.statistics),
              completedAt: new Date(),
            });

            console.log('✅ Обработка файла завершена успешно');

            // Cleanup original file
            cleanupFile(req.file!.path);
          } catch (error) {
            console.error('❌ Ошибка при обработке файла:', error);
            await storage.updateProcessedFile(processedFile.id, {
              status: 'error',
              errorMessage: error instanceof Error ? error.message : 'Неизвестная ошибка при обработке файла',
              completedAt: new Date(),
            });
            
            // Cleanup original file
            if (req.file) {
              cleanupFile(req.file.path);
            }
          }
        });

        res.json({ 
          message: "Файл загружен и отправлен на обработку",
          fileId: processedFile.id,
          file: processedFile
        });

      } catch (error) {
        console.error('Upload error:', error);
        if (req.file) {
          cleanupFile(req.file.path);
        }
        res.status(500).json({ 
          message: error instanceof Error ? error.message : "Ошибка при загрузке файла" 
        });
      }
    });
  });

  // Import from Google Sheets
  app.post("/api/import-google-sheets", async (req: Request, res: Response) => {
    try {
      const { url } = req.body;
      
      if (!url) {
        return res.status(400).json({ message: "URL Google Таблиц обязателен" });
      }
      
      if (!validateGoogleSheetsUrl(url)) {
        return res.status(400).json({ message: "Неверный URL Google Таблиц" });
      }
      
      console.log('Google Sheets import request:', url);
      
      // Создаем запись о файле
      const processedFile = await storage.createProcessedFile({
                  originalName: `Google Sheets Import - ${new Date().toISOString()}`,
          processedName: getOutputFileName(`google_sheets_import_${Date.now()}.xlsx`),
          status: 'processing',
          fileSize: 0, // Неизвестен пока
          rowsProcessed: 0,
          statistics: null,
          errorMessage: null,
      });

      // Обрабатываем асинхронно с отслеживанием прогресса
      setImmediate(async () => {
        try {
          console.log('🔄 Начинаем импорт из Google Таблиц');
          
          // Update to show download stage
          console.log('📊 PROGRESS: Setting downloading stage (30%)');
          await storage.updateProcessedFile(processedFile.id, {
            status: 'processing',
            rowsProcessed: 1,
            statistics: JSON.stringify({ stage: 'downloading', message: 'Загрузка данных из Google Таблиц...' }),
          });

          // Импортируем данные из Google Таблиц
          console.log('🔄 PROGRESS: Starting Google Sheets import...');
          const fileBuffer = await importFromGoogleSheets(url);
          console.log('✅ PROGRESS: Google Sheets import completed, size:', fileBuffer.length);
          
          // Update to show processing stage
          console.log('📊 PROGRESS: Setting processing stage (70%)');
          await storage.updateProcessedFile(processedFile.id, {
            status: 'processing',
            rowsProcessed: 2,
            statistics: JSON.stringify({ stage: 'processing', message: 'Обработка данных...' }),
          });
          
          // Сохраняем во временный файл
          const tempFileName = `temp_google_sheets_${Date.now()}.xlsx`;
          const tempPath = path.join(process.cwd(), 'uploads', tempFileName);
          fs.writeFileSync(tempPath, fileBuffer);
          console.log('📁 PROGRESS: Temp file saved:', tempPath);
          
          // Обрабатываем данные с объединенным процессором
          console.log('🔄 PROGRESS: Starting file processing with unified processor...');
          const result = await unifiedProcessor.processExcelFile(tempPath);
          console.log('✅ PROGRESS: File processing completed, path:', result.outputPath);
          
          // Удаляем временный файл
          fs.unlinkSync(tempPath);
          console.log('🗑️ PROGRESS: Temp file deleted');
          
          // Обновляем запись об успешном завершении
          const actualFileName = path.basename(result.outputPath);
          console.log('✅ PROGRESS: Setting completed stage (100%), file:', actualFileName);
          await storage.updateProcessedFile(processedFile.id, {
            processedName: actualFileName,
            status: 'completed',
            fileSize: fileBuffer.length,
            rowsProcessed: result.statistics.totalRows,
            statistics: JSON.stringify(result.statistics),
            completedAt: new Date(),
          });

          console.log('✅ Импорт из Google Таблиц завершен успешно');
          
        } catch (error) {
          console.error('❌ Ошибка при импорте из Google Таблиц:', error);
          await storage.updateProcessedFile(processedFile.id, {
            status: 'error',
            errorMessage: error instanceof Error ? error.message : 'Ошибка при импорте из Google Таблиц',
            completedAt: new Date(),
          });
        }
      });

      res.json({ 
        message: "Импорт из Google Таблиц запущен",
        fileId: processedFile.id,
        file: processedFile
      });

    } catch (error) {
      console.error('Google Sheets import error:', error);
      res.status(500).json({ 
        message: error instanceof Error ? error.message : "Ошибка при импорте из Google Таблиц" 
      });
    }
  });

  // Get processing status
  app.get("/api/files/:id", async (req: Request, res: Response) => {
    try {
      const id = parseInt(req.params.id);
      const file = await storage.getProcessedFile(id);
      
      if (!file) {
        return res.status(404).json({ message: "Файл не найден" });
      }

      res.json(file);
    } catch (error) {
      console.error('Get file error:', error);
      res.status(500).json({ message: "Ошибка при получении информации о файле" });
    }
  });

  // Download processed file
  app.get("/api/files/:id/download", async (req: Request, res: Response) => {
    try {
      const id = parseInt(req.params.id);
      const file = await storage.getProcessedFile(id);
      
      if (!file) {
        return res.status(404).json({ message: "Файл не найден" });
      }

      if (file.status !== 'completed') {
        return res.status(400).json({ message: "Файл еще не готов для скачивания" });
      }

      const filePath = path.join(process.cwd(), 'uploads', file.processedName);
      
      if (!fs.existsSync(filePath)) {
        return res.status(404).json({ message: "Обработанный файл не найден на диске" });
      }

      res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(file.processedName)}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      
      const fileStream = fs.createReadStream(filePath);
      fileStream.pipe(res);

    } catch (error) {
      console.error('Download error:', error);
      res.status(500).json({ message: "Ошибка при скачивании файла" });
    }
  });

  // Get all processed files
  app.get("/api/files", async (req: Request, res: Response) => {
    try {
      const files = await storage.getAllProcessedFiles();
      res.json(files);
    } catch (error) {
      console.error('Get files error:', error);
      res.status(500).json({ message: "Ошибка при получении списка файлов" });
    }
  });

  // Delete processed file
  app.delete("/api/files/:id", async (req: Request, res: Response) => {
    try {
      const id = parseInt(req.params.id);
      const file = await storage.getProcessedFile(id);
      
      if (!file) {
        return res.status(404).json({ message: "Файл не найден" });
      }

      // Delete physical files
      const originalPath = path.join(process.cwd(), 'uploads', file.originalName);
      const processedPath = path.join(process.cwd(), 'uploads', file.processedName);
      
      cleanupFile(originalPath);
      cleanupFile(processedPath);

      // Delete from storage
      const deleted = await storage.deleteProcessedFile(id);
      
      if (deleted) {
        res.json({ message: "Файл удален" });
      } else {
        res.status(500).json({ message: "Ошибка при удалении файла" });
      }

    } catch (error) {
      console.error('Delete error:', error);
      res.status(500).json({ message: "Ошибка при удалении файла" });
    }
  });

  const httpServer = createServer(app);
  return httpServer;
}
