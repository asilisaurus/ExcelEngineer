import type { Express, Request, Response } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { upload, cleanupFile, getOutputFileName } from "./services/file-handler";
import { ExcelProcessor } from "./services/excel-processor";
import { insertProcessedFileSchema, processingStatsSchema } from "@shared/schema";
import fs from 'fs';
import path from 'path';

export async function registerRoutes(app: Express): Promise<Server> {
  const excelProcessor = new ExcelProcessor();

  // Upload and process Excel file
  app.post("/api/upload", upload.single('file'), async (req: Request, res: Response) => {
    try {
      if (!req.file) {
        return res.status(400).json({ message: "Файл не был загружен" });
      }

      // Create initial record
      const processedFile = await storage.createProcessedFile({
        originalName: req.file.originalname,
        processedName: getOutputFileName(req.file.originalname),
        status: 'processing',
        fileSize: req.file.size,
        rowsProcessed: null,
        statistics: null,
        errorMessage: null,
      });

      // Process file asynchronously
      setImmediate(async () => {
        try {
          const fileBuffer = fs.readFileSync(req.file!.path);
          const { workbook, statistics } = await excelProcessor.processExcelFile(fileBuffer, req.file!.originalname);
          
          // Save processed file
          const outputPath = path.join(path.dirname(req.file!.path), processedFile.processedName);
          await workbook.xlsx.writeFile(outputPath);

          // Update record with success
          await storage.updateProcessedFile(processedFile.id, {
            status: 'completed',
            rowsProcessed: statistics.totalRows,
            statistics: statistics,
            completedAt: new Date(),
          });

          // Cleanup original file
          cleanupFile(req.file!.path);
        } catch (error) {
          console.error('Processing error:', error);
          await storage.updateProcessedFile(processedFile.id, {
            status: 'error',
            errorMessage: error instanceof Error ? error.message : 'Неизвестная ошибка',
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
