import { processedFiles, type ProcessedFile, type InsertProcessedFile } from "@shared/schema";

export interface IStorage {
  createProcessedFile(file: InsertProcessedFile): Promise<ProcessedFile>;
  getProcessedFile(id: number): Promise<ProcessedFile | undefined>;
  getAllProcessedFiles(): Promise<ProcessedFile[]>;
  updateProcessedFile(id: number, updates: Partial<ProcessedFile>): Promise<ProcessedFile | undefined>;
  deleteProcessedFile(id: number): Promise<boolean>;
}

export class MemStorage implements IStorage {
  private files: Map<number, ProcessedFile>;
  private currentId: number;

  constructor() {
    this.files = new Map();
    this.currentId = 1;
  }

  async createProcessedFile(insertFile: InsertProcessedFile): Promise<ProcessedFile> {
    const id = this.currentId++;
    const file: ProcessedFile = {
      ...insertFile,
      id,
      createdAt: new Date(),
      completedAt: null,
    };
    this.files.set(id, file);
    return file;
  }

  async getProcessedFile(id: number): Promise<ProcessedFile | undefined> {
    return this.files.get(id);
  }

  async getAllProcessedFiles(): Promise<ProcessedFile[]> {
    return Array.from(this.files.values()).sort((a, b) => 
      new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
    );
  }

  async updateProcessedFile(id: number, updates: Partial<ProcessedFile>): Promise<ProcessedFile | undefined> {
    const existingFile = this.files.get(id);
    if (!existingFile) return undefined;

    const updatedFile = { ...existingFile, ...updates };
    this.files.set(id, updatedFile);
    return updatedFile;
  }

  async deleteProcessedFile(id: number): Promise<boolean> {
    return this.files.delete(id);
  }
}

export const storage = new MemStorage();
