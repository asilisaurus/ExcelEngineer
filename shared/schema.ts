import { pgTable, text, serial, integer, boolean, timestamp, json } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const processedFiles = pgTable("processed_files", {
  id: serial("id").primaryKey(),
  originalName: text("original_name").notNull(),
  processedName: text("processed_name").notNull(),
  status: text("status").notNull(), // 'processing', 'completed', 'error'
  fileSize: integer("file_size").notNull(),
  rowsProcessed: integer("rows_processed"),
  statistics: json("statistics"), // JSON object with processing stats
  errorMessage: text("error_message"),
  createdAt: timestamp("created_at").defaultNow().notNull(),
  completedAt: timestamp("completed_at"),
});

export const insertProcessedFileSchema = createInsertSchema(processedFiles).pick({
  originalName: true,
  processedName: true,
  status: true,
  fileSize: true,
  rowsProcessed: true,
  statistics: true,
  errorMessage: true,
});

export type InsertProcessedFile = z.infer<typeof insertProcessedFileSchema>;
export type ProcessedFile = typeof processedFiles.$inferSelect;

// Processing statistics schema
export const processingStatsSchema = z.object({
  totalRows: z.number(),
  reviewsCount: z.number(),
  commentsCount: z.number(),
  activeDiscussionsCount: z.number(),
  totalViews: z.number(),
  engagementRate: z.number(),
  platformsWithData: z.number(),
});

export type ProcessingStats = z.infer<typeof processingStatsSchema>;

// File upload schema
export const fileUploadSchema = z.object({
  file: z.any(), // File object from multer
});

export type FileUpload = z.infer<typeof fileUploadSchema>;
