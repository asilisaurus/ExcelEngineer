# Excel Processing Application

## Overview
This is a full-stack web application built for processing Excel files containing social media analytics data. The application provides a user-friendly interface for uploading Excel files, processing them to extract analytics data, and generating formatted reports with engagement metrics.

## System Architecture

### Frontend Architecture
- **Framework**: React 18 with TypeScript
- **Build Tool**: Vite for fast development and optimized builds
- **Styling**: Tailwind CSS with shadcn/ui component library
- **Routing**: Wouter for lightweight client-side routing
- **State Management**: TanStack Query for server state management
- **File Upload**: react-dropzone for drag-and-drop file uploads

### Backend Architecture
- **Runtime**: Node.js with Express.js framework
- **Language**: TypeScript with ES modules
- **File Processing**: xlsx and ExcelJS libraries for Excel manipulation
- **File Upload**: Multer for handling multipart/form-data
- **Database**: PostgreSQL with Drizzle ORM (currently using in-memory storage)
- **Session Management**: connect-pg-simple for PostgreSQL session store

## Key Components

### Data Storage
- **Schema**: Defined in `shared/schema.ts` using Drizzle ORM
- **Tables**: `processed_files` table for tracking file processing status
- **Current Implementation**: In-memory storage (`MemStorage` class) for development
- **Production Ready**: Configured for PostgreSQL with Neon Database

### File Processing Pipeline
1. **Upload**: Files are uploaded to `/uploads` directory with unique names
2. **Validation**: Only `.xls` and `.xlsx` files up to 10MB are accepted
3. **Processing**: ExcelProcessor extracts data from specific sheets and ranges
4. **Transformation**: Cleans data, calculates engagement metrics, and formats output
5. **Storage**: Processed files are saved and metadata is stored in database

### API Endpoints
- `POST /api/upload` - Upload and process Excel files
- `GET /api/files/:id` - Get processing status and file details
- `GET /api/files` - List all processed files
- `GET /api/download/:filename` - Download processed files

## Data Flow

1. **File Upload**: User drags/drops or selects Excel file through the UI
2. **Validation**: Client-side validation for file type and size
3. **Processing**: Server processes file asynchronously with status updates
4. **Real-time Updates**: Client polls for processing status every 2 seconds
5. **Results**: User can view statistics and download processed file

### Processing Steps
1. **Data Extraction**: Reads specific ranges from Excel sheets for different platforms
2. **Data Cleaning**: Removes unnecessary columns and cleans view counts
3. **Metric Calculation**: Calculates engagement rates using formula: (comments + likes + reposts) / views * 100
4. **Report Generation**: Creates formatted Excel output with grouped data and statistics

## External Dependencies

### Frontend Dependencies
- React ecosystem (React, React DOM, React Query)
- UI components (@radix-ui/* components)
- Styling (Tailwind CSS, class-variance-authority)
- Utilities (date-fns, clsx, cmdk)

### Backend Dependencies
- Express.js and middleware
- Excel processing (xlsx, exceljs)
- Database (drizzle-orm, @neondatabase/serverless)
- File handling (multer)
- Validation (zod, drizzle-zod)

### Development Dependencies
- TypeScript and build tools
- Vite with React plugin
- Replit-specific plugins for development

## Deployment Strategy

### Development Environment
- **Command**: `npm run dev`
- **Port**: 5000 (configured in .replit)
- **Features**: Hot reloading, runtime error overlay, development middleware

### Production Build
- **Build Command**: `npm run build`
- **Serves**: Static files from `dist/public`
- **Start Command**: `npm run start`
- **Platform**: Configured for Replit autoscale deployment

### Database Configuration
- **Development**: In-memory storage for quick development
- **Production**: PostgreSQL with Drizzle migrations
- **Migrations**: Run with `npm run db:push`

## Changelog
- June 27, 2025. Initial setup

## User Preferences
Preferred communication style: Simple, everyday language.