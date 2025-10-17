# Excel Processor UI

A modern web application for processing Excel files built with Angular frontend and Node.js backend. This application provides the same functionality as the Spring Boot version but uses modern web technologies.

## 🚀 Features

- **Client-Side Processing**: All Excel processing happens in the browser for privacy and performance
- **Drag & Drop Support**: Easy file selection with drag and drop functionality
- **File Format Support**: Supports both .xls and .xlsx files
- **Optimized Output**: Generates smaller, Excel-compatible processed files
- **Choose New File**: Select different files at any time during the workflow
- **Responsive Design**: Works on desktop and mobile devices
- **Modern UI**: Clean, professional interface with smooth animations

## 🛠️ Technology Stack

### Frontend
- **Angular 17**: Modern component-based framework
- **TypeScript**: Type-safe JavaScript
- **SheetJS (XLSX)**: Client-side Excel processing
- **CSS3**: Modern styling with gradients and animations

### Backend
- **Node.js**: JavaScript runtime
- **Express.js**: Web framework
- **CORS**: Cross-origin resource sharing

## 📋 Prerequisites

- **Node.js** (v18 or higher)
- **npm** (v8 or higher)

## 🚀 Quick Start

### 1. Install Node.js and npm

Download and install Node.js from [https://nodejs.org/](https://nodejs.org/)
This will also install npm automatically.

### 2. Install Dependencies

```bash
# Install backend dependencies
npm install

# Install frontend dependencies
cd frontend
npm install
cd ..
```

### 3. Development Mode

```bash
# Start the Angular development server (frontend)
cd frontend
npm start
# Frontend will be available at http://localhost:4200

# In another terminal, start the Node.js server (backend)
cd ..
npm run dev
# Backend will be available at http://localhost:3000
```

### 4. Production Build

```bash
# Build the Angular application
npm run build

# Start the production server
npm start
# Application will be available at http://localhost:3000
```

## 📁 Project Structure

```
excel-processor-ui/
├── frontend/                 # Angular frontend
│   ├── src/
│   │   ├── app/
│   │   │   └── app.component.ts
│   │   ├── assets/
│   │   ├── index.html
│   │   ├── main.ts
│   │   └── styles.css
│   ├── angular.json
│   ├── package.json
│   └── tsconfig.json
├── server.js                 # Node.js backend
├── package.json              # Backend dependencies
└── README.md
```

## 🎯 Usage

1. **Select File**: Click "Browse Files" or drag and drop an Excel file
2. **Process File**: Click "Process File" to process the uploaded file
3. **Download**: Click "Save Processed File" to download the processed file
4. **Choose New File**: Click "Choose New File" to select a different file

## 🔧 Configuration

### Port Configuration
- **Frontend (Development)**: Port 4200
- **Backend**: Port 3000 (configurable via PORT environment variable)

### File Limits
- **Maximum file size**: 50MB
- **Supported formats**: .xls, .xlsx

## 🚀 Deployment

### Production Build
```bash
npm run build
npm start
```

### Environment Variables
- `PORT`: Server port (default: 3000)

## 🔍 Key Differences from Spring Boot Version

| Feature | Spring Boot | Angular + Node.js |
|---------|-------------|-------------------|
| Frontend | HTML/CSS/JS | Angular + TypeScript |
| Backend | Spring Boot | Node.js + Express |
| Build Tool | Maven | npm |
| Processing | Client-side | Client-side |
| Deployment | JAR file | Node.js server |

## 🐛 Troubleshooting

### Common Issues

1. **Port already in use**
   ```bash
   # Kill process using port 3000
   npx kill-port 3000
   ```

2. **Angular build fails**
   ```bash
   # Clear Angular cache
   cd frontend
   npm run ng cache clean
   ```

3. **File not processing**
   - Ensure file is .xls or .xlsx format
   - Check file size is under 50MB
   - Verify browser supports FileReader API

## 📝 License

MIT License - see LICENSE file for details.

## 👨‍💻 Author

Built by Titix

---

**Note**: This Angular version provides the same functionality as the Spring Boot version but uses modern web technologies for better performance and developer experience.