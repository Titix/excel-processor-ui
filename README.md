# Excel Processor UI

A modern web application for processing Excel files built with React frontend and Node.js backend. This application provides client-side Excel processing with a clean, responsive interface.

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
- **React 18**: Modern component-based framework with hooks
- **TypeScript**: Type-safe JavaScript
- **SheetJS (XLSX)**: Client-side Excel processing
- **CSS3**: Modern styling with gradients and animations

### Backend
- **Node.js**: JavaScript runtime
- **Express.js**: Web framework
- **CORS**: Cross-origin resource sharing

## 📋 Prerequisites

- **Node.js** (v14 or higher)
- **npm** (v6 or higher)

## 🚀 Quick Start

### 1. Install Node.js and npm

Download and install Node.js from [https://nodejs.org/](https://nodejs.org/)
This will also install npm automatically.

### 2. Install Dependencies

```bash
# Install all dependencies
npm install
```

### 3. Development Mode

```bash
# Start the React development server
npm start
# Application will be available at http://localhost:3000 with hot reloading
```

### 4. Production Build

```bash
# Build the React application
npm run build

# Start the production server
npm run server
# Application will be available at http://localhost:3000
```

## 📁 Project Structure

```
excel-processor-ui/
├── src/
│   ├── frontend/             # React frontend source
│   │   ├── App.tsx          # Main React component
│   │   ├── App.css          # Component styles
│   │   ├── index.tsx        # React entry point
│   │   ├── index.css        # Global styles
│   │   └── public/          # Static assets
│   │       ├── index.html   # HTML template
│   │       └── manifest.json # PWA manifest
│   └── backend/             # Node.js backend source
│       ├── server.js        # Express server
│       └── build.js         # Build script
├── src/                     # React build files (symlinked)
├── public/                  # React public files (symlinked)
├── build/                   # Production build (generated)
├── package.json             # Dependencies and scripts
├── tsconfig.json            # React TypeScript config
└── README.md
```

## 🎯 Usage

1. **Select File**: Click "Browse Files" or drag and drop an Excel file
2. **Process File**: Click "Process File" to process the uploaded file
3. **Download**: Click "Save Processed File" to download the processed file
4. **Choose New File**: Click "Choose New File" to select a different file

## 🔧 Configuration

### Port Configuration
- **Development**: Port 3000 (React dev server)
- **Production**: Port 3000 (configurable via PORT environment variable)

### File Limits
- **Maximum file size**: 50MB
- **Supported formats**: .xls, .xlsx

## 🚀 Deployment

### Production Build
```bash
npm run build
npm run server
```

### Environment Variables
- `PORT`: Server port (default: 3000)

## 📝 Available Scripts

- `npm start` - Start React development server
- `npm run build` - Build React app for production
- `npm run server` - Start Node.js production server
- `npm run dev` - Start server with nodemon (auto-restart)
- `npm run build-server` - Run build script
- `npm test` - Run React tests
- `npm run eject` - Eject from Create React App

## 🔍 Key Features

| Feature | Description |
|---------|-------------|
| Frontend | React 18 with TypeScript |
| Backend | Node.js + Express |
| Build Tool | Create React App |
| Processing | Client-side with SheetJS |
| Deployment | Node.js server |
| Structure | Organized src/frontend and src/backend |

## 🐛 Troubleshooting

### Common Issues

1. **Port already in use**
   ```bash
   # Kill process using port 3000
   taskkill /F /IM node.exe
   ```

2. **Build fails**
   ```bash
   # Clear npm cache and reinstall
   npm cache clean --force
   npm install
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

**Note**: This React version provides modern web technologies for better performance and developer experience with hot reloading and TypeScript support. The project structure is organized with clear separation between frontend and backend code.