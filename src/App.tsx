import React, { useState, useRef } from 'react';
import './App.css';

declare global {
  interface Window {
    XLSX: any;
  }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
interface FileData {
  name: string;
  size: number;
  data: any;
}

type MessageType = 'success' | 'error' | 'info';

const App: React.FC = () => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [workbookData, setWorkbookData] = useState<any>(null);
  const [processedWorkbookData, setProcessedWorkbookData] = useState<any>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [message, setMessage] = useState('');
  const [messageType, setMessageType] = useState<MessageType>('info');
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  const showMessage = (messageText: string, type: MessageType) => {
    setMessage(messageText);
    setMessageType(type);
    
    // Auto-hide success messages after 8 seconds
    if (type === 'success') {
      setTimeout(() => {
        setMessage('');
      }, 8000);
    }
  };

  const onDragOver = (event: React.DragEvent) => {
    event.preventDefault();
    setIsDragOver(true);
  };

  const onDragLeave = (event: React.DragEvent) => {
    event.preventDefault();
    setIsDragOver(false);
  };

  const onDrop = (event: React.DragEvent) => {
    event.preventDefault();
    setIsDragOver(false);
    
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      handleFile(files[0]);
    }
  };

  const onFileSelected = (event: React.ChangeEvent<HTMLInputElement>) => {
    const target = event.target;
    if (target.files && target.files.length > 0) {
      handleFile(target.files[0]);
    }
  };

  const handleFile = (file: File) => {
    // Validate file type
    if (!file.name.match(/\.(xlsx?)$/i)) {
      showMessage('Please select an Excel file (.xls or .xlsx)', 'error');
      return;
    }
    
    // Validate file size (50MB limit)
    if (file.size > 50 * 1024 * 1024) {
      showMessage('File size must be less than 50MB', 'error');
      return;
    }
    
    setSelectedFile(file);
    readExcelFile(file);
  };

  const readExcelFile = (file: File) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = window.XLSX.read(data, { type: 'array' });
        setWorkbookData(workbook);
        showMessage('📁 File loaded successfully! Ready to process.', 'success');
      } catch (error) {
        console.error('Error reading Excel file:', error);
        showMessage('Error reading Excel file: ' + (error as Error).message, 'error');
        setWorkbookData(null);
      }
    };
    
    reader.onerror = () => {
      showMessage('Error reading file', 'error');
      setWorkbookData(null);
    };
    
    reader.readAsArrayBuffer(file);
  };

  const processFile = () => {
    if (!workbookData) {
      showMessage('Please select a file first', 'error');
      return;
    }
    
    try {
      // Create a copy of the workbook for processing (safer approach)
      const processedWorkbook = window.XLSX.read(
        window.XLSX.write(workbookData, { type: 'array' }), 
        { type: 'array' }
      );
      
      // Process each sheet
      Object.keys(processedWorkbook.Sheets).forEach(sheetName => {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const sheet = processedWorkbook.Sheets[sheetName];
        
        // Here you can add your processing logic
        // For now, we'll just copy the data as-is
        // You can modify cells, add formulas, etc.
      });
      
      setProcessedWorkbookData(processedWorkbook);
      showMessage('✅ File processed successfully! Ready to download.', 'success');
      
    } catch (error) {
      console.error('Processing error:', error);
      showMessage('Processing failed: ' + (error as Error).message, 'error');
    }
  };

  const downloadFile = () => {
    if (!processedWorkbookData) {
      showMessage('Please process a file first', 'error');
      return;
    }
    
    try {
      // Determine the original file format
      const originalFormat = selectedFile?.name.toLowerCase().endsWith('.xls') ? 'xls' : 'xlsx';
      
      // Convert workbook to Excel file with optimized settings for compatibility and size
      const wbout = window.XLSX.write(processedWorkbookData, { 
        bookType: originalFormat,
        type: 'array',
        compression: true,  // Enable compression for smaller files
        cellStyles: false,  // Disable cell styles to reduce size
        cellNF: false,      // Disable number formats to reduce size
        cellHTML: false     // Disable HTML in cells to reduce size
      });
      
      // Create blob with appropriate MIME type
      const mimeType = originalFormat === 'xls' 
        ? 'application/vnd.ms-excel' 
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      
      const blob = new Blob([wbout], { type: mimeType });
      const url = window.URL.createObjectURL(blob);
      
      // Create download link with original format extension
      const a = document.createElement('a');
      a.href = url;
      a.download = selectedFile ? selectedFile.name.replace(/\.(xlsx?)$/i, '_processed.$1') : `processed_file.${originalFormat}`;
      document.body.appendChild(a);
      a.click();
      
      // Cleanup
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      // Add a small delay to ensure download has started
      setTimeout(() => {
        showMessage('📥 File downloaded successfully! Check your downloads folder.', 'success');
      }, 500);
      
    } catch (error) {
      console.error('Download error:', error);
      showMessage('Download failed: ' + (error as Error).message, 'error');
    }
  };

  const chooseNewFile = () => {
    // Reset the application state
    setSelectedFile(null);
    setWorkbookData(null);
    setProcessedWorkbookData(null);
    setMessage('');
    
    // Show message
    showMessage('📁 Ready to select a new file', 'info');
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div className="container">
      <header>
        <h1>Excel Processor</h1>
        <p>Upload, process, and download Excel files with ease</p>
      </header>

      <main>
        <div className="upload-section">
          <div 
            className={`upload-area ${isDragOver ? 'dragover' : ''}`}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
          >
            <div className="upload-content">
              <div className="upload-icon">📁</div>
              <h3>Choose Excel File</h3>
              <p>Drag and drop your Excel file here or click to browse</p>
              <input 
                type="file" 
                ref={fileInputRef}
                accept=".xls,.xlsx" 
                style={{ display: 'none' }}
                onChange={onFileSelected}
              />
              <button 
                type="button" 
                className="btn btn-primary" 
                onClick={() => fileInputRef.current?.click()}
              >
                Browse Files
              </button>
            </div>
          </div>
          
          {selectedFile && (
            <div className="file-info">
              <div className="file-details">
                <span className="file-icon">📄</span>
                <span className="file-name">{selectedFile.name}</span>
                <span className="file-size">{formatFileSize(selectedFile.size)}</span>
              </div>
              <button 
                type="button" 
                className="btn btn-secondary btn-small" 
                onClick={chooseNewFile}
              >
                Choose New File
              </button>
            </div>
          )}
        </div>

        {selectedFile && (
          <div className="process-section">
            <div className="section-header">
              <h3>Process File</h3>
              <p>Click the button below to process your uploaded file</p>
            </div>
            <button 
              type="button" 
              className="btn btn-warning" 
              disabled={!workbookData}
              onClick={processFile}
            >
              Process File
            </button>
          </div>
        )}

        {processedWorkbookData && (
          <div className="download-section">
            <div className="section-header">
              <h3>Download Processed File</h3>
              <p>Your file has been processed successfully</p>
            </div>
            <button 
              type="button" 
              className="btn btn-success" 
              onClick={downloadFile}
            >
              Save Processed File
            </button>
          </div>
        )}

        {message && (
          <div className={`message-area ${messageType}`}>
            {message}
          </div>
        )}
      </main>

      <footer>
        <p>&copy; 2025 Excel Processor. Built with React and Node.js.</p>
      </footer>
    </div>
  );
};

export default App;
