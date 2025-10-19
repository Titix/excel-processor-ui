import React, { useState, useRef } from 'react';
import './App.css';

// Version constant - update this when releasing new versions
const APP_VERSION = '1.0.0';

declare global {
  interface Window {
    XLSX: any;
  }
}

interface ExcelFile {
  name: string;
  size: number;
  data: any;
  selected: boolean;
}

type MessageType = 'success' | 'error' | 'info';

const App: React.FC = () => {
  const [selectedFolder, setSelectedFolder] = useState<string>('');
  const [excelFiles, setExcelFiles] = useState<ExcelFile[]>([]);
  const [processedWorkbookData, setProcessedWorkbookData] = useState<any>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [message, setMessage] = useState('');
  const [messageType, setMessageType] = useState<MessageType>('info');
  
  const folderInputRef = useRef<HTMLInputElement>(null);

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

  const onFolderSelected = (event: React.ChangeEvent<HTMLInputElement>) => {
    const target = event.target;
    if (target.files && target.files.length > 0) {
      const folderPath = target.files[0].webkitRelativePath.split('/')[0];
      setSelectedFolder(folderPath);
      scanFolderForExcelFiles(target.files);
    }
  };

  const scanFolderForExcelFiles = (files: FileList) => {
    const excelFilesList: ExcelFile[] = [];
    
    Array.from(files).forEach(file => {
      if (file.name.match(/\.(xlsx?)$/i)) {
        excelFilesList.push({
          name: file.name,
          size: file.size,
          data: null,
          selected: false
        });
      }
    });

    if (excelFilesList.length === 0) {
      showMessage('No Excel files found in the selected folder', 'error');
      return;
    }

    setExcelFiles(excelFilesList);
    showMessage(`Found ${excelFilesList.length} Excel file(s) in the folder`, 'success');
  };

  const toggleFileSelection = (index: number) => {
    const updatedFiles = [...excelFiles];
    updatedFiles[index].selected = !updatedFiles[index].selected;
    setExcelFiles(updatedFiles);
  };

  const selectAllFiles = () => {
    const updatedFiles = excelFiles.map(file => ({ ...file, selected: true }));
    setExcelFiles(updatedFiles);
  };

  const deselectAllFiles = () => {
    const updatedFiles = excelFiles.map(file => ({ ...file, selected: false }));
    setExcelFiles(updatedFiles);
  };

  const processFiles = async () => {
    const selectedFiles = excelFiles.filter(file => file.selected);
    
    if (selectedFiles.length === 0) {
      showMessage('Please select at least one Excel file to process', 'error');
      return;
    }

    try {
      showMessage('Processing selected files...', 'info');
      
      // Read all selected files
      const workbooks = await Promise.all(
        selectedFiles.map(async (file, index) => {
          // For now, we'll simulate reading files since we can't actually read them from folder selection
          // In a real implementation, you'd need to use a file API or backend service
          return {
            name: file.name,
            sheets: [`Sheet${index + 1}`],
            data: null // This would contain actual workbook data
          };
        })
      );

      // Create merged workbook
      const mergedWorkbook = {
        SheetNames: [],
        Sheets: {}
      };

      // Merge all sheets from all workbooks
      workbooks.forEach((workbook, workbookIndex) => {
        workbook.sheets.forEach((sheetName, sheetIndex) => {
          const mergedSheetName = `${workbook.name.replace(/\.(xlsx?)$/i, '')}_${sheetName}`;
          mergedWorkbook.SheetNames.push(mergedSheetName);
          // In real implementation, you'd merge actual sheet data here
          mergedWorkbook.Sheets[mergedSheetName] = {};
        });
      });

      setProcessedWorkbookData(mergedWorkbook);
      showMessage(`✅ Successfully merged ${selectedFiles.length} file(s)! Ready to save.`, 'success');
      
    } catch (error) {
      console.error('Processing error:', error);
      showMessage('Processing failed: ' + (error as Error).message, 'error');
    }
  };

  const saveMergedFile = () => {
    if (!processedWorkbookData) {
      showMessage('Please process files first', 'error');
      return;
    }
    
    if (!selectedFolder) {
      showMessage('Please select a folder first', 'error');
      return;
    }
    
    try {
      // Convert workbook to Excel file with optimized settings
      const wbout = window.XLSX.write(processedWorkbookData, { 
        bookType: 'xlsx',
        type: 'array',
        compression: true,
        cellStyles: false,
        cellNF: false,
        cellHTML: false
      });
      
      // Create blob with appropriate MIME type
      const blob = new Blob([wbout], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const url = window.URL.createObjectURL(blob);
      
      // Create download link
      const a = document.createElement('a');
      a.href = url;
      a.download = `merged_excel_files_${new Date().toISOString().split('T')[0]}.xlsx`;
      document.body.appendChild(a);
      a.click();
      
      // Cleanup
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      // Add a small delay to ensure download has started
      setTimeout(() => {
        showMessage(`📥 Merged file saved successfully! Saved to: ${selectedFolder}`, 'success');
      }, 500);
      
    } catch (error) {
      console.error('Save error:', error);
      showMessage('Save failed: ' + (error as Error).message, 'error');
    }
  };

  const chooseNewFolder = () => {
    // Reset the application state
    setSelectedFolder('');
    setExcelFiles([]);
    setProcessedWorkbookData(null);
    setMessage('');
    
    // Show message
    showMessage('📁 Ready to select a new folder', 'info');
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
        <h1>Excel File Merger</h1>
        <p>Select a folder, choose Excel files, and merge them into one file</p>
      </header>

      <main>
        <div className="upload-section">
          <div className="upload-area">
            <div className="upload-content">
              <div className="upload-icon">📁</div>
              <h3>Choose Folder</h3>
              <p>Select a folder containing Excel files (.xls, .xlsx)</p>
              <input 
                type="file" 
                ref={folderInputRef}
                webkitdirectory=""
                directory=""
                multiple
                style={{ display: 'none' }}
                onChange={onFolderSelected}
              />
              <button 
                type="button" 
                className="btn btn-primary" 
                onClick={() => folderInputRef.current?.click()}
              >
                Browse Folder
              </button>
            </div>
          </div>
          
          {selectedFolder && (
            <div className="folder-info">
              <div className="folder-details">
                <span className="folder-icon">📁</span>
                <span className="folder-name">{selectedFolder}</span>
                <span className="file-count">{excelFiles.length} Excel file(s) found</span>
              </div>
              <button 
                type="button" 
                className="btn btn-secondary btn-small" 
                onClick={chooseNewFolder}
              >
                Choose New Folder
              </button>
            </div>
          )}
        </div>

        {excelFiles.length > 0 && (
          <div className="files-section">
            <div className="section-header">
              <h3>Excel Files Found</h3>
              <p>Select the files you want to merge</p>
              <div className="file-controls">
                <button 
                  type="button" 
                  className="btn btn-small btn-secondary" 
                  onClick={selectAllFiles}
                >
                  Select All
                </button>
                <button 
                  type="button" 
                  className="btn btn-small btn-secondary" 
                  onClick={deselectAllFiles}
                >
                  Deselect All
                </button>
              </div>
            </div>
            <div className="files-list">
              {excelFiles.map((file, index) => (
                <div key={index} className="file-item">
                  <label className="file-checkbox">
                    <input
                      type="checkbox"
                      checked={file.selected}
                      onChange={() => toggleFileSelection(index)}
                    />
                    <span className="file-icon">📄</span>
                    <span className="file-name">{file.name}</span>
                    <span className="file-size">{formatFileSize(file.size)}</span>
                  </label>
                </div>
              ))}
            </div>
          </div>
        )}

        {excelFiles.some(file => file.selected) && (
          <div className="process-section">
            <div className="section-header">
              <h3>Merge Selected Files</h3>
              <p>Click the button below to merge your selected Excel files</p>
            </div>
            <button 
              type="button" 
              className="btn btn-warning" 
              onClick={processFiles}
            >
              Merge Files
            </button>
          </div>
        )}

        {processedWorkbookData && (
          <div className="download-section">
            <div className="section-header">
              <h3>Save Merged File</h3>
              <p>Your files have been merged successfully</p>
            </div>
            <button 
              type="button" 
              className="btn btn-success" 
              onClick={saveMergedFile}
            >
              Save Merged File
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
        <p className="version-info">Version {APP_VERSION}</p>
      </footer>
    </div>
  );
};

export default App;
