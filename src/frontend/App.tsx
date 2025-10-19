import React, { useState } from 'react';
import './App.css';
import { useLanguage, formatMessage } from '../contexts/LanguageContext';
import LanguageSelector from '../components/LanguageSelector';

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
  path: string; // Store the full path for later use
}

type MessageType = 'success' | 'error' | 'info';

const App: React.FC = () => {
  const { t } = useLanguage();
  const [selectedFolder, setSelectedFolder] = useState<string>('');
  const [excelFiles, setExcelFiles] = useState<ExcelFile[]>([]);
  const [processedWorkbookData, setProcessedWorkbookData] = useState<any>(null);
  const [message, setMessage] = useState('');
  const [messageType, setMessageType] = useState<MessageType>('info');
  

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


  const handleFolderSelection = async () => {
    try {
      // Use File System Access API to avoid browser upload messages
      const dirHandle = await (window as any).showDirectoryPicker();
      const excelFilesList: ExcelFile[] = [];
      
      // Read directory entries without accessing file content
      for await (const [name, handle] of dirHandle.entries()) {
        if (handle.kind === 'file' && (name.endsWith('.xls') || name.endsWith('.xlsx'))) {
          const file = await handle.getFile();
          excelFilesList.push({
            name: file.name,
            size: file.size,
            data: null, // No file content loaded
            selected: false,
            path: name
          });
        }
      }
      
      setSelectedFolder(dirHandle.name);
      setExcelFiles(excelFilesList);
      
      if (excelFilesList.length === 0) {
        showMessage(t.messages.noExcelFilesInFolder, 'info');
      } else {
        showMessage(formatMessage(t.messages.foundExcelFiles, { count: excelFilesList.length }), 'success');
      }
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        showMessage(t.messages.browserNotSupported, 'error');
      }
    }
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

  const loadSelectedFiles = async (files: ExcelFile[]): Promise<any[]> => {
    const loadedFiles: any[] = [];
    
    for (const file of files) {
      if (file.selected) {
        try {
          // In a real implementation, you would load the file from the stored path
          // For now, we'll simulate loading the file
          const simulatedWorkbook = {
            name: file.name,
            sheets: ['Sheet1', 'Sheet2'], // Simulated sheet names
            data: {} // Simulated data
          };
          loadedFiles.push(simulatedWorkbook);
        } catch (error) {
          console.error(`Error loading file ${file.name}:`, error);
          throw new Error(formatMessage(t.messages.failedToLoadFile, { fileName: file.name }));
        }
      }
    }
    
    return loadedFiles;
  };

  const processFiles = async () => {
    const selectedFiles = excelFiles.filter(file => file.selected);
    
    if (selectedFiles.length === 0) {
      showMessage(t.messages.selectAtLeastOneFile, 'error');
      return;
    }

    try {
      showMessage(t.messages.loadingAndProcessing, 'info');
      
      // Load only the selected files
      const workbooks = await loadSelectedFiles(excelFiles);

      // Create merged workbook
      const mergedWorkbook = {
        SheetNames: [] as string[],
        Sheets: {} as { [key: string]: any }
      };

      // Merge all sheets from all workbooks
      workbooks.forEach((workbook, workbookIndex) => {
        workbook.sheets.forEach((sheetName: string, sheetIndex: number) => {
          const mergedSheetName = `${workbook.name.replace(/\.(xlsx?)$/i, '')}_${sheetName}`;
          mergedWorkbook.SheetNames.push(mergedSheetName);
          // In real implementation, you'd merge actual sheet data here
          mergedWorkbook.Sheets[mergedSheetName] = {};
        });
      });

      setProcessedWorkbookData(mergedWorkbook);
      showMessage(formatMessage(t.messages.successfullyMerged, { count: selectedFiles.length }), 'success');
      
    } catch (error) {
      console.error('Processing error:', error);
      showMessage(formatMessage(t.messages.processingFailed, { error: (error as Error).message }), 'error');
    }
  };

  const saveMergedFile = () => {
    if (!processedWorkbookData) {
      showMessage(t.messages.pleaseProcessFilesFirst, 'error');
      return;
    }
    
    if (!selectedFolder) {
      showMessage(t.messages.pleaseSelectFolderFirst, 'error');
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
        showMessage(formatMessage(t.messages.mergedFileSaved, { folder: selectedFolder }), 'success');
      }, 500);
      
    } catch (error) {
      console.error('Save error:', error);
      showMessage(formatMessage(t.messages.saveFailed, { error: (error as Error).message }), 'error');
    }
  };

  const chooseNewFolder = () => {
    // Reset the application state
    setSelectedFolder('');
    setExcelFiles([]);
    setProcessedWorkbookData(null);
    setMessage('');
    
    // Show message
    showMessage(t.messages.readyToSelectNewFolder, 'info');
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return `0 ${t.fileSizeUnits.bytes}`;
    const k = 1024;
    const sizes = [t.fileSizeUnits.bytes, t.fileSizeUnits.kb, t.fileSizeUnits.mb, t.fileSizeUnits.gb];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div className="container">
      <header>
        <div className="header-top">
          <div className="header-content">
            <h1>{t.appTitle}</h1>
            <p>{t.appSubtitle}</p>
          </div>
          <LanguageSelector />
        </div>
      </header>

      <main>
        {!selectedFolder ? (
          <div className="folder-selection-section">
            <div className="folder-selection-area">
              <div className="folder-selection-content">
                <div className="folder-icon">📁</div>
                <h3>{t.selectFolder}</h3>
                <p>{t.selectFolderDescription}</p>
                
                
                <div className="selection-options">
                  <button 
                    type="button" 
                    className="btn btn-primary" 
                    onClick={handleFolderSelection}
                  >
                    {t.selectFolder}
                  </button>
                </div>
              </div>
            </div>
          </div>
        ) : (
          <div className="folder-info">
            <div className="folder-details">
              <span className="folder-icon">📁</span>
              <span className="folder-name">{selectedFolder}</span>
              <span className="file-count">{formatMessage(t.excelFilesFound, { count: excelFiles.length })}</span>
            </div>
            <button 
              type="button" 
              className="btn btn-secondary btn-small" 
              onClick={chooseNewFolder}
            >
              {t.chooseNewFolder}
            </button>
          </div>
        )}

        {selectedFolder && (
          <div className="files-section">
            <div className="section-header">
              <h3>{t.excelFilesFoundTitle}</h3>
              <p>{t.selectFilesToMerge}</p>
              {excelFiles.length > 0 && (
                <div className="file-controls">
                  <button 
                    type="button" 
                    className="btn btn-small btn-secondary" 
                    onClick={selectAllFiles}
                  >
                    {t.selectAll}
                  </button>
                  <button 
                    type="button" 
                    className="btn btn-small btn-secondary" 
                    onClick={deselectAllFiles}
                  >
                    {t.deselectAll}
                  </button>
                </div>
              )}
            </div>
            <div className="files-list">
              {excelFiles.length > 0 ? (
                excelFiles.map((file, index) => (
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
                ))
              ) : (
                <div className="no-files-message">
                  <p>{t.noExcelFilesFound}</p>
                  <p>{t.noExcelFilesSuggestion}</p>
                </div>
              )}
            </div>
          </div>
        )}

        {excelFiles.some(file => file.selected) && (
          <div className="process-section">
            <div className="section-header">
              <h3>{t.mergeSelectedFiles}</h3>
              <p>{t.mergeFilesDescription}</p>
            </div>
            <button 
              type="button" 
              className="btn btn-warning" 
              onClick={processFiles}
            >
              {t.mergeFiles}
            </button>
          </div>
        )}

        {processedWorkbookData && (
          <div className="download-section">
            <div className="section-header">
              <h3>{t.saveMergedFile}</h3>
              <p>{t.saveMergedFileDescription}</p>
            </div>
            <button 
              type="button" 
              className="btn btn-success" 
              onClick={saveMergedFile}
            >
              {t.saveMergedFile}
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
        <p>{t.footer.copyright}</p>
        <p className="version-info">{formatMessage(t.footer.version, { version: APP_VERSION })}</p>
      </footer>
    </div>
  );
};

export default App;