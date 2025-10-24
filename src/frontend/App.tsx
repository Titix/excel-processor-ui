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
  const [directoryHandle, setDirectoryHandle] = useState<any>(null);
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
            setDirectoryHandle(dirHandle);
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
          // Load the actual Excel file using the stored directory handle
          const fileHandle = await directoryHandle.getFileHandle(file.path);
          const fileData = await fileHandle.getFile();
          const arrayBuffer = await fileData.arrayBuffer();
          
          const workbook = window.XLSX.read(arrayBuffer, { type: 'array' });
          
          loadedFiles.push({
            name: file.name,
            path: file.path,
            workbook: workbook,
            sheets: workbook.SheetNames
          });
        } catch (error) {
          console.error(`Error loading file ${file.name}:`, error);
          throw new Error(formatMessage(t.messages.failedToLoadFile, { fileName: file.name }));
        }
      }
    }
    
    return loadedFiles;
  };

  // Function to validate column consistency across files
  const validateColumnConsistency = (loadedFiles: any[]): { isValid: boolean; message: string } => {
    if (loadedFiles.length <= 1) {
      return { isValid: true, message: '' };
    }

    const allColumns: { [fileName: string]: string[] } = {};
    
    // Extract column headers from each file
    loadedFiles.forEach(file => {
      const columns: string[] = [];
      file.workbook.SheetNames.forEach((sheetName: string) => {
        const worksheet = file.workbook.Sheets[sheetName];
        if (worksheet && worksheet['!ref']) {
          const range = window.XLSX.utils.decode_range(worksheet['!ref']);
          // Get first row as headers
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = window.XLSX.utils.encode_cell({ r: range.s.r, c: col });
            const cell = worksheet[cellAddress];
            if (cell && cell.v !== undefined) {
              columns.push(String(cell.v));
            }
          }
        }
      });
      allColumns[file.name] = columns;
    });

    // Compare columns across files
    const fileNames = Object.keys(allColumns);
    const firstFileColumns = allColumns[fileNames[0]];
    const mismatchedFiles: string[] = [];

    for (let i = 1; i < fileNames.length; i++) {
      const currentFileColumns = allColumns[fileNames[i]];
      if (JSON.stringify(firstFileColumns) !== JSON.stringify(currentFileColumns)) {
        mismatchedFiles.push(fileNames[i]);
      }
    }

    if (mismatchedFiles.length > 0) {
      const message = `Column mismatch detected! Files with different columns: ${mismatchedFiles.join(', ')}. Expected columns: ${firstFileColumns.join(', ')}`;
      return { isValid: false, message };
    }

    return { isValid: true, message: `All files have consistent columns: ${firstFileColumns.join(', ')}` };
  };

  const processFiles = async () => {
    const selectedFiles = excelFiles.filter(file => file.selected);
    
    if (selectedFiles.length === 0) {
      showMessage(t.messages.selectAtLeastOneFile, 'error');
      return;
    }

    try {
      showMessage(t.messages.loadingAndProcessing, 'info');
      
      // Load the actual Excel files and read their data
      const loadedFiles = await loadSelectedFiles(selectedFiles);
      
      // Validate column consistency
      const validation = validateColumnConsistency(loadedFiles);
      if (!validation.isValid) {
        showMessage(validation.message, 'error');
        return;
      }
      
      // Show validation success message
      if (validation.message) {
        showMessage(validation.message, 'info');
      }
      
      // Create merged workbook
      const mergedWorkbook = {
        SheetNames: [] as string[],
        Sheets: {} as { [key: string]: any }
      };

      // Merge all sheets from all workbooks into a single sheet named "processed"
      const mergedSheetName = "processed";
      mergedWorkbook.SheetNames.push(mergedSheetName);
      
      // Step 1: Collect all data from all files
      const allRows: any[][] = [];
      let headers: any[] = [];
      let totalRowsCollected = 0;
      
      console.log('=== COLLECTING DATA FROM ALL FILES ===');
      
      loadedFiles.forEach((file, fileIndex) => {
        console.log(`Processing file ${fileIndex + 1}: ${file.name}`);
        
        const workbook = file.workbook;
        
        workbook.SheetNames.forEach((sheetName: string) => {
          const worksheet = workbook.Sheets[sheetName];
          
          if (worksheet && worksheet['!ref']) {
            try {
              // Convert sheet to array of arrays for easier processing
              const jsonData = window.XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: '', 
                raw: false 
              }) as any[][];
              
              if (jsonData.length === 0) return;
              
              // Use headers from first file
              if (fileIndex === 0 && allRows.length === 0) {
                headers = jsonData[0] || [];
                console.log('Headers detected:', headers);
              }
              
              // Add all data rows (skip header)
              const dataRows = jsonData.slice(1);
              allRows.push(...dataRows);
              totalRowsCollected += dataRows.length;
              
              console.log(`Added ${dataRows.length} rows from ${file.name} sheet ${sheetName}`);
            } catch (error) {
              console.warn(`Error processing sheet ${sheetName} in file ${file.name}:`, error);
            }
          }
        });
      });
      
      console.log(`Total rows collected: ${totalRowsCollected}`);
      
      // Step 2: Remove duplicates across all files
      const uniqueRows: any[][] = [];
      const seenRows = new Set<string>();
      let duplicatesRemoved = 0;
      
      console.log('=== REMOVING DUPLICATES ===');
      
      allRows.forEach((row, index) => {
        // Convert row to string for comparison
        const rowString = row.map(cell => 
          cell === null || cell === undefined ? '' : String(cell).trim()
        ).join('|');
        
        if (!seenRows.has(rowString)) {
          seenRows.add(rowString);
          uniqueRows.push(row);
        } else {
          duplicatesRemoved++;
          console.log(`Duplicate removed:`, row);
        }
      });
      
      console.log(`Duplicates removed: ${duplicatesRemoved}`);
      console.log(`Unique rows: ${uniqueRows.length}`);
      
      // Step 3: Create merged sheet with unique rows only
      const mergedSheetData: any = {};
      const finalData = [headers, ...uniqueRows];
      
      // Convert final data back to Excel format
      finalData.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (cell !== undefined && cell !== null && cell !== '') {
            const cellAddress = window.XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            mergedSheetData[cellAddress] = { v: cell, t: typeof cell === 'number' ? 'n' : 's' };
          }
        });
      });
      
      // Set the range for the merged sheet
      const maxRow = finalData.length - 1;
      const maxCol = headers.length - 1;
      
      mergedSheetData['!ref'] = `A1:${window.XLSX.utils.encode_cell({ r: maxRow, c: maxCol })}`;
      
      mergedWorkbook.Sheets[mergedSheetName] = mergedSheetData;

      setProcessedWorkbookData(mergedWorkbook);
      
      // Show success message with duplicate removal results
      const successMessage = `✅ Files processed successfully! 
        📊 Total rows processed: ${totalRowsCollected}
        🚫 Duplicates removed: ${duplicatesRemoved}
        ✅ Unique rows in result: ${uniqueRows.length}
        📁 Ready to download merged file.`;
      
      showMessage(successMessage, 'success');
      
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
      
      // Generate filename with current timestamp in YYYYMMDDhhss format
      const now = new Date();
      const year = now.getFullYear();
      const month = String(now.getMonth() + 1).padStart(2, '0');
      const day = String(now.getDate()).padStart(2, '0');
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const seconds = String(now.getSeconds()).padStart(2, '0');
      const timestamp = `${year}${month}${day}${hours}${minutes}${seconds}`;
      
      // Create download link with shorter filename to avoid Excel limitations
      const a = document.createElement('a');
      a.href = url;
      a.download = `proc_${timestamp}.xlsx`;
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