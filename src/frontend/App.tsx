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
  const [duplicatesWorkbookData, setDuplicatesWorkbookData] = useState<any>(null);
  const [message, setMessage] = useState('');
  const [messageType, setMessageType] = useState<MessageType>('info');
  const [filterDuplicates, setFilterDuplicates] = useState<boolean>(false); // Keep disabled for now
  

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
      }
      // Removed the success message - folder info already shows the count in the UI
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

  // Function to remove duplicates from rows
  const removeDuplicates = (rows: any[][]): { uniqueRows: any[][], duplicatesRemoved: number } => {
    const uniqueRows: any[][] = [];
    const seenRows = new Set<string>();
    let duplicatesRemoved = 0;
    
    console.log('=== REMOVING DUPLICATES ===');
    
    rows.forEach((row, index) => {
      // Convert row to string for comparison
      const rowString = row.map(cell => 
        cell === null || cell === undefined ? '' : String(cell).trim()
      ).join('|');
      
      if (!seenRows.has(rowString)) {
        seenRows.add(rowString);
        uniqueRows.push(row);
      } else {
        duplicatesRemoved++;
        console.log(`Duplicate removed at index ${index}:`, row);
      }
    });
    
    console.log(`Duplicates removed: ${duplicatesRemoved}`);
    console.log(`Unique rows: ${uniqueRows.length}`);
    
    return { uniqueRows, duplicatesRemoved };
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
    console.log('=== processFiles called ===');
    const selectedFiles = excelFiles.filter(file => file.selected);
    console.log('Selected files:', selectedFiles.length);
    
    if (selectedFiles.length === 0) {
      showMessage(t.messages.selectAtLeastOneFile, 'error');
      return;
    }

    try {
      console.log('Filter duplicates enabled:', filterDuplicates);
      showMessage(t.messages.loadingAndProcessing, 'info');
      
      // Load the actual Excel files and read their data
      const loadedFiles = await loadSelectedFiles(selectedFiles);
      
      // Validate column consistency
      const validation = validateColumnConsistency(loadedFiles);
      if (!validation.isValid) {
        // Show translated error message
        showMessage(t.messages.headerMismatchError, 'error');
        
        // Uncheck all selected files instead of refreshing the page
        const updatedFiles = excelFiles.map(file => ({ ...file, selected: false }));
        setExcelFiles(updatedFiles);
        setProcessedWorkbookData(null);
        setDuplicatesWorkbookData(null);
        
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
              
              // Use headers from first file/sheet only
              if (fileIndex === 0 && headers.length === 0) {
                headers = jsonData[0] || [];
                // Remove empty columns from the end of headers
                let lastNonEmptyIndex = headers.length - 1;
                while (lastNonEmptyIndex >= 0 && (!headers[lastNonEmptyIndex] || headers[lastNonEmptyIndex] === '')) {
                  lastNonEmptyIndex--;
                }
                headers = headers.slice(0, lastNonEmptyIndex + 1);
                console.log('Headers detected:', headers);
              }
              
              // Add all data rows (skip header) and filter out completely empty rows
              const dataRows = jsonData.slice(1)
                .filter(row => {
                  // Check if row has any non-empty cell
                  return row.some(cell => cell !== null && cell !== undefined && cell !== '');
                })
                .map(row => {
                  // Trim row to match header length (remove extra empty columns)
                  return headers.length > 0 ? row.slice(0, headers.length) : row;
                });
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
      console.log('Headers:', headers);
      
      // Verify headers were captured
      if (headers.length === 0) {
        console.error('ERROR: No headers were captured!');
        showMessage('Error: Could not detect headers in the files. Please ensure the files contain valid data.', 'error');
        return;
      }
      
      // Step 2: Remove duplicates across all files (if enabled)
      let uniqueRows: any[][];
      let duplicatesRemoved = 0;
      
      if (filterDuplicates) {
        console.log('Filtering duplicates enabled');
        const result = removeDuplicates(allRows);
        uniqueRows = result.uniqueRows;
        duplicatesRemoved = result.duplicatesRemoved;
      } else {
        console.log('Duplicate filtering is disabled. All rows kept.');
        uniqueRows = allRows;
      }
      
      console.log(`Final uniqueRows length: ${uniqueRows.length}`);
      
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
      
      console.log('Merged workbook created successfully');
      console.log('Final data rows:', uniqueRows.length);
      console.log('Final headers:', headers.length);

      console.log('Setting processedWorkbookData state...');
      setProcessedWorkbookData(mergedWorkbook);
      console.log('✓ STATE SET: processedWorkbookData is now truthy. Download button should appear below.');
      
      // Show success message with duplicate removal results
      let successMessage = `✅ ${t.messages.filesProcessedSuccessfully}
        📊 ${t.messages.totalRowsProcessed} ${totalRowsCollected}`;
      
      if (filterDuplicates) {
        successMessage += `
        🚫 ${t.messages.duplicatesRemoved} ${duplicatesRemoved}
        ✅ ${t.messages.uniqueRowsInResult} ${uniqueRows.length}`;
      } else {
        successMessage += `
        ✅ ${t.messages.allRowsKept} ${uniqueRows.length} ${t.messages.rowsKeptInResult}`;
      }
      
      successMessage += `
        📁 ${t.messages.readyToDownloadMergedFile}`;
      
      console.log('Showing success message:', successMessage);
      showMessage(successMessage, 'success');
      
    } catch (error) {
      console.error('Processing error:', error);
      showMessage(formatMessage(t.messages.processingFailed, { error: (error as Error).message }), 'error');
    }
  };

  const findDuplicatesInFile = async () => {
    const selectedFiles = excelFiles.filter(file => file.selected);
    
    if (selectedFiles.length === 0) {
      showMessage(t.messages.selectAtLeastOneFile, 'error');
      return;
    }
    
    if (selectedFiles.length > 1) {
      showMessage(t.messages.selectExactlyOneFile, 'error');
      return;
    }

    try {
      showMessage(t.messages.loadingAndProcessing, 'info');
      
      // Load the selected file
      const loadedFiles = await loadSelectedFiles(selectedFiles);
      const file = loadedFiles[0];
      
      // Create workbook for duplicates
      const duplicatesWorkbook = {
        SheetNames: [] as string[],
        Sheets: {} as { [key: string]: any }
      };

      const duplicatesSheetName = "duplicates";
      duplicatesWorkbook.SheetNames.push(duplicatesSheetName);
      
      // Collect all rows from the single file with row tracking
      const allRowsWithInfo: { row: any[], originalRowNumber: number, sheetName: string }[] = [];
      let headers: any[] = [];
      
      console.log('=== FINDING DUPLICATES IN SINGLE FILE ===');
      console.log(`Processing file: ${file.name}`);
      
      const workbook = file.workbook;
      let globalRowCounter = 2; // Start from 2 (after header row)
      
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
            
            // Use headers from first sheet
            if (headers.length === 0) {
              headers = jsonData[0] || [];
              console.log('Headers detected:', headers);
            }
            
            // Add all data rows (skip header) with row tracking
            const dataRows = jsonData.slice(1);
            dataRows.forEach((row, localIndex) => {
              allRowsWithInfo.push({
                row: row,
                originalRowNumber: globalRowCounter + localIndex,
                sheetName: sheetName
              });
            });
            
            globalRowCounter += dataRows.length;
            console.log(`Added ${dataRows.length} rows from sheet ${sheetName}`);
          } catch (error) {
            console.warn(`Error processing sheet ${sheetName}:`, error);
          }
        }
      });
      
      console.log(`Total rows to check for duplicates: ${allRowsWithInfo.length}`);
      
      // Find duplicates within the file and group them
      const duplicateGroups: { 
        row: any[], 
        rowNumbers: number[], 
        sheetNames: string[] 
      }[] = [];
      const seenRows = new Map<string, { row: any[], rowNumbers: number[], sheetNames: string[] }>();
      
      console.log('=== FINDING DUPLICATES ===');
      
      allRowsWithInfo.forEach((rowInfo, index) => {
        // Convert row to string for comparison
        const rowString = rowInfo.row.map(cell => 
          cell === null || cell === undefined ? '' : String(cell).trim()
        ).join('|');
        
        if (seenRows.has(rowString)) {
          // This is a duplicate - add row number to existing group
          const existingGroup = seenRows.get(rowString)!;
          existingGroup.rowNumbers.push(rowInfo.originalRowNumber);
          existingGroup.sheetNames.push(rowInfo.sheetName);
          console.log(`Duplicate found at row ${rowInfo.originalRowNumber} in sheet ${rowInfo.sheetName}:`, rowInfo.row);
        } else {
          // First occurrence - create new group
          seenRows.set(rowString, {
            row: rowInfo.row,
            rowNumbers: [rowInfo.originalRowNumber],
            sheetNames: [rowInfo.sheetName]
          });
        }
      });
      
      // Convert to array and filter only groups with duplicates (more than 1 occurrence)
      seenRows.forEach((group, rowString) => {
        if (group.rowNumbers.length > 1) {
          duplicateGroups.push(group);
        }
      });
      
      console.log(`Duplicate groups found: ${duplicateGroups.length}`);
      
      if (duplicateGroups.length === 0) {
        showMessage(t.messages.noDuplicatesFound, 'info');
        setDuplicatesWorkbookData(null);
        return;
      }
      
      // Create duplicates sheet with grouped duplicate information
      const duplicatesSheetData: any = {};
      
      // Add new header with row numbers column
      const newHeaders = ['Duplicate Row Numbers', 'Sheets', ...headers];
      const finalData = [newHeaders];
      
      // Add duplicate groups with all row numbers
      duplicateGroups.forEach(group => {
        const rowNumbersStr = group.rowNumbers.join(', ');
        const uniqueSheets = Array.from(new Set(group.sheetNames));
        const sheetsStr = uniqueSheets.join(', ');
        const rowWithInfo = [rowNumbersStr, sheetsStr, ...group.row];
        finalData.push(rowWithInfo);
      });
      
      // Convert final data back to Excel format
      finalData.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (cell !== undefined && cell !== null && cell !== '') {
            const cellAddress = window.XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            duplicatesSheetData[cellAddress] = { v: cell, t: typeof cell === 'number' ? 'n' : 's' };
          }
        });
      });
      
      // Set the range for the duplicates sheet
      const maxRow = finalData.length - 1;
      const maxCol = headers.length - 1;
      
      duplicatesSheetData['!ref'] = `A1:${window.XLSX.utils.encode_cell({ r: maxRow, c: maxCol })}`;
      
      duplicatesWorkbook.Sheets[duplicatesSheetName] = duplicatesSheetData;

      setDuplicatesWorkbookData(duplicatesWorkbook);
      
      // Show success message with duplicate results
      const totalDuplicateInstances = duplicateGroups.reduce((sum, group) => sum + group.rowNumbers.length, 0);
      const successMessage = `Found ${duplicateGroups.length} duplicate pattern(s) with ${totalDuplicateInstances} total instances`;
      showMessage(successMessage, 'success');
      
    } catch (error) {
      console.error('Duplicate finding error:', error);
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

  const saveDuplicatesFile = () => {
    if (!duplicatesWorkbookData) {
      showMessage(t.messages.pleaseProcessFilesFirst, 'error');
      return;
    }
    
    if (!selectedFolder) {
      showMessage(t.messages.pleaseSelectFolderFirst, 'error');
      return;
    }
    
    try {
      // Convert workbook to Excel file with optimized settings
      const wbout = window.XLSX.write(duplicatesWorkbookData, { 
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
      a.download = `duplicates_${timestamp}.xlsx`;
      document.body.appendChild(a);
      a.click();
      
      // Cleanup
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      // Add a small delay to ensure download has started
      setTimeout(() => {
        showMessage(t.messages.duplicatesFileSaved, 'success');
      }, 500);
      
    } catch (error) {
      console.error('Save duplicates error:', error);
      showMessage(formatMessage(t.messages.saveFailed, { error: (error as Error).message }), 'error');
    }
  };

  const chooseNewFolder = () => {
    // Reset the application state
    setSelectedFolder('');
    setExcelFiles([]);
    setProcessedWorkbookData(null);
    setDuplicatesWorkbookData(null);
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
            {/* Filter duplicates checkbox - temporarily hidden */}
            {/* <div className="filter-option">
              <label className="filter-checkbox">
                <input
                  type="checkbox"
                  checked={filterDuplicates}
                  onChange={(e) => setFilterDuplicates(e.target.checked)}
                />
                <span>{t.filterDuplicates}</span>
              </label>
            </div> */}
            <div className="button-group">
              <button 
                type="button" 
                className="btn btn-warning" 
                onClick={processFiles}
              >
                {t.mergeFiles}
              </button>
              <button 
                type="button" 
                className="btn btn-info" 
                onClick={findDuplicatesInFile}
                disabled={excelFiles.filter(file => file.selected).length !== 1}
              >
                {t.saveDuplicates}
              </button>
            </div>
            {excelFiles.filter(file => file.selected).length !== 1 && (
              <p className="button-hint">{t.findDuplicatesDescription}</p>
            )}
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

        {duplicatesWorkbookData && (
          <div className="download-section">
            <div className="section-header">
              <h3>{t.findDuplicates}</h3>
              <p>{t.findDuplicatesDescription}</p>
            </div>
            <button 
              type="button" 
              className="btn btn-success" 
              onClick={saveDuplicatesFile}
            >
              {t.saveDuplicates}
            </button>
          </div>
        )}

        {message && (
          <div className={`message-area ${messageType}`}>
            <div className="message-content">{message}</div>
            {messageType === 'error' && (
              <button 
                type="button" 
                className="btn btn-primary message-ok-button" 
                onClick={() => setMessage('')}
              >
                {t.ok}
              </button>
            )}
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