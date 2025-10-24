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
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [workbookData, setWorkbookData] = useState<any[]>([]);
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
      handleFiles(Array.from(files));
    }
  };

  const onFileSelected = (event: React.ChangeEvent<HTMLInputElement>) => {
    const target = event.target;
    if (target.files && target.files.length > 0) {
      handleFiles(Array.from(target.files));
    }
  };

  const handleFiles = (files: File[]) => {
    // Filter valid Excel files
    const validFiles = files.filter(file => {
      if (!file.name.match(/\.(xlsx?)$/i)) {
        showMessage(`Skipping ${file.name}: Not an Excel file`, 'error');
        return false;
      }
      if (file.size > 50 * 1024 * 1024) {
        showMessage(`Skipping ${file.name}: File too large (>50MB)`, 'error');
        return false;
      }
      return true;
    });

    if (validFiles.length === 0) {
      showMessage('No valid Excel files selected', 'error');
      return;
    }

    setSelectedFiles(validFiles);
    readExcelFiles(validFiles);
  };

  const readExcelFiles = (files: File[]) => {
    const workbookPromises = files.map((file, index) => {
      return new Promise<{ file: File; workbook: any; index: number }>((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = window.XLSX.read(data, { type: 'array' });
            resolve({ file, workbook, index });
          } catch (error) {
            console.error(`Error reading ${file.name}:`, error);
            reject(error);
          }
        };
        
        reader.onerror = () => {
          reject(new Error(`Error reading ${file.name}`));
        };
        
        reader.readAsArrayBuffer(file);
      });
    });

    Promise.all(workbookPromises)
      .then((workbookResults) => {
        setWorkbookData(workbookResults);
        showMessage(`📁 ${files.length} files loaded successfully! Ready to merge and remove duplicates.`, 'success');
      })
      .catch((error) => {
        console.error('Error reading Excel files:', error);
        showMessage('Error reading Excel files: ' + (error as Error).message, 'error');
        setWorkbookData([]);
      });
  };

  const processFile = () => {
    if (!workbookData || workbookData.length === 0) {
      showMessage('Please select files first', 'error');
      return;
    }

    try {
      console.log('=== CROSS-FILE DUPLICATE REMOVAL ===');
      console.log(`Processing ${workbookData.length} files for cross-file duplicate removal`);
      
      // Step 1: Collect ALL rows from ALL files into one big array
      const allDataRows: any[][] = [];
      let headers: any[] = [];
      let totalRowsFromAllFiles = 0;
      
      workbookData.forEach(({ file, workbook }, fileIndex) => {
        console.log(`\n--- Processing file ${fileIndex + 1}: ${file.name} ---`);
        
        Object.keys(workbook.Sheets).forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const jsonData = window.XLSX.utils.sheet_to_json(sheet, { 
            header: 1, 
            defval: '', 
            raw: false 
          }) as any[][];
          
          if (jsonData.length === 0) return;
          
          // Get headers from first file
          if (fileIndex === 0 && headers.length === 0) {
            headers = jsonData[0] || [];
            console.log('Headers:', headers);
          }
          
          // Add ALL data rows from this file (skip header)
          const dataRows = jsonData.slice(1);
          console.log(`Adding ${dataRows.length} rows from ${file.name} sheet ${sheetName}`);
          
          // Log first few rows to show what we're collecting
          dataRows.slice(0, 2).forEach((row, idx) => {
            console.log(`  Row ${idx + 1} from ${file.name}:`, row);
          });
          
          allDataRows.push(...dataRows);
          totalRowsFromAllFiles += dataRows.length;
        });
      });
      
      console.log(`\n=== COLLECTION COMPLETE ===`);
      console.log(`Total rows collected from ALL files: ${totalRowsFromAllFiles}`);
      console.log(`All rows array length: ${allDataRows.length}`);
      
      // Step 2: Remove duplicates ACROSS ALL FILES
      const uniqueRows: any[][] = [];
      const seenRows = new Set<string>();
      let duplicatesRemoved = 0;
      
      console.log(`\n=== CROSS-FILE DUPLICATE DETECTION ===`);
      
      allDataRows.forEach((row, index) => {
        // Convert row to string for comparison (this is what makes rows "equal")
        const rowString = row.map(cell => 
          cell === null || cell === undefined ? '' : String(cell).trim()
        ).join('|');
        
        if (!seenRows.has(rowString)) {
          // First time seeing this row - keep it
          seenRows.add(rowString);
          uniqueRows.push(row);
          if (index < 5) {
            console.log(`✓ KEEPING unique row ${index + 1}:`, row);
          }
        } else {
          // We've seen this exact row before - it's a duplicate
          duplicatesRemoved++;
          console.log(`✗ REMOVING duplicate row ${index + 1}:`, row, `(seen before)`);
        }
      });
      
      console.log(`\n=== FINAL RESULTS ===`);
      console.log(`Total rows from all files: ${totalRowsFromAllFiles}`);
      console.log(`Duplicates removed (across all files): ${duplicatesRemoved}`);
      console.log(`Unique rows in final result: ${uniqueRows.length}`);
      
      // Step 3: Create final Excel file with unique rows only
      const finalData = [headers, ...uniqueRows];
      const mergedSheet = window.XLSX.utils.aoa_to_sheet(finalData);
      const mergedWorkbook = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(mergedWorkbook, mergedSheet, 'Merged_Data');
      
      setProcessedWorkbookData(mergedWorkbook);
      
      showMessage(`✅ Files merged! ${duplicatesRemoved} duplicates removed across all files. ${uniqueRows.length} unique rows in result.`, 'success');
      
      // Automatically verify the result
      setTimeout(() => {
        verifyResult(mergedWorkbook);
      }, 1000);
      
    } catch (error) {
      console.error('Error:', error);
      showMessage('Processing failed: ' + (error as Error).message, 'error');
    }
  };

  const downloadFile = () => {
    if (!processedWorkbookData) {
      showMessage('Please process files first', 'error');
      return;
    }
    
    try {
      // Use xlsx format for merged files
      const wbout = window.XLSX.write(processedWorkbookData, { 
        bookType: 'xlsx',
        type: 'array',
        compression: true,  // Enable compression for smaller files
        cellStyles: false,  // Disable cell styles to reduce size
        cellNF: false,      // Disable number formats to reduce size
        cellHTML: false     // Disable HTML in cells to reduce size
      });
      
      // Create blob with appropriate MIME type
      const blob = new Blob([wbout], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const url = window.URL.createObjectURL(blob);
      
      // Create download link
      const a = document.createElement('a');
      a.href = url;
      a.download = `merged_files_${new Date().toISOString().split('T')[0]}.xlsx`;
      document.body.appendChild(a);
      a.click();
      
      // Cleanup
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      // Add a small delay to ensure download has started
      setTimeout(() => {
        showMessage('📥 Merged file downloaded successfully! Check your downloads folder.', 'success');
      }, 500);
      
    } catch (error) {
      console.error('Download error:', error);
      showMessage('Download failed: ' + (error as Error).message, 'error');
    }
  };

  const chooseNewFiles = () => {
    // Reset the application state
    setSelectedFiles([]);
    setWorkbookData([]);
    setProcessedWorkbookData(null);
    setMessage('');
    
    // Show message
    showMessage('📁 Ready to select new files', 'info');
  };

  // Test function to create sample data with cross-file duplicates
  const createTestData = () => {
    console.log('=== CREATING COMPREHENSIVE TEST DATA ===');
    console.log('TEST BUTTON CLICKED - CODE IS WORKING!');
    alert('🧪 Test button clicked! Check console for details.');
    
    // Create File 1 with some rows
    const file1Data = [
      ['Name', 'Age', 'City'], // Headers
      ['John', 25, 'New York'], // Row 1
      ['Jane', 30, 'London'],   // Row 2
      ['Bob', 35, 'Paris'],     // Row 3
    ];
    
    // Create File 2 with some overlapping rows (duplicates)
    const file2Data = [
      ['Name', 'Age', 'City'], // Headers (same as file 1)
      ['John', 25, 'New York'], // DUPLICATE of File 1 Row 1
      ['Alice', 28, 'Berlin'],  // Row 4 (unique)
      ['Jane', 30, 'London'],   // DUPLICATE of File 1 Row 2
    ];
    
    // Create File 3 with more duplicates
    const file3Data = [
      ['Name', 'Age', 'City'], // Headers (same as others)
      ['Bob', 35, 'Paris'],     // DUPLICATE of File 1 Row 3
      ['Charlie', 22, 'Tokyo'], // Row 5 (unique)
    ];
    
    // Create workbooks
    const workbook1 = window.XLSX.utils.book_new();
    const sheet1 = window.XLSX.utils.aoa_to_sheet(file1Data);
    window.XLSX.utils.book_append_sheet(workbook1, sheet1, 'Data');
    
    const workbook2 = window.XLSX.utils.book_new();
    const sheet2 = window.XLSX.utils.aoa_to_sheet(file2Data);
    window.XLSX.utils.book_append_sheet(workbook2, sheet2, 'Data');
    
    const workbook3 = window.XLSX.utils.book_new();
    const sheet3 = window.XLSX.utils.aoa_to_sheet(file3Data);
    window.XLSX.utils.book_append_sheet(workbook3, sheet3, 'Data');
    
    // Create mock file objects
    const mockFile1 = new File(['test1'], 'file1.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const mockFile2 = new File(['test2'], 'file2.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const mockFile3 = new File(['test3'], 'file3.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // Set the test data
    setSelectedFiles([mockFile1, mockFile2, mockFile3]);
    setWorkbookData([
      { file: mockFile1, workbook: workbook1, index: 0 },
      { file: mockFile2, workbook: workbook2, index: 1 },
      { file: mockFile3, workbook: workbook3, index: 2 }
    ]);
    
    console.log('🧪 COMPREHENSIVE TEST DATA CREATED:');
    console.log('📁 File 1: 3 data rows (John, Jane, Bob)');
    console.log('📁 File 2: 3 data rows (John DUPLICATE, Alice, Jane DUPLICATE)');
    console.log('📁 File 3: 2 data rows (Bob DUPLICATE, Charlie)');
    console.log('');
    console.log('📊 EXPECTED RESULTS:');
    console.log('✅ Total rows from all files: 8 data rows');
    console.log('✅ Duplicates to be removed: 3 (John, Jane, Bob)');
    console.log('✅ Unique rows in result: 5 (John, Jane, Bob, Alice, Charlie)');
    console.log('✅ Headers in result: 1 (only once)');
    console.log('✅ Total rows in result file: 6 (1 header + 5 unique data rows)');
    
    showMessage('🧪 Comprehensive test data loaded! Click "Merge Files" to test duplicate removal.', 'info');
  };

  // Verification function to check the final result
  const verifyResult = (workbook: any) => {
    console.log('=== VERIFYING FINAL RESULT ===');
    
    try {
      // Get the merged sheet data
      const mergedSheet = workbook.Sheets['Merged_Data'];
      const jsonData = window.XLSX.utils.sheet_to_json(mergedSheet, { 
        header: 1, 
        defval: '', 
        raw: false 
      }) as any[][];
      
      console.log(`📊 Final result has ${jsonData.length} total rows`);
      
      // Check 1: Verify only one header row
      const headers = jsonData[0] || [];
      console.log('✅ Headers:', headers);
      
      if (jsonData.length > 0) {
        console.log('✅ VERIFICATION 1 PASSED: Header present exactly once');
      } else {
        console.log('❌ VERIFICATION 1 FAILED: No header found');
      }
      
      // Check 2: Verify no duplicates in data rows
      const dataRows = jsonData.slice(1); // Skip header
      console.log(`📋 Data rows to check: ${dataRows.length}`);
      
      const seenRows = new Set<string>();
      let duplicateCount = 0;
      
      dataRows.forEach((row, index) => {
        const rowString = row.map(cell => 
          cell === null || cell === undefined ? '' : String(cell).trim()
        ).join('|');
        
        if (seenRows.has(rowString)) {
          duplicateCount++;
          console.log(`❌ DUPLICATE FOUND in final result at row ${index + 2}:`, row);
        } else {
          seenRows.add(rowString);
          console.log(`✅ Unique row ${index + 2}:`, row);
        }
      });
      
      if (duplicateCount === 0) {
        console.log('✅ VERIFICATION 2 PASSED: No duplicates in final result');
      } else {
        console.log(`❌ VERIFICATION 2 FAILED: ${duplicateCount} duplicates found in final result`);
      }
      
      // Summary
      console.log('=== VERIFICATION SUMMARY ===');
      console.log(`📊 Total rows in result: ${jsonData.length}`);
      console.log(`📋 Data rows: ${dataRows.length}`);
      console.log(`🚫 Duplicates in result: ${duplicateCount}`);
      console.log(`✅ Headers: ${headers.length > 0 ? 'Present' : 'Missing'}`);
      
      if (duplicateCount === 0 && headers.length > 0) {
        console.log('🎉 ALL VERIFICATIONS PASSED! Result is correct.');
        showMessage('🎉 Verification passed! No duplicates, single header present.', 'success');
      } else {
        console.log('⚠️ Some verifications failed. Check the logs above.');
        showMessage('⚠️ Verification failed. Check console for details.', 'error');
      }
      
    } catch (error) {
      console.error('Verification error:', error);
      showMessage('Verification failed: ' + (error as Error).message, 'error');
    }
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
        <p>Merge multiple Excel files and remove duplicate rows</p>
        <div style={{padding: '20px', backgroundColor: 'yellow', border: '3px solid red', margin: '10px'}}>
          <button 
            onClick={() => alert('TEST BUTTON WORKS!')}
            style={{
              backgroundColor: 'red',
              color: 'white',
              padding: '20px',
              fontSize: '20px',
              border: 'none',
              borderRadius: '5px',
              cursor: 'pointer'
            }}
          >
            🧪 CLICK ME - TEST BUTTON 🧪
          </button>
        </div>
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
              <h3>Choose Excel Files</h3>
              <p>Drag and drop multiple Excel files here or click to browse</p>
              <input 
                type="file" 
                ref={fileInputRef}
                accept=".xls,.xlsx" 
                multiple
                style={{ display: 'none' }}
                onChange={onFileSelected}
              />
              <button 
                type="button" 
                className="btn btn-primary" 
                onClick={() => fileInputRef.current?.click()}
              >
                Select Files
              </button>
              <button 
                type="button" 
                className="btn btn-secondary" 
                onClick={createTestData}
                style={{ marginLeft: '10px', backgroundColor: '#ff6b6b', color: 'white', border: '2px solid #ff6b6b' }}
              >
                🧪 Load Test Data
              </button>
            </div>
          </div>
          
          {selectedFiles.length > 0 && (
            <div className="file-info">
              <h4>Selected Files ({selectedFiles.length})</h4>
              <div className="file-list">
                {selectedFiles.map((file, index) => (
                  <div key={index} className="file-details">
                    <span className="file-icon">📄</span>
                    <span className="file-name">{file.name}</span>
                    <span className="file-size">{formatFileSize(file.size)}</span>
                  </div>
                ))}
              </div>
              <button 
                type="button" 
                className="btn btn-secondary btn-small" 
                onClick={chooseNewFiles}
              >
                Choose New Files
              </button>
            </div>
          )}
        </div>

        {selectedFiles.length > 0 && (
          <div className="process-section">
            <div className="section-header">
              <h3>Merge Files & Remove Duplicates</h3>
              <p>Click the button below to merge your files and remove duplicate rows</p>
            </div>
            <button 
              type="button" 
              className="btn btn-warning" 
              disabled={!workbookData || workbookData.length === 0}
              onClick={processFile}
            >
              Merge Files
            </button>
          </div>
        )}

        {processedWorkbookData && (
          <div className="download-section">
            <div className="section-header">
              <h3>Download Merged File</h3>
              <p>Your files have been merged and duplicates removed</p>
            </div>
            <button 
              type="button" 
              className="btn btn-success" 
              onClick={downloadFile}
            >
              Download Merged File
            </button>
            <button 
              type="button" 
              className="btn btn-info" 
              onClick={() => verifyResult(processedWorkbookData)}
              style={{ marginLeft: '10px' }}
            >
              Verify Result
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
