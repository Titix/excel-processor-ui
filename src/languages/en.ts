export const en = {
  // App title and header
  appTitle: 'Excel File Processor',
  appSubtitle: 'Select a folder',
  
  // Folder selection
  selectFolder: 'Select Folder',
  selectFolderDescription: 'Opens a folder selection dialog to scan for Excel files',
  
  // Folder info
  excelFilesFound: '{count} Excel file(s) found',
  chooseNewFolder: 'Choose New Folder',
  
  // File section
  excelFilesFoundTitle: 'Excel Files Found',
  selectFilesToMerge: 'Select the files you want to merge',
  selectAll: 'Select All',
  deselectAll: 'Deselect All',
  
  // No files message
  noExcelFilesFound: 'No Excel files (.xlsx, .xls) found in this folder.',
  noExcelFilesSuggestion: 'Please select a different folder or add Excel files to this folder.',
  
  // Process section
  mergeSelectedFiles: 'Merge Selected Files',
  mergeFilesDescription: 'Merges selected Excel files into a single file with all rows from all files',
  mergeFiles: 'Merge Files',
  filterAndMerge: 'Filter & Merge',
  filterAndMergeDescription: 'Filters and merges files based on predefined column rules and adds calculated Month and Week columns',
  filterDuplicates: 'Filter duplicates',
  retailWeekly: 'Retail weekly',
  
  // Duplicate section
  findDuplicates: 'Find Duplicates',
  findDuplicatesDescription: 'Identifies and extracts duplicate rows from a single selected Excel file',
  saveDuplicates: 'Save Duplicates',
  
  // Save section
  saveMergedFile: 'Save Processed File',
  saveMergedFileDescription: 'Downloads the processed Excel file to your computer',
  
  // Messages
  messages: {
    noExcelFilesInFolder: 'No Excel files found in the selected folder',
    foundExcelFiles: 'Found {count} Excel file(s) in the folder',
    browserNotSupported: 'Browser does not support this feature or user cancelled selection',
    selectAtLeastOneFile: 'Please select at least one Excel file to process',
    loadingAndProcessing: 'Loading and processing selected files...',
    successfullyMerged: '✅ Successfully merged {count} file(s)! Ready to save.',
    processingFailed: 'Processing failed: {error}',
    pleaseProcessFilesFirst: 'Please process files first',
    pleaseSelectFolderFirst: 'Please select a folder first',
    mergedFileSaved: '📥 Merged file saved successfully! Saved to: {folder}',
    saveFailed: 'Save failed: {error}',
    readyToSelectNewFolder: '📁 Ready to select a new folder',
    failedToLoadFile: 'Failed to load file: {fileName}',
    selectExactlyOneFile: 'Please select exactly one Excel file to find duplicates',
    duplicatesFound: 'Found {count} duplicate row(s) in the file',
    noDuplicatesFound: 'No duplicate rows found in the selected file',
    duplicatesFileSaved: '📥 Duplicates file saved successfully!',
    headerMismatchError: '❌ HEADER MISMATCH DETECTED!\n\n⚠️ The selected files have different column structures and cannot be merged.\n\nAll file selections have been cleared. Please select files with identical column headers.',
    filesProcessedSuccessfully: 'Files processed successfully!',
    totalRowsProcessed: 'Total rows processed:',
    duplicatesRemoved: 'Duplicates removed:',
    uniqueRowsInResult: 'Unique rows in result:',
    allRowsKept: 'All',
    rowsKeptInResult: 'rows kept in result',
    readyToDownloadMergedFile: 'Ready to download merged file.',
    bothHOLEmptyError: '❌ VALUE ERROR!\n\nBoth HOL_1 and HOL_2 are set to "Üres" in the following rows: {rowNumbers}\n\nFile processing stopped. Please fix the data in the Excel files!',
    checkboxSelected: 'Checkbox is selected.'
  },
  
  // File size units
  fileSizeUnits: {
    bytes: 'Bytes',
    kb: 'KB',
    mb: 'MB',
    gb: 'GB'
  },
  
  // Footer
  footer: {
    copyright: '© 2025 Excel Processor. Built with React and Node.js.',
    version: 'Version {version}'
  },
  
  // Language selector
  language: 'Language',
  selectLanguage: 'Select Language',
  
  // Common buttons
  ok: 'OK'
};
