export const hu = {
  // App title and header
  appTitle: 'Excel Fájl Feldolgozó',
  appSubtitle: 'Válassz egy mappát',
  
  // Folder selection
  selectFolder: 'Mappa Kiválasztása',
  selectFolderDescription: 'Válassz egy mappát az Excel fájlok kereséséhez',
  
  // Folder info
  excelFilesFound: 'Léteznek Excel fájlok',
  chooseNewFolder: 'Új Mappa Kiválasztása',
  
  // File section
  excelFilesFoundTitle: 'Excel Fájlok Listája',
  selectFilesToMerge: 'Válaszd ki az összevonni kívánt fájlokat',
  selectAll: 'Összes Kijelölése',
  deselectAll: 'Összes Kijelölés törlése',
  
  // No files message
  noExcelFilesFound: 'Nem található Excel fájl (.xlsx, .xls) ebben a mappában.',
  noExcelFilesSuggestion: 'Kérjük, válassz másik mappát vagy adj hozzá Excel fájlokat ehhez a mappához.',
  
  // Process section
  mergeSelectedFiles: 'Kiválasztott Fájlok Összevonása',
  mergeFilesDescription: 'Kattints a gombra a kiválasztott Excel fájlok összevonásához',
  mergeFiles: 'Fájlok Összevonása',
  
  // Save section
  saveMergedFile: 'Összevont Fájl Mentése',
  saveMergedFileDescription: 'A fájlok sikeresen össze lettek vonva',
  
  // Messages
  messages: {
    noExcelFilesInFolder: 'Nem található Excel fájl a kiválasztott mappában',
    foundExcelFiles: '{count} Excel fájl található a mappában',
    browserNotSupported: 'A böngésző nem támogatja ezt a funkciót vagy a felhasználó megszakította a kiválasztást',
    selectAtLeastOneFile: 'Kérjük, válassz ki legalább egy Excel fájlt a feldolgozáshoz',
    loadingAndProcessing: 'Kiválasztott fájlok betöltése és feldolgozása...',
    successfullyMerged: '✅ {count} fájl sikeresen összevonva! Készen áll a mentésre.',
    processingFailed: 'Feldolgozás sikertelen: {error}',
    pleaseProcessFilesFirst: 'Kérjük, először dolgozd fel a fájlokat',
    pleaseSelectFolderFirst: 'Kérjük, először válassz egy mappát',
    mergedFileSaved: '📥 Összevont fájl sikeresen mentve! Mentve ide: {folder}',
    saveFailed: 'Mentés sikertelen: {error}',
    readyToSelectNewFolder: '📁 Készen áll az új mappa kiválasztására',
    failedToLoadFile: 'Fájl betöltése sikertelen: {fileName}'
  },
  
  // File size units
  fileSizeUnits: {
    bytes: 'Bájt',
    kb: 'KB',
    mb: 'MB',
    gb: 'GB'
  },
  
  // Footer
  footer: {
    copyright: '© 2025 Excel Feldolgozó. React és Node.js segítségével készítve.',
    version: 'Verzió {version}'
  },
  
  // Language selector
  language: 'Nyelv',
  selectLanguage: 'Nyelv Kiválasztása'
};
