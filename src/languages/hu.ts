export const hu = {
  // App title and header
  appTitle: 'Excel Fájl Feldolgozó',
  appSubtitle: 'Válassz egy mappát',
  
  // Folder selection
  selectFolder: 'Mappa Kiválasztása',
  selectFolderDescription: 'Válassz egy mappát az Excel fájlok kereséséhez',
  
  // Folder info
  excelFilesFound: '{count} Excel fájl található',
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
  filterAndMerge: 'Szűrés & Összevonás',
  filterAndMergeDescription: 'Fájlok összevonása meghatározott oszlopokkal',
  filterDuplicates: 'Duplikátumok szűrése',
  
  // Duplicate section
  findDuplicates: 'Duplikátumok Keresése',
  findDuplicatesDescription: 'Duplikátum sorok keresése és mentése egy Excel fájlból',
  saveDuplicates: 'Duplikátumok Mentése',
  
  // Save section
  saveMergedFile: 'Feldolgozott Fájl Mentése',
  saveMergedFileDescription: 'A fájlok feldolgozása sikeres volt',
  
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
    failedToLoadFile: 'Fájl betöltése sikertelen: {fileName}',
    selectExactlyOneFile: 'Kérjük, válassz ki pontosan egy Excel fájlt a duplikátumok kereséséhez',
    duplicatesFound: '{count} duplikátum sor található a fájlban',
    noDuplicatesFound: 'Nem található duplikátum sor a kiválasztott fájlban',
    duplicatesFileSaved: '📥 Duplikátumok fájl sikeresen mentve!',
    headerMismatchError: '❌ FEJLÉC ELTÉRÉS ÉSZLELVE!\n\n⚠️ A kiválasztott fájlok eltérő oszlopstruktúrával rendelkeznek és nem egyesíthetők.\n\nMinden fájl kijelölés törölve. Kérjük, válassz ki azonos fejlécű fájlokat.',
    filesProcessedSuccessfully: 'Fájlok sikeresen feldolgozva!',
    totalRowsProcessed: 'Összes feldolgozott sor:',
    duplicatesRemoved: 'Eltávolított duplikátumok:',
    uniqueRowsInResult: 'Egyedi sorok az eredményben:',
    allRowsKept: 'Összes',
    rowsKeptInResult: 'sor megtartva az eredményben',
    readyToDownloadMergedFile: 'Az összevont fájl letölthető.',
    bothHOLEmptyError: '❌ ÉRTÉKHIBA!\n\nHOL_1 és HOL_2 egyaránt "Üres"-re van állítva a következő sorokban: {rowNumbers}\n\nA fájlok feldolgozása megszakadt. Kérjük, javítsa ki az Excel fájlokban az adatokat!'
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
  selectLanguage: 'Nyelv Kiválasztása',
  
  // Common buttons
  ok: 'OK'
};
