/**
 * Utility functions for Excel Processor
 */

export const formatFileSize = (bytes: number): string => {
  if (bytes === 0) return '0 Bytes';
  if (bytes < 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  if (i >= sizes.length) return '0 Bytes';
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

export const isValidExcelFile = (fileName: string): boolean => {
  return /\.(xlsx?)$/i.test(fileName);
};

export const isValidFileSize = (fileSize: number, maxSizeMB: number = 50): boolean => {
  return fileSize > 0 && fileSize <= maxSizeMB * 1024 * 1024;
};

export const getSheetNames = (workbook: any): string[] => {
  if (!workbook || !workbook.Sheets) return [];
  return Object.keys(workbook.Sheets);
};

export const getSheetData = (workbook: any, sheetName: string): any => {
  if (!workbook || !workbook.Sheets) return undefined;
  return workbook.Sheets[sheetName];
};

export const processSheet = (sheetData: any): any => {
  if (!sheetData) return {};
  
  const processedSheet = { ...sheetData };
  Object.keys(processedSheet).forEach(key => {
    if (processedSheet[key] && processedSheet[key].v && typeof processedSheet[key].v === 'string') {
      processedSheet[key].v = processedSheet[key].v.toUpperCase();
    }
  });
  return processedSheet;
};

export const createDownloadLink = (data: ArrayBuffer, fileName: string, mimeType: string): HTMLAnchorElement => {
  const blob = new Blob([data], { type: mimeType });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  return a;
};

export const getMimeType = (fileExtension: string): string => {
  switch (fileExtension.toLowerCase()) {
    case 'xls':
      return 'application/vnd.ms-excel';
    case 'xlsx':
      return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    default:
      return 'application/octet-stream';
  }
};

export const generateProcessedFileName = (originalFileName: string): string => {
  return originalFileName.replace(/\.(xlsx?)$/i, '_processed.$1');
};

export const messageTypes = ['success', 'error', 'info'] as const;
export type MessageType = typeof messageTypes[number];

export const isValidMessageType = (type: string): type is MessageType => {
  return messageTypes.includes(type as MessageType);
};

export const getMessageIcon = (type: MessageType): string => {
  switch (type) {
    case 'success':
      return '✅';
    case 'error':
      return '❌';
    case 'info':
      return 'ℹ️';
    default:
      return 'ℹ️';
  }
};

export const formatMessage = (message: string, type: MessageType): string => {
  const icon = getMessageIcon(type);
  return `${icon} ${message}`;
};

