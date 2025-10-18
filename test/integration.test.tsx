/**
 * Integration Tests for Excel Processor Application
 */

import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import App from '../src/frontend/App';

// Mock CSS imports
jest.mock('../src/frontend/App.css', () => ({}));

// Mock XLSX library with more comprehensive functionality
const mockXLSX = {
  read: jest.fn(),
  write: jest.fn(),
  utils: {
    sheet_to_json: jest.fn(),
    json_to_sheet: jest.fn(),
    book_new: jest.fn(),
    book_append_sheet: jest.fn(),
  }
};

Object.defineProperty(window, 'XLSX', {
  value: mockXLSX,
  writable: true,
});

// Mock FileReader with comprehensive functionality
const createMockFileReader = () => {
  const mockFileReader = {
    readAsArrayBuffer: jest.fn(),
    onload: null,
    onerror: null,
    result: new ArrayBuffer(8),
    readyState: 0,
    error: null,
  };

  // Simulate async file reading
  mockFileReader.readAsArrayBuffer.mockImplementation(function() {
    setTimeout(() => {
      if (this.onload) {
        this.onload({ target: { result: new ArrayBuffer(8) } });
      }
    }, 100);
  });

  return mockFileReader;
};

Object.defineProperty(window, 'FileReader', {
  value: jest.fn(() => createMockFileReader()),
  writable: true,
});

// Mock URL methods
Object.defineProperty(window.URL, 'createObjectURL', {
  value: jest.fn(() => 'mock-url'),
  writable: true,
});

Object.defineProperty(window.URL, 'revokeObjectURL', {
  value: jest.fn(),
  writable: true,
});

describe('Excel Processor Integration Tests', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    
    // Setup default mock responses
    mockXLSX.read.mockReturnValue({
      Sheets: {
        'Sheet1': {
          'A1': { v: 'Name' },
          'B1': { v: 'Age' },
          'A2': { v: 'John' },
          'B2': { v: 25 },
          'A3': { v: 'Jane' },
          'B3': { v: 30 }
        }
      }
    });
    
    mockXLSX.write.mockReturnValue(new ArrayBuffer(8));
    mockXLSX.utils.sheet_to_json.mockReturnValue([
      { Name: 'John', Age: 25 },
      { Name: 'Jane', Age: 30 }
    ]);
    mockXLSX.utils.json_to_sheet.mockReturnValue({});
    mockXLSX.utils.book_new.mockReturnValue({});
    mockXLSX.utils.book_append_sheet.mockReturnValue({});
  });

  describe('Complete File Processing Workflow', () => {
    test('handles complete workflow from file upload to download', async () => {
      render(<App />);
      
      // Step 1: Upload file
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      const file = new File(['test content'], 'test.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file, 'size', { value: 1024 });
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      await waitFor(() => {
        expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
      });

      // Step 2: Process file
      const processButton = screen.getByText('Process File');
      fireEvent.click(processButton);

      await waitFor(() => {
        expect(screen.getByText('✅ File processed successfully! Ready to download.')).toBeInTheDocument();
      });

      // Step 3: Download file
      const downloadButton = screen.getByText('Save Processed File');
      fireEvent.click(downloadButton);

      // Verify download was triggered
      expect(window.URL.createObjectURL).toHaveBeenCalled();
    });

    test('handles multiple file uploads sequentially', async () => {
      render(<App />);
      
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      
      // First file
      const file1 = new File(['test content 1'], 'test1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file1, 'size', { value: 1024 });
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file1] }
      });

      await waitFor(() => {
        expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
      });

      // Process first file
      const processButton = screen.getByText('Process File');
      fireEvent.click(processButton);

      await waitFor(() => {
        expect(screen.getByText('✅ File processed successfully! Ready to download.')).toBeInTheDocument();
      });

      // Choose new file
      const chooseNewFileButton = screen.getByText('Choose New File');
      fireEvent.click(chooseNewFileButton);

      await waitFor(() => {
        expect(screen.getByText('📁 Ready to select a new file')).toBeInTheDocument();
      });

      // Second file
      const file2 = new File(['test content 2'], 'test2.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file2, 'size', { value: 2048 });
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file2] }
      });

      await waitFor(() => {
        expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
      });
    });
  });

  describe('Error Handling Integration', () => {
    test('handles file reading errors gracefully', async () => {
      render(<App />);
      
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      const file = new File(['test content'], 'test.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file, 'size', { value: 1024 });
      
      // Mock file reading error
      const mockFileReader = createMockFileReader();
      mockFileReader.onerror = jest.fn();
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      // Simulate error
      if (mockFileReader.onerror) {
        mockFileReader.onerror({ target: { error: new Error('File read error') } });
      }

      await waitFor(() => {
        expect(screen.getByText(/Error reading Excel file/)).toBeInTheDocument();
      });
    });

    test('handles Excel processing errors gracefully', async () => {
      render(<App />);
      
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      const file = new File(['test content'], 'test.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file, 'size', { value: 1024 });
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      await waitFor(() => {
        expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
      });

      // Mock processing error
      mockXLSX.write.mockImplementation(() => {
        throw new Error('Processing failed');
      });

      const processButton = screen.getByText('Process File');
      fireEvent.click(processButton);

      await waitFor(() => {
        expect(screen.getByText(/Processing failed/)).toBeInTheDocument();
      });
    });

    test('handles invalid file types', async () => {
      render(<App />);
      
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      const file = new File(['test content'], 'test.txt', { type: 'text/plain' });
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      await waitFor(() => {
        expect(screen.getByText('Please select an Excel file (.xls or .xlsx)')).toBeInTheDocument();
      });
    });

    test('handles oversized files', async () => {
      render(<App />);
      
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      const file = new File(['test content'], 'test.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file, 'size', { value: 51 * 1024 * 1024 }); // 51MB
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      await waitFor(() => {
        expect(screen.getByText('File size must be less than 50MB')).toBeInTheDocument();
      });
    });
  });

  describe('User Interface Integration', () => {
    test('maintains consistent UI state throughout workflow', async () => {
      render(<App />);
      
      // Initial state
      expect(screen.getByText('Excel Processor')).toBeInTheDocument();
      expect(screen.getByText('Choose Excel File')).toBeInTheDocument();
      expect(screen.getByText('Version 1.0.0')).toBeInTheDocument();
      
      // After file upload
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      const file = new File(['test content'], 'test.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file, 'size', { value: 1024 });
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      await waitFor(() => {
        expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
      });

      // UI elements should still be present
      expect(screen.getByText('Excel Processor')).toBeInTheDocument();
      expect(screen.getByText('Version 1.0.0')).toBeInTheDocument();
    });

    test('handles drag and drop visual feedback correctly', async () => {
      render(<App />);
      
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      
      // Drag over
      fireEvent.dragOver(uploadArea!);
      expect(uploadArea).toHaveClass('dragover');
      
      // Drag leave
      fireEvent.dragLeave(uploadArea!);
      expect(uploadArea).not.toHaveClass('dragover');
      
      // Drag over again
      fireEvent.dragOver(uploadArea!);
      expect(uploadArea).toHaveClass('dragover');
      
      // Drop
      const file = new File(['test content'], 'test.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file, 'size', { value: 1024 });
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      await waitFor(() => {
        expect(uploadArea).not.toHaveClass('dragover');
      });
    });
  });

  describe('Performance Integration', () => {
    test('handles large Excel files efficiently', async () => {
      render(<App />);
      
      // Mock large workbook
      const largeWorkbook = {
        Sheets: {
          'Sheet1': {}
        }
      };
      
      // Create a large sheet with many cells
      for (let i = 1; i <= 1000; i++) {
        largeWorkbook.Sheets['Sheet1'][`A${i}`] = { v: `Value ${i}` };
        largeWorkbook.Sheets['Sheet1'][`B${i}`] = { v: i };
      }
      
      mockXLSX.read.mockReturnValue(largeWorkbook);
      
      const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
      const file = new File(['test content'], 'large-test.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      Object.defineProperty(file, 'size', { value: 1024 });
      
      const startTime = Date.now();
      
      fireEvent.drop(uploadArea!, {
        dataTransfer: { files: [file] }
      });

      await waitFor(() => {
        expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
      });

      const processButton = screen.getByText('Process File');
      fireEvent.click(processButton);

      await waitFor(() => {
        expect(screen.getByText('✅ File processed successfully! Ready to download.')).toBeInTheDocument();
      });

      const endTime = Date.now();
      const processingTime = endTime - startTime;
      
      // Processing should complete within reasonable time (5 seconds)
      expect(processingTime).toBeLessThan(5000);
    });
  });
});
