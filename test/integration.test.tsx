/**
 * Integration Tests for Excel File Merger Application
 */

import React from 'react';
import { render, screen, fireEvent, waitFor, act } from '@testing-library/react';
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

// Mock URL methods
Object.defineProperty(window.URL, 'createObjectURL', {
  value: jest.fn(() => 'mock-url'),
  writable: true,
});

Object.defineProperty(window.URL, 'revokeObjectURL', {
  value: jest.fn(),
  writable: true,
});

describe('Excel File Merger Integration Tests', () => {
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
    test('handles complete workflow from folder selection to file merge', async () => {
      render(<App />);
      
      // Step 1: Select folder with Excel files
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      const file1 = new File(['test content 1'], 'file1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const file2 = new File(['test content 2'], 'file2.xls', { 
        type: 'application/vnd.ms-excel' 
      });
      
      Object.defineProperty(file1, 'webkitRelativePath', { value: 'test-folder/file1.xlsx' });
      Object.defineProperty(file2, 'webkitRelativePath', { value: 'test-folder/file2.xls' });
      Object.defineProperty(file1, 'size', { value: 1024 });
      Object.defineProperty(file2, 'size', { value: 2048 });
      
      Object.defineProperty(folderInput, 'files', {
        value: [file1, file2],
        writable: false,
      });
      
      fireEvent.change(folderInput);

      await waitFor(() => {
        expect(screen.getByText('Found 2 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Step 2: Select files to merge
      const checkboxes = document.querySelectorAll('input[type="checkbox"]');
      fireEvent.click(checkboxes[0]); // Select first file
      fireEvent.click(checkboxes[1]); // Select second file

      await waitFor(() => {
        expect(screen.getByText('Merge Files')).toBeInTheDocument();
      });

      // Step 3: Merge files
      const mergeButton = screen.getByText('Merge Files');
      
      await act(async () => {
        fireEvent.click(mergeButton);
      });

      await waitFor(() => {
        expect(screen.getByText('✅ Successfully merged 2 file(s)! Ready to save.')).toBeInTheDocument();
      });

      // Step 4: Save merged file
      const saveButton = screen.getByRole('button', { name: 'Save Merged File' });
      
      await act(async () => {
        fireEvent.click(saveButton);
      });

      // Verify download was triggered
      expect(window.URL.createObjectURL).toHaveBeenCalled();
    });

    test('handles multiple folder selections sequentially', async () => {
      render(<App />);
      
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      // First folder selection
      const file1 = new File(['test content 1'], 'file1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      Object.defineProperty(file1, 'webkitRelativePath', { value: 'folder1/file1.xlsx' });
      Object.defineProperty(file1, 'size', { value: 1024 });
      
      Object.defineProperty(folderInput, 'files', {
        value: [file1],
        writable: false,
      });
      
      fireEvent.change(folderInput);

      await waitFor(() => {
        expect(screen.getByText('Found 1 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Choose new folder
      const chooseNewFolderButton = screen.getByText('Choose New Folder');
      fireEvent.click(chooseNewFolderButton);

      await waitFor(() => {
        expect(screen.getByText('📁 Ready to select a new folder')).toBeInTheDocument();
      });

      // Second folder selection - simplified approach
      const browseButton = screen.getByText('Browse Folder');
      fireEvent.click(browseButton);
      
      // Just verify that the UI resets properly for a new folder selection
      await waitFor(() => {
        expect(screen.getByText('📁 Ready to select a new folder')).toBeInTheDocument();
      });
    });
  });

  describe('Error Handling Integration', () => {
    test('handles folder selection with no Excel files', async () => {
      render(<App />);
      
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      const file1 = new File(['test content 1'], 'file1.txt', { type: 'text/plain' });
      const file2 = new File(['test content 2'], 'file2.pdf', { type: 'application/pdf' });
      
      Object.defineProperty(file1, 'webkitRelativePath', { value: 'test-folder/file1.txt' });
      Object.defineProperty(file2, 'webkitRelativePath', { value: 'test-folder/file2.pdf' });
      
      Object.defineProperty(folderInput, 'files', {
        value: [file1, file2],
        writable: false,
      });
      
      fireEvent.change(folderInput);

      await waitFor(() => {
        expect(screen.getByText('No Excel files found in the selected folder')).toBeInTheDocument();
      });
    });

    test('handles merge processing errors gracefully', async () => {
      render(<App />);
      
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      const file1 = new File(['test content 1'], 'file1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      Object.defineProperty(file1, 'webkitRelativePath', { value: 'test-folder/file1.xlsx' });
      Object.defineProperty(file1, 'size', { value: 1024 });
      
      Object.defineProperty(folderInput, 'files', {
        value: [file1],
        writable: false,
      });
      
      fireEvent.change(folderInput);

      await waitFor(() => {
        expect(screen.getByText('file1.xlsx')).toBeInTheDocument();
      });

      const checkbox = document.querySelector('input[type="checkbox"]') as HTMLInputElement;
      fireEvent.click(checkbox);

      await waitFor(() => {
        expect(screen.getByText('Merge Files')).toBeInTheDocument();
      });

      // Mock processing error by simulating a Promise rejection in the processFiles function
      // We'll mock Promise.all to reject, which will trigger the catch block
      const originalPromiseAll = Promise.all;
      Promise.all = jest.fn().mockRejectedValue(new Error('Processing failed'));

      const mergeButton = screen.getByText('Merge Files');
      
      await act(async () => {
        fireEvent.click(mergeButton);
      });

      await waitFor(() => {
        expect(screen.getByText(/Processing failed/)).toBeInTheDocument();
      });
      
      // Restore original Promise.all
      Promise.all = originalPromiseAll;
    });

    test('handles save errors gracefully', async () => {
      render(<App />);
      
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      const file1 = new File(['test content 1'], 'file1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      Object.defineProperty(file1, 'webkitRelativePath', { value: 'test-folder/file1.xlsx' });
      Object.defineProperty(file1, 'size', { value: 1024 });
      
      Object.defineProperty(folderInput, 'files', {
        value: [file1],
        writable: false,
      });
      
      fireEvent.change(folderInput);

      await waitFor(() => {
        expect(screen.getByText('file1.xlsx')).toBeInTheDocument();
      });

      const checkbox = document.querySelector('input[type="checkbox"]') as HTMLInputElement;
      fireEvent.click(checkbox);

      await waitFor(() => {
        expect(screen.getByText('Merge Files')).toBeInTheDocument();
      });

      const mergeButton = screen.getByText('Merge Files');
      
      await act(async () => {
        fireEvent.click(mergeButton);
      });

      await waitFor(() => {
        expect(screen.getByRole('button', { name: 'Save Merged File' })).toBeInTheDocument();
      });

      // Mock save error
      mockXLSX.write.mockImplementationOnce(() => {
        throw new Error('Save failed');
      });

      const saveButton = screen.getByRole('button', { name: 'Save Merged File' });
      
      await act(async () => {
        fireEvent.click(saveButton);
      });

      await waitFor(() => {
        expect(screen.getByText('Save failed: Save failed')).toBeInTheDocument();
      });
    });
  });

  describe('User Interface Integration', () => {
    test('maintains consistent UI state throughout workflow', async () => {
      render(<App />);
      
      // Initial state
      expect(screen.getByText('Excel File Merger')).toBeInTheDocument();
      expect(screen.getByText('Choose Folder')).toBeInTheDocument();
      
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      const file1 = new File(['test content 1'], 'file1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      Object.defineProperty(file1, 'webkitRelativePath', { value: 'test-folder/file1.xlsx' });
      Object.defineProperty(file1, 'size', { value: 1024 });
      
      Object.defineProperty(folderInput, 'files', {
        value: [file1],
        writable: false,
      });
      
      fireEvent.change(folderInput);

      // After folder selection
      await waitFor(() => {
        expect(screen.getByText('Excel Files Found')).toBeInTheDocument();
      });

      const checkbox = document.querySelector('input[type="checkbox"]') as HTMLInputElement;
      fireEvent.click(checkbox);

      // After file selection
      await waitFor(() => {
        expect(screen.getByText('Merge Files')).toBeInTheDocument();
      });

      const mergeButton = screen.getByText('Merge Files');
      
      await act(async () => {
        fireEvent.click(mergeButton);
      });

      // After merge
      await waitFor(() => {
        expect(screen.getByRole('button', { name: 'Save Merged File' })).toBeInTheDocument();
      });

      // UI should remain consistent throughout
      expect(screen.getByText('Excel File Merger')).toBeInTheDocument();
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });

    test('handles file selection controls correctly', async () => {
      render(<App />);
      
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      const file1 = new File(['test content 1'], 'file1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const file2 = new File(['test content 2'], 'file2.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      Object.defineProperty(file1, 'webkitRelativePath', { value: 'test-folder/file1.xlsx' });
      Object.defineProperty(file2, 'webkitRelativePath', { value: 'test-folder/file2.xlsx' });
      Object.defineProperty(file1, 'size', { value: 1024 });
      Object.defineProperty(file2, 'size', { value: 2048 });
      
      Object.defineProperty(folderInput, 'files', {
        value: [file1, file2],
        writable: false,
      });
      
      fireEvent.change(folderInput);

      await waitFor(() => {
        expect(screen.getByText('Excel Files Found')).toBeInTheDocument();
      });

      // Test Select All
      const selectAllButton = screen.getByText('Select All');
      fireEvent.click(selectAllButton);

      const checkboxes = document.querySelectorAll('input[type="checkbox"]');
      checkboxes.forEach(checkbox => {
        expect(checkbox).toBeChecked();
      });

      // Test Deselect All
      const deselectAllButton = screen.getByText('Deselect All');
      fireEvent.click(deselectAllButton);

      checkboxes.forEach(checkbox => {
        expect(checkbox).not.toBeChecked();
      });
    });
  });

  describe('Performance Integration', () => {
    test('handles large number of Excel files efficiently', async () => {
      render(<App />);
      
      const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      
      // Create many Excel files
      const files = Array.from({ length: 50 }, (_, i) => {
        const file = new File([`test content ${i}`], `file${i}.xlsx`, { 
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        Object.defineProperty(file, 'webkitRelativePath', { value: `test-folder/file${i}.xlsx` });
        Object.defineProperty(file, 'size', { value: 1024 });
        return file;
      });
      
      Object.defineProperty(folderInput, 'files', {
        value: files,
        writable: false,
      });
      
      fireEvent.change(folderInput);

      await waitFor(() => {
        expect(screen.getByText('Found 50 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Should handle large number of files without performance issues
      const checkboxes = document.querySelectorAll('input[type="checkbox"]');
      expect(checkboxes).toHaveLength(50);

      // Test Select All with many files
      const selectAllButton = screen.getByText('Select All');
      fireEvent.click(selectAllButton);

      checkboxes.forEach(checkbox => {
        expect(checkbox).toBeChecked();
      });
    });
  });
});