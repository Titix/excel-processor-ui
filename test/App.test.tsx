import React from 'react';
import { render, screen, fireEvent, waitFor, act } from '@testing-library/react';
import '@testing-library/jest-dom';
import App from '../src/frontend/App';

// Mock CSS imports
jest.mock('../src/frontend/App.css', () => ({}));

// Mock XLSX library
const mockWorkbook = {
  Sheets: {
    'Sheet1': {
      'A1': { v: 'Name' },
      'B1': { v: 'Age' },
      'A2': { v: 'John' },
      'B2': { v: 25 }
    }
  },
  SheetNames: ['Sheet1']
};

const mockXLSX = {
  read: jest.fn(() => mockWorkbook),
  write: jest.fn(() => new ArrayBuffer(8)),
  utils: {
    sheet_to_json: jest.fn(() => [{ col1: 'data1' }]),
    json_to_sheet: jest.fn(),
    book_new: jest.fn(),
    book_append_sheet: jest.fn(),
  },
};

// Mock window.XLSX
Object.defineProperty(window, 'XLSX', {
  value: mockXLSX,
  writable: true,
});

// Mock FileReader
const mockFileReader = {
  readAsArrayBuffer: jest.fn(),
  onload: null,
  onerror: null,
  result: new ArrayBuffer(8),
};

// Mock global FileReader
Object.defineProperty(window, 'FileReader', {
  value: jest.fn(() => mockFileReader),
  writable: true,
});

// Mock URL.createObjectURL and URL.revokeObjectURL
Object.defineProperty(window.URL, 'createObjectURL', {
  value: jest.fn(() => 'mock-url'),
  writable: true,
});

Object.defineProperty(window.URL, 'revokeObjectURL', {
  value: jest.fn(),
  writable: true,
});

describe('App Component - Excel File Merger', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    mockXLSX.read.mockReturnValue(mockWorkbook);
    mockXLSX.write.mockReturnValue(new ArrayBuffer(8));
    mockFileReader.readAsArrayBuffer.mockClear();
    mockFileReader.onload = null;
    mockFileReader.onerror = null;
  });

  test('renders main heading', () => {
    render(<App />);
    expect(screen.getByText('Excel File Merger')).toBeInTheDocument();
  });

  test('renders folder selection area', () => {
    render(<App />);
    expect(screen.getByText('Choose Folder')).toBeInTheDocument();
    expect(screen.getByText('Select a folder containing Excel files (.xls, .xlsx)')).toBeInTheDocument();
  });

  test('renders browse folder button', () => {
    render(<App />);
    expect(screen.getByText('Browse Folder')).toBeInTheDocument();
  });

  test('renders version in footer', () => {
    render(<App />);
    expect(screen.getByText('Version 1.0.0')).toBeInTheDocument();
  });

  test('renders copyright in footer', () => {
    render(<App />);
    expect(screen.getByText(/© 2025 Excel Processor/)).toBeInTheDocument();
  });

  test('shows folder input when browse button is clicked', () => {
    render(<App />);
    const browseButton = screen.getByText('Browse Folder');
    fireEvent.click(browseButton);
    
    // The folder input should be triggered (though hidden)
    const folderInput = document.querySelector('input[type="file"]');
    expect(folderInput).toBeInTheDocument();
    expect(folderInput).toHaveAttribute('webkitdirectory');
    expect(folderInput).toHaveAttribute('multiple');
  });

  test('handles folder selection with Excel files', async () => {
    render(<App />);
    const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
    
    // Create mock files with webkitRelativePath
    const file1 = new File(['test content 1'], 'file1.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    const file2 = new File(['test content 2'], 'file2.xls', { 
      type: 'application/vnd.ms-excel' 
    });
    const file3 = new File(['test content 3'], 'file3.txt', { 
      type: 'text/plain' 
    });
    
    Object.defineProperty(file1, 'webkitRelativePath', { value: 'test-folder/file1.xlsx' });
    Object.defineProperty(file2, 'webkitRelativePath', { value: 'test-folder/file2.xls' });
    Object.defineProperty(file3, 'webkitRelativePath', { value: 'test-folder/file3.txt' });
    Object.defineProperty(file1, 'size', { value: 1024 });
    Object.defineProperty(file2, 'size', { value: 2048 });
    Object.defineProperty(file3, 'size', { value: 512 });
    
    Object.defineProperty(folderInput, 'files', {
      value: [file1, file2, file3],
      writable: false,
    });
    
    fireEvent.change(folderInput);

    await waitFor(() => {
      expect(screen.getByText('Found 2 Excel file(s) in the folder')).toBeInTheDocument();
    });

    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });

    await waitFor(() => {
      expect(screen.getByText('2 Excel file(s) found')).toBeInTheDocument();
    });
  });

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

  test('shows Excel files list with checkboxes', async () => {
    render(<App />);
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
      expect(screen.getByText('Excel Files Found')).toBeInTheDocument();
    });

    await waitFor(() => {
      expect(screen.getByText('file1.xlsx')).toBeInTheDocument();
    });

    await waitFor(() => {
      expect(screen.getByText('file2.xls')).toBeInTheDocument();
    });

    // Check for checkboxes
    const checkboxes = document.querySelectorAll('input[type="checkbox"]');
    expect(checkboxes).toHaveLength(2);
  });

  test('handles file selection toggle', async () => {
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
    expect(checkbox).not.toBeChecked();

    fireEvent.click(checkbox);
    expect(checkbox).toBeChecked();
  });

  test('handles select all files', async () => {
    render(<App />);
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
      expect(screen.getByText('Select All')).toBeInTheDocument();
    });

    const selectAllButton = screen.getByText('Select All');
    fireEvent.click(selectAllButton);

    const checkboxes = document.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(checkbox => {
      expect(checkbox).toBeChecked();
    });
  });

  test('handles deselect all files', async () => {
    render(<App />);
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
      expect(screen.getByText('Select All')).toBeInTheDocument();
    });

    // First select all
    const selectAllButton = screen.getByText('Select All');
    fireEvent.click(selectAllButton);

    // Then deselect all
    const deselectAllButton = screen.getByText('Deselect All');
    fireEvent.click(deselectAllButton);

    const checkboxes = document.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(checkbox => {
      expect(checkbox).not.toBeChecked();
    });
  });

  test('shows merge button when files are selected', async () => {
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
  });

  test('handles merge files process', async () => {
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
      expect(screen.getByText('✅ Successfully merged 1 file(s)! Ready to save.')).toBeInTheDocument();
    });
  });

  test('shows error when trying to merge without selecting files', async () => {
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

    // Don't select any files, just click merge button directly
    // This should not be possible since the button only appears when files are selected
    // But let's test the error handling in the processFiles function
    const checkbox = document.querySelector('input[type="checkbox"]') as HTMLInputElement;
    expect(checkbox).not.toBeChecked();

    // The merge button should not be visible
    expect(screen.queryByText('Merge Files')).not.toBeInTheDocument();
  });

  test('shows save button after merging files', async () => {
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
  });

  test('handles save merged file', async () => {
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

    const saveButton = screen.getByRole('button', { name: 'Save Merged File' });
    
    await act(async () => {
      fireEvent.click(saveButton);
    });

    // Verify download was triggered
    expect(window.URL.createObjectURL).toHaveBeenCalled();
  });

  test('handles choose new folder button click', async () => {
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
      expect(screen.getByText('Choose New Folder')).toBeInTheDocument();
    });

    const chooseNewFolderButton = screen.getByText('Choose New Folder');
    fireEvent.click(chooseNewFolderButton);

    await waitFor(() => {
      expect(screen.getByText('📁 Ready to select a new folder')).toBeInTheDocument();
    });
  });

  test('handles empty folder selection', async () => {
    render(<App />);
    const folderInput = document.querySelector('input[type="file"]') as HTMLInputElement;
    
    Object.defineProperty(folderInput, 'files', {
      value: [],
      writable: false,
    });
    
    fireEvent.change(folderInput);

    // Should not show any error or success message for empty selection
    expect(screen.queryByText(/Found.*Excel file/)).not.toBeInTheDocument();
    expect(screen.queryByText(/No Excel files found/)).not.toBeInTheDocument();
  });

  test('auto-hides success messages after 8 seconds', async () => {
    jest.useFakeTimers();
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
      expect(screen.getByText('Found 1 Excel file(s) in the folder')).toBeInTheDocument();
    });

    // Fast-forward time by 8 seconds
    act(() => {
      jest.advanceTimersByTime(8000);
    });

    await waitFor(() => {
      expect(screen.queryByText('Found 1 Excel file(s) in the folder')).not.toBeInTheDocument();
    });

    jest.useRealTimers();
  });

  test('handles save error gracefully', async () => {
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

    // Mock XLSX.write to throw an error
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
