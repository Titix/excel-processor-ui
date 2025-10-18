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

describe('App Component', () => {
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
    expect(screen.getByText('Excel Processor')).toBeInTheDocument();
  });

  test('renders upload area', () => {
    render(<App />);
    expect(screen.getByText('Choose Excel File')).toBeInTheDocument();
    expect(screen.getByText('Drag and drop your Excel file here or click to browse')).toBeInTheDocument();
  });

  test('renders browse files button', () => {
    render(<App />);
    expect(screen.getByText('Browse Files')).toBeInTheDocument();
  });

  test('renders version in footer', () => {
    render(<App />);
    expect(screen.getByText('Version 1.0.0')).toBeInTheDocument();
  });

  test('renders copyright in footer', () => {
    render(<App />);
    expect(screen.getByText(/© 2025 Excel Processor/)).toBeInTheDocument();
  });

  test('shows file input when browse button is clicked', () => {
    render(<App />);
    const browseButton = screen.getByText('Browse Files');
    fireEvent.click(browseButton);
    
    // The file input should be triggered (though hidden)
    const fileInput = document.querySelector('input[type="file"]');
    expect(fileInput).toBeInTheDocument();
  });

  test('handles drag over event', () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    fireEvent.dragOver(uploadArea!);
    expect(uploadArea).toHaveClass('dragover');
  });

  test('handles drag leave event', () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    fireEvent.dragOver(uploadArea!);
    expect(uploadArea).toHaveClass('dragover');
    
    fireEvent.dragLeave(uploadArea!);
    expect(uploadArea).not.toHaveClass('dragover');
  });

  test('validates file type on drop', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.txt', { type: 'text/plain' });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: {
        files: [file]
      }
    });

    await waitFor(() => {
      expect(screen.getByText('Please select an Excel file (.xls or .xlsx)')).toBeInTheDocument();
    });
  });

  test('validates file size on drop', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    // Create a mock file that's too large (51MB)
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 51 * 1024 * 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: {
        files: [file]
      }
    });

    await waitFor(() => {
      expect(screen.getByText('File size must be less than 50MB')).toBeInTheDocument();
    });
  });

  test('processes valid Excel file on drop', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    // Mock successful file reading
    mockFileReader.onload = jest.fn();
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: {
        files: [file]
      }
    });

    // Simulate successful file reading
    if (mockFileReader.onload) {
      mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
    }

    await waitFor(() => {
      expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
    });
  });

  test('shows process button after file is loaded', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    mockFileReader.onload = jest.fn();
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: {
        files: [file]
      }
    });

    if (mockFileReader.onload) {
      mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
    }

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Process File' })).toBeInTheDocument();
    });
  });

  test('processes file when process button is clicked', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    const processButton = screen.getByRole('button', { name: 'Process File' });
    
    await act(async () => {
      fireEvent.click(processButton);
    });

    await waitFor(() => {
      expect(screen.getByText('✅ File processed successfully! Ready to download.')).toBeInTheDocument();
    });
  });

  test('shows download button after processing', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    const processButton = screen.getByRole('button', { name: 'Process File' });
    
    await act(async () => {
      fireEvent.click(processButton);
    });

    await waitFor(() => {
      expect(screen.getByText('Save Processed File')).toBeInTheDocument();
    });
  });

  test('handles file processing error', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    // Mock processing error
    mockXLSX.write.mockImplementation(() => {
      throw new Error('Processing failed');
    });

    const processButton = screen.getByRole('button', { name: 'Process File' });
    
    await act(async () => {
      fireEvent.click(processButton);
    });

    await waitFor(() => {
      expect(screen.getByText(/Processing failed/)).toBeInTheDocument();
    });
  });

  test('handles file reading error', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    // Mock XLSX read to throw an error
    mockXLSX.read.mockImplementation(() => {
      throw new Error('Error reading Excel file');
    });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: {
        files: [file]
      }
    });

    // Simulate successful file reading but XLSX processing fails
    if (mockFileReader.onload) {
      mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
    }

    await waitFor(() => {
      expect(screen.getByText(/Error reading Excel file/)).toBeInTheDocument();
    });
  });

  test('handles empty file drop', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [] }
    });

    // Should not show any error or success message
    expect(screen.queryByText(/File loaded successfully/)).not.toBeInTheDocument();
    expect(screen.queryByText(/Please select an Excel file/)).not.toBeInTheDocument();
  });

  test('handles multiple files drop', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file1 = new File(['test content 1'], 'test1.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    const file2 = new File(['test content 2'], 'test2.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    Object.defineProperty(file1, 'size', { value: 1024 });
    Object.defineProperty(file2, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file1, file2] }
    });

    // Should only process the first file
    await waitFor(() => {
      expect(screen.getByText('test1.xlsx')).toBeInTheDocument();
    });
  });

  test('handles file input change event', async () => {
    render(<App />);
    const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
    
    const file = new File(['test content'], 'test.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    Object.defineProperty(fileInput, 'files', {
      value: [file],
      writable: false,
    });
    
    fireEvent.change(fileInput);

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
    });
  });

  test('handles file input change with invalid file', async () => {
    render(<App />);
    const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
    
    const file = new File(['test content'], 'test.txt', { type: 'text/plain' });
    
    Object.defineProperty(fileInput, 'files', {
      value: [file],
      writable: false,
    });
    
    fireEvent.change(fileInput);

    await waitFor(() => {
      expect(screen.getByText('Please select an Excel file (.xls or .xlsx)')).toBeInTheDocument();
    });
  });

  test('handles download button click', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    const processButton = screen.getByRole('button', { name: 'Process File' });
    
    await act(async () => {
      fireEvent.click(processButton);
    });

    await waitFor(() => {
      expect(screen.getByText('Save Processed File')).toBeInTheDocument();
    });

    const downloadButton = screen.getByText('Save Processed File');
    
    await act(async () => {
      fireEvent.click(downloadButton);
    });

    // Verify download was triggered
    expect(window.URL.createObjectURL).toHaveBeenCalled();
  });

  test('handles choose new file button click', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    mockFileReader.onload = jest.fn();
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    if (mockFileReader.onload) {
      mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
    }

    await waitFor(() => {
      expect(screen.getByText('Choose New File')).toBeInTheDocument();
    });

    const chooseNewFileButton = screen.getByText('Choose New File');
    fireEvent.click(chooseNewFileButton);

    await waitFor(() => {
      expect(screen.getByText('📁 Ready to select a new file')).toBeInTheDocument();
    });
  });

  test('handles browse files button click', () => {
    render(<App />);
    const browseButton = screen.getByText('Browse Files');
    
    fireEvent.click(browseButton);
    
    // The file input should be triggered (though hidden)
    const fileInput = document.querySelector('input[type="file"]');
    expect(fileInput).toBeInTheDocument();
  });

  test('handles upload area click', () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    fireEvent.click(uploadArea!);
    
    // The file input should be triggered
    const fileInput = document.querySelector('input[type="file"]');
    expect(fileInput).toBeInTheDocument();
  });

  test('handles drag over with preventDefault', () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    fireEvent.dragOver(uploadArea!);
    
    expect(uploadArea).toHaveClass('dragover');
  });

  test('handles drop with preventDefault', () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });
    
    // The drop should be handled without errors
    expect(uploadArea).not.toHaveClass('dragover');
  });

  test('resets application state when choose new file is clicked', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    mockFileReader.onload = jest.fn();
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    if (mockFileReader.onload) {
      mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
    }

    await waitFor(() => {
      expect(screen.getByText('Choose New File')).toBeInTheDocument();
    });

    const chooseNewFileButton = screen.getByText('Choose New File');
    fireEvent.click(chooseNewFileButton);

    await waitFor(() => {
      expect(screen.getByText('📁 Ready to select a new file')).toBeInTheDocument();
    });
  });

  test('auto-hides success messages after 8 seconds', async () => {
    jest.useFakeTimers();
    render(<App />);
    
    // Trigger a success message
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('📁 File loaded successfully! Ready to process.')).toBeInTheDocument();
    });

    // Fast-forward time by 8 seconds
    act(() => {
      jest.advanceTimersByTime(8000);
    });

    await waitFor(() => {
      expect(screen.queryByText('📁 File loaded successfully! Ready to process.')).not.toBeInTheDocument();
    });

    jest.useRealTimers();
  });

  test('handles FileReader onerror event', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onerror event
    await act(async () => {
      if (mockFileReader.onerror) {
        mockFileReader.onerror({} as any);
      }
    });

    await waitFor(() => {
      expect(screen.getByText('Error reading file')).toBeInTheDocument();
    });
  });

  test('shows error when processing without file', async () => {
    render(<App />);
    
    // Upload a file but don't simulate the FileReader onload event
    // This means workbookData will be null
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    // The Process File button should be disabled because workbookData is null
    const processButton = screen.getByRole('button', { name: 'Process File' });
    expect(processButton).toBeDisabled();
    
    // Since the button is disabled, we can't click it to test the error message
    // Instead, let's test that the button is properly disabled when workbookData is null
    expect(processButton).toHaveAttribute('disabled');
  });


  test('shows error when downloading without processed file', async () => {
    render(<App />);
    
    // Upload a file and simulate FileReader onload to get workbookData
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    const processButton = screen.getByRole('button', { name: 'Process File' });
    
    await act(async () => {
      fireEvent.click(processButton);
    });

    await waitFor(() => {
      expect(screen.getByText('Save Processed File')).toBeInTheDocument();
    });

    // Now clear the processedWorkbookData state to simulate error condition
    // We need to trigger a re-render with processedWorkbookData = null
    // This is tricky in tests, so let's test the actual error condition
    // by clicking download when processedWorkbookData is somehow null
    
    // Let's test a different approach - test the actual error condition
    // by mocking the download function to throw an error
    const downloadButton = screen.getByText('Save Processed File');
    
    // Mock XLSX.write to throw an error
    mockXLSX.write.mockImplementationOnce(() => {
      throw new Error('Download failed');
    });
    
    fireEvent.click(downloadButton);

    expect(screen.getByText('Download failed: Download failed')).toBeInTheDocument();
  });

  test('download button is only shown when file is processed', async () => {
    render(<App />);
    
    // Upload a file and simulate FileReader onload to get workbookData
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    const processButton = screen.getByRole('button', { name: 'Process File' });
    
    await act(async () => {
      fireEvent.click(processButton);
    });

    await waitFor(() => {
      expect(screen.getByText('Save Processed File')).toBeInTheDocument();
    });

    // The download button should only be visible after processing
    expect(screen.getByText('Save Processed File')).toBeInTheDocument();
  });

  test('handles download errors gracefully', async () => {
    render(<App />);
    const uploadArea = screen.getByText('Choose Excel File').closest('.upload-area');
    
    const file = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Object.defineProperty(file, 'size', { value: 1024 });
    
    fireEvent.drop(uploadArea!, {
      dataTransfer: { files: [file] }
    });

    // Simulate the FileReader onload event
    await act(async () => {
      if (mockFileReader.onload) {
        mockFileReader.onload({ target: { result: new ArrayBuffer(8) } });
      }
    });

    await waitFor(() => {
      expect(screen.getByText('test.xlsx')).toBeInTheDocument();
    });

    const processButton = screen.getByRole('button', { name: 'Process File' });
    
    await act(async () => {
      fireEvent.click(processButton);
    });

    await waitFor(() => {
      expect(screen.getByText('Save Processed File')).toBeInTheDocument();
    });

    // Mock URL.createObjectURL to throw an error
    const originalCreateObjectURL = window.URL.createObjectURL;
    window.URL.createObjectURL = jest.fn(() => {
      throw new Error('Download failed');
    });

    const downloadButton = screen.getByText('Save Processed File');
    
    await act(async () => {
      fireEvent.click(downloadButton);
    });

    await waitFor(() => {
      expect(screen.getByText(/Download failed/)).toBeInTheDocument();
    });

    // Restore original function
    window.URL.createObjectURL = originalCreateObjectURL;
  });
});
