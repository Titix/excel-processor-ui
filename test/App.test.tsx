import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import App from '../src/frontend/App';

// Mock CSS imports
jest.mock('../src/frontend/App.css', () => ({}));

// Mock XLSX library
const mockXLSX = {
  read: jest.fn(),
  write: jest.fn(),
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
    mockXLSX.read.mockReturnValue({
      Sheets: {
        'Sheet1': {
          'A1': { v: 'Name' },
          'B1': { v: 'Age' },
          'A2': { v: 'John' },
          'B2': { v: 25 }
        }
      }
    });
    mockXLSX.write.mockReturnValue(new ArrayBuffer(8));
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
      expect(screen.getByText('Process File')).toBeInTheDocument();
    });
  });

  test('processes file when process button is clicked', async () => {
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
      expect(screen.getByText('Process File')).toBeInTheDocument();
    });

    const processButton = screen.getByText('Process File');
    fireEvent.click(processButton);

    await waitFor(() => {
      expect(screen.getByText('✅ File processed successfully! Ready to download.')).toBeInTheDocument();
    });
  });

  test('shows download button after processing', async () => {
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
      expect(screen.getByText('Process File')).toBeInTheDocument();
    });

    const processButton = screen.getByText('Process File');
    fireEvent.click(processButton);

    await waitFor(() => {
      expect(screen.getByText('Save Processed File')).toBeInTheDocument();
    });
  });

  test('handles file processing error', async () => {
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
      expect(screen.getByText('Process File')).toBeInTheDocument();
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

  test('handles file reading error', async () => {
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

    // Mock file reading error
    if (mockFileReader.onload) {
      mockFileReader.onload({ target: { result: null } });
    }

    await waitFor(() => {
      expect(screen.getByText(/Error reading Excel file/)).toBeInTheDocument();
    });
  });

  test('resets application state when choose new file is clicked', async () => {
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
      expect(screen.getByText('Choose New File')).toBeInTheDocument();
    });

    const chooseNewFileButton = screen.getByText('Choose New File');
    fireEvent.click(chooseNewFileButton);

    await waitFor(() => {
      expect(screen.getByText('📁 Ready to select a new file')).toBeInTheDocument();
    });
  });
});
