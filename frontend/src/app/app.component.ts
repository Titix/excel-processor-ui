import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';

declare const XLSX: any;

interface FileData {
  name: string;
  size: number;
  data: any;
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule],
  template: `
    <div class="container">
      <header>
        <h1>Excel Processor</h1>
        <p>Upload, process, and download Excel files with ease</p>
      </header>

      <main>
        <div class="upload-section">
          <div class="upload-area" 
               [class.dragover]="isDragOver"
               (dragover)="onDragOver($event)"
               (dragleave)="onDragLeave($event)"
               (drop)="onDrop($event)">
            <div class="upload-content">
              <div class="upload-icon">📁</div>
              <h3>Choose Excel File</h3>
              <p>Drag and drop your Excel file here or click to browse</p>
              <input type="file" 
                     #fileInput 
                     accept=".xls,.xlsx" 
                     style="display: none;"
                     (change)="onFileSelected($event)">
              <button type="button" 
                      class="btn btn-primary" 
                      (click)="fileInput.click()">
                Browse Files
              </button>
            </div>
          </div>
          
          <div class="file-info" *ngIf="selectedFile">
            <div class="file-details">
              <span class="file-icon">📄</span>
              <span class="file-name">{{ selectedFile.name }}</span>
              <span class="file-size">{{ formatFileSize(selectedFile.size) }}</span>
            </div>
            <button type="button" 
                    class="btn btn-secondary btn-small" 
                    (click)="chooseNewFile()">
              Choose New File
            </button>
          </div>
        </div>

        <div class="process-section" *ngIf="selectedFile">
          <div class="section-header">
            <h3>Process File</h3>
            <p>Click the button below to process your uploaded file</p>
          </div>
          <button type="button" 
                  class="btn btn-warning" 
                  [disabled]="!workbookData"
                  (click)="processFile()">
            Process File
          </button>
        </div>

        <div class="download-section" *ngIf="processedWorkbookData">
          <div class="section-header">
            <h3>Download Processed File</h3>
            <p>Your file has been processed successfully</p>
          </div>
          <button type="button" 
                  class="btn btn-success" 
                  (click)="downloadFile()">
            Save Processed File
          </button>
        </div>

        <div class="message-area" 
             [class.success]="messageType === 'success'"
             [class.error]="messageType === 'error'"
             [class.info]="messageType === 'info'"
             *ngIf="message">
          {{ message }}
        </div>
      </main>

      <footer>
        <p>&copy; 2025 Excel Processor. Built with Angular and Node.js.</p>
      </footer>
    </div>
  `,
  styles: []
})
export class AppComponent implements OnInit {
  selectedFile: File | null = null;
  workbookData: any = null;
  processedWorkbookData: any = null;
  isDragOver = false;
  message = '';
  messageType: 'success' | 'error' | 'info' = 'info';

  ngOnInit() {
    // Initialize component
  }

  onDragOver(event: DragEvent) {
    event.preventDefault();
    this.isDragOver = true;
  }

  onDragLeave(event: DragEvent) {
    event.preventDefault();
    this.isDragOver = false;
  }

  onDrop(event: DragEvent) {
    event.preventDefault();
    this.isDragOver = false;
    
    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.handleFile(files[0]);
    }
  }

  onFileSelected(event: Event) {
    const target = event.target as HTMLInputElement;
    if (target.files && target.files.length > 0) {
      this.handleFile(target.files[0]);
    }
  }

  handleFile(file: File) {
    // Validate file type
    if (!file.name.match(/\.(xlsx?)$/i)) {
      this.showMessage('Please select an Excel file (.xls or .xlsx)', 'error');
      return;
    }
    
    // Validate file size (50MB limit)
    if (file.size > 50 * 1024 * 1024) {
      this.showMessage('File size must be less than 50MB', 'error');
      return;
    }
    
    this.selectedFile = file;
    this.readExcelFile(file);
  }

  readExcelFile(file: File) {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        this.workbookData = XLSX.read(data, { type: 'array' });
        this.showMessage('📁 File loaded successfully! Ready to process.', 'success');
      } catch (error) {
        console.error('Error reading Excel file:', error);
        this.showMessage('Error reading Excel file: ' + (error as Error).message, 'error');
        this.workbookData = null;
      }
    };
    
    reader.onerror = () => {
      this.showMessage('Error reading file', 'error');
      this.workbookData = null;
    };
    
    reader.readAsArrayBuffer(file);
  }

  processFile() {
    if (!this.workbookData) {
      this.showMessage('Please select a file first', 'error');
      return;
    }
    
    try {
      // Create a copy of the workbook for processing (safer approach)
      this.processedWorkbookData = XLSX.read(XLSX.write(this.workbookData, { type: 'array' }), { type: 'array' });
      
      // Process each sheet
      Object.keys(this.processedWorkbookData.Sheets).forEach(sheetName => {
        const sheet = this.processedWorkbookData.Sheets[sheetName];
        
        // Here you can add your processing logic
        // For now, we'll just copy the data as-is
        // You can modify cells, add formulas, etc.
      });
      
      this.showMessage('✅ File processed successfully! Ready to download.', 'success');
      
    } catch (error) {
      console.error('Processing error:', error);
      this.showMessage('Processing failed: ' + (error as Error).message, 'error');
    }
  }

  downloadFile() {
    if (!this.processedWorkbookData) {
      this.showMessage('Please process a file first', 'error');
      return;
    }
    
    try {
      // Determine the original file format
      const originalFormat = this.selectedFile?.name.toLowerCase().endsWith('.xls') ? 'xls' : 'xlsx';
      
      // Convert workbook to Excel file with optimized settings for compatibility and size
      const wbout = XLSX.write(this.processedWorkbookData, { 
        bookType: originalFormat,
        type: 'array',
        compression: true,  // Enable compression for smaller files
        cellStyles: false,  // Disable cell styles to reduce size
        cellNF: false,      // Disable number formats to reduce size
        cellHTML: false     // Disable HTML in cells to reduce size
      });
      
      // Create blob with appropriate MIME type
      const mimeType = originalFormat === 'xls' 
        ? 'application/vnd.ms-excel' 
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      
      const blob = new Blob([wbout], { type: mimeType });
      const url = window.URL.createObjectURL(blob);
      
      // Create download link with original format extension
      const a = document.createElement('a');
      a.href = url;
      a.download = this.selectedFile ? this.selectedFile.name.replace(/\.(xlsx?)$/i, '_processed.$1') : `processed_file.${originalFormat}`;
      document.body.appendChild(a);
      a.click();
      
      // Cleanup
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      // Add a small delay to ensure download has started
      setTimeout(() => {
        this.showMessage('📥 File downloaded successfully! Check your downloads folder.', 'success');
      }, 500);
      
    } catch (error) {
      console.error('Download error:', error);
      this.showMessage('Download failed: ' + (error as Error).message, 'error');
    }
  }

  chooseNewFile() {
    // Reset the application state
    this.selectedFile = null;
    this.workbookData = null;
    this.processedWorkbookData = null;
    this.message = '';
    
    // Show message
    this.showMessage('📁 Ready to select a new file', 'info');
  }

  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  showMessage(message: string, type: 'success' | 'error' | 'info') {
    this.message = message;
    this.messageType = type;
    
    // Auto-hide success messages after 8 seconds
    if (type === 'success') {
      setTimeout(() => {
        this.message = '';
      }, 8000);
    }
  }
}
