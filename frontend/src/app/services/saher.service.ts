import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { SaherStatistics } from '../models/saher-statistics.model';

/**
 * Response from the extract-dates endpoint
 */
export interface ExtractDatesResponse {
  success: boolean;
  minDate: string | null;
  maxDate: string | null;
  availableDates: string[];
  totalDates: number;
  error?: string;
}

@Injectable({
  providedIn: 'root'
})
export class SaherService {
  private readonly http = inject(HttpClient);
  private readonly apiUrl = 'http://localhost:3000';

  /**
   * Process Saher file and download cleaned Excel
   */
  processAndDownload(file: File): Observable<Blob> {
    const formData = new FormData();
    formData.append('file', file);

    return this.http.post(`${this.apiUrl}/api/process-saher`, formData, {
      responseType: 'blob'
    });
  }

  /**
   * Get Saher statistics as JSON with date range filtering
   * Returns structured data suitable for tables and future charts
   * @param file - The Excel file to process
   * @param startDate - Start date in YYYY-MM-DD format
   * @param endDate - End date in YYYY-MM-DD format
   */
  getStatistics(file: File, startDate: string, endDate: string): Observable<SaherStatistics> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('startDate', startDate);
    formData.append('endDate', endDate);

    return this.http.post<SaherStatistics>(`${this.apiUrl}/api/process-saher`, formData);
  }

  /**
   * Export cleaned Saher data as Excel file with statistics
   * @param file - The Excel file to process
   * @param startDate - Start date in YYYY-MM-DD format
   * @param endDate - End date in YYYY-MM-DD format
   */
  exportExcel(file: File, startDate: string, endDate: string): Observable<Blob> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('startDate', startDate);
    formData.append('endDate', endDate);

    return this.http.post(`${this.apiUrl}/api/export-saher`, formData, {
      responseType: 'blob'
    });
  }

  /**
   * Export Saher data as PowerPoint file
   * @param file - The Excel file to process
   * @param startDate - Start date in YYYY-MM-DD format
   * @param endDate - End date in YYYY-MM-DD format
   */
  exportPpt(file: File, startDate: string, endDate: string): Observable<Blob> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('startDate', startDate);
    formData.append('endDate', endDate);

    return this.http.post(`${this.apiUrl}/api/export-saher-ppt`, formData, {
      responseType: 'blob'
    });
  }

  /**
   * Extract available dates from an uploaded file
   * Returns the date range and list of valid dates present in the file
   * @param file - The Excel file to extract dates from
   * @param fileType - Type of file: 'saher' or 'mva' (accidents)
   */
  extractDates(file: File, fileType: 'saher' | 'mva'): Observable<ExtractDatesResponse> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('fileType', fileType);

    return this.http.post<ExtractDatesResponse>(`${this.apiUrl}/api/extract-dates`, formData);
  }

  /**
   * Export weekly SAHER data as PDF report
   * Returns a PDF with weekly window statistics and table
   * @param file - The Excel file to process
   * @param startDate - Start date in YYYY-MM-DD format
   * @param endDate - End date in YYYY-MM-DD format
   */
  exportWeeklyPdf(file: File, startDate: string, endDate: string): Observable<Blob> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('startDate', startDate);
    formData.append('endDate', endDate);

    return this.http.post(`${this.apiUrl}/api/export-weekly-pdf`, formData, {
      responseType: 'blob'
    });
  }
}













