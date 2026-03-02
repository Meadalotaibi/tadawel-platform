import { Component, signal, inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { FormGroup, FormControl, ReactiveFormsModule, Validators } from '@angular/forms';
import { MatDatepickerModule } from '@angular/material/datepicker';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatMenuModule } from '@angular/material/menu';
import { MatSnackBar, MatSnackBarModule } from '@angular/material/snack-bar';
import { LanguageService } from './services/language.service';
import { SaherService } from './services/saher.service';
import { SaherStatistics, BusinessLineRegionRow, HsseRegionRow, RegionKey, REGION_COLUMNS } from './models/saher-statistics.model';
import { SaherChartComponent } from './components/saher-chart/saher-chart.component';
import { HsseChartComponent } from './components/hsse-chart/hsse-chart.component';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { exportElementToPdf } from './utils/pdf-export';

@Component({
  selector: 'app-root',
  imports: [
    CommonModule,
    ReactiveFormsModule,
    MatDatepickerModule,
    MatFormFieldModule,
    MatInputModule,
    MatButtonModule,
    MatIconModule,
    MatMenuModule,
    MatSnackBarModule,
    SaherChartComponent,
    HsseChartComponent
  ],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App implements OnInit {
  private http = inject(HttpClient);
  private saherService = inject(SaherService);
  private snackBar = inject(MatSnackBar);
  languageService = inject(LanguageService);

  saherFile = signal<File | null>(null);
  mvaFile = signal<File | null>(null);
  analyzing = signal<boolean>(false);
  exporting = signal<boolean>(false);
  exportingPpt = signal<boolean>(false);
  exportingZip = signal<boolean>(false);
  exportingPdf = signal<boolean>(false);
  exportingWeeklyPdf = signal<boolean>(false);
  error = signal<string | null>(null);
  success = signal<string | null>(null);

  // UI state: controls whether to show upload section or results section
  analysisCompleted = signal<boolean>(false);

  // Date range form group with required validators
  dateRangeForm = new FormGroup({
    start: new FormControl<Date | null>(null, Validators.required),
    end: new FormControl<Date | null>(null, Validators.required)
  });

  // Minimum selectable date: January 1, 2026
  // Dates before this are disabled in the calendar UI
  readonly minSelectableDate = new Date(2026, 0, 1); // Month is 0-indexed (0 = January)

  // Statistics data - structured for tables and future charts
  statistics = signal<SaherStatistics | null>(null);
  hasData = signal<boolean>(false);

  // Store latest exported blobs for future ZIP bundling
  latestExcelBlob: Blob | null = null;
  latestPptBlob: Blob | null = null;

  private readonly apiUrl = 'http://localhost:3000';

  // Region columns for the SAHER table - typed as RegionKey[]
  readonly regionColumns: readonly RegionKey[] = REGION_COLUMNS;

  ngOnInit() {
    // Subscribe to form changes to clear errors when dates change
    this.dateRangeForm.valueChanges.subscribe(() => {
      if (this.error()) {
        this.error.set(null);
      }
    });
  }

  get translations() {
    return this.languageService.getTranslations();
  }

  get currentLang() {
    return this.languageService.currentLanguage();
  }

  // Check if date range is complete (both start and end dates selected)
  isDateRangeComplete(): boolean {
    const { start, end } = this.dateRangeForm.value;
    return !!(start && end);
  }

  // Check if analyze button should be enabled
  canAnalyze(): boolean {
    return !!(this.saherFile() && this.isDateRangeComplete() && !this.analyzing());
  }

  // Check if any export is currently in progress
  isAnyExportInProgress(): boolean {
    return this.exporting() || this.exportingPpt() || this.exportingZip() || this.exportingPdf() || this.exportingWeeklyPdf();
  }

  // Format date to YYYY-MM-DD string
  private formatDateToYMD(date: Date): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  // Get formatted start date
  getStartDateString(): string | null {
    const start = this.dateRangeForm.value.start;
    return start ? this.formatDateToYMD(start) : null;
  }

  // Get formatted end date
  getEndDateString(): string | null {
    const end = this.dateRangeForm.value.end;
    return end ? this.formatDateToYMD(end) : null;
  }

  // Clear date range
  clearDateRange() {
    this.dateRangeForm.reset();
  }

  onSaherFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      const file = input.files[0];
      if (this.isExcelFile(file)) {
        this.saherFile.set(file);
        this.error.set(null);
        this.success.set(null);
      } else {
        this.error.set('Please select a valid Excel file (.xlsx or .xls)');
      }
    }
  }

  onMvaFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      const file = input.files[0];
      if (this.isExcelFile(file)) {
        this.mvaFile.set(file);
        this.error.set(null);
        this.success.set(null);
      } else {
        this.error.set('Please select a valid Excel file (.xlsx or .xls)');
      }
    }
  }

  isExcelFile(file: File): boolean {
    const validTypes = [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];
    return validTypes.includes(file.type) || file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
  }

  analyzeData() {
    const saherFile = this.saherFile();
    const startDate = this.getStartDateString();
    const endDate = this.getEndDateString();
    
    if (!saherFile) {
      this.error.set(this.translations.uploadPage.fileValidation);
      return;
    }

    if (!startDate || !endDate) {
      this.error.set(this.translations.dateRange.validation);
      return;
    }

    this.analyzing.set(true);
    this.error.set(null);
    this.success.set(null);
    this.statistics.set(null);
    this.hasData.set(false);

    // Fetch statistics with date range (YYYY-MM-DD format)
    this.saherService.getStatistics(saherFile, startDate, endDate).subscribe({
      next: (response: any) => {
        // Check for noData response
        if (response.noData) {
          this.statistics.set(null);
          this.hasData.set(false);
          this.analyzing.set(false);
          this.error.set(response.message || this.translations.statistics.noData);
          return;
        }

        // Map the new response structure to our model
        const stats: SaherStatistics = {
          noData: response.noData,
          selectedTotals: response.selectedTotals,
          weeklyTotals: response.weeklyTotals,
          statisticsTables: response.statisticsTables,
          // Map for backward compatibility with existing UI
          summary: {
            cancelCount: response.selectedTotals.cancels,
            totalViolations: response.selectedTotals.totalViolations
          },
          businessLineByRegion: response.statisticsTables.businessLineByRegion,
          hsseViolations: response.statisticsTables.hsseViolations
        };

        this.statistics.set(stats);
        // Check if there's actual data
        this.hasData.set(stats.selectedTotals.totalViolations > 0 || stats.selectedTotals.cancels > 0);
        
        this.analyzing.set(false);
        this.analysisCompleted.set(true);
        this.success.set(this.currentLang === 'en' 
          ? 'Data analyzed successfully!' 
          : 'تم تحليل البيانات بنجاح!');
      },
      error: (err) => {
        console.error('Error fetching statistics:', err);
        this.analyzing.set(false);
        this.error.set(this.currentLang === 'en' 
          ? 'Error analyzing data. Please try again.' 
          : 'خطأ في تحليل البيانات. يرجى المحاولة مرة أخرى.');
      }
    });
  }

  switchLanguage() {
    this.languageService.toggleLanguage();
  }

  removeSaherFile() {
    this.saherFile.set(null);
    this.success.set(null);
    this.statistics.set(null);
    this.hasData.set(false);
    this.analysisCompleted.set(false);
    this.clearDateRange();
  }

  removeMvaFile() {
    this.mvaFile.set(null);
    this.success.set(null);
  }

  // Go back to upload section (new analysis)
  startNewAnalysis() {
    this.analysisCompleted.set(false);
    this.statistics.set(null);
    this.hasData.set(false);
    this.success.set(null);
    this.error.set(null);
    // Keep the file and date range so user can quickly re-analyze
  }

  // Helper method to get business line data for table display
  getBusinessLineData(): BusinessLineRegionRow[] {
    return this.statistics()?.businessLineByRegion.data || [];
  }

  // Helper method to get HSSE data (single table with Group × Region)
  getHsseData(): HsseRegionRow[] {
    return this.statistics()?.hsseViolations.data || [];
  }

  // Type-safe accessor for region values in BusinessLineRegionRow
  getRegionValue(row: BusinessLineRegionRow, region: RegionKey): number {
    return row[region];
  }

  // Type-safe accessor for region values in HsseRegionRow
  getHsseRegionValue(row: HsseRegionRow, region: RegionKey): number {
    return row[region];
  }

  /**
   * Calculate total HSSE violations from the chart data
   * Dynamically computed from the filtered dataset
   */
  getHsseTotalViolations(): number {
    const hsseData = this.getHsseData();
    if (!hsseData || hsseData.length === 0) return 0;
    
    // Find the "Total" row which contains the sum
    const totalRow = hsseData.find(row => row['Group'] === 'Total');
    if (totalRow) {
      return totalRow['Total'] ?? 0;
    }
    
    // Fallback: sum all Total values from non-Total rows
    return hsseData
      .filter(row => row['Group'] !== 'Total')
      .reduce((sum, row) => sum + (row['Total'] ?? 0), 0);
  }

  /**
   * Get weekly HSSE violations count
   * Returns the actual count from the backend (filtered by weekly date range, excluding CANCEL)
   */
  getWeeklyHsseViolations(): number {
    const stats = this.statistics();
    if (!stats) return 0;
    
    // Use the actual weekly HSSE count from the backend
    return stats.weeklyTotals?.hsseViolations ?? 0;
  }

  /**
   * Get the weekly date range for display
   */
  getWeeklyDateRange(): string {
    const stats = this.statistics();
    if (!stats?.weeklyTotals) return '';
    return `${stats.weeklyTotals.weeklyStartDate} → ${stats.weeklyTotals.weeklyEndDate}`;
  }

  // Export cleaned data as Excel file
  exportData() {
    const saherFile = this.saherFile();
    const startDate = this.getStartDateString();
    const endDate = this.getEndDateString();
    
    if (!saherFile || !startDate || !endDate) {
      this.error.set(this.translations.dateRange.validation);
      return;
    }

    this.exporting.set(true);
    this.error.set(null);

    this.saherService.exportExcel(saherFile, startDate, endDate).subscribe({
      next: (blob: Blob) => {
        // Store blob for future ZIP bundling
        this.latestExcelBlob = blob;
        
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `saher_cleaned_${startDate}_${endDate}.xlsx`;
        link.click();
        window.URL.revokeObjectURL(url);
        
        this.exporting.set(false);
        this.success.set(this.currentLang === 'en' 
          ? 'Excel file exported successfully!' 
          : 'تم تصدير ملف Excel بنجاح!');
      },
      error: (err) => {
        console.error('Error exporting file:', err);
        this.exporting.set(false);
        this.error.set(this.currentLang === 'en' 
          ? 'Error exporting file. Please try again.' 
          : 'خطأ في تصدير الملف. يرجى المحاولة مرة أخرى.');
      }
    });
  }

  // Export data as PowerPoint file
  exportPptData() {
    const saherFile = this.saherFile();
    const startDate = this.getStartDateString();
    const endDate = this.getEndDateString();
    
    if (!saherFile || !startDate || !endDate) {
      this.error.set(this.translations.dateRange.validation);
      return;
    }

    this.exportingPpt.set(true);
    this.error.set(null);

    this.saherService.exportPpt(saherFile, startDate, endDate).subscribe({
      next: (blob: Blob) => {
        // Store blob for future ZIP bundling
        this.latestPptBlob = blob;
        
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `saher_weekly_report_${startDate}_${endDate}.pptx`;
        link.click();
        window.URL.revokeObjectURL(url);
        
        this.exportingPpt.set(false);
        this.success.set(this.currentLang === 'en' 
          ? 'PowerPoint file exported successfully!' 
          : 'تم تصدير ملف PowerPoint بنجاح!');
      },
      error: (err) => {
        console.error('Error exporting PowerPoint:', err);
        this.exportingPpt.set(false);
        this.error.set(this.currentLang === 'en' 
          ? 'Error exporting PowerPoint. Please try again.' 
          : 'خطأ في تصدير PowerPoint. يرجى المحاولة مرة أخرى.');
      }
    });
  }

  // Export all files as ZIP (Excel + PPT)
  exportAllAsZip() {
    const saherFile = this.saherFile();
    const startDate = this.getStartDateString();
    const endDate = this.getEndDateString();
    
    if (!saherFile || !startDate || !endDate) {
      this.snackBar.open(
        this.currentLang === 'en' ? 'Please select file and date range first' : 'يرجى اختيار الملف والنطاق الزمني أولاً',
        this.currentLang === 'en' ? 'Close' : 'إغلاق',
        { duration: 3000 }
      );
      return;
    }

    this.exportingZip.set(true);

    // Always fetch fresh PPT blob for ZIP
    this.saherService.exportPpt(saherFile, startDate, endDate).subscribe({
      next: async (pptBlob: Blob) => {
        try {
          // Store the PPT blob
          this.latestPptBlob = pptBlob;
          
          // Create ZIP
          const zip = new JSZip();
          
          // Add Excel blob if exists
          if (this.latestExcelBlob) {
            zip.file(`saher_cleaned_${startDate}_${endDate}.xlsx`, this.latestExcelBlob);
          }
          
          // Add PPT blob (always required)
          zip.file(`Saher_for_Week_${startDate}_${endDate}.pptx`, pptBlob);
          
          // Generate ZIP
          const zipBlob = await zip.generateAsync({ type: 'blob' });
          
          // Download ZIP with today's date
          const today = new Date().toISOString().split('T')[0];
          saveAs(zipBlob, `Saher_Export_${today}.zip`);
          
          this.exportingZip.set(false);
          this.snackBar.open(
            this.currentLang === 'en' ? 'ZIP exported successfully!' : 'تم تصدير الملف المضغوط بنجاح!',
            this.currentLang === 'en' ? 'Close' : 'إغلاق',
            { duration: 3000, panelClass: ['success-snackbar'] }
          );
        } catch (zipError) {
          console.error('Error creating ZIP:', zipError);
          this.exportingZip.set(false);
          this.snackBar.open(
            this.currentLang === 'en' ? 'Error creating ZIP file' : 'خطأ في إنشاء الملف المضغوط',
            this.currentLang === 'en' ? 'Close' : 'إغلاق',
            { duration: 3000, panelClass: ['error-snackbar'] }
          );
        }
      },
      error: (err) => {
        console.error('Error fetching PPT for ZIP:', err);
        this.exportingZip.set(false);
        this.snackBar.open(
          this.currentLang === 'en' ? 'Error generating PowerPoint for ZIP' : 'خطأ في إنشاء PowerPoint للملف المضغوط',
          this.currentLang === 'en' ? 'Close' : 'إغلاق',
          { duration: 3000, panelClass: ['error-snackbar'] }
        );
      }
    });
  }

  // Export report as PDF
  async exportPdf() {
    const startDate = this.getStartDateString();
    const endDate = this.getEndDateString();
    
    if (!startDate || !endDate) {
      this.snackBar.open(
        this.currentLang === 'en' ? 'Please select a date range first' : 'يرجى اختيار النطاق الزمني أولاً',
        this.currentLang === 'en' ? 'Close' : 'إغلاق',
        { duration: 3000 }
      );
      return;
    }

    // Get the export root element
    const exportRoot = document.getElementById('report-export-root');
    if (!exportRoot) {
      this.snackBar.open(
        this.currentLang === 'en' ? 'Report content not found' : 'لم يتم العثور على محتوى التقرير',
        this.currentLang === 'en' ? 'Close' : 'إغلاق',
        { duration: 3000, panelClass: ['error-snackbar'] }
      );
      return;
    }

    this.exportingPdf.set(true);

    try {
      // Export to PDF
      await exportElementToPdf({
        element: exportRoot,
        fileName: `Saher_Weekly_Report_${startDate}_${endDate}`,
        orientation: 'portrait',
        scale: 2,
        backgroundColor: '#0f172a'
      });

      this.exportingPdf.set(false);
      this.snackBar.open(
        this.currentLang === 'en' ? 'PDF exported successfully!' : 'تم تصدير ملف PDF بنجاح!',
        this.currentLang === 'en' ? 'Close' : 'إغلاق',
        { duration: 3000, panelClass: ['success-snackbar'] }
      );
    } catch (error) {
      console.error('Error exporting PDF:', error);
      this.exportingPdf.set(false);
      this.snackBar.open(
        this.currentLang === 'en' ? 'Error exporting PDF. Please try again.' : 'خطأ في تصدير PDF. يرجى المحاولة مرة أخرى.',
        this.currentLang === 'en' ? 'Close' : 'إغلاق',
        { duration: 3000, panelClass: ['error-snackbar'] }
      );
    }
  }

  // Export Weekly PDF report (server-generated)
  exportWeeklyPdf() {
    const saherFile = this.saherFile();
    const startDate = this.getStartDateString();
    const endDate = this.getEndDateString();
    
    if (!saherFile || !startDate || !endDate) {
      this.snackBar.open(
        this.currentLang === 'en' ? 'Please select file and date range first' : 'يرجى اختيار الملف والنطاق الزمني أولاً',
        this.currentLang === 'en' ? 'Close' : 'إغلاق',
        { duration: 3000 }
      );
      return;
    }

    this.exportingWeeklyPdf.set(true);
    this.error.set(null);

    this.saherService.exportWeeklyPdf(saherFile, startDate, endDate).subscribe({
      next: (blob: Blob) => {
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        // Filename is set by Content-Disposition header, but we provide a fallback
        link.download = `WeeklyReport_${startDate}_${endDate}.pdf`;
        link.click();
        window.URL.revokeObjectURL(url);
        
        this.exportingWeeklyPdf.set(false);
        this.snackBar.open(
          this.currentLang === 'en' ? 'Weekly PDF exported successfully!' : 'تم تصدير تقرير PDF الأسبوعي بنجاح!',
          this.currentLang === 'en' ? 'Close' : 'إغلاق',
          { duration: 3000, panelClass: ['success-snackbar'] }
        );
      },
      error: (err) => {
        console.error('Error exporting Weekly PDF:', err);
        this.exportingWeeklyPdf.set(false);
        this.snackBar.open(
          this.currentLang === 'en' ? 'Error exporting Weekly PDF. Please try again.' : 'خطأ في تصدير تقرير PDF الأسبوعي. يرجى المحاولة مرة أخرى.',
          this.currentLang === 'en' ? 'Close' : 'إغلاق',
          { duration: 3000, panelClass: ['error-snackbar'] }
        );
      }
    });
  }
}
