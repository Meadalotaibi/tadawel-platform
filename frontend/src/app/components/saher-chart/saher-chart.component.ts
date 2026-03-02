import { Component, Input, OnChanges, SimpleChanges, ViewChild, ElementRef, AfterViewInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Chart, ChartData } from 'chart.js';
import { BusinessLineRegionRow, RegionKey, REGION_COLUMNS, ChartDatasetObject } from '../../models/saher-statistics.model';
import { GROUPED_BAR_OPTIONS, createRegionDataset } from '../shared/chart-config';

/**
 * SAHER Violations Chart Component
 * 
 * Consumes data as Object Array format directly from backend statistics.
 * 
 * DATA FORMAT (Object Array):
 * Input: BusinessLineRegionRow[] - Array of objects from statisticsTables.businessLineByRegion.data
 * Each object: { 'Business Line Org Description': string, COA: number, WOA: number, EOA: number, SOA: number, 'N/A': number, Total: number }
 * 
 * CHART STRUCTURE:
 * - X-axis labels: Business Line names (from 'Business Line Org Description')
 * - Datasets: Object Array of region data (COA, WOA, EOA, SOA, N/A)
 * - Each dataset object: { label: string, data: number[], backgroundColor: string, ... }
 */
@Component({
  selector: 'app-saher-chart',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './saher-chart.component.html',
  styleUrl: './saher-chart.component.css'
})
export class SaherChartComponent implements OnChanges, AfterViewInit, OnDestroy {
  @ViewChild('chartCanvas') chartCanvas!: ElementRef<HTMLCanvasElement>;

  /**
   * Input: Object Array of business line data from backend
   * Mapped directly from: response.statisticsTables.businessLineByRegion.data
   * No hardcoded values - data comes from API response
   */
  @Input() data: BusinessLineRegionRow[] = [];
  
  /**
   * Chart title - passed from parent component
   */
  @Input() title: string = 'SAHER Violations by Business Line & Region';

  // Chart.js instance
  private chart: Chart<'bar'> | null = null;

  // Region columns from model - consistent with backend response
  private readonly regions: readonly RegionKey[] = REGION_COLUMNS;

  ngAfterViewInit(): void {
    this.createChart();
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['data'] && this.data && this.chartCanvas) {
      this.updateChart();
    }
  }

  ngOnDestroy(): void {
    if (this.chart) {
      this.chart.destroy();
      this.chart = null;
    }
  }

  /**
   * Create the Chart.js instance
   */
  private createChart(): void {
    if (!this.chartCanvas?.nativeElement) return;

    const ctx = this.chartCanvas.nativeElement.getContext('2d');
    if (!ctx) return;

    const chartData = this.buildChartData();

    this.chart = new Chart(ctx, {
      type: 'bar' as const,
      data: chartData,
      options: GROUPED_BAR_OPTIONS
    });
  }

  /**
   * Update existing chart with new data
   */
  private updateChart(): void {
    if (!this.chart) {
      this.createChart();
      return;
    }

    const chartData = this.buildChartData();
    this.chart.data = chartData;
    this.chart.update();
  }

  /**
   * Transform Object Array input into Chart.js format
   * 
   * INPUT (Object Array from backend):
   * [
   *   { 'Business Line Org Description': 'Generation', COA: 5, WOA: 3, EOA: 2, SOA: 1, 'N/A': 0, Total: 11 },
   *   { 'Business Line Org Description': 'Distribution', COA: 8, WOA: 4, EOA: 1, SOA: 2, 'N/A': 1, Total: 16 },
   *   ...
   * ]
   * 
   * OUTPUT (Chart.js format with Object Array datasets):
   * {
   *   labels: ['Generation', 'Distribution', ...],
   *   datasets: [
   *     { label: 'COA', data: [5, 8, ...], backgroundColor: '...', ... },
   *     { label: 'WOA', data: [3, 4, ...], backgroundColor: '...', ... },
   *     ...
   *   ]
   * }
   */
  private buildChartData(): ChartData<'bar'> {
    // Filter out the "Total" row - only display actual business lines
    const filteredData = this.data.filter(
      row => row['Business Line Org Description'] !== 'Total'
    );

    // X-axis labels: Business Line names mapped from Object Array
    const labels: string[] = filteredData.map(
      row => row['Business Line Org Description']
    );

    // Datasets: Object Array - one dataset object per region
    // Each dataset maps values from the corresponding region column in input data
    const datasets: ChartDatasetObject[] = this.regions.map((region, index) => 
      createRegionDataset(
        region,                                           // Dataset label (region name)
        filteredData.map(row => row[region] || 0),       // Data array from Object Array
        index                                             // Color index
      ) as ChartDatasetObject
    );

    return {
      labels,
      datasets
    };
  }
}
