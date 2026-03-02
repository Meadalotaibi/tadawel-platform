import { Component, Input, OnChanges, SimpleChanges, ViewChild, ElementRef, AfterViewInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Chart, ChartData } from 'chart.js';
import { HsseRegionRow, RegionKey, REGION_COLUMNS, ChartDatasetObject } from '../../models/saher-statistics.model';
import { GROUPED_BAR_OPTIONS, createRegionDataset } from '../shared/chart-config';

/**
 * HSSE Violations Chart Component
 * 
 * Consumes data as Object Array format directly from backend statistics.
 * 
 * DATA FORMAT (Object Array):
 * Input: HsseRegionRow[] - Array of objects from statisticsTables.hsseViolations.data
 * Each object: { Group: string, COA: number, WOA: number, EOA: number, SOA: number, 'N/A': number, Total: number }
 * 
 * CHART STRUCTURE:
 * - X-axis labels: HSSE Group names (from 'Group')
 * - Datasets: Object Array of region data (COA, WOA, EOA, SOA, N/A)
 * - Each dataset object: { label: string, data: number[], backgroundColor: string, ... }
 * 
 * CONSISTENCY:
 * - Uses same REGION_COLUMNS as SAHER chart
 * - Uses same createRegionDataset helper for styling
 * - Uses same GROUPED_BAR_OPTIONS for chart configuration
 */
@Component({
  selector: 'app-hsse-chart',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './hsse-chart.component.html',
  styleUrl: './hsse-chart.component.css'
})
export class HsseChartComponent implements OnChanges, AfterViewInit, OnDestroy {
  @ViewChild('chartCanvas') chartCanvas!: ElementRef<HTMLCanvasElement>;

  /**
   * Input: Object Array of HSSE group data from backend
   * Mapped directly from: response.statisticsTables.hsseViolations.data
   * No hardcoded values - data comes from API response
   */
  @Input() data: HsseRegionRow[] = [];
  
  /**
   * Chart title - passed from parent component
   */
  @Input() title: string = 'HSSE Violations by Group & Region';

  // Chart.js instance
  private chart: Chart<'bar'> | null = null;

  // Region columns from model - consistent with backend response and SAHER chart
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
   *   { Group: 'PPE Violation', COA: 3, WOA: 2, EOA: 1, SOA: 0, 'N/A': 0, Total: 6 },
   *   { Group: 'Unsafe Act', COA: 5, WOA: 1, EOA: 2, SOA: 1, 'N/A': 0, Total: 9 },
   *   ...
   * ]
   * 
   * OUTPUT (Chart.js format with Object Array datasets):
   * {
   *   labels: ['PPE Violation', 'Unsafe Act', ...],
   *   datasets: [
   *     { label: 'COA', data: [3, 5, ...], backgroundColor: '...', ... },
   *     { label: 'WOA', data: [2, 1, ...], backgroundColor: '...', ... },
   *     ...
   *   ]
   * }
   */
  private buildChartData(): ChartData<'bar'> {
    // Filter out the "Total" row - only display actual HSSE groups
    const filteredData = this.data.filter(
      row => row['Group'] !== 'Total'
    );

    // X-axis labels: HSSE Group names mapped from Object Array
    const labels: string[] = filteredData.map(
      row => row['Group']
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
