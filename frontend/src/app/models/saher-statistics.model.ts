// Saher Statistics Data Models
// Designed for reusability - can be used for both tables and future charts

// Selected range totals
export interface SelectedTotals {
  totalViolations: number;
  cancels: number;
}

// Weekly totals with date range
export interface WeeklyTotals {
  weeklyStartDate: string;
  weeklyEndDate: string;
  totalViolations: number;
  hsseViolations: number;
}

// Legacy summary format (for backward compatibility)
export interface SaherSummary {
  cancelCount: number;
  totalViolations: number;
  totalRowsProcessed?: number;
}

// Region column keys type
export type RegionKey = 'COA' | 'WOA' | 'EOA' | 'SOA' | 'N/A';

// All region keys as a readonly tuple for iteration
export const REGION_COLUMNS: readonly RegionKey[] = ['COA', 'WOA', 'EOA', 'SOA', 'N/A'] as const;

export interface BusinessLineRegionRow {
  'Business Line Org Description': string;
  'COA': number;
  'WOA': number;
  'EOA': number;
  'SOA': number;
  'N/A': number;
  'Total': number;
}

export interface BusinessLineByRegion {
  columns: string[];
  data: BusinessLineRegionRow[];
}

// HSSE table row (Group × Region)
export interface HsseRegionRow {
  'Group': string;
  'COA': number;
  'WOA': number;
  'EOA': number;
  'SOA': number;
  'N/A': number;
  'Total': number;
}

export interface HsseViolations {
  columns: string[];
  data: HsseRegionRow[];
}

// Statistics tables container
export interface StatisticsTables {
  businessLineByRegion: BusinessLineByRegion;
  hsseViolations: HsseViolations;
}

// Main response interface from /api/process-saher
export interface SaherStatistics {
  noData: boolean;
  selectedTotals: SelectedTotals;
  weeklyTotals: WeeklyTotals;
  statisticsTables: StatisticsTables;
  // Computed summary for backward compatibility with existing UI
  summary: SaherSummary;
  // Direct access to tables (for convenience)
  businessLineByRegion: BusinessLineByRegion;
  hsseViolations: HsseViolations;
}

/**
 * ============================================
 * CHART DATA FORMATS - Object Array Pattern
 * ============================================
 * Charts consume data as Object Arrays directly from backend statistics.
 * No hardcoded values - all data is mapped from API response.
 */

/**
 * Chart.js Dataset Object format
 * Each dataset represents a region (COA, WOA, EOA, SOA, N/A)
 */
export interface ChartDatasetObject {
  label: string;           // Region name (e.g., 'COA', 'WOA')
  data: number[];          // Values for each business line/group
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
  borderRadius?: number;
}

/**
 * Chart.js Data structure (labels + datasets)
 * Labels = Business Lines (SAHER) or Groups (HSSE)
 * Datasets = Object Array of region data
 */
export interface ChartDataStructure {
  labels: string[];               // X-axis labels (Business Lines or Groups)
  datasets: ChartDatasetObject[]; // Object Array of region datasets
}

/**
 * Mapping function type for transforming statistics to chart format
 * Input: Object Array from backend (BusinessLineRegionRow[] or HsseRegionRow[])
 * Output: ChartDataStructure
 */
export type ChartDataMapper<T> = (data: T[], regions: readonly RegionKey[]) => ChartDataStructure;

/**
 * Legacy chart-ready data format (retained for backward compatibility)
 */
export interface ChartDataPoint {
  label: string;
  value: number;
  color?: string;
}

export interface ChartSeries {
  name: string;
  data: ChartDataPoint[];
}



















