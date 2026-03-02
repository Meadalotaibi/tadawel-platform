import { ChartConfiguration } from 'chart.js';
import { Context } from 'chartjs-plugin-datalabels';

/**
 * Shared chart colors - emerald/teal theme matching the app
 */
export const CHART_COLORS = {
  backgrounds: [
    'rgba(16, 185, 129, 0.85)',   // COA - Emerald 500
    'rgba(52, 211, 153, 0.85)',   // WOA - Emerald 400
    'rgba(6, 182, 212, 0.85)',    // EOA - Cyan 500
    'rgba(59, 130, 246, 0.85)',   // SOA - Blue 500
    'rgba(139, 92, 246, 0.85)',   // N/A - Violet 500
  ],
  borders: [
    'rgba(16, 185, 129, 1)',
    'rgba(52, 211, 153, 1)',
    'rgba(6, 182, 212, 1)',
    'rgba(59, 130, 246, 1)',
    'rgba(139, 92, 246, 1)',
  ],
  // High contrast label colors for data labels
  labelColors: {
    light: 'rgba(255, 255, 255, 1)',      // White for dark backgrounds
    dark: 'rgba(15, 23, 42, 1)',          // Slate-900 for light backgrounds
    shadow: 'rgba(0, 0, 0, 0.4)'          // Subtle shadow for depth
  }
};

/**
 * Calculate relative luminance of an rgba color string
 * Used to determine if text should be light or dark for contrast
 */
function getColorLuminance(rgbaString: string): number {
  const match = rgbaString.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!match) return 0.5;
  
  const r = parseInt(match[1]) / 255;
  const g = parseInt(match[2]) / 255;
  const b = parseInt(match[3]) / 255;
  
  // sRGB luminance formula
  const toLinear = (c: number) => c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  return 0.2126 * toLinear(r) + 0.7152 * toLinear(g) + 0.0722 * toLinear(b);
}

/**
 * Get contrasting text color based on background luminance
 */
function getContrastColor(bgColor: string): string {
  const luminance = getColorLuminance(bgColor);
  // Use white text for darker backgrounds, dark text for lighter ones
  return luminance > 0.45 ? CHART_COLORS.labelColors.dark : CHART_COLORS.labelColors.light;
}

/**
 * Shared chart options for grouped bar charts
 * Used by both SAHER and HSSE violation charts
 */
export const GROUPED_BAR_OPTIONS: ChartConfiguration<'bar'>['options'] = {
  responsive: true,
  maintainAspectRatio: false,
  indexAxis: 'x',
  plugins: {
    legend: {
      display: true,
      position: 'top',
      align: 'center',
      labels: {
        color: 'rgba(255, 255, 255, 0.9)',
        font: {
          family: "'Outfit', 'Cairo', sans-serif",
          size: 12,
          weight: 500
        },
        padding: 20,
        usePointStyle: true,
        pointStyle: 'rectRounded'
      }
    },
    tooltip: {
      enabled: true,
      backgroundColor: 'rgba(6, 78, 59, 0.97)',
      titleColor: 'rgba(255, 255, 255, 1)',
      bodyColor: 'rgba(255, 255, 255, 0.9)',
      borderColor: 'rgba(16, 185, 129, 0.5)',
      borderWidth: 1,
      cornerRadius: 10,
      padding: 14,
      titleFont: {
        family: "'Outfit', 'Cairo', sans-serif",
        size: 14,
        weight: 600
      },
      bodyFont: {
        family: "'Outfit', 'Cairo', sans-serif",
        size: 13
      },
      // Show color box for region identification
      displayColors: true,
      boxWidth: 12,
      boxHeight: 12,
      boxPadding: 6,
      // Use point interaction for precise hover
      mode: 'nearest' as const,
      intersect: true,
      // Custom callbacks for smart tooltips
      callbacks: {
        // Title: Show Business Line / Group name (x-axis label)
        title: (tooltipItems) => {
          if (tooltipItems.length > 0) {
            const label = tooltipItems[0].label;
            return label || '';
          }
          return '';
        },
        // Label: Show Region name and exact value
        label: (tooltipItem) => {
          const datasetLabel = tooltipItem.dataset.label || 'Region';
          const value = tooltipItem.parsed.y;
          // Safe formatting: fallback to '0' if value is null/undefined
          const formattedValue = (value ?? 0).toLocaleString();
          return `${datasetLabel}: ${formattedValue}`;
        },
        // Footer: Optional summary (can be extended)
        footer: () => ''
      }
    },
    // Data Labels: Clean & minimal - just the numeric value
    // Tooltips provide detailed context on hover
    datalabels: {
      // Only show labels when value > 0
      display: (context: Context) => {
        const rawValue = context.dataset.data[context.dataIndex];
        // Safe check: treat null/undefined as 0
        const value = typeof rawValue === 'number' ? rawValue : 0;
        return value > 0;
      },
      // Position: on top of each bar
      anchor: 'end',
      align: 'top',
      offset: 2,
      // High contrast white color for visibility
      color: 'rgba(255, 255, 255, 0.95)',
      // Clean, readable font
      font: {
        family: "'Outfit', 'Cairo', sans-serif",
        size: 10,
        weight: 700
      },
      // Formatter: Just the numeric value, no extra text
      formatter: (value: number | null) => {
        // Safe formatting: return empty string if null/undefined or <= 0
        if (value == null || value <= 0) return '';
        return value.toLocaleString();
      },
      // Minimal padding
      padding: {
        top: 2,
        bottom: 2,
        left: 3,
        right: 3
      },
      // Subtle background for readability on glass UI
      backgroundColor: 'rgba(0, 0, 0, 0.35)',
      borderRadius: 3,
      // Prevent overlap
      clamp: true,
      clip: false
    }
  },
  scales: {
    x: {
      grid: {
        color: 'rgba(16, 185, 129, 0.1)',
        lineWidth: 1
      },
      ticks: {
        color: 'rgba(255, 255, 255, 0.8)',
        font: {
          family: "'Outfit', 'Cairo', sans-serif",
          size: 11,
          weight: 500
        },
        maxRotation: 45,
        minRotation: 45
      },
      border: {
        color: 'rgba(16, 185, 129, 0.2)'
      }
    },
    y: {
      beginAtZero: true,
      grid: {
        color: 'rgba(16, 185, 129, 0.1)',
        lineWidth: 1
      },
      ticks: {
        color: 'rgba(255, 255, 255, 0.8)',
        font: {
          family: "'Outfit', 'Cairo', sans-serif",
          size: 12
        },
        stepSize: 1,
        precision: 0
      },
      border: {
        color: 'rgba(16, 185, 129, 0.2)'
      }
    }
  },
  animation: {
    duration: 800,
    easing: 'easeOutQuart'
  },
  interaction: {
    mode: 'index',
    intersect: false
  }
};

/**
 * Create dataset configuration for a region
 */
export function createRegionDataset(
  label: string,
  data: number[],
  colorIndex: number
) {
  return {
    label,
    data,
    backgroundColor: CHART_COLORS.backgrounds[colorIndex],
    borderColor: CHART_COLORS.borders[colorIndex],
    borderWidth: 1,
    borderRadius: 4,
    borderSkipped: false as const,
    barPercentage: 0.8,
    categoryPercentage: 0.85
  };
}






