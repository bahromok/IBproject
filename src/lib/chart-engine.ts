import { ChartJSNodeCanvas } from 'chartjs-node-canvas';
import sharp from 'sharp';

// ─── TYPES ────────────────────────────────────────────────────────────────────

export interface ChartConfig {
  type: 'pie' | 'bar' | 'line' | 'doughnut' | 'horizontalBar' | 'stacked-bar' | 'area' | 'radar' | 'scatter' | 'combo';
  labels: string[];
  values: number[];
  title?: string;
  colors?: string[];
  width?: number;
  height?: number;
  showLegend?: boolean;
  showValues?: boolean;
  currency?: boolean;
  percentage?: boolean;
  seriesName?: string;
  borderColor?: string;
  borderWidth?: number;
  showGrid?: boolean;
  showAxis?: boolean;
  backgroundColor?: string;
  fontSize?: number;
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  dataLabels?: boolean;
  donutHole?: number;
  startAngle?: number;
}

export interface MultiSeriesChartConfig {
  type: 'bar' | 'line' | 'stacked-bar' | 'area' | 'combo' | 'radar';
  labels: string[];
  datasets: Array<{ name: string; values: number[]; color?: string; type?: string; fill?: boolean }>;
  title?: string;
  width?: number;
  height?: number;
  showLegend?: boolean;
  currency?: boolean;
  showGrid?: boolean;
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
}

// ─── COLOR PALETTES ───────────────────────────────────────────────────────────

const PALETTES: Record<string, string[]> = {
  professional: ['#1E40AF', '#3B82F6', '#60A5FA', '#93C5FD', '#2563EB', '#1D4ED8', '#1E3A8A', '#DBEAFE'],
  vibrant: ['#7C3AED', '#EC4899', '#F59E0B', '#10B981', '#3B82F6', '#EF4444', '#8B5CF6', '#14B8A6'],
  warm: ['#D97706', '#EA580C', '#DC2626', '#B45309', '#F59E0B', '#FBBF24', '#F97316', '#C2410C'],
  cool: ['#0EA5E9', '#06B6D4', '#14B8A6', '#10B981', '#3B82F6', '#6366F1', '#8B5CF6', '#2563EB'],
  pastel: ['#A5B4FC', '#FCA5A5', '#FCD34D', '#6EE7B7', '#93C5FD', '#F9A8D4', '#C4B5FD', '#67E8F9'],
  grayscale: ['#1F2937', '#374151', '#4B5563', '#6B7280', '#9CA3AF', '#D1D5DB', '#E5E7EB', '#F3F4F6'],
  earth: ['#92400E', '#78350F', '#365314', '#164E63', '#1E3A5F', '#4C1D95', '#831843', '#7F1D1D'],
  ocean: ['#0C4A6E', '#0369A1', '#0284C7', '#0EA5E9', '#38BDF8', '#7DD3FC', '#BAE6FD', '#E0F2FE'],
  sunset: ['#9D174D', '#BE185D', '#DB2777', '#F472B6', '#F59E0B', '#FBBF24', '#FCD34D', '#FDE68A'],
  neon: ['#00FF87', '#00F0FF', '#FF00E5', '#FFD700', '#FF3366', '#33FF57', '#3399FF', '#FF6633'],
};

// ─── CHART RENDERER ───────────────────────────────────────────────────────────

/**
 * Render a chart as a PNG buffer. Uses chartjs-node-canvas with sharp fallback.
 */
export async function renderChart(config: ChartConfig | MultiSeriesChartConfig): Promise<Buffer> {
  const width = config.width || 600;
  const height = config.height || 400;

  try {
    const canvas = new ChartJSNodeCanvas({
      width,
      height,
      backgroundColour: 'white',
      plugins: {
        modern: ['chartjs-plugin-datalabels'],
      },
    });

    let chartConfig: any;

    if ('datasets' in config) {
      chartConfig = buildMultiSeriesConfig(config as MultiSeriesChartConfig);
    } else {
      chartConfig = buildSingleSeriesConfig(config as ChartConfig);
    }

    const buffer = await canvas.renderToBuffer(chartConfig);
    return buffer;
  } catch (error: any) {
    // Fallback: render a simple chart image using sharp
    return renderFallbackChart(config, width, height);
  }
}

/**
 * Fallback chart renderer using sharp (always works, no native deps)
 */
async function renderFallbackChart(config: ChartConfig | MultiSeriesChartConfig, width: number, height: number): Promise<Buffer> {
  const isMulti = 'datasets' in config;
  const title = config.title || 'Chart';
  const labels = config.labels;
  const palette = PALETTES.professional;

  // Build SVG chart
  let svg = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">`;
  svg += `<rect width="100%" height="100%" fill="white"/>`;

  // Title
  if (title) {
    svg += `<text x="${width / 2}" y="30" text-anchor="middle" font-size="16" font-weight="bold" font-family="Arial, sans-serif" fill="#1F2937">${escapeXmlSvg(title)}</text>`;
  }

  const chartTop = 50;
  const chartBottom = height - 40;
  const chartHeight = chartBottom - chartTop;
  const chartLeft = 60;
  const chartRight = width - 40;
  const chartWidth = chartRight - chartLeft;

  if (!isMulti) {
    const singleConfig = config as ChartConfig;
    const values = singleConfig.values;
    const type = singleConfig.type;

    if (type === 'pie' || type === 'doughnut') {
      const cx = width / 2;
      const cy = (chartTop + chartBottom) / 2;
      const radius = Math.min(chartWidth, chartHeight) / 2 - 20;
      const innerRadius = type === 'doughnut' ? radius * 0.5 : 0;
      const total = values.reduce((a, b) => a + b, 0);

      let startAngle = -Math.PI / 2;
      values.forEach((value, i) => {
        const sliceAngle = (value / total) * 2 * Math.PI;
        const endAngle = startAngle + sliceAngle;
        const color = singleConfig.colors?.[i] || palette[i % palette.length];

        const x1 = cx + radius * Math.cos(startAngle);
        const y1 = cy + radius * Math.sin(startAngle);
        const x2 = cx + radius * Math.cos(endAngle);
        const y2 = cy + radius * Math.sin(endAngle);
        const largeArc = sliceAngle > Math.PI ? 1 : 0;

        if (innerRadius > 0) {
          const ix1 = cx + innerRadius * Math.cos(startAngle);
          const iy1 = cy + innerRadius * Math.sin(startAngle);
          const ix2 = cx + innerRadius * Math.cos(endAngle);
          const iy2 = cy + innerRadius * Math.sin(endAngle);
          svg += `<path d="M${x1},${y1} A${radius},${radius} 0 ${largeArc},1 ${x2},${y2} L${ix2},${iy2} A${innerRadius},${innerRadius} 0 ${largeArc},0 ${ix1},${iy1} Z" fill="${color}" stroke="white" stroke-width="2"/>`;
        } else {
          svg += `<path d="M${cx},${cy} L${x1},${y1} A${radius},${radius} 0 ${largeArc},1 ${x2},${y2} Z" fill="${color}" stroke="white" stroke-width="2"/>`;
        }

        // Label
        const midAngle = startAngle + sliceAngle / 2;
        const labelRadius = radius * 0.7;
        const lx = cx + labelRadius * Math.cos(midAngle);
        const ly = cy + labelRadius * Math.sin(midAngle);
        const pct = Math.round((value / total) * 100);
        if (pct >= 5) {
          svg += `<text x="${lx}" y="${ly}" text-anchor="middle" dominant-baseline="middle" font-size="11" font-weight="bold" font-family="Arial, sans-serif" fill="white">${pct}%</text>`;
        }

        startAngle = endAngle;
      });

      // Legend
      const legendY = chartBottom + 5;
      labels.forEach((label, i) => {
        const lx = 20 + (i % 4) * (width / 4);
        const ly = legendY + Math.floor(i / 4) * 16;
        const color = singleConfig.colors?.[i] || palette[i % palette.length];
        svg += `<rect x="${lx}" y="${ly - 8}" width="10" height="10" fill="${color}" rx="2"/>`;
        svg += `<text x="${lx + 14}" y="${ly}" font-size="10" font-family="Arial, sans-serif" fill="#374151">${escapeXmlSvg(label)}</text>`;
      });
    } else {
      // Bar/Line chart
      const maxVal = Math.max(...values, 1);
      const minVal = Math.min(0, Math.min(...values));
      const range = maxVal - minVal || 1;
      const barWidth = chartWidth / labels.length * 0.7;
      const barGap = chartWidth / labels.length * 0.3;

      // Grid lines
      for (let i = 0; i <= 5; i++) {
        const y = chartTop + (chartHeight / 5) * i;
        svg += `<line x1="${chartLeft}" y1="${y}" x2="${chartRight}" y2="${y}" stroke="#e5e7eb" stroke-width="1"/>`;
        const val = maxVal - (range / 5) * i;
        const label = singleConfig.currency ? '$' + Math.round(val).toLocaleString() : Math.round(val).toString();
        svg += `<text x="${chartLeft - 8}" y="${y + 4}" text-anchor="end" font-size="10" font-family="Arial, sans-serif" fill="#6B7280">${label}</text>`;
      }

      // Bars or line
      if (type === 'line' || type === 'area') {
        let pathD = '';
        values.forEach((value, i) => {
          const x = chartLeft + (chartWidth / (values.length - 1 || 1)) * i;
          const y = chartBottom - ((value - minVal) / range) * chartHeight;
          pathD += (i === 0 ? 'M' : 'L') + `${x},${y} `;
        });
        if (type === 'area') {
          const lastX = chartLeft + chartWidth;
          svg += `<path d="${pathD}L${lastX},${chartBottom} L${chartLeft},${chartBottom} Z" fill="${palette[0]}" fill-opacity="0.15"/>`;
        }
        svg += `<path d="${pathD}" fill="none" stroke="${palette[0]}" stroke-width="2.5"/>`;
        values.forEach((value, i) => {
          const x = chartLeft + (chartWidth / (values.length - 1 || 1)) * i;
          const y = chartBottom - ((value - minVal) / range) * chartHeight;
          svg += `<circle cx="${x}" cy="${y}" r="4" fill="${palette[0]}" stroke="white" stroke-width="2"/>`;
        });
      } else {
        const isH = type === 'horizontalBar';
        values.forEach((value, i) => {
          const color = singleConfig.colors?.[i] || palette[i % palette.length];
          if (isH) {
            const barH = chartHeight / labels.length * 0.7;
            const y = chartTop + (chartHeight / labels.length) * i + barGap / 2;
            const barW = ((value - minVal) / range) * chartWidth;
            svg += `<rect x="${chartLeft}" y="${y}" width="${barW}" height="${barH}" fill="${color}" rx="3"/>`;
            if (singleConfig.showValues) {
              svg += `<text x="${chartLeft + barW + 5}" y="${y + barH / 2 + 4}" font-size="10" font-family="Arial, sans-serif" fill="#374151">${value}</text>`;
            }
          } else {
            const x = chartLeft + (chartWidth / labels.length) * i + barGap / 2;
            const barH = ((value - minVal) / range) * chartHeight;
            const y = chartBottom - barH;
            svg += `<rect x="${x}" y="${y}" width="${barWidth}" height="${barH}" fill="${color}" rx="3"/>`;
            if (singleConfig.showValues) {
              svg += `<text x="${x + barWidth / 2}" y="${y - 5}" text-anchor="middle" font-size="10" font-family="Arial, sans-serif" fill="#374151">${value}</text>`;
            }
          }
        });
      }

      // X axis labels
      labels.forEach((label, i) => {
        const x = chartLeft + (chartWidth / labels.length) * i + (chartWidth / labels.length) / 2;
        svg += `<text x="${x}" y="${chartBottom + 16}" text-anchor="middle" font-size="10" font-family="Arial, sans-serif" fill="#6B7280">${escapeXmlSvg(label.substring(0, 12))}</text>`;
      });

      // Axes
      svg += `<line x1="${chartLeft}" y1="${chartTop}" x2="${chartLeft}" y2="${chartBottom}" stroke="#9CA3AF" stroke-width="1"/>`;
      svg += `<line x1="${chartLeft}" y1="${chartBottom}" x2="${chartRight}" y2="${chartBottom}" stroke="#9CA3AF" stroke-width="1"/>`;
    }
  } else {
    // Multi-series bar chart
    const multiConfig = config as MultiSeriesChartConfig;
    const datasets = multiConfig.datasets;
    const allValues = datasets.flatMap(d => d.values);
    const maxVal = Math.max(...allValues, 1);
    const minVal = Math.min(0, Math.min(...allValues));
    const range = maxVal - minVal || 1;

    // Grid
    for (let i = 0; i <= 5; i++) {
      const y = chartTop + (chartHeight / 5) * i;
      svg += `<line x1="${chartLeft}" y1="${y}" x2="${chartRight}" y2="${y}" stroke="#e5e7eb" stroke-width="1"/>`;
      const val = maxVal - (range / 5) * i;
      svg += `<text x="${chartLeft - 8}" y="${y + 4}" text-anchor="end" font-size="10" font-family="Arial, sans-serif" fill="#6B7280">${Math.round(val)}</text>`;
    }

    const groupWidth = chartWidth / labels.length;
    const barWidth = (groupWidth * 0.8) / datasets.length;

    datasets.forEach((ds, di) => {
      const color = ds.color || palette[di % palette.length];
      ds.values.forEach((value, li) => {
        const x = chartLeft + groupWidth * li + (groupWidth * 0.1) + barWidth * di;
        const barH = ((value - minVal) / range) * chartHeight;
        const y = chartBottom - barH;
        svg += `<rect x="${x}" y="${y}" width="${barWidth - 2}" height="${barH}" fill="${color}" rx="2"/>`;
      });
    });

    // Legend
    datasets.forEach((ds, i) => {
      const color = ds.color || palette[i % palette.length];
      const lx = chartLeft + i * 120;
      svg += `<rect x="${lx}" y="${chartBottom + 24}" width="10" height="10" fill="${color}" rx="2"/>`;
      svg += `<text x="${lx + 14}" y="${chartBottom + 33}" font-size="10" font-family="Arial, sans-serif" fill="#374151">${escapeXmlSvg(ds.name)}</text>`;
    });

    // X axis labels
    labels.forEach((label, i) => {
      const x = chartLeft + groupWidth * i + groupWidth / 2;
      svg += `<text x="${x}" y="${chartBottom + 16}" text-anchor="middle" font-size="10" font-family="Arial, sans-serif" fill="#6B7280">${escapeXmlSvg(label.substring(0, 10))}</text>`;
    });

    svg += `<line x1="${chartLeft}" y1="${chartTop}" x2="${chartLeft}" y2="${chartBottom}" stroke="#9CA3AF" stroke-width="1"/>`;
    svg += `<line x1="${chartLeft}" y1="${chartBottom}" x2="${chartRight}" y2="${chartBottom}" stroke="#9CA3AF" stroke-width="1"/>`;
  }

  svg += '</svg>';

  // Convert SVG to PNG using sharp
  const buffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return buffer;
}

function escapeXmlSvg(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Render chart and return as base64 string
 */
export async function renderChartBase64(config: ChartConfig | MultiSeriesChartConfig): Promise<string> {
  const buffer = await renderChart(config);
  return buffer.toString('base64');
}

// ─── SINGLE SERIES CONFIG ─────────────────────────────────────────────────────

function buildSingleSeriesConfig(config: ChartConfig): any {
  const { type, labels, values, title, colors, showValues, currency, percentage, seriesName, showGrid, fontSize, legendPosition, dataLabels } = config;
  const palette = colors || PALETTES.professional;
  const backgroundColors = labels.map((_, i) => palette[i % palette.length]);

  const isHorizontal = type === 'horizontalBar';
  const chartType = isHorizontal ? 'bar' : type;
  const showGridLines = showGrid !== false;

  const datasets: any[] = [{
    label: seriesName || 'Data',
    data: values,
    backgroundColor: type === 'line' ? palette[0] : (type === 'area' ? hexToRgba(palette[0], 0.3) : backgroundColors),
    borderColor: type === 'line' ? palette[0] : (type === 'bar' ? backgroundColors : (config.borderColor || '#ffffff')),
    borderWidth: type === 'pie' || type === 'doughnut' ? (config.borderWidth || 2) : (config.borderWidth || 1),
    fill: type === 'area',
    tension: type === 'line' ? 0.4 : 0,
    pointBackgroundColor: type === 'line' || type === 'area' ? palette[0] : undefined,
    pointRadius: type === 'line' || type === 'area' ? 4 : 0,
    pointHoverRadius: type === 'line' || type === 'area' ? 6 : 0,
  }];

  const scales: any = {};
  if (type === 'bar' || type === 'line' || type === 'area' || isHorizontal || type === 'scatter') {
    const valueAxis = isHorizontal ? 'x' : 'y';
    const labelAxis = isHorizontal ? 'y' : 'x';

    scales[labelAxis] = {
      grid: { display: showGridLines, color: '#f3f4f6' },
      ticks: { font: { size: fontSize || 11, family: 'Arial, sans-serif' }, color: '#6B7280' },
    };

    scales[valueAxis] = {
      beginAtZero: true,
      grid: { display: showGridLines, color: '#f3f4f6' },
      ticks: {
        font: { size: fontSize || 11, family: 'Arial, sans-serif' },
        color: '#6B7280',
        callback: function (value: any) {
          if (currency) return '$' + Number(value).toLocaleString();
          if (percentage) return value + '%';
          return value;
        }
      }
    };
  }

  return {
    type: chartType,
    data: { labels, datasets },
    options: {
      responsive: false,
      animation: false,
      plugins: {
        title: {
          display: !!title,
          text: title || '',
          font: { size: fontSize ? fontSize + 4 : 16, weight: 'bold', family: 'Arial, sans-serif' },
          padding: { bottom: 16 },
          color: '#1F2937',
        },
        legend: {
          display: config.showLegend !== false && (type === 'pie' || type === 'doughnut'),
          position: legendPosition || 'right',
          labels: { font: { size: fontSize || 11, family: 'Arial, sans-serif' }, padding: 12, color: '#374151' },
        },
        datalabels: (showValues || dataLabels) ? {
          display: true,
          color: type === 'pie' || type === 'doughnut' ? '#ffffff' : '#374151',
          font: { size: fontSize || 11, weight: 'bold', family: 'Arial, sans-serif' },
          formatter: (value: number) => {
            if (currency) return '$' + value.toLocaleString();
            if (percentage) return value + '%';
            return value;
          },
        } : { display: false },
      },
      indexAxis: isHorizontal ? 'y' : 'x',
      scales: Object.keys(scales).length > 0 ? scales : undefined,
    },
  };
}

// ─── MULTI SERIES CONFIG ──────────────────────────────────────────────────────

function buildMultiSeriesConfig(config: MultiSeriesChartConfig): any {
  const { type, labels, datasets, title, currency, showGrid, legendPosition } = config;
  const showGridLines = showGrid !== false;

  const chartType = type === 'area' ? 'line' : (type === 'stacked-bar' ? 'bar' : type);
  const isStacked = type === 'stacked-bar' || type === 'area';

  const chartDatasets = datasets.map((ds, i) => {
    const color = ds.color || PALETTES.professional[i % PALETTES.professional.length];
    const dsType = ds.type || (type === 'combo' ? (i === 0 ? 'bar' : 'line') : chartType);
    return {
      label: ds.name,
      data: ds.values,
      type: dsType,
      backgroundColor: dsType === 'line' ? 'transparent' : (isStacked ? color : hexToRgba(color, 0.8)),
      borderColor: color,
      borderWidth: dsType === 'line' ? 2.5 : 1,
      fill: ds.fill || (type === 'area' ? 'origin' : false),
      tension: 0.4,
      pointRadius: dsType === 'line' ? 4 : 0,
      pointBackgroundColor: color,
      order: dsType === 'line' ? 1 : 2,
    };
  });

  return {
    type: chartType,
    data: { labels, datasets: chartDatasets },
    options: {
      responsive: false,
      animation: false,
      plugins: {
        title: {
          display: !!title,
          text: title || '',
          font: { size: 16, weight: 'bold', family: 'Arial, sans-serif' },
          padding: { bottom: 16 },
          color: '#1F2937',
        },
        legend: {
          display: config.showLegend !== false,
          position: legendPosition || 'top',
          labels: { font: { size: 11, family: 'Arial, sans-serif' }, padding: 12, color: '#374151' },
        },
      },
      scales: {
        x: {
          stacked: isStacked,
          grid: { display: showGridLines, color: '#f3f4f6' },
          ticks: { font: { size: 11, family: 'Arial, sans-serif' }, color: '#6B7280' },
        },
        y: {
          stacked: isStacked,
          beginAtZero: true,
          grid: { display: showGridLines, color: '#f3f4f6' },
          ticks: {
            font: { size: 11, family: 'Arial, sans-serif' },
            color: '#6B7280',
            callback: (value: any) => currency ? '$' + Number(value).toLocaleString() : value,
          },
        },
      },
    },
  };
}

function hexToRgba(hex: string, alpha: number): string {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// ─── AUTO-CHART DETECTION ─────────────────────────────────────────────────────

export function detectChartType(labels: string[], values: number[]): 'pie' | 'bar' | 'line' {
  const labelCount = labels.length;
  const allNumeric = values.every(v => typeof v === 'number' && !isNaN(v));

  if (!allNumeric) return 'bar';

  if (labelCount <= 6 && values.every(v => v >= 0)) {
    const sum = values.reduce((a, b) => a + b, 0);
    const avg = sum / values.length;
    const hasVariation = values.some(v => Math.abs(v - avg) > avg * 0.1);
    if (hasVariation) return 'pie';
  }

  const timePatterns = /jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|q[1-4]|20\d{2}/i;
  const timeMatches = labels.filter(l => timePatterns.test(l)).length;
  if (timeMatches >= labelCount * 0.5) return 'line';

  return 'bar';
}

export function buildChartFromData(
  headers: string[],
  rows: any[][],
  options?: {
    type?: 'pie' | 'bar' | 'line';
    labelCol?: number;
    valueCol?: number;
    title?: string;
  }
): ChartConfig {
  const labelCol = options?.labelCol ?? 0;
  const valueCol = options?.valueCol ?? 1;

  const labels = rows.map(row => String(row[labelCol] || ''));
  const values = rows.map(row => {
    const val = row[valueCol];
    if (typeof val === 'number') return val;
    const parsed = parseFloat(String(val || '0').replace(/[^0-9.-]/g, ''));
    return isNaN(parsed) ? 0 : parsed;
  });

  const type = options?.type || detectChartType(labels, values);

  return {
    type,
    labels,
    values,
    title: options?.title || headers[valueCol] || 'Chart',
    seriesName: headers[valueCol] || 'Value',
  };
}
