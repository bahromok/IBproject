/**
 * Style Engine - Reusable style templates for professional documents
 */

// ═══════════════════════════════════════════════════════════════════════════════════
// STYLE DEFINITIONS
// ═══════════════════════════════════════════════════════════════════════════════════

export const STYLE_TEMPLATES = {
  // Professional business styles
  professional: {
    header: {
      bgColor: '1E40AF',
      fontColor: 'FFFFFF',
      bold: true,
      fontSize: 12,
      alignment: 'center'
    },
    subheader: {
      bgColor: '3B82F6',
      fontColor: 'FFFFFF',
      bold: true,
      fontSize: 11
    },
    data: {
      fontSize: 11,
      borderColor: 'E5E7EB'
    },
    alternateRow: {
      bgColor: 'F3F4F6'
    }
  },
  
  // Financial/report styles
  financial: {
    header: {
      bgColor: '2D3748',
      fontColor: 'FFFFFF',
      bold: true,
      fontSize: 11
    },
    currency: {
      numberFormat: '$#,##0.00'
    },
    percentage: {
      numberFormat: '0.00%'
    },
    positive: {
      fontColor: '059669',
      bold: true
    },
    negative: {
      fontColor: 'DC2626',
      bold: true
    }
  },
  
  // Modern/minimal styles
  modern: {
    header: {
      bgColor: '111827',
      fontColor: 'F9FAFB',
      bold: true,
      fontSize: 11
    },
    accent: {
      bgColor: '10B981',
      fontColor: 'FFFFFF'
    },
    highlight: {
      bgColor: 'FEF3C7',
      fontColor: '92400E'
    }
  },
  
  // Colorful/vibrant styles
  vibrant: {
    header: {
      bgColor: '7C3AED',
      fontColor: 'FFFFFF',
      bold: true
    },
    accent1: {
      bgColor: 'EC4899',
      fontColor: 'FFFFFF'
    },
    accent2: {
      bgColor: 'F59E0B',
      fontColor: 'FFFFFF'
    },
    accent3: {
      bgColor: '10B981',
      fontColor: 'FFFFFF'
    }
  },
  
  // Word document styles
  document: {
    title: {
      font: 'Calibri',
      size: 36,
      color: '1F2937',
      bold: true,
      alignment: 'center'
    },
    heading1: {
      font: 'Calibri',
      size: 28,
      color: '1E40AF',
      bold: true
    },
    heading2: {
      font: 'Calibri',
      size: 24,
      color: '374151',
      bold: true
    },
    body: {
      font: 'Calibri',
      size: 22,
      color: '374151',
      lineSpacing: 1.15
    },
    bullet: {
      font: 'Calibri',
      size: 22,
      color: '4B5563'
    }
  }
};

// ═══════════════════════════════════════════════════════════════════════════════════
// STYLE HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════════

/**
 * Get a style template by name
 */
export function getStyleTemplate(name: string): any {
  return STYLE_TEMPLATES[name as keyof typeof STYLE_TEMPLATES] || STYLE_TEMPLATES.professional;
}

/**
 * Apply a style template to Excel operations
 */
export function applyExcelStyle(styleName: string, range: string): any[] {
  const style = getStyleTemplate(styleName);
  const operations: any[] = [];
  
  if (style.header) {
    operations.push({
      type: 'set_range_style',
      range: range.split(':')[0] + ':' + range.split(':')[0].replace(/[0-9]+/, '1'),
      style: style.header
    });
  }
  
  return operations;
}

/**
 * Generate Excel style operations for a table
 */
export function generateTableStyleOperations(
  startCell: string,
  endCell: string,
  styleName: string = 'professional'
): any[] {
  const style = getStyleTemplate(styleName);
  const operations: any[] = [];
  
  // Header row styling
  operations.push({
    type: 'set_range_style',
    range: startCell + ':' + startCell.replace(/[0-9]+/, '1'),
    style: style.header
  });
  
  // Alternate row styling
  const startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
  const endRow = parseInt(endCell.replace(/[A-Z]/g, ''));
  
  for (let row = startRow + 1; row <= endRow; row++) {
    if (row % 2 === 0) {
      const cellRef = startCell.replace(/[0-9]+/, '') + row;
      operations.push({
        type: 'set_cell_style',
        cell: cellRef,
        style: { bgColor: style.alternateRow?.bgColor || 'F3F4F6' }
      });
    }
  }
  
  return operations;
}

/**
 * Generate Word style operations for a section
 */
export function generateWordStyleOperations(
  styleName: string = 'document'
): any[] {
  const style = getStyleTemplate(styleName);
  return []; // Word styling is applied during creation
}

/**
 * Get color by name
 */
export function getColor(name: string): string {
  const colors: Record<string, string> = {
    red: 'EF4444',
    green: '10B981',
    blue: '3B82F6',
    yellow: 'F59E0B',
    purple: '7C3AED',
    pink: 'EC4899',
    orange: 'F97316',
    gray: '6B7280',
    dark: '1F2937',
    white: 'FFFFFF',
    black: '000000'
  };
  return colors[name.toLowerCase()] || name;
}

/**
 * Get number format by type
 */
export function getNumberFormat(type: string): string {
  const formats: Record<string, string> = {
    currency: '$#,##0.00',
    percentage: '0.00%',
    integer: '#,##0',
    decimal: '#,##0.00',
    date: 'MM/DD/YYYY',
    time: 'HH:MM',
    text: '@'
  };
  return formats[type.toLowerCase()] || type;
}
