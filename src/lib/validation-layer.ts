import ExcelJS from 'exceljs';
import { readFileBuffer, fileExists, writeFileBuffer } from './file-storage';
import { readDocx } from './doc-reader';

export interface ValidationIssue {
  type: 'error' | 'warning' | 'info';
  location: string;
  message: string;
  fix?: string;
}

/**
 * Validate an Excel file and return issues
 */
export async function validateExcel(filename: string): Promise<ValidationIssue[]> {
  const issues: ValidationIssue[] = [];
  
  if (!fileExists(filename)) {
    issues.push({ type: 'error', location: filename, message: 'File not found' });
    return issues;
  }
  
  const buffer = readFileBuffer(filename);
  if (!buffer) {
    issues.push({ type: 'error', location: filename, message: 'Cannot read file' });
    return issues;
  }
  
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer as any);
    
    for (const ws of wb.worksheets) {
      // Check for empty sheets
      if (ws.rowCount === 0) {
        issues.push({ type: 'warning', location: ws.name, message: 'Sheet is empty' });
        continue;
      }
      
      // Check for missing headers
      const headerRow = ws.getRow(1);
      let hasHeaders = false;
      headerRow.eachCell(cell => {
        if (cell.value) hasHeaders = true;
      });
      if (!hasHeaders) {
        issues.push({ type: 'warning', location: ws.name, message: 'No headers found in first row' });
      }
      
      // Check for broken formulas
      ws.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          if (typeof cell.value === 'object' && cell.value && 'formula' in cell.value) {
            const formula = (cell.value as any).formula;
            if (!formula || formula.trim() === '') {
              issues.push({
                type: 'error',
                location: `${ws.name}!${cell.address}`,
                message: 'Empty formula',
                fix: 'Remove or complete the formula'
              });
            }
          }
        });
      });
      
      // Check for inconsistent column widths
      if (ws.columns) {
        const widths = ws.columns.map(c => c.width || 0);
        const avgWidth = widths.reduce((a, b) => a + b, 0) / widths.length;
        widths.forEach((w, i) => {
          if (w && w < 5) {
            issues.push({
              type: 'warning',
              location: `${ws.name} column ${i + 1}`,
              message: 'Column too narrow (may truncate content)'
            });
          }
        });
      }
    }
    
    // Check for no sheets
    if (wb.worksheets.length === 0) {
      issues.push({ type: 'error', location: filename, message: 'No worksheets found' });
    }
    
  } catch (e: any) {
    issues.push({ type: 'error', location: filename, message: 'Invalid Excel format: ' + e.message });
  }
  
  return issues;
}

/**
 * Validate a Word file and return issues
 */
export async function validateWord(filename: string): Promise<ValidationIssue[]> {
  const issues: ValidationIssue[] = [];
  
  if (!fileExists(filename)) {
    issues.push({ type: 'error', location: filename, message: 'File not found' });
    return issues;
  }
  
  try {
    const model = await readDocx(filename);
    
    // Check for empty document
    if (model.elements.length === 0) {
      issues.push({ type: 'warning', location: filename, message: 'Document is empty' });
    }
    
    // Check for missing title
    const hasTitle = model.elements.some(e => e.type === 'heading' && e.level === 1);
    if (!hasTitle && model.elements.length > 0) {
      issues.push({ type: 'info', location: filename, message: 'No main title (H1) found' });
    }
    
    // Check for very long paragraphs (hard to read)
    model.elements.forEach((el, i) => {
      if (el.type === 'paragraph' && el.text && el.text.length > 500) {
        issues.push({
          type: 'warning',
          location: `Element ${i + 1}`,
          message: 'Very long paragraph (consider breaking it up)'
        });
      }
    });
    
    // Check for empty tables
    model.elements.forEach((el, i) => {
      if (el.type === 'table') {
        if (!el.rows || el.rows.length === 0) {
          issues.push({
            type: 'warning',
            location: `Table ${i + 1}`,
            message: 'Table has no data rows'
          });
        }
      }
    });
    
  } catch (e: any) {
    issues.push({ type: 'error', location: filename, message: 'Cannot read document: ' + e.message });
  }
  
  return issues;
}

/**
 * Validate any file
 */
export async function validateFile(filename: string): Promise<ValidationIssue[]> {
  if (filename.endsWith('.xlsx')) {
    return validateExcel(filename);
  } else if (filename.endsWith('.docx')) {
    return validateWord(filename);
  }
  return [{ type: 'info', location: filename, message: 'Unknown file type' }];
}

/**
 * Attempt to fix issues in a file
 */
export async function fixIssues(filename: string, issues: ValidationIssue[]): Promise<string[]> {
  const fixes: string[] = [];
  
  if (filename.endsWith('.xlsx')) {
    const buffer = readFileBuffer(filename);
    if (!buffer) return fixes;
    
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer as any);
      
      for (const issue of issues) {
        if (issue.message === 'Column too narrow (may truncate content)') {
          // Auto-adjust column widths
          wb.worksheets.forEach(ws => {
            ws.columns?.forEach(col => {
              if (col.width && col.width < 10) {
                col.width = 15;
              }
            });
          });
          fixes.push('Adjusted narrow column widths');
        }
      }
      
      if (fixes.length > 0) {
        const output = await wb.xlsx.writeBuffer();
        writeFileBuffer(filename, Buffer.from(output as any));
      }
    } catch (e) {
      // Could not fix
    }
  }
  
  return fixes;
}

/**
 * Get validation summary
 */
export function getValidationSummary(issues: ValidationIssue[]): string {
  if (issues.length === 0) {
    return 'No issues found.';
  }
  
  const errors = issues.filter(i => i.type === 'error').length;
  const warnings = issues.filter(i => i.type === 'warning').length;
  const infos = issues.filter(i => i.type === 'info').length;
  
  let summary = '';
  if (errors > 0) summary += `${errors} error(s), `;
  if (warnings > 0) summary += `${warnings} warning(s), `;
  if (infos > 0) summary += `${infos} info`;
  
  return summary.trim().replace(/,$/, '');
}
