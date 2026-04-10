import ExcelJS from 'exceljs';
import { readDocx, countImagesInDocx, countTablesInDocx } from './doc-reader';
import { readFileBuffer, fileExists } from './file-storage';

/**
 * Extract Excel file structure for AI understanding — ENHANCED
 */
export async function extractExcelStructure(filename: string): Promise<{
  sheets: Array<{
    name: string;
    headers: string[];
    rowCount: number;
    columnCount: number;
    data: any[][];
    formulas: Record<string, string>;
    styles: {
      headerRow: boolean;
      frozenPanes: boolean;
      autoFilter: boolean;
      imageCount: number;
      conditionalFormats: number;
      dataValidations: number;
      mergedCells: string[];
      columnWidths: number[];
    };
  }>;
  summary: string;
}> {
  if (!fileExists(filename)) {
    throw new Error('File not found: ' + filename);
  }

  const buffer = readFileBuffer(filename);
  if (!buffer) throw new Error('Could not read file');

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);

  const sheets = wb.worksheets.map(ws => {
    const headers: string[] = [];
    const data: any[][] = [];
    const formulas: Record<string, string> = {};

    // Extract headers
    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      headers.push(String(cell.value || ''));
    });

    // Extract data (skip header, up to 20 rows)
    ws.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && rowNumber <= 21) {
        const rowData: any[] = [];
        row.eachCell((cell, colNumber) => {
          rowData.push(cell.value);
          if (typeof cell.value === 'object' && cell.value && 'formula' in cell.value) {
            formulas[`${String.fromCharCode(64 + colNumber)}${rowNumber}`] = '=' + (cell.value as any).formula;
          }
        });
        data.push(rowData);
      }
    });

    // Image count
    const imageCount = ((ws as any)._images || []).length;

    // Conditional formatting count
    const condFormatCount = ((ws as any)._conditionalFormatting || []).length || 0;

    // Merged cells
    const mergedCells: string[] = [];
    try {
      for (const range of (ws as any).merges || []) {
        mergedCells.push(typeof range === 'string' ? range : String(range));
      }
    } catch { }

    // Column widths
    const columnWidths: number[] = [];
    for (let c = 1; c <= ws.columnCount; c++) {
      try {
        columnWidths.push(ws.getColumn(c).width || 0);
      } catch {
        columnWidths.push(0);
      }
    }

    return {
      name: ws.name,
      headers,
      rowCount: ws.rowCount - 1,
      columnCount: ws.columnCount,
      data,
      formulas,
      styles: {
        headerRow: headers.length > 0,
        frozenPanes: ws.views && ws.views.length > 0 && ws.views[0].state === 'frozen',
        autoFilter: !!ws.autoFilter,
        imageCount,
        conditionalFormats: condFormatCount,
        dataValidations: 0, // ExcelJS doesn't expose count easily
        mergedCells,
        columnWidths: columnWidths.slice(0, 10),
      }
    };
  });

  const summary = sheets.map(s => {
    const parts = [`Sheet "${s.name}": ${s.rowCount} rows x ${s.columnCount} cols`];
    if (s.headers.length > 0) parts.push('Headers: ' + s.headers.join(', '));
    if (s.styles.imageCount > 0) parts.push(s.styles.imageCount + ' image(s)');
    if (s.styles.frozenPanes) parts.push('frozen panes');
    if (s.styles.autoFilter) parts.push('auto-filter');
    if (s.styles.mergedCells.length > 0) parts.push(s.styles.mergedCells.length + ' merged range(s)');
    if (Object.keys(s.formulas).length > 0) parts.push(Object.keys(s.formulas).length + ' formula(s)');
    if (s.data.length > 0) parts.push('Data preview: ' + s.data.slice(0, 3).map(r => r.join(',')).join(' | '));
    return parts.join('. ');
  }).join('\n');

  return { sheets, summary };
}

/**
 * Extract Word document structure for AI understanding — ENHANCED
 */
export async function extractWordStructure(filename: string): Promise<{
  title: string;
  elements: Array<{
    type: string;
    content: string;
    level?: number;
    items?: string[];
    headers?: string[];
    rows?: string[][];
  }>;
  wordCount: number;
  imageCount: number;
  tableCount: number;
  structure: string;
}> {
  if (!fileExists(filename)) {
    throw new Error('File not found: ' + filename);
  }

  const model = await readDocx(filename);

  // Count images and tables via XML analysis
  let imageCount = 0;
  let tableCount = 0;
  try {
    imageCount = await countImagesInDocx(filename);
    tableCount = await countTablesInDocx(filename);
  } catch { }

  const elements = model.elements.map((el, i) => {
    const eid = i + 1;
    switch (el.type) {
      case 'heading':
        return { type: 'heading', content: el.text || '', level: el.level, id: '@h' + eid };
      case 'paragraph':
        return { type: 'paragraph', content: el.text || '', id: '@p' + eid };
      case 'list':
        return { type: 'list', content: 'Bullet list', items: el.items, id: '@p' + eid };
      case 'table':
        return {
          type: 'table',
          content: 'Table: ' + (el.headers?.join(' | ') || ''),
          headers: el.headers,
          rows: el.rows?.slice(0, 5),
          id: '@t' + eid,
          rowCount: el.rows?.length || 0,
        };
      default:
        return { type: el.type, content: el.text || '', id: '@p' + eid };
    }
  });

  const structure = elements.map((el) => {
    const prefix = el.id;
    if (el.type === 'heading') return `  ${prefix} [H${el.level}] ${el.content}`;
    if (el.type === 'table') return `  ${prefix} [TABLE ${el.rowCount || 0} rows] ${el.content}`;
    if (el.type === 'list') return `  ${prefix} [LIST ${(el.items || []).length} items] ${(el.items || []).slice(0, 3).join(', ')}...`;
    return `  ${prefix} [P] ${el.content.substring(0, 60)}${el.content.length > 60 ? '...' : ''}`;
  }).join('\n');

  const extra: string[] = [];
  if (imageCount > 0) extra.push(imageCount + ' image(s)');
  if (tableCount > 0) extra.push(tableCount + ' table(s)');

  return {
    title: model.title,
    elements: elements as any,
    wordCount: model.wordCount,
    imageCount,
    tableCount,
    structure: `Document: ${filename}\nTitle: ${model.title}\nWords: ${model.wordCount}${extra.length ? '\nContains: ' + extra.join(', ') : ''}\nElements (${elements.length}):\n${structure}`
  };
}

/**
 * Get simplified file summary for AI context
 */
export async function getFileSummary(filename: string): Promise<string> {
  if (!fileExists(filename)) {
    return `File "${filename}" does not exist yet.`;
  }

  try {
    if (filename.endsWith('.xlsx')) {
      const structure = await extractExcelStructure(filename);
      return structure.summary;
    } else if (filename.endsWith('.docx')) {
      const structure = await extractWordStructure(filename);
      return structure.structure;
    }
  } catch (e: any) {
    return `Could not analyze ${filename}: ${e.message}`;
  }

  return `File: ${filename}`;
}
