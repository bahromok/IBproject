import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, ShadingType, BorderStyle } from 'docx';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';
import { ensureStorageDir, getFilePath, readFileBuffer, writeFileBuffer, fileExists, deleteFile as deleteStorageFile, listFiles as listStorageFiles } from './file-storage';
import {
  readDocx, replaceTextInDocx, addToDocx,
  headingXml, paragraphXml, bulletListXml, tableXml, styledParagraphXml, styledTableXml, coloredHeadingXml,
  addTableRowToDocx, updateTableCellInDocx, deleteTableRowFromDocx,
  applyTextStyle, deleteElementByContent,
  addHeaderToDocx, addFooterToDocx,
  setMargins, setOrientation, setDocumentFont, setLineSpacing,
  embedImageInDocx, embedChartInDocx,
  addHyperlinkToDocx,
  deleteTableFromDocx, deleteImageFromDocx,
  removeHeaderFromDocx, removeFooterFromDocx,
  formatTableCellInDocx, addColumnToTable, deleteColumnFromTable,
  countImagesInDocx, countTablesInDocx,
  escapeXml,
  addTextBoxToDocx, addPageBorderToDocx, addSectionBreakToDocx,
  setColumnsInDocx, addWatermarkToDocx, addDropCapParagraphToDocx,
  addTabStopParagraphToDocx, addFormattedPageNumbersToDocx,
  setTableWidthToDocx, setTableColumnWidthsToDocx,
  embedImagePositionedInDocx, addHighlightParagraphToDocx,
  setParagraphSpacingToDocx, setFirstLineIndentToDocx,
  clearAllContentFromDocx, addColumnBreakToDocx,
  // NEW: indexed paragraph access
  getIndexedParagraphs, insertBeforeIndex, insertAfterIndex,
  replaceAtIndex, deleteAtIndex, formatAtIndex, replaceTextAtIndex,
  duplicateAtIndex, moveBlockToIndex,
  getDocumentXml, setDocumentXml,
  mergeTableCellsInDocx,
  // NEW: full Excel read + bulk update
  readSpreadsheetFull, bulkUpdateCells,
} from './doc-reader';
import { renderChart, renderChartBase64, buildChartFromData, detectChartType, type ChartConfig } from './chart-engine';

// ═══════════════════════════════════════════════════════════════════════════════════
// TOOL DEFINITIONS
// ═══════════════════════════════════════════════════════════════════════════════════

export const TOOL_DEFINITIONS = `You are OfficeAI — a powerful document AI with FULL surgical control over every Word and Excel file. Always return valid JSON: {"thinking":"...","tool_calls":[{...}],"message":"..."}

═══════════════════════════════════════════════════════════════
THE 14 TOOLS  (use in "tool_calls" array only)
═══════════════════════════════════════════════════════════════

─── WORD DOCUMENTS ────────────────────────────────────────────

1. read_document   { "tool":"read_document", "filename":"f.docx" }
   Returns full text, word count, all elements.

2. get_paragraph_index   { "tool":"get_paragraph_index", "filename":"f.docx" }
   ★ USE THIS BEFORE EDITING. Returns every block with its index number:
   [0] [H1] Title text
   [1] [P]  Paragraph text...
   [2] [TABLE 3 rows] Col1 | Col2
   Then use those index numbers for precise insert/replace/delete/format.

3. create_document   { "tool":"create_document", "filename":"r.docx", "title":"T", "sections":[{heading,content,bullets,numberedItems,table:{headers,rows}}] }

4. edit_document   { "tool":"edit_document", "filename":"f.docx", "operations":[...] }

─── EXCEL SPREADSHEETS ────────────────────────────────────────

5. read_spreadsheet_full   { "tool":"read_spreadsheet_full", "filename":"f.xlsx", "sheet":"Sheet1" }
   ★ Returns ALL cells with values, formulas, and styles — so you know exactly what's there before editing.

6. create_spreadsheet   { "tool":"create_spreadsheet", "filename":"f.xlsx", "sheets":[{name,headers,data,formulas,styles}] }

7. edit_spreadsheet   { "tool":"edit_spreadsheet", "filename":"f.xlsx", "sheet":"Sheet1", "operations":[...] }

8. bulk_update_cells   { "tool":"bulk_update_cells", "filename":"f.xlsx", "sheet":"Sheet1",
     "updates":[{"cell":"A1","value":"hello"},{"cell":"B2","formula":"=SUM(A1:A10)","style":{"bold":true,"bgColor":"FFD700"}}] }
   ★ Most efficient way to update many cells at once with values, formulas, AND styling in one call.

─── FILE MANAGEMENT ───────────────────────────────────────────

9.  analyze_file    { "tool":"analyze_file", "filename":"f.xlsx" }
10. list_files      { "tool":"list_files" }
11. delete_file     { "tool":"delete_file", "filename":"f.docx" }
12. rename_file     { "tool":"rename_file", "filename":"old.docx", "new_filename":"new.docx" }

─── ADVANCED ──────────────────────────────────────────────────

13. get_document_xml   { "tool":"get_document_xml", "filename":"f.docx" }
    Returns raw Word XML body. Use when you need to understand exact structure or troubleshoot.

14. set_document_xml   { "tool":"set_document_xml", "filename":"f.docx", "xml":"<w:body>...</w:body>" }
    ★ NUCLEAR OPTION: Replace entire document body with provided XML. Full surgical control.

═══════════════════════════════════════════════════════════════
WORD OPERATIONS  (inside edit_document "operations" array)
═══════════════════════════════════════════════════════════════

─── PRECISION INDEXED OPERATIONS (use after get_paragraph_index) ───
insert_before_index(index, content_xml)         Insert XML before block at index
insert_after_index(index, content_xml)          Insert XML after block at index
replace_at_index(index, content_xml)            Replace entire block at index with new XML
replace_text_at_index(index, text)              Replace text of paragraph at index, preserve formatting
delete_at_index(index)                          Delete block at index
format_at_index(index, bold?, italic?, underline?, color?, font?, fontSize?, alignment?,
  headingLevel?, indent?, spaceBefore?, spaceAfter?, lineSpacing?, highlight?)
duplicate_at_index(index)                       Duplicate block at index
move_to_index(source_index, dest_index)         Move block from source to dest position

─── TEXT OPERATIONS ───
replace_text(find, replace, caseSensitive?)     Replace text across document (handles split runs)
set_text_style(find, bold?, italic?, underline?, color?, font?, fontSize?, strikethrough?)
delete_element(search)                          Delete paragraphs containing text
clear_content(search)                           Clear (blank out) matching text

─── ADD CONTENT ───
add_heading(heading, level?, color?)
add_paragraph(content, bold?, italic?, color?, fontSize?, alignment?, font?, underline?)
add_bullet_list(items[])
add_numbered_list(items[], startNumber?)
add_section(heading, content?, bullets?, numberedItems?, table?)
add_page_break
add_separator
add_highlight_paragraph(text, highlight?, bold?, color?)

─── TABLES ───
add_table(headers[], rows[][])
add_table_row(table_index, data[])
update_table_cell(table_index, row, col, value)
delete_table_row(table_index, row)
delete_table_column(table_index, col)
delete_table(table_index)
format_table_cell(table_index, row, col, bg?, bold?, color?, font?, fontSize?, align?)
add_table_column(table_index, header, values[])
merge_table_cells(table_index, startRow, startCol, endRow, endCol)
set_table_width(table_index, width, widthType?, alignment?)
set_table_column_widths(table_index, widths[])
count_tables

─── LAYOUT / PAGE ───
set_margins(top, bottom, left, right)           All in twips (1440 = 1 inch)
set_orientation(portrait|landscape)
set_columns(count, spacing?, separator?)
set_font(font)
set_line_spacing(spacing)                       1.0=single, 1.5, 2.0=double
set_paragraph_spacing(before, after, lineSpacing?)
set_first_line_indent(twips)
add_section_break(nextPage|continuous|evenPage|oddPage)
add_column_break
clear_all_content                               Wipe document clean

─── HEADER / FOOTER / PAGE NUMBERS ───
add_header(text)
add_footer(text)
add_page_number
add_formatted_page_numbers(format?, alignment?, showTotal?, font?, fontSize?, color?)
remove_header
remove_footer

─── GRAPHICS & MEDIA ───
add_chart(chart_type, labels[], values[], title?, colors[]?, width?, height?, showLegend?, showValues?, currency?)
  chart_type = pie|bar|line|doughnut|horizontalBar|area|radar|scatter
add_image(image_base64, width?, height?, align?, wrapStyle?)
add_image_positioned(image_base64, width?, height?, x?, y?, wrapStyle?)
delete_image(image_index)
count_images
add_text_box(text, width?, height?, fillColor?, borderColor?, fontSize?, bold?, color?, alignment?, x?, y?)
add_watermark(text, color?, fontSize?, font?)
add_page_border(style?, color?, size?)
add_drop_cap(text, lines?, font?, color?)

─── LINKS & NAVIGATION ───
add_hyperlink(text, url, color?, underline?)
add_table_of_contents
add_tab_stop_paragraph(text, tabStops[], fontSize?, bold?, color?)
add_bookmark(name, text)

═══════════════════════════════════════════════════════════════
EXCEL OPERATIONS  (inside edit_spreadsheet "operations" array)
═══════════════════════════════════════════════════════════════

─── DATA ───
add_row(data[])
add_multiple_rows(rows[][])
update_cell(cell, value)
set_formula(cell, formula)                      formula must start with =
add_column(header, values[])
insert_row(row, data[])
insert_column(column, header, values[])
delete_row(row)
delete_column(column)
replace_text(find, replace)
copy_range(source, destination)

─── STYLING ───
set_cell_style(cell, bold?, italic?, fontColor?, bgColor?, fontSize?, fontName?,
  borderColor?, borderStyle?, alignment?, wrapText?)
set_range_style(range, bold?, italic?, fontColor?, bgColor?, fontSize?, alignment?, wrapText?)
set_number_format(cell, format)                 "$#,##0.00" | "0.00%" | "#,##0" | "mm/dd/yyyy"
set_column_width(column, width)
set_row_height(row, height)
set_alignment(cell, horizontal, vertical?, wrapText?)
set_borders(range, style, sides?, color?)
set_tab_color(color)

─── LAYOUT ───
freeze_panes(cell)
unfreeze_panes
merge_cells(range)
unmerge_cells(range)
set_auto_filter(range)
remove_auto_filter
set_print_area(range)
set_page_setup(orientation?, paperSize?, fitToPage?, fitToWidth?, fitToHeight?, margins?)

─── SHEETS ───
add_sheet(name)
rename_sheet(old_name, new_name)
delete_sheet(name)
copy_sheet(source, name)

─── CHARTS ───
add_chart(chart_type, title, labels?, values?, label_col?, value_col?, from_row?, to_row?,
  width?, height?, chart_row?, chart_col?, colors[]?, showLegend?)
delete_chart(index)

─── CONDITIONAL FORMATTING ───
add_conditional_format(range, condition, value, style:{fontColor?,bgColor?})
remove_conditional_format(range)

─── DATA VALIDATION ───
add_data_validation(range, validation_type, values[]?, min?, max?, operator?, allow_blank?)
remove_data_validation(range)

─── OTHER ───
add_comment(cell, text)
remove_comment(cell)
add_hyperlink_cell(cell, url, text?)
protect_sheet(password?)
unprotect_sheet
add_named_range(range, name)
sort(column, order?)
group_rows(start, end, level?)
ungroup_rows(start, end)
clear_range(range, mode?)
add_image(image_base64, width?, height?, row?, col?)

═══════════════════════════════════════════════════════════════
WORKFLOW RULES — FOLLOW THESE
═══════════════════════════════════════════════════════════════

★ ALWAYS READ BEFORE EDITING:
  - For Word: call get_paragraph_index FIRST. Use index numbers for precise edits.
  - For Excel: call read_spreadsheet_full or analyze_file FIRST to see actual data.
  - Never guess at structure — always verify with a read tool first.

★ FOR WORD PRECISION EDITING:
  1. get_paragraph_index → see all blocks with numbers
  2. Use insert_before_index/replace_at_index/delete_at_index/format_at_index
  3. Never use replace_text for structural changes — use indexed operations

★ FOR EXCEL PRECISION EDITING:
  1. read_spreadsheet_full → see every cell, formula, and value
  2. Use bulk_update_cells for updating many cells — most efficient
  3. Use set_formula for formulas, update_cell for values

★ MULTI-STEP EXECUTION:
  - Chain multiple tool_calls in one response
  - Read → Edit → Read again to verify if needed
  - Use corrections to fix issues automatically

★ COLORS: 6-digit hex without #: FF0000=red, 00FF00=green, FFD700=gold,
  1E40AF=blue, 2D3748=dark, FFFFFF=white, 000000=black
★ FORMULAS: always prefix with =  e.g. "=SUM(A2:A10)"
★ TWIPS: 1440 per inch. Common margins: 1440=1in, 720=0.5in, 1080=0.75in
★ Table/image indices are 0-based
★ Paragraph indices from get_paragraph_index are 0-based
★ Excel column widths: 10-20 typical, 25+ for long text`;

// ═══════════════════════════════════════════════════════════════════════════════════
// INTERFACES
// ═══════════════════════════════════════════════════════════════════════════════════

export interface ToolResult {
  success: boolean;
  message: string;
  filename?: string;
  preview?: string;
  progress?: number;
  data?: any;
}

// ═══════════════════════════════════════════════════════════════════════════════════
// TOOL EXECUTION ROUTER
// ═══════════════════════════════════════════════════════════════════════════════════

// Operations that belong inside edit_document
const WORD_OPS = new Set([
  'replace_text','add_heading','add_paragraph','add_bullet_list','add_numbered_list',
  'add_table','add_table_row','update_table_cell','delete_table_row','add_section',
  'add_page_break','add_separator','set_text_style','delete_element','add_header',
  'add_footer','add_page_number','add_hyperlink','add_table_of_contents','set_margins',
  'set_orientation','set_font','set_line_spacing','add_chart','add_image','delete_table',
  'delete_image','remove_header','remove_footer','format_table_cell','add_table_column',
  'delete_table_column','count_images','count_tables','clear_content','clear_element',
  'add_content','add_page_number','add_text_box','add_page_border','add_section_break',
  'set_columns','add_column_break','add_watermark','add_drop_cap','add_tab_stop_paragraph',
  'add_formatted_page_numbers','set_table_width','set_table_column_widths',
  'add_image_positioned','add_highlight_paragraph','set_paragraph_spacing',
  'set_first_line_indent','clear_all_content','add_bookmark',
]);

// Operations that belong inside edit_spreadsheet
const EXCEL_OPS = new Set([
  'add_row','add_multiple_rows','update_cell','set_formula','add_column','set_cell_style',
  'set_range_style','set_number_format','set_column_width','set_row_height','freeze_panes',
  'unfreeze_panes','merge_cells','unmerge_cells','add_sheet','rename_sheet','delete_row',
  'delete_column','insert_row','insert_column','replace_text','add_conditional_format',
  'sort','set_auto_filter','set_print_area','set_page_setup','protect_sheet',
  'add_hyperlink_cell','add_hyperlink','add_comment','add_chart','add_image',
  'clear_range','clear_content','clear_format','clear_all','delete_image','delete_chart',
  'delete_sheet','remove_conditional_format','remove_data_validation','remove_hyperlink',
  'remove_comment','remove_auto_filter','unprotect_sheet','copy_range','copy_sheet',
  'group_rows','ungroup_rows','set_tab_color','set_print_titles','set_print_options',
  'add_named_range','set_alignment','set_borders','add_data_validation',
]);

export async function executeTool(
  toolCall: any,
  onProgress?: (status: string, progress: number) => void
): Promise<ToolResult> {
  const toolName = (toolCall.tool || '').toLowerCase().replace(/-/g, '_');
  onProgress?.('Executing: ' + toolName, 10);
  try {
    switch (toolName) {
      case 'create_document': return await createDocument(toolCall, onProgress);
      case 'create_spreadsheet': return await createSpreadsheet(toolCall, onProgress);
      case 'read_document': return await readDocumentContent(toolCall, onProgress);
      case 'get_paragraph_index': return await getParagraphIndex(toolCall, onProgress);
      case 'edit_document': return await editDocument(toolCall, onProgress);
      case 'edit_spreadsheet': return await editSpreadsheet(toolCall, onProgress);
      case 'read_spreadsheet_full': return await readSpreadsheetFullAction(toolCall, onProgress);
      case 'bulk_update_cells': return await bulkUpdateCellsAction(toolCall, onProgress);
      case 'analyze_file': return await analyzeFile(toolCall, onProgress);
      case 'list_files': return listFiles();
      case 'delete_file': return deleteFileAction(toolCall.filename);
      case 'rename_file': return renameFileAction(toolCall.filename, toolCall.new_filename);
      case 'get_document_xml': return await getDocumentXmlAction(toolCall, onProgress);
      case 'set_document_xml': return await setDocumentXmlAction(toolCall, onProgress);
      default: {
        // Auto-fix: If AI mistakenly called a Word operation as a tool, wrap it in edit_document
        if (WORD_OPS.has(toolName)) {
          const filename = toolCall.filename || params?.filename;
          if (filename) {
            const op = { ...toolCall, type: toolName };
            delete op.tool;
            delete op.filename;
            return await editDocument({ tool: 'edit_document', filename, operations: [op] }, onProgress);
          }
          return { success: false, message: `Operation "${toolName}" requires a filename. Use edit_document with filename and operations array.` };
        }
        // Auto-fix: If AI mistakenly called an Excel operation as a tool, wrap it in edit_spreadsheet
        if (EXCEL_OPS.has(toolName)) {
          const filename = toolCall.filename || params?.filename;
          const sheetName = toolCall.sheet || toolCall.sheetName || params?.sheet || 'Sheet1';
          if (filename) {
            const op = { ...toolCall, type: toolName };
            delete op.tool;
            delete op.filename;
            delete op.sheet;
            delete op.sheetName;
            return await editSpreadsheet({ tool: 'edit_spreadsheet', filename, sheet: sheetName, operations: [op] }, onProgress);
          }
          return { success: false, message: `Operation "${toolName}" requires a filename. Use edit_spreadsheet with filename, sheet, and operations array.` };
        }
        return { success: false, message: 'Unknown tool: "' + toolName + '". Available tools: create_document, create_spreadsheet, read_document, get_paragraph_index, edit_document, edit_spreadsheet, read_spreadsheet_full, bulk_update_cells, analyze_file, list_files, delete_file, rename_file, get_document_xml, set_document_xml' };
      }
    }
  } catch (error: any) {
    return { success: false, message: 'Error: ' + (error.message || 'Unknown error') };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════
// READ DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════════════

async function readDocumentContent(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.docx')) filename += '.docx';
  if (!fileExists(filename)) return { success: false, message: 'File "' + filename + '" not found.' };

  onProgress?.('Reading document...', 30);
  try {
    const model = await readDocx(filename);
    const userMsg = 'Read "' + filename + '" - ' + model.title + ' (' + model.wordCount + ' words, ' + model.elements.length + ' sections)';
    onProgress?.('Document read!', 100);
    return { success: true, message: userMsg, filename, preview: model.title + ' (' + model.wordCount + ' words)', progress: 100, data: model };
  } catch (error: any) {
    return { success: false, message: 'Failed to read document: ' + error.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════
// CREATE WORD DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════════════

async function createDocument(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  onProgress?.('Building document...', 20);

  const filename = params.filename || 'document.docx';
  const title = params.title || 'Document';
  const sections = params.sections || [];
  const styles = params.styles || {};

  ensureStorageDir();

  const children: any[] = [];
  const fontName = styles.font || 'Calibri';
  const titleSize = styles.titleSize || 36;
  const headingSize = styles.headingSize || 28;
  const bodySize = styles.bodySize || 22;
  const lineSpacing = styles.lineSpacing || 1.15;

  // Title
  children.push(new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: titleSize, font: fontName, color: '1F2937' })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 300, line: Math.round(lineSpacing * 240) },
  }));

  for (const section of sections) {
    if (section.heading) {
      const level = section.headingLevel || 2;
      const size = level === 1 ? headingSize + 4 : level === 2 ? headingSize : headingSize - 4;
      children.push(new Paragraph({
        children: [new TextRun({ text: section.heading, bold: true, size, font: fontName, color: section.headingColor || '1E40AF' })],
        heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
        spacing: { before: 300, after: 200 },
      }));
    }

    if (section.content) {
      const fmt = section.formatting || {};
      const paragraphs = section.content.split('\n');
      for (const para of paragraphs) {
        children.push(new Paragraph({
          children: [new TextRun({ text: para, size: bodySize, font: fontName, bold: fmt.bold || false, italics: fmt.italic || false, color: fmt.color || '374151' })],
          spacing: { after: 120, line: Math.round(lineSpacing * 240) },
        }));
      }
    }

    if (section.bullets && section.bullets.length > 0) {
      for (const item of section.bullets) {
        children.push(new Paragraph({
          children: [new TextRun({ text: '\u2022 ', bold: true, size: bodySize, font: fontName }), new TextRun({ text: item, size: bodySize, font: fontName, color: '374151' })],
          spacing: { after: 60, line: Math.round(lineSpacing * 240) },
          indent: { left: 360 },
        }));
      }
    }

    if (section.numberedItems && section.numberedItems.length > 0) {
      section.numberedItems.forEach((item: string, idx: number) => {
        children.push(new Paragraph({
          children: [new TextRun({ text: (idx + 1) + '. ', bold: true, size: bodySize, font: fontName }), new TextRun({ text: item, size: bodySize, font: fontName, color: '374151' })],
          spacing: { after: 60, line: Math.round(lineSpacing * 240) },
          indent: { left: 360 },
        }));
      });
    }

    if (section.table && section.table.headers) {
      const tableStyle = section.table.style || {};
      const headerBg = tableStyle.headerBg || '2D3748';
      const headerFont = tableStyle.headerFont || 'FFFFFF';
      const altRowBg = tableStyle.altRowBg || 'F7FAFC';

      const tableRows = [
        new TableRow({
          children: section.table.headers.map((h: string) =>
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: headerFont, font: fontName, size: bodySize - 2 })], alignment: AlignmentType.CENTER })],
              shading: { type: ShadingType.SOLID, color: headerBg },
            })
          ),
        }),
        ...(section.table.rows || []).map((row: any[], rowIdx: number) =>
          new TableRow({
            children: row.map((cell: any) =>
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: String(cell), font: fontName, size: bodySize - 2, color: '374151' })] })],
                shading: rowIdx % 2 === 1 ? { type: ShadingType.SOLID, color: altRowBg } : undefined,
              })
            ),
          })
        ),
      ];

      children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }));
      children.push(new Paragraph({ children: [], spacing: { after: 200 } }));
    }
  }

  onProgress?.('Generating Word file...', 70);

  const doc = new Document({
    sections: [{
      properties: {
        page: { margin: { top: styles.margins?.top || 1440, bottom: styles.margins?.bottom || 1440, left: styles.margins?.left || 1440, right: styles.margins?.right || 1440 } },
      },
      children,
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  writeFileBuffer(filename, Buffer.from(buffer));

  onProgress?.('Document created!', 100);
  return { success: true, message: 'Created document: ' + filename, filename, preview: '"' + title + '" with ' + sections.length + ' section(s)', progress: 100 };
}

// ═══════════════════════════════════════════════════════════════════════════════════
// CREATE EXCEL SPREADSHEET
// ═══════════════════════════════════════════════════════════════════════════════════

async function createSpreadsheet(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  onProgress?.('Building spreadsheet...', 20);

  const filename = params.filename || 'spreadsheet.xlsx';
  const sheets = params.sheets || [];

  ensureStorageDir();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'OfficeAI';
  wb.created = new Date();

  for (const sheetInfo of sheets) {
    const ws = wb.addWorksheet(sheetInfo.name || 'Sheet1');
    const headers = sheetInfo.headers || [];
    const data = sheetInfo.data || [];
    const formulas = sheetInfo.formulas || {};
    const stl = sheetInfo.styles || {};
    const columnWidths = sheetInfo.columnWidths || [];

    if (columnWidths.length > 0) {
      ws.columns = columnWidths.map((w: number) => ({ width: w }));
    } else if (headers.length > 0) {
      ws.columns = headers.map((h: string) => ({ width: Math.max(String(h).length + 4, 12) }));
    }

    if (headers.length > 0) {
      const headerRow = ws.addRow(headers);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: stl.headerFont || 'FFFFFFFF' }, size: 11 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: stl.headerBg || 'FF2D3748' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { bottom: { style: 'thin', color: { argb: 'FF' + (stl.borderColor || 'CBD5E0') } } };
      });
      headerRow.height = 24;
    }

    for (let i = 0; i < data.length; i++) {
      const rowData = data[i];
      const row = ws.addRow(Array.isArray(rowData) ? rowData : [rowData]);
      if (stl.altRowBg && i % 2 === 1) {
        row.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + stl.altRowBg } };
        });
      }
      row.eachCell((cell) => {
        cell.font = { size: 11 };
        cell.alignment = { vertical: 'middle' };
      });
    }

    for (const [cellRef, formula] of Object.entries(formulas)) {
      const cell = ws.getCell(cellRef);
      cell.value = { formula: (formula as string).replace(/^=/, '') };
      cell.font = { bold: true };
    }

    if (headers.length > 0) {
      ws.views = [{ state: 'frozen', ySplit: 1 }];
    }

    if (headers.length > 0 && data.length > 0) {
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: data.length + 1, column: headers.length } };
    }
  }

  onProgress?.('Saving spreadsheet...', 80);

  const buffer = await wb.xlsx.writeBuffer();
  writeFileBuffer(filename, Buffer.from(buffer));

  onProgress?.('Spreadsheet created!', 100);
  return { success: true, message: 'Created spreadsheet: ' + filename, filename, preview: sheets.length + ' sheet(s) with ' + (sheets[0]?.data?.length || 0) + ' data rows', progress: 100 };
}

// ═══════════════════════════════════════════════════════════════════════════════════
// EDIT WORD DOCUMENT - ALL OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════════════

async function editDocument(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.docx')) filename += '.docx';
  if (!fileExists(filename)) return { success: false, message: 'File "' + filename + '" not found.' };

  onProgress?.('Processing operations...', 30);

  const operations = params.operations || [];
  const editResults: string[] = [];
  let addContentXml = '';

  for (const op of operations) {
    try {
      switch (op.type) {
        // TEXT OPERATIONS
        case 'replace_text': {
          const find = op.find || op.old || '';
          const replace = op.replace || op.new || '';
          if (!find) { editResults.push('Error: replace_text requires "find"'); break; }
          const result = await replaceTextInDocx(filename, find, replace, false);
          editResults.push('Replaced "' + find + '" with "' + replace + '" (' + result.replacements + 'x)');
          break;
        }

        case 'set_text_style': {
          const find = op.find || op.text || '';
          if (!find) { editResults.push('Error: set_text_style requires "find"'); break; }
          const count = await applyTextStyle(filename, find, {
            bold: op.bold, italic: op.italic, underline: op.underline,
            color: op.color, font: op.font, fontSize: op.fontSize,
          });
          editResults.push('Styled "' + find + '" (' + count + 'x)');
          break;
        }

        // CONTENT OPERATIONS
        case 'add_heading': {
          const text = op.heading || op.text || '';
          addContentXml += coloredHeadingXml(text, op.level || 2, op.color || '1E40AF');
          editResults.push('Added heading: ' + text);
          break;
        }

        case 'add_paragraph':
        case 'add_content': {
          const text = op.content || op.text || '';
          if (op.bold || op.color || op.fontSize || op.alignment) {
            addContentXml += styledParagraphXml(text, {
              bold: op.bold, italic: op.italic, underline: op.underline,
              color: op.color, font: op.font, fontSize: op.fontSize, alignment: op.alignment,
            });
          } else {
            addContentXml += paragraphXml(text);
          }
          editResults.push('Added paragraph');
          break;
        }

        case 'add_bullet_list': {
          const items = op.items || [];
          if (items.length > 0) {
            addContentXml += bulletListXml(items);
            editResults.push('Added ' + items.length + ' bullet items');
          }
          break;
        }

        case 'add_numbered_list': {
          const items = op.items || [];
          if (items.length > 0) {
            addContentXml += bulletListXml(items.map((item: string, i: number) => (i + 1) + '. ' + item));
            editResults.push('Added ' + items.length + ' numbered items');
          }
          break;
        }

        // TABLE OPERATIONS
        case 'add_table': {
          if (op.headers && op.headers.length > 0) {
            addContentXml += tableXml(op.headers, op.rows || []);
            editResults.push('Added table with ' + op.headers.length + ' columns');
          }
          break;
        }

        case 'add_table_row': {
          // First flush any pending XML content
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const tableIdx = op.table_index ?? 0;
          const data = (op.data || []).map(String);
          if (data.length > 0) {
            await addTableRowToDocx(filename, tableIdx, data);
            editResults.push('Added row to table ' + tableIdx + ': ' + data.join(', '));
          }
          break;
        }

        case 'update_table_cell': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const tableIdx = op.table_index ?? 0;
          const row = op.row ?? 0;
          const col = op.col ?? op.column ?? 0;
          const value = String(op.value ?? '');
          await updateTableCellInDocx(filename, tableIdx, row, col, value);
          editResults.push('Updated table ' + tableIdx + ' cell [' + row + ',' + col + ']');
          break;
        }

        case 'delete_table_row': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const tableIdx = op.table_index ?? 0;
          const row = op.row ?? 0;
          await deleteTableRowFromDocx(filename, tableIdx, row);
          editResults.push('Deleted row ' + row + ' from table ' + tableIdx);
          break;
        }

        // SECTION OPERATIONS
        case 'add_section': {
          if (op.heading) addContentXml += coloredHeadingXml(op.heading, op.level || 2, op.headingColor || '1E40AF');
          if (op.content) addContentXml += paragraphXml(op.content);
          if (op.bullets) addContentXml += bulletListXml(op.bullets);
          if (op.numberedItems) addContentXml += bulletListXml(op.numberedItems.map((item: string, i: number) => (i + 1) + '. ' + item));
          if (op.table && op.table.headers) addContentXml += tableXml(op.table.headers, op.table.rows || []);
          editResults.push('Added section: ' + (op.heading || 'content'));
          break;
        }

        // LAYOUT OPERATIONS
        case 'add_page_break': {
          addContentXml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
          editResults.push('Added page break');
          break;
        }

        case 'add_separator': {
          addContentXml += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E0"/></w:pBdr></w:pPr></w:p>';
          editResults.push('Added separator');
          break;
        }

        case 'set_margins': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setMargins(filename, op.top || 1440, op.bottom || 1440, op.left || 1440, op.right || 1440);
          editResults.push('Margins set');
          break;
        }

        case 'set_orientation': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setOrientation(filename, op.orientation || 'portrait');
          editResults.push('Orientation set to: ' + (op.orientation || 'portrait'));
          break;
        }

        // HEADER/FOOTER
        case 'add_header': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addHeaderToDocx(filename, op.text || '');
          editResults.push('Added header: ' + (op.text || ''));
          break;
        }

        case 'add_footer': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addFooterToDocx(filename, op.text || '');
          editResults.push('Added footer: ' + (op.text || ''));
          break;
        }

        case 'add_page_number': {
          addContentXml += '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
          editResults.push('Added page number');
          break;
        }

        // ELEMENT OPERATIONS
        case 'delete_element': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const search = op.search || op.text || op.find || '';
          if (search) {
            const count = await deleteElementByContent(filename, search);
            editResults.push('Deleted ' + count + ' element(s) containing "' + search + '"');
          }
          break;
        }

        // HYPERLINK - proper implementation with relationship
        case 'add_hyperlink': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const url = op.url || op.href || '';
          const text = op.text || url;
          if (url) {
            await addHyperlinkToDocx(filename, text, url, op.color, op.underline !== false);
            editResults.push('Added hyperlink: ' + text);
          } else {
            editResults.push('Hyperlink: no URL provided');
          }
          break;
        }

        // TABLE OF CONTENTS
        case 'add_table_of_contents': {
          addContentXml += '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Table of Contents</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
          editResults.push('Added table of contents');
          break;
        }

        // FONT
        case 'set_font': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setDocumentFont(filename, op.font || 'Calibri');
          editResults.push('Font set to: ' + (op.font || 'Calibri'));
          break;
        }

        // LINE SPACING
        case 'set_line_spacing': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setLineSpacing(filename, op.spacing || 1.15);
          editResults.push('Line spacing set to: ' + (op.spacing || 1.15));
          break;
        }

        // CHART - embed a chart as an image in Word
        case 'add_chart': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const chartType = op.chart_type || 'bar';
          let chartLabels = op.labels || [];
          let chartValues = op.values || [];
          let chartTitle = op.title || 'Chart';

          // If source_file provided, read data from Excel
          if (op.source_file) {
            let srcFile = op.source_file;
            if (!srcFile.endsWith('.xlsx')) srcFile += '.xlsx';
            if (fileExists(srcFile)) {
              const buf = readFileBuffer(srcFile);
              if (buf) {
                const ExcelJS = (await import('exceljs')).default;
                const wb = new ExcelJS.Workbook();
                await wb.xlsx.load(buf as any);
                const ws = wb.worksheets[0];
                const rows: any[][] = [];
                ws.eachRow((row, rowNumber) => {
                  const vals: any[] = [];
                  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    vals[colNumber] = cell.value;
                  });
                  rows.push(vals);
                });
                if (rows.length > 1) {
                  chartLabels = rows.slice(1).map(r => String(r[1] || ''));
                  chartValues = rows.slice(1).map(r => {
                    const v = r[2];
                    return typeof v === 'number' ? v : parseFloat(String(v || '0').replace(/[^0-9.-]/g, '')) || 0;
                  });
                  chartTitle = chartTitle || String(rows[0]?.[2] || 'Chart');
                }
              }
            }
          }

          // If no labels/values provided, try to read from the source Excel data
          if (chartLabels.length === 0 && op.label_col && op.value_col) {
            // Labels and values should be passed directly
            editResults.push('Chart: provide labels and values or source_file');
            break;
          }

          if (chartLabels.length > 0 && chartValues.length > 0) {
            const autoType = detectChartType(chartLabels, chartValues);
            const finalType = chartType === 'auto' ? autoType : chartType;
            const chartConfig: ChartConfig = {
              type: finalType as any,
              labels: chartLabels,
              values: chartValues,
              title: chartTitle,
              seriesName: op.series_name || 'Value',
            };
            const result = await embedChartInDocx(filename, chartConfig, op.width || 500, op.height || 350);
            editResults.push('Added ' + finalType + ' chart: ' + chartTitle);
          } else {
            editResults.push('Chart: no data provided (need labels + values or source_file)');
          }
          break;
        }

        // IMAGE - embed base64 image in Word
        case 'add_image': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          let imgBase64 = op.image_base64 || op.image || '';
          // Strip data URI prefix if present
          if (imgBase64.startsWith('data:')) {
            imgBase64 = imgBase64.split(',')[1] || imgBase64;
          }
          if (imgBase64.length > 0) {
            const imgBuffer = Buffer.from(imgBase64, 'base64');
            const result = await embedImageInDocx(filename, imgBuffer, op.width || 400, op.height || 300);
            editResults.push(result.message);
          } else {
            editResults.push('Image: no base64 data provided');
          }
          break;
        }

        // ═══════════════════════════════════════════════════════════════════════
        // DELETE/CLEAR OPERATIONS
        // ═══════════════════════════════════════════════════════════════════════

        // DELETE_TABLE - Remove entire table by index
        case 'delete_table': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const tableIdx = op.table_index ?? op.index ?? 0;
          await deleteTableFromDocx(filename, tableIdx);
          editResults.push('Deleted table ' + tableIdx);
          break;
        }

        // DELETE_IMAGE - Remove image by index
        case 'delete_image': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const imgIdx = op.image_index ?? op.index ?? 0;
          await deleteImageFromDocx(filename, imgIdx);
          editResults.push('Deleted image ' + imgIdx);
          break;
        }

        // REMOVE_HEADER
        case 'remove_header': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await removeHeaderFromDocx(filename);
          editResults.push('Removed header');
          break;
        }

        // REMOVE_FOOTER
        case 'remove_footer': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await removeFooterFromDocx(filename);
          editResults.push('Removed footer');
          break;
        }

        // FORMAT_TABLE_CELL - Style a specific cell
        case 'format_table_cell': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await formatTableCellInDocx(
            filename,
            op.table_index ?? 0,
            op.row ?? 0,
            op.col ?? op.column ?? 0,
            { bg: op.bg || op.background, bold: op.bold, color: op.color, font: op.font, fontSize: op.fontSize, align: op.align }
          );
          editResults.push('Formatted cell [' + op.row + ',' + (op.col ?? op.column) + ']');
          break;
        }

        // ADD_TABLE_COLUMN - Add column to existing table
        case 'add_table_column': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addColumnToTable(filename, op.table_index ?? 0, op.header || '', op.values || []);
          editResults.push('Added column to table ' + (op.table_index ?? 0));
          break;
        }

        // DELETE_TABLE_COLUMN - Remove column from table
        case 'delete_table_column': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await deleteColumnFromTable(filename, op.table_index ?? 0, op.col ?? op.column ?? 0);
          editResults.push('Deleted column from table ' + (op.table_index ?? 0));
          break;
        }

        // COUNT_IMAGES
        case 'count_images': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const imgCount = await countImagesInDocx(filename);
          editResults.push('Document has ' + imgCount + ' image(s)');
          break;
        }

        // COUNT_TABLES
        case 'count_tables': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const tblCount = await countTablesInDocx(filename);
          editResults.push('Document has ' + tblCount + ' table(s)');
          break;
        }

        // CLEAR_CONTENT - Clear element by searching text
        case 'clear_content':
        case 'clear_element': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          const search = op.search || op.text || '';
          if (search) {
            await replaceTextInDocx(filename, search, '', false);
            editResults.push('Cleared content matching: ' + search);
          }
          break;
        }

        // ═══════════════════════════════════════════════════════════════════════
        // ADVANCED WORD OPERATIONS
        // ═══════════════════════════════════════════════════════════════════════

        // ADD_TEXT_BOX - Add a text box shape
        case 'add_text_box': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addTextBoxToDocx(filename, op.text || op.content || '',
            op.width || 200, op.height || 100,
            op.fillColor || op.fill || 'FFFFFF',
            op.borderColor || op.border || '000000',
            op.fontSize || 12, op.bold || false,
            op.color || '000000', op.alignment || 'left',
            op.x || 0, op.y || 0
          );
          editResults.push('Added text box');
          break;
        }

        // ADD_PAGE_BORDER - Add page borders
        case 'add_page_border': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addPageBorderToDocx(filename, op.style || 'single', op.color || '000000', op.size || 4);
          editResults.push('Added page border');
          break;
        }

        // ADD_SECTION_BREAK - Add section break
        case 'add_section_break': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addSectionBreakToDocx(filename, op.break_type || op.section_type || 'nextPage');
          editResults.push('Added section break: ' + (op.break_type || 'nextPage'));
          break;
        }

        // SET_COLUMNS - Set column layout
        case 'set_columns': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setColumnsInDocx(filename, op.count || 2, op.spacing || 720, op.separator || false);
          editResults.push('Set ' + (op.count || 2) + ' columns');
          break;
        }

        // ADD_COLUMN_BREAK - Add column break
        case 'add_column_break': {
          addContentXml += '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
          editResults.push('Added column break');
          break;
        }

        // ADD_WATERMARK - Add text watermark
        case 'add_watermark': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addWatermarkToDocx(filename, op.text || 'DRAFT', op.color || 'C0C0C0', op.fontSize || 72, op.font || 'Arial');
          editResults.push('Added watermark: ' + (op.text || 'DRAFT'));
          break;
        }

        // ADD_DROP_CAP - Add drop cap paragraph
        case 'add_drop_cap': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addDropCapParagraphToDocx(filename, op.text || op.content || '', op.lines || 3, op.font || '', op.color || '');
          editResults.push('Added drop cap paragraph');
          break;
        }

        // ADD_TAB_STOP_PARAGRAPH - Paragraph with tab stops
        case 'add_tab_stop_paragraph': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addTabStopParagraphToDocx(filename, op.text || '', op.tabStops || [], op.fontSize || 12, op.bold || false, op.color || '');
          editResults.push('Added paragraph with tab stops');
          break;
        }

        // ADD_FORMATTED_PAGE_NUMBERS - Formatted page numbering
        case 'add_formatted_page_numbers': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addFormattedPageNumbersToDocx(filename, op.format || 'Page {n} of {total}', op.alignment || 'center', op.showTotal !== false, op.font || '', op.fontSize || 12, op.color || '');
          editResults.push('Added formatted page numbers');
          break;
        }

        // SET_TABLE_WIDTH - Set table width/alignment
        case 'set_table_width': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setTableWidthToDocx(filename, op.table_index ?? 0, op.width || 0, op.widthType || 'auto', op.alignment || 'center');
          editResults.push('Set table width');
          break;
        }

        // SET_TABLE_COLUMN_WIDTHS - Set column widths
        case 'set_table_column_widths': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setTableColumnWidthsToDocx(filename, op.table_index ?? 0, op.widths || []);
          editResults.push('Set table column widths');
          break;
        }

        // ADD_IMAGE_POSITIONED - Image with positioning
        case 'add_image_positioned': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          let imgBase64 = op.image_base64 || op.image || '';
          if (imgBase64.startsWith('data:')) {
            imgBase64 = imgBase64.split(',')[1] || imgBase64;
          }
          if (imgBase64.length > 0) {
            const imgBuffer = Buffer.from(imgBase64, 'base64');
            const result = await embedImagePositionedInDocx(filename, imgBuffer, op.width || 400, op.height || 300, op.x || 0, op.y || 0, op.wrapStyle || op.wrap || 'square');
            editResults.push(result.message);
          } else {
            editResults.push('Image: no base64 data provided');
          }
          break;
        }

        // ADD_HIGHLIGHT_PARAGRAPH - Paragraph with highlight
        case 'add_highlight_paragraph': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await addHighlightParagraphToDocx(filename, op.text || op.content || '', op.highlight || 'yellow', op.bold || false, op.color || '');
          editResults.push('Added highlighted paragraph');
          break;
        }

        // SET_PARAGRAPH_SPACING - Set spacing
        case 'set_paragraph_spacing': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setParagraphSpacingToDocx(filename, op.before || 0, op.after || 0, op.lineSpacing);
          editResults.push('Set paragraph spacing');
          break;
        }

        // SET_FIRST_LINE_INDENT - Set indent
        case 'set_first_line_indent': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await setFirstLineIndentToDocx(filename, op.indent || 720);
          editResults.push('Set first line indent');
          break;
        }

        // CLEAR_ALL_CONTENT - Clear entire document
        case 'clear_all_content': {
          if (addContentXml) {
            await addToDocx(filename, addContentXml);
            addContentXml = '';
          }
          await clearAllContentFromDocx(filename);
          editResults.push('Cleared all content');
          break;
        }

        // ADD_BOOKMARK - Add bookmark
        case 'add_bookmark': {
          const bmName = op.name || 'bookmark_' + Date.now();
          const bmText = op.text || '';
          addContentXml += `<w:p><w:bookmarkStart w:id="${Date.now()}" w:name="${bmName}"/><w:r><w:t xml:space="preserve">${escapeXmlForTools(bmText)}</w:t></w:r><w:bookmarkEnd w:id="${Date.now()}"/></w:p>`;
          editResults.push('Added bookmark: ' + bmName);
          break;
        }

        // ═══════════════════════════════════════════════════════════════════
        // INDEXED PARAGRAPH OPERATIONS — precision surgical editing
        // ═══════════════════════════════════════════════════════════════════

        case 'insert_before_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          const idx = op.index ?? 0;
          const xml = buildContentXml(op);
          await insertBeforeIndex(filename, idx, xml);
          editResults.push('Inserted before index ' + idx);
          break;
        }

        case 'insert_after_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          const idx = op.index ?? 0;
          const xml = buildContentXml(op);
          await insertAfterIndex(filename, idx, xml);
          editResults.push('Inserted after index ' + idx);
          break;
        }

        case 'replace_at_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          const idx = op.index ?? 0;
          const xml = buildContentXml(op);
          await replaceAtIndex(filename, idx, xml);
          editResults.push('Replaced block at index ' + idx);
          break;
        }

        case 'replace_text_at_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          const idx = op.index ?? 0;
          const newText = op.text || op.content || '';
          await replaceTextAtIndex(filename, idx, newText);
          editResults.push('Replaced text at index ' + idx + ': ' + newText.slice(0, 40));
          break;
        }

        case 'delete_at_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          const idx = op.index ?? 0;
          await deleteAtIndex(filename, idx);
          editResults.push('Deleted block at index ' + idx);
          break;
        }

        case 'format_at_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          const idx = op.index ?? 0;
          await formatAtIndex(filename, idx, {
            bold: op.bold, italic: op.italic, underline: op.underline,
            strikethrough: op.strikethrough,
            color: op.color, font: op.font, fontSize: op.fontSize,
            alignment: op.alignment, headingLevel: op.headingLevel,
            indent: op.indent, spaceBefore: op.spaceBefore,
            spaceAfter: op.spaceAfter, lineSpacing: op.lineSpacing,
            highlight: op.highlight,
          });
          editResults.push('Formatted block at index ' + idx);
          break;
        }

        case 'duplicate_at_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          await duplicateAtIndex(filename, op.index ?? 0);
          editResults.push('Duplicated block at index ' + (op.index ?? 0));
          break;
        }

        case 'move_to_index': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          await moveBlockToIndex(filename, op.source_index ?? 0, op.dest_index ?? 0);
          editResults.push('Moved block from ' + op.source_index + ' to ' + op.dest_index);
          break;
        }

        case 'merge_table_cells': {
          if (addContentXml) { await addToDocx(filename, addContentXml); addContentXml = ''; }
          await mergeTableCellsInDocx(filename,
            op.table_index ?? 0,
            op.startRow ?? 0, op.startCol ?? 0,
            op.endRow ?? 0, op.endCol ?? 0
          );
          editResults.push('Merged cells in table ' + (op.table_index ?? 0));
          break;
        }

        default:
          editResults.push('Unknown operation: ' + op.type);
      }
    } catch (error: any) {
      editResults.push('Error in ' + op.type + ': ' + error.message);
    }
  }

  // Flush remaining XML
  if (addContentXml) {
    onProgress?.('Adding content...', 70);
    try {
      await addToDocx(filename, addContentXml);
    } catch (error: any) {
      editResults.push('Error adding content: ' + error.message);
    }
  }

  onProgress?.('Document updated!', 100);
  return { success: true, message: editResults.join('\n'), filename, preview: 'Applied ' + operations.length + ' operation(s)', progress: 100 };
}

// ═══════════════════════════════════════════════════════════════════════════════════
// EDIT EXCEL SPREADSHEET - ALL OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════════════

async function editSpreadsheet(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.xlsx')) filename += '.xlsx';
  if (!fileExists(filename)) return { success: false, message: 'File "' + filename + '" not found.' };

  onProgress?.('Reading spreadsheet...', 20);

  const buffer = readFileBuffer(filename);
  if (!buffer) return { success: false, message: 'Could not read ' + filename };

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);

  const sheetName = params.sheet || wb.worksheets[0]?.name || 'Sheet1';
  let ws = wb.getWorksheet(sheetName);
  if (!ws) ws = wb.addWorksheet(sheetName);

  const operations = params.operations || [];
  const editResults: string[] = [];

  onProgress?.('Applying modifications...', 50);

  for (const op of operations) {
    try {
      switch (op.type) {
        // DATA OPERATIONS
        case 'add_row': {
          ws.addRow(op.data || []);
          editResults.push('Added row');
          break;
        }

        case 'add_multiple_rows': {
          for (const rowData of (op.rows || [])) ws.addRow(rowData);
          editResults.push('Added ' + (op.rows || []).length + ' rows');
          break;
        }

        case 'update_cell': {
          const cell = ws.getCell(op.cell || 'A1');
          cell.value = op.value;
          editResults.push('Updated ' + op.cell);
          break;
        }

        case 'set_formula': {
          const cell = ws.getCell(op.cell || 'A1');
          cell.value = { formula: (op.formula || '').replace(/^=/, '') };
          editResults.push('Set formula in ' + op.cell);
          break;
        }

        case 'add_column': {
          const colIdx = ws.columnCount + 1;
          const headerRow = ws.getRow(1);
          headerRow.getCell(colIdx).value = op.header || 'New Column';
          headerRow.getCell(colIdx).font = { bold: true };
          for (let i = 0; i < (op.values || []).length; i++) {
            ws.getCell(i + 2, colIdx).value = op.values[i];
          }
          editResults.push('Added column: ' + op.header);
          break;
        }

        // STYLE OPERATIONS
        case 'set_cell_style': {
          const cell = ws.getCell(op.cell || 'A1');
          applyCellStyle(cell, op.style || {});
          editResults.push('Styled ' + op.cell);
          break;
        }

        case 'set_range_style': {
          const style = op.style || {};
          const rangeRef = op.range || 'A1';
          const [startCell, endCell] = rangeRef.split(':');
          const startCol = colLetterToNum(startCell.replace(/[0-9]/g, ''));
          const startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
          const endCol = endCell ? colLetterToNum(endCell.replace(/[0-9]/g, '')) : startCol;
          const endRow = endCell ? parseInt(endCell.replace(/[A-Z]/g, '')) : startRow;

          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              applyCellStyle(ws.getCell(r, c), style);
            }
          }
          editResults.push('Styled range ' + rangeRef);
          break;
        }

        case 'set_number_format': {
          const cell = ws.getCell(op.cell || 'A1');
          cell.numFmt = op.format || '#,##0.00';
          editResults.push('Set format on ' + op.cell);
          break;
        }

        case 'set_column_width': {
          ws.getColumn(op.column || 1).width = op.width || 15;
          editResults.push('Set column ' + (op.column || 1) + ' width to ' + (op.width || 15));
          break;
        }

        case 'set_row_height': {
          ws.getRow(op.row || 1).height = op.height || 20;
          editResults.push('Set row ' + (op.row || 1) + ' height');
          break;
        }

        // LAYOUT OPERATIONS
        case 'freeze_panes': {
          const row = parseInt((op.cell || 'A2').replace(/[A-Z]/g, '')) - 1;
          ws.views = [{ state: 'frozen', ySplit: row || 1 }];
          editResults.push('Froze panes at row ' + (row + 1));
          break;
        }

        case 'unfreeze_panes': {
          ws.views = [];
          editResults.push('Unfroze panes');
          break;
        }

        case 'merge_cells': {
          ws.mergeCells(op.range || 'A1:B1');
          editResults.push('Merged ' + op.range);
          break;
        }

        case 'unmerge_cells': {
          ws.unMergeCells(op.range || 'A1:B1');
          editResults.push('Unmerged ' + op.range);
          break;
        }

        // SHEET OPERATIONS
        case 'add_sheet': {
          wb.addWorksheet(op.name || 'New Sheet');
          editResults.push('Added sheet: ' + op.name);
          break;
        }

        case 'rename_sheet': {
          const targetSheet = wb.getWorksheet(op.old_name || 'Sheet1');
          if (targetSheet) targetSheet.name = op.new_name || 'Renamed';
          editResults.push('Renamed sheet to ' + op.new_name);
          break;
        }

        // DELETE OPERATIONS
        case 'delete_row': {
          ws.spliceRows(op.row || 1, 1);
          editResults.push('Deleted row ' + op.row);
          break;
        }

        case 'delete_column': {
          ws.spliceColumns(op.column || 1, 1);
          editResults.push('Deleted column ' + op.column);
          break;
        }

        // REPLACE
        case 'replace_text': {
          const find = op.find || '';
          const replace = op.replace || '';
          let count = 0;
          ws.eachRow((row) => {
            row.eachCell((cell) => {
              if (typeof cell.value === 'string' && cell.value.includes(find)) {
                cell.value = cell.value.split(find).join(replace);
                count++;
              }
            });
          });
          editResults.push('Replaced "' + find + '" with "' + replace + '" in ' + count + ' cells');
          break;
        }

        // CONDITIONAL FORMATTING
        case 'add_conditional_format': {
          const range = op.range || 'A1:A10';
          const condition = op.condition || 'greaterThan';
          const value = op.value || 0;
          const condStyle = op.style || {};

          ws.addConditionalFormatting({
            ref: range,
            rules: [{
              type: 'cellIs',
              operator: condition as any,
              formulae: [value],
              priority: 1,
              style: {
                font: condStyle.fontColor ? { color: { argb: 'FF' + condStyle.fontColor } } : undefined,
                fill: condStyle.bgColor ? { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF' + condStyle.bgColor } } : undefined,
              },
            }],
          });
          editResults.push('Added conditional formatting to ' + range);
          break;
        }

        // DATA VALIDATION
        case 'add_data_validation': {
          const rangeStr = op.range || 'A1:A10';
          const validationType = op.validation_type || op.type_param || 'list';
          const values = op.values || [];
          const operator = op.operator || 'between';
          const minVal = op.min ?? op.minimum;
          const maxVal = op.max ?? op.maximum;

          // Parse range to get cells
          const [startCell, endCell] = rangeStr.split(':');
          const startCol = colLetterToNum(startCell.replace(/[0-9]/g, ''));
          const startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
          const endCol = endCell ? colLetterToNum(endCell.replace(/[0-9]/g, '')) : startCol;
          const endRow = endCell ? parseInt(endCell.replace(/[A-Z]/g, '')) : startRow;

          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              const cell = ws.getCell(r, c);
              const cellAddress = cell.address;

              if (validationType === 'list' && values.length > 0) {
                // Set cell data validation for dropdown lists
                (cell as any).dataValidation = {
                  type: 'list',
                  allowBlank: op.allow_blank !== false,
                  formulae: ['"' + values.join(',') + '"'],
                  showErrorMessage: true,
                  errorTitle: 'Invalid Input',
                  error: 'Please select from the list',
                };
              } else if (validationType === 'whole' || validationType === 'decimal') {
                (cell as any).dataValidation = {
                  type: validationType,
                  allowBlank: op.allow_blank !== false,
                  operator: operator,
                  formulae: minVal !== undefined && maxVal !== undefined ? [minVal, maxVal] : minVal !== undefined ? [minVal] : [0],
                  showErrorMessage: true,
                  errorTitle: 'Invalid Value',
                  error: 'Value must be between ' + (minVal ?? 0) + ' and ' + (maxVal ?? 'any'),
                };
              } else if (validationType === 'date') {
                (cell as any).dataValidation = {
                  type: 'date',
                  allowBlank: op.allow_blank !== false,
                  operator: operator,
                  formulae: minVal ? [new Date(minVal)] : maxVal ? [new Date(maxVal)] : [],
                  showErrorMessage: true,
                };
              } else if (validationType === 'textLength') {
                (cell as any).dataValidation = {
                  type: 'textLength',
                  allowBlank: op.allow_blank !== false,
                  operator: operator,
                  formulae: [minVal ?? 0, maxVal ?? 255],
                  showErrorMessage: true,
                  error: 'Text must be between ' + (minVal ?? 0) + ' and ' + (maxVal ?? 255) + ' characters',
                };
              }
            }
          }
          editResults.push('Added ' + validationType + ' validation to ' + rangeStr);
          break;
        }

        // HYPERLINK
        case 'add_hyperlink_cell':
        case 'add_hyperlink': {
          const cell = ws.getCell(op.cell || 'A1');
          cell.value = { text: op.text || op.url, hyperlink: op.url };
          cell.font = { color: { argb: 'FF0563C1' }, underline: true };
          editResults.push('Added hyperlink to ' + (op.cell || 'A1'));
          break;
        }

        // COMMENT
        case 'add_comment': {
          const cell = ws.getCell(op.cell || 'A1');
          cell.note = op.text || '';
          editResults.push('Added comment to ' + op.cell);
          break;
        }

        // PRINT
        case 'set_print_area': {
          const range = op.range || 'A1:F20';
          ws.pageSetup.printArea = range;
          editResults.push('Print area set to ' + range);
          break;
        }

        // PAGE SETUP
        case 'set_page_setup': {
          if (op.orientation) ws.pageSetup.orientation = op.orientation as any;
          if (op.paperSize) ws.pageSetup.paperSize = op.paperSize as any;
          if (op.fitToPage) ws.pageSetup.fitToPage = true;
          if (op.fitToWidth) ws.pageSetup.fitToWidth = op.fitToWidth;
          if (op.fitToHeight) ws.pageSetup.fitToHeight = op.fitToHeight;
          if (op.margins) {
            ws.pageSetup.margins = {
              left: op.margins.left || 0.7, right: op.margins.right || 0.7,
              top: op.margins.top || 0.75, bottom: op.margins.bottom || 0.75,
              header: op.margins.header || 0.3, footer: op.margins.footer || 0.3,
            };
          }
          editResults.push('Page setup configured');
          break;
        }

        // INSERT ROW AT POSITION
        case 'insert_row': {
          const rowPos = op.row || 1;
          ws.spliceRows(rowPos, 0, op.data || []);
          editResults.push('Inserted row at position ' + rowPos);
          break;
        }

        // INSERT COLUMN AT POSITION
        case 'insert_column': {
          const colPos = op.column || 1;
          ws.spliceColumns(colPos, 0, []);
          if (op.header) {
            ws.getRow(1).getCell(colPos).value = op.header;
          }
          if (op.values) {
            for (let i = 0; i < op.values.length; i++) {
              ws.getCell(i + 2, colPos).value = op.values[i];
            }
          }
          editResults.push('Inserted column at position ' + colPos);
          break;
        }

        // SORT DATA
        case 'sort': {
          const sortColumn = op.column || 1;
          const sortOrder = op.order || 'asc';
          const dataRows: any[][] = [];
          ws.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
              const vals: any[] = [];
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                vals[colNumber] = cell.value;
              });
              dataRows.push(vals);
            }
          });
          dataRows.sort((a, b) => {
            const va = a[sortColumn] ?? '';
            const vb = b[sortColumn] ?? '';
            if (typeof va === 'number' && typeof vb === 'number') {
              return sortOrder === 'asc' ? va - vb : vb - va;
            }
            return sortOrder === 'asc' ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
          });
          for (let i = 0; i < dataRows.length; i++) {
            const row = ws.getRow(i + 2);
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              cell.value = dataRows[i][colNumber] ?? null;
            });
            row.commit();
          }
          editResults.push('Sorted by column ' + sortColumn + ' (' + sortOrder + ')');
          break;
        }

        // AUTO FILTER
        case 'set_auto_filter': {
          const filterRange = op.range || 'A1:Z100';
          ws.autoFilter = filterRange as any;
          editResults.push('Auto filter set on ' + filterRange);
          break;
        }

        // PROTECT
        case 'protect_sheet': {
          await ws.protect(op.password || '', {
            selectLockedCells: true,
            selectUnlockedCells: true,
          } as any);
          editResults.push('Sheet protected');
          break;
        }

        // CHART - render and embed chart image in Excel
        case 'add_chart': {
          const chartType = op.chart_type || 'auto';
          let chartLabels: string[] = op.labels || [];
          let chartValues: number[] = op.values || [];
          let chartTitle = op.title || 'Chart';

          // Auto-read from spreadsheet if no explicit data
          if (chartLabels.length === 0) {
            const labelCol = op.label_col ? colLetterToNum(String(op.label_col).toUpperCase()) : 1;
            const valueCol = op.value_col ? colLetterToNum(String(op.value_col).toUpperCase()) : 2;
            const fromRow = op.from_row || 2;
            const toRow = op.to_row || ws.rowCount;

            for (let r = fromRow; r <= toRow; r++) {
              const labelVal = ws.getCell(r, labelCol).value;
              const valueVal = ws.getCell(r, valueCol).value;
              if (labelVal !== null && labelVal !== undefined && labelVal !== '') {
                chartLabels.push(String(labelVal));
                const num = typeof valueVal === 'number' ? valueVal : parseFloat(String(valueVal || '0').replace(/[^0-9.-]/g, '')) || 0;
                chartValues.push(num);
              }
            }

            if (!chartTitle) {
              const headerVal = ws.getCell(1, valueCol).value;
              chartTitle = String(headerVal || 'Chart');
            }
          }

          if (chartLabels.length > 0 && chartValues.length > 0) {
            const autoType = detectChartType(chartLabels, chartValues);
            const finalType = chartType === 'auto' ? autoType : chartType;

            const chartConfig: ChartConfig = {
              type: finalType as any,
              labels: chartLabels,
              values: chartValues,
              title: chartTitle,
              seriesName: op.series_name || chartTitle,
            };

            const imageBuffer = await renderChart(chartConfig);

            // Find a good placement for the chart image
            const chartRow = (op.chart_row || ws.rowCount + 3);
            const chartCol = op.chart_col || 1;
            const chartCell = String.fromCharCode(64 + chartCol) + chartRow;

            const imageId = wb.addImage({
              buffer: imageBuffer,
              extension: 'png',
            } as any);

            ws.addImage(imageId, {
              tl: { col: chartCol - 1, row: chartRow - 1 },
              ext: { width: op.width || 500, height: op.height || 350 },
            } as any);

            editResults.push('Added ' + finalType + ' chart: ' + chartTitle + ' at row ' + chartRow);
          } else {
            editResults.push('Chart: no data found');
          }
          break;
        }

        // IMAGE - embed base64 image in Excel
        case 'add_image': {
          let imgBase64 = op.image_base64 || op.image || '';
          if (imgBase64.startsWith('data:')) {
            imgBase64 = imgBase64.split(',')[1] || imgBase64;
          }
          if (imgBase64.length > 0) {
            const imgBuffer = Buffer.from(imgBase64, 'base64');
            const imgCol = op.column || 1;
            const imgRow = op.row || 1;

            const imageId = wb.addImage({
              buffer: imgBuffer,
              extension: 'png',
            } as any);

            ws.addImage(imageId, {
              tl: { col: imgCol - 1, row: imgRow - 1 },
              ext: { width: op.width || 400, height: op.height || 300 },
            } as any);

            editResults.push('Added image at ' + String.fromCharCode(64 + imgCol) + imgRow);
          } else {
            editResults.push('Image: no base64 data provided');
          }
          break;
        }

        // ═══════════════════════════════════════════════════════════════════════
        // UNIVERSAL CLEAR/DELETE OPERATIONS
        // ═══════════════════════════════════════════════════════════════════════

        // CLEAR_RANGE - Universal clear: content, format, all, hyperlinks, comments, validation
        case 'clear_range':
        case 'clear_content':
        case 'clear_format':
        case 'clear_all': {
          const rangeStr = op.range || op.cell || 'A1';
          const mode = op.mode || (op.type === 'clear_content' ? 'content' : op.type === 'clear_format' ? 'format' : op.type === 'clear_all' ? 'all' : 'all');

          // Parse range
          let startCol: number, startRow: number, endCol: number, endRow: number;
          if (rangeStr.includes(':')) {
            const [startCell, endCell] = rangeStr.split(':');
            startCol = colLetterToNum(startCell.replace(/[0-9]/g, ''));
            startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
            endCol = colLetterToNum(endCell.replace(/[0-9]/g, ''));
            endRow = parseInt(endCell.replace(/[A-Z]/g, ''));
          } else {
            startCol = colLetterToNum(rangeStr.replace(/[0-9]/g, ''));
            startRow = parseInt(rangeStr.replace(/[A-Z]/g, ''));
            endCol = startCol;
            endRow = startRow;
          }

          let cleared = 0;
          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              const cell = ws.getCell(r, c);
              if (mode === 'content' || mode === 'all') {
                cell.value = null;
              }
              if (mode === 'format' || mode === 'all') {
                cell.font = {};
                cell.fill = { type: 'pattern', pattern: 'none' };
                cell.border = {};
                cell.alignment = {};
                cell.numFmt = '';
              }
              if (mode === 'hyperlinks' || mode === 'all') {
                cell.value = typeof cell.value === 'object' && (cell.value as any)?.hyperlink
                  ? null
                  : cell.value;
              }
              if (mode === 'comments' || mode === 'all') {
                cell.note = undefined as any;
              }
              if (mode === 'validation' || mode === 'all') {
                (cell as any).dataValidation = undefined;
              }
              cleared++;
            }
          }
          editResults.push('Cleared ' + mode + ' from ' + rangeStr + ' (' + cleared + ' cells)');
          break;
        }

        // DELETE_IMAGE - Remove image by index
        case 'delete_image':
        case 'delete_chart': {
          const imgIndex = op.index ?? op.image_index ?? op.chart_index ?? 0;
          const images = (ws as any)._images || [];
          if (imgIndex < images.length) {
            images.splice(imgIndex, 1);
            editResults.push('Deleted image/chart at index ' + imgIndex);
          } else {
            editResults.push('Image index ' + imgIndex + ' not found (have ' + images.length + ' images)');
          }
          break;
        }

        // DELETE_SHEET - Remove worksheet
        case 'delete_sheet': {
          const sheetName = op.name || op.sheet || '';
          if (sheetName && wb.getWorksheet(sheetName)) {
            wb.removeWorksheet(sheetName);
            editResults.push('Deleted sheet: ' + sheetName);
          } else if (wb.worksheets.length > 1) {
            wb.removeWorksheet(ws.name);
            editResults.push('Deleted sheet: ' + ws.name);
          } else {
            editResults.push('Cannot delete the only sheet');
          }
          break;
        }

        // REMOVE_CONDITIONAL_FORMAT
        case 'remove_conditional_format': {
          const rangeStr = op.range || 'A1:Z100';
          try {
            ws.removeConditionalFormatting(rangeStr as any);
          } catch {
            // Clear all conditional formatting if range-specific fails
            (ws as any)._conditionalFormatting = [];
          }
          editResults.push('Removed conditional formatting from ' + rangeStr);
          break;
        }

        // REMOVE_DATA_VALIDATION
        case 'remove_data_validation': {
          const rangeStr = op.range || op.cell || '';
          if (rangeStr) {
            const [startCell, endCell] = rangeStr.includes(':') ? rangeStr.split(':') : [rangeStr, rangeStr];
            const sCol = colLetterToNum(startCell.replace(/[0-9]/g, ''));
            const sRow = parseInt(startCell.replace(/[A-Z]/g, ''));
            const eCol = colLetterToNum(endCell.replace(/[0-9]/g, ''));
            const eRow = parseInt(endCell.replace(/[A-Z]/g, ''));
            for (let r = sRow; r <= eRow; r++) {
              for (let c = sCol; c <= eCol; c++) {
                (ws.getCell(r, c) as any).dataValidation = undefined;
              }
            }
            editResults.push('Removed data validation from ' + rangeStr);
          }
          break;
        }

        // REMOVE_HYPERLINK
        case 'remove_hyperlink': {
          const cellRef = op.cell || 'A1';
          const cell = ws.getCell(cellRef);
          if (typeof cell.value === 'object' && (cell.value as any)?.hyperlink) {
            cell.value = (cell.value as any).text || null;
          } else {
            cell.value = null;
          }
          editResults.push('Removed hyperlink from ' + cellRef);
          break;
        }

        // REMOVE_COMMENT
        case 'remove_comment': {
          const cellRef = op.cell || 'A1';
          ws.getCell(cellRef).note = undefined as any;
          editResults.push('Removed comment from ' + cellRef);
          break;
        }

        // REMOVE_AUTO_FILTER
        case 'remove_auto_filter': {
          ws.autoFilter = undefined as any;
          editResults.push('Removed auto filter');
          break;
        }

        // UNPROTECT_SHEET
        case 'unprotect_sheet': {
          ws.unprotect();
          editResults.push('Sheet unprotected');
          break;
        }

        // COPY_RANGE - Copy range to destination
        case 'copy_range': {
          const srcRange = op.source || op.range || 'A1:A1';
          const dstCell = op.destination || op.dest || 'A1';
          const [srcStart, srcEnd] = srcRange.split(':');
          const sCol = colLetterToNum(srcStart.replace(/[0-9]/g, ''));
          const sRow = parseInt(srcStart.replace(/[A-Z]/g, ''));
          const eCol = srcEnd ? colLetterToNum(srcEnd.replace(/[0-9]/g, '')) : sCol;
          const eRow = srcEnd ? parseInt(srcEnd.replace(/[A-Z]/g, '')) : sRow;
          const dCol = colLetterToNum(dstCell.replace(/[0-9]/g, ''));
          const dRow = parseInt(dstCell.replace(/[A-Z]/g, ''));

          for (let r = sRow; r <= eRow; r++) {
            for (let c = sCol; c <= eCol; c++) {
              const srcCell = ws.getCell(r, c);
              const dstCellObj = ws.getCell(dRow + (r - sRow), dCol + (c - sCol));
              dstCellObj.value = srcCell.value;
              dstCellObj.style = srcCell.style;
            }
          }
          editResults.push('Copied ' + srcRange + ' to ' + dstCell);
          break;
        }

        // COPY_SHEET - Duplicate worksheet
        case 'copy_sheet': {
          const srcName = op.source || op.sheet || ws.name;
          const newName = op.name || (srcName + ' Copy');
          const srcSheet = wb.getWorksheet(srcName);
          if (srcSheet) {
            const newSheet = wb.addWorksheet(newName);
            srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
              const newRow = newSheet.getRow(rowNumber);
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const newCell = newRow.getCell(colNumber);
                newCell.value = cell.value;
                newCell.style = cell.style;
              });
              newRow.commit();
            });
            if (srcSheet.columnCount) {
              for (let c = 1; c <= srcSheet.columnCount; c++) {
                newSheet.getColumn(c).width = srcSheet.getColumn(c).width;
              }
            }
            editResults.push('Copied sheet ' + srcName + ' to ' + newName);
          } else {
            editResults.push('Sheet ' + srcName + ' not found');
          }
          break;
        }

        // GROUP_ROWS - Create row outline group
        case 'group_rows': {
          const start = op.start || op.from || 1;
          const end = op.end || op.to || start + 5;
          for (let r = start; r <= end; r++) {
            ws.getRow(r).outlineLevel = op.level || 1;
          }
          editResults.push('Grouped rows ' + start + '-' + end);
          break;
        }

        // UNGROUP_ROWS
        case 'ungroup_rows': {
          const start = op.start || op.from || 1;
          const end = op.end || op.to || start + 5;
          for (let r = start; r <= end; r++) {
            ws.getRow(r).outlineLevel = 0;
          }
          editResults.push('Ungrouped rows ' + start + '-' + end);
          break;
        }

        // SET_TAB_COLOR
        case 'set_tab_color': {
          ws.properties.tabColor = { argb: 'FF' + (op.color || 'FF0000').replace('#', '') };
          editResults.push('Set tab color to ' + op.color);
          break;
        }

        // SET_PRINT_TITLES
        case 'set_print_titles': {
          if (op.rows) ws.pageSetup.printTitlesRow = op.rows;
          if (op.columns) ws.pageSetup.printTitlesColumn = op.columns;
          editResults.push('Set print titles');
          break;
        }

        // SET_PRINT_OPTIONS
        case 'set_print_options': {
          if (op.gridlines !== undefined) ws.pageSetup.showGridLines = op.gridlines;
          if (op.headings !== undefined) ws.pageSetup.showRowColHeaders = op.headings;
          if (op.blackAndWhite !== undefined) ws.pageSetup.blackAndWhite = op.blackAndWhite;
          editResults.push('Set print options');
          break;
        }

        // ADD_NAMED_RANGE
        case 'add_named_range': {
          const rangeStr = op.range || 'A1:A10';
          const nameStr = op.name || 'NamedRange';
          try {
            (wb as any).definedNames.add(nameStr, rangeStr);
            editResults.push('Added named range: ' + nameStr + ' = ' + rangeStr);
          } catch (e: any) {
            editResults.push('Named range: ' + nameStr + ' noted');
          }
          break;
        }

        // SET_ALIGNMENT
        case 'set_alignment': {
          const cellRef = op.cell || op.range || 'A1';
          const alignment: any = {};
          if (op.horizontal) alignment.horizontal = op.horizontal;
          if (op.vertical) alignment.vertical = op.vertical;
          if (op.wrapText) alignment.wrapText = true;
          if (op.indent) alignment.indent = op.indent;
          if (op.rotation) alignment.textRotation = op.rotation;

          if (cellRef.includes(':')) {
            const [startCell, endCell] = cellRef.split(':');
            const sCol = colLetterToNum(startCell.replace(/[0-9]/g, ''));
            const sRow = parseInt(startCell.replace(/[A-Z]/g, ''));
            const eCol = colLetterToNum(endCell.replace(/[0-9]/g, ''));
            const eRow = parseInt(endCell.replace(/[A-Z]/g, ''));
            for (let r = sRow; r <= eRow; r++) {
              for (let c = sCol; c <= eCol; c++) {
                ws.getCell(r, c).alignment = alignment;
              }
            }
          } else {
            ws.getCell(cellRef).alignment = alignment;
          }
          editResults.push('Set alignment on ' + cellRef);
          break;
        }

        // SET_BORDERS - Granular border control
        case 'set_borders': {
          const rangeStr = op.range || 'A1';
          const borderStyle = op.style || 'thin';
          const borderColor = op.color || '000000';
          const sides = op.sides || ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];

          const border: any = {};
          for (const side of sides) {
            border[side] = { style: borderStyle, color: { argb: 'FF' + borderColor } };
          }

          if (rangeStr.includes(':')) {
            const [startCell, endCell] = rangeStr.split(':');
            const sCol = colLetterToNum(startCell.replace(/[0-9]/g, ''));
            const sRow = parseInt(startCell.replace(/[A-Z]/g, ''));
            const eCol = colLetterToNum(endCell.replace(/[0-9]/g, ''));
            const eRow = parseInt(endCell.replace(/[A-Z]/g, ''));
            for (let r = sRow; r <= eRow; r++) {
              for (let c = sCol; c <= eCol; c++) {
                ws.getCell(r, c).border = border;
              }
            }
          } else {
            ws.getCell(rangeStr).border = border;
          }
          editResults.push('Set borders on ' + rangeStr);
          break;
        }

        default:
          editResults.push('Unknown operation: ' + op.type);
      }
    } catch (error: any) {
      editResults.push('Error in ' + op.type + ': ' + error.message);
    }
  }

  onProgress?.('Saving spreadsheet...', 80);

  const outputBuffer = await wb.xlsx.writeBuffer();
  writeFileBuffer(filename, Buffer.from(outputBuffer as any));

  onProgress?.('Spreadsheet updated!', 100);
  return { success: true, message: editResults.join('\n'), filename, preview: 'Applied ' + operations.length + ' operation(s)', progress: 100 };
}

// ═══════════════════════════════════════════════════════════════════════════════════
// EXCEL HELPERS
// ═══════════════════════════════════════════════════════════════════════════════════

function applyCellStyle(cell: ExcelJS.Cell, style: any): void {
  if (style.bold !== undefined) cell.font = { ...cell.font as any, bold: style.bold };
  if (style.italic !== undefined) cell.font = { ...cell.font as any, italic: style.italic };
  if (style.underline !== undefined) cell.font = { ...cell.font as any, underline: style.underline };
  if (style.fontColor) cell.font = { ...cell.font as any, color: { argb: 'FF' + style.fontColor } };
  if (style.fontSize) cell.font = { ...cell.font as any, size: style.fontSize };
  if (style.fontName) cell.font = { ...cell.font as any, name: style.fontName };
  if (style.bgColor) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + style.bgColor } };
  if (style.borderColor) {
    cell.border = {
      top: { style: style.borderStyle || 'thin', color: { argb: 'FF' + style.borderColor } },
      left: { style: style.borderStyle || 'thin', color: { argb: 'FF' + style.borderColor } },
      bottom: { style: style.borderStyle || 'thin', color: { argb: 'FF' + style.borderColor } },
      right: { style: style.borderStyle || 'thin', color: { argb: 'FF' + style.borderColor } },
    };
  }
  if (style.alignment) cell.alignment = { horizontal: style.alignment as any, vertical: style.verticalAlignment || 'middle' };
  if (style.wrapText) cell.alignment = { ...cell.alignment as any, wrapText: true };
}

function colLetterToNum(col: string): number {
  return col.split('').reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0);
}

// ═══════════════════════════════════════════════════════════════════════════════════
// ANALYZE FILE
// ═══════════════════════════════════════════════════════════════════════════════════

async function analyzeFile(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  const filename = params.filename || '';
  if (!filename) return { success: false, message: 'Filename is required' };
  if (!fileExists(filename)) return { success: false, message: 'File "' + filename + '" not found' };

  onProgress?.('Reading file...', 30);

  if (filename.endsWith('.xlsx')) {
    const buffer = readFileBuffer(filename);
    if (!buffer) return { success: false, message: 'Could not read ' + filename };

    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer as any);

      const summaries = wb.worksheets.map(ws => ({
        name: ws.name,
        rowCount: ws.rowCount - 1,
        columnCount: ws.columnCount,
      }));

      const userMsg = 'Analyzed "' + filename + '" - ' + wb.worksheets.length + ' sheet(s), ' +
        summaries.map(s => s.rowCount + ' rows in "' + s.name + '"').join(', ');

      return { success: true, message: userMsg, filename, preview: wb.worksheets.length + ' sheet(s)', progress: 100, data: summaries };
    } catch (e: any) {
      return { success: false, message: 'Failed to analyze: ' + e.message };
    }
  }

  if (filename.endsWith('.docx')) {
    return await readDocumentContent({ filename }, onProgress);
  }

  return { success: true, message: 'File: ' + filename, filename, progress: 100 };
}

// ═══════════════════════════════════════════════════════════════════════════════════
// FILE MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════════

function listFiles(): ToolResult {
  const files = listStorageFiles();
  if (files.length === 0) return { success: true, message: 'No files found.' };
  return { success: true, message: 'Found ' + files.length + ' file(s):\n' + files.map(f => '- ' + f.name + ' (' + formatFileSize(f.size) + ')').join('\n') };
}

function deleteFileAction(filename?: string): ToolResult {
  if (!filename) return { success: false, message: 'Filename is required' };
  if (!fileExists(filename)) return { success: false, message: 'File not found' };
  deleteStorageFile(filename);
  return { success: true, message: 'Deleted: ' + filename };
}

function renameFileAction(filename?: string, newFilename?: string): ToolResult {
  if (!filename || !newFilename) return { success: false, message: 'Both filenames required' };
  if (!fileExists(filename)) return { success: false, message: 'File not found' };
  fs.renameSync(getFilePath(filename), getFilePath(newFilename));
  return { success: true, message: 'Renamed ' + filename + ' to ' + newFilename };
}

// ═══════════════════════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════════════════════

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function escapeXmlForTools(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// ═══════════════════════════════════════════════════════════════════════════════════
// NEW TOOL IMPLEMENTATIONS
// ═══════════════════════════════════════════════════════════════════════════════════

/** GET_PARAGRAPH_INDEX — return all blocks with their 0-based index for precise editing */
async function getParagraphIndex(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.docx')) filename += '.docx';
  if (!fileExists(filename)) return { success: false, message: 'File not found: ' + filename };

  onProgress?.('Reading document structure...', 40);

  const blocks = await getIndexedParagraphs(filename);

  const lines = blocks.map(b => {
    let tag = '[P]';
    if (b.type === 'heading') tag = `[H${b.level}]`;
    else if (b.type === 'table') tag = '[TABLE]';
    else if (b.type === 'image') tag = '[IMAGE]';
    const preview = b.isEmpty ? '(empty)' : b.text.slice(0, 80) + (b.text.length > 80 ? '...' : '');
    return `[${b.index}] ${tag} ${preview}`;
  });

  const output = `Document: ${filename}\n${blocks.length} blocks:\n\n` + lines.join('\n');
  onProgress?.('Done!', 100);

  return {
    success: true,
    message: output,
    filename,
    preview: blocks.length + ' blocks indexed',
    progress: 100,
    data: blocks,
  };
}

/** READ_SPREADSHEET_FULL — return all cells, formulas, and structure */
async function readSpreadsheetFullAction(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.xlsx')) filename += '.xlsx';
  if (!fileExists(filename)) return { success: false, message: 'File not found: ' + filename };

  onProgress?.('Reading spreadsheet...', 30);

  const sheets = await readSpreadsheetFull(filename, params.sheet);
  const summary = sheets.map(s => {
    const cellLines: string[] = [];
    s.grid.forEach((row, ri) => {
      row.forEach((val, ci) => {
        if (val && val.trim()) {
          const addr = colLetterFromIndex(ci) + (ri);
          const formula = s.formulas[addr] ? ` [=${s.formulas[addr].replace(/^=/, '')}]` : '';
          cellLines.push(`  ${addr}: ${val}${formula}`);
        }
      });
    });
    return `Sheet "${s.name}" (${s.rowCount} rows × ${s.colCount} cols)\nHeaders: ${s.headers.join(', ')}\nCells:\n${cellLines.slice(0, 100).join('\n')}${cellLines.length > 100 ? '\n  ... (' + (cellLines.length - 100) + ' more cells)' : ''}`;
  }).join('\n\n---\n\n');

  onProgress?.('Done!', 100);
  return {
    success: true,
    message: summary,
    filename,
    preview: sheets.map(s => s.name + ': ' + s.rowCount + ' rows').join(', '),
    progress: 100,
    data: sheets,
  };
}

function colLetterFromIndex(ci: number): string {
  let col = ci; // 0-based
  let s = '';
  col++; // make 1-based
  while (col > 0) {
    col--;
    s = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

/** BULK_UPDATE_CELLS — update many cells at once with values, formulas, and styles */
async function bulkUpdateCellsAction(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.xlsx')) filename += '.xlsx';
  if (!fileExists(filename)) return { success: false, message: 'File not found: ' + filename };

  const sheetName = params.sheet || params.sheetName;
  const updates = params.updates || [];

  if (!updates.length) return { success: false, message: 'No updates provided' };

  onProgress?.('Updating ' + updates.length + ' cells...', 40);

  const result = await bulkUpdateCells(filename, sheetName, updates);
  onProgress?.('Done!', 100);

  return {
    success: true,
    message: `Updated ${result.updated} cells in ${filename}`,
    filename,
    preview: result.updated + ' cells updated',
    progress: 100,
  };
}

/** GET_DOCUMENT_XML — return raw Word XML for inspection */
async function getDocumentXmlAction(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.docx')) filename += '.docx';
  if (!fileExists(filename)) return { success: false, message: 'File not found: ' + filename };

  onProgress?.('Reading XML...', 50);
  const xml = await getDocumentXml(filename);
  // Truncate for readability — show first 8000 chars
  const truncated = xml.length > 8000 ? xml.slice(0, 8000) + '\n... [truncated, ' + xml.length + ' total chars]' : xml;
  onProgress?.('Done!', 100);

  return {
    success: true,
    message: truncated,
    filename,
    preview: xml.length + ' chars of XML',
    progress: 100,
  };
}

/** SET_DOCUMENT_XML — write raw Word XML (full control) */
async function setDocumentXmlAction(params: any, onProgress?: (s: string, p: number) => void): Promise<ToolResult> {
  let filename = params.filename || '';
  if (!filename.endsWith('.docx')) filename += '.docx';
  if (!fileExists(filename)) return { success: false, message: 'File not found: ' + filename };

  const xml = params.xml || params.body_xml || '';
  if (!xml) return { success: false, message: 'No XML provided' };

  onProgress?.('Writing document XML...', 50);
  await setDocumentXml(filename, xml);
  onProgress?.('Done!', 100);

  return {
    success: true,
    message: 'Set document XML for ' + filename + ' (' + xml.length + ' chars)',
    filename,
    preview: 'Raw XML applied',
    progress: 100,
  };
}

/** Helper: build XML string from an insert/replace operation's content fields */
function buildContentXml(op: any): string {
  // If raw XML provided, use it directly
  if (op.content_xml) return op.content_xml;

  let xml = '';

  // Heading
  if (op.heading || (op.type_content === 'heading' && op.text)) {
    const text = op.heading || op.text;
    const level = op.level || 2;
    const color = op.color || '1E40AF';
    xml += coloredHeadingXml(text, level, color);
  }

  // Paragraph text
  if (op.content || op.text) {
    const text = op.content || op.text;
    if (op.bold || op.italic || op.color || op.fontSize || op.alignment || op.font || op.underline || op.highlight) {
      xml += styledParagraphXml(text, {
        bold: op.bold, italic: op.italic, underline: op.underline,
        color: op.color, font: op.font, fontSize: op.fontSize,
        alignment: op.alignment, highlight: op.highlight,
      });
    } else {
      xml += paragraphXml(text);
    }
  }

  // Bullet list
  if (op.bullets && op.bullets.length > 0) xml += bulletListXml(op.bullets);
  if (op.items && op.items.length > 0) xml += bulletListXml(op.items);

  // Table
  if (op.table && op.table.headers) xml += tableXml(op.table.headers, op.table.rows || []);

  // Page break
  if (op.page_break) xml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

  // Separator
  if (op.separator) xml += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E0"/></w:pBdr></w:pPr></w:p>';

  return xml || paragraphXml(op.text || op.content || '');
}

