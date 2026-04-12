import JSZip from 'jszip';
import mammoth from 'mammoth';
import { readFileBuffer, writeFileBuffer } from './file-storage';

// ─── TYPES ────────────────────────────────────────────────────────────────────

export interface DocElement {
  id: string;
  type: 'heading' | 'paragraph' | 'list' | 'table';
  level?: number;
  text?: string;
  items?: string[];
  headers?: string[];
  rows?: string[][];
}

export interface DocumentModel {
  title: string;
  elements: DocElement[];
  fullText: string;
  wordCount: number;
}

export interface DocxZip {
  zip: JSZip;
  changed: boolean;
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

export function escapeXml(text: string): string {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ').trim();
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function loadDocx(filename: string): Promise<DocxZip> {
  const buffer = readFileBuffer(filename);
  if (!buffer) throw new Error('Could not read file: ' + filename);
  const zip = await JSZip.loadAsync(buffer);
  return { zip, changed: false };
}

async function saveDocx(dx: DocxZip, filename: string): Promise<void> {
  if (!dx.changed) return;
  const buf = await dx.zip.generateAsync({ type: 'nodebuffer' });
  writeFileBuffer(filename, buf);
}

async function getXml(dx: DocxZip, path: string): Promise<string | null> {
  const file = dx.zip.file(path);
  if (!file) return null;
  return file.async('text');
}

function setXml(dx: DocxZip, path: string, content: string): void {
  dx.zip.file(path, content);
  dx.changed = true;
}

// ─── PARAGRAPH-LEVEL TEXT EXTRACTION ─────────────────────────────────────────
//
// CORE FIX: Word splits text across multiple <w:r> runs.
// "Hello World" may be stored as: <w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r>
// We MUST extract all <w:t> content concatenated to find/replace text reliably.

function extractParaText(paraXml: string): string {
  const pieces: string[] = [];
  const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(paraXml)) !== null) pieces.push(m[1]);
  return pieces.join('');
}

function extractPPr(paraXml: string): string {
  const m = paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : '';
}

function extractFirstRunRPr(paraXml: string): string {
  const m = paraXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

function buildRun(rPr: string, text: string): string {
  const attr = text.includes(' ') || text.startsWith(' ') || text.endsWith(' ') ? ' xml:space="preserve"' : '';
  return `<w:r>${rPr}<w:t${attr}>${escapeXml(text)}</w:t></w:r>`;
}

// ─── TABLE HELPERS ────────────────────────────────────────────────────────────

interface Range { start: number; end: number; content: string }

function findTables(xml: string): Range[] {
  const results: Range[] = [];
  let depth = 0, start = -1, i = 0;
  while (i < xml.length) {
    if (xml.startsWith('<w:tbl>', i) || xml.startsWith('<w:tbl ', i)) {
      if (depth === 0) start = i;
      depth++;
    } else if (xml.startsWith('</w:tbl>', i)) {
      depth--;
      if (depth === 0 && start !== -1) {
        const end = i + '</w:tbl>'.length;
        results.push({ start, end, content: xml.slice(start, end) });
        start = -1;
      }
    }
    i++;
  }
  return results;
}

function findTableRows(tableXml: string): Range[] {
  const results: Range[] = [];
  let i = 0;
  while (i < tableXml.length) {
    let rowStart = tableXml.indexOf('<w:tr>', i);
    const rowStart2 = tableXml.indexOf('<w:tr ', i);
    if (rowStart === -1 && rowStart2 === -1) break;
    if (rowStart === -1) rowStart = rowStart2;
    else if (rowStart2 !== -1) rowStart = Math.min(rowStart, rowStart2);
    const rowEnd = tableXml.indexOf('</w:tr>', rowStart);
    if (rowEnd === -1) break;
    const end = rowEnd + '</w:tr>'.length;
    results.push({ start: rowStart, end, content: tableXml.slice(rowStart, end) });
    i = end;
  }
  return results;
}

function findTableCells(rowXml: string): Range[] {
  const results: Range[] = [];
  let i = 0;
  while (i < rowXml.length) {
    const s = rowXml.indexOf('<w:tc>', i);
    if (s === -1) break;
    const e = rowXml.indexOf('</w:tc>', s);
    if (e === -1) break;
    const end = e + '</w:tc>'.length;
    results.push({ start: s, end, content: rowXml.slice(s, end) });
    i = end;
  }
  return results;
}

// ─── READ DOCUMENT ────────────────────────────────────────────────────────────

export async function readDocx(filename: string): Promise<DocumentModel> {
  const buffer = readFileBuffer(filename);
  if (!buffer) throw new Error('Could not read file: ' + filename);
  const textResult = await mammoth.extractRawText({ buffer });
  const fullText = textResult.value.trim();
  const htmlResult = await mammoth.convertToHtml({ buffer });
  const elements = parseHtmlToElements(htmlResult.value);
  const title = elements.find(e => e.type === 'heading' && e.level === 1)?.text || filename.replace('.docx', '');
  return { title, elements, fullText, wordCount: fullText.split(/\s+/).filter(w => w.length > 0).length };
}

function parseHtmlToElements(html: string): DocElement[] {
  const elements: DocElement[] = [];
  let id = 0;
  let m: RegExpExecArray | null;

  const headingRe = /<h([1-6])[^>]*>([\s\S]*?)<\/h[1-6]>/gi;
  while ((m = headingRe.exec(html)) !== null)
    elements.push({ id: 'h' + id++, type: 'heading', level: parseInt(m[1]), text: stripHtml(m[2]) });

  const paraRe = /<p[^>]*>([\s\S]*?)<\/p>/gi;
  while ((m = paraRe.exec(html)) !== null) {
    const text = stripHtml(m[1]);
    if (text.trim()) elements.push({ id: 'p' + id++, type: 'paragraph', text });
  }

  const tableRe = /<table[^>]*>([\s\S]*?)<\/table>/gi;
  while ((m = tableRe.exec(html)) !== null) {
    const tableHtml = m[1];
    const headers: string[] = [];
    const rows: string[][] = [];
    const thRe = /<th[^>]*>([\s\S]*?)<\/th>/gi;
    let tm: RegExpExecArray | null;
    while ((tm = thRe.exec(tableHtml)) !== null) headers.push(stripHtml(tm[1]));
    const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    while ((tm = trRe.exec(tableHtml)) !== null) {
      if (tm[1].includes('<th')) continue;
      const row: string[] = [];
      const tdRe = /<td[^>]*>([\s\S]*?)<\/td>/gi;
      let tdm: RegExpExecArray | null;
      while ((tdm = tdRe.exec(tm[1])) !== null) row.push(stripHtml(tdm[1]));
      if (row.length > 0) rows.push(row);
    }
    if (headers.length > 0 || rows.length > 0)
      elements.push({ id: 't' + id++, type: 'table', headers, rows });
  }

  const ulRe = /<ul[^>]*>([\s\S]*?)<\/ul>/gi;
  while ((m = ulRe.exec(html)) !== null) {
    const items: string[] = [];
    const liRe = /<li[^>]*>([\s\S]*?)<\/li>/gi;
    let lm: RegExpExecArray | null;
    while ((lm = liRe.exec(m[1])) !== null) items.push(stripHtml(lm[1]));
    if (items.length > 0) elements.push({ id: 'l' + id++, type: 'list', items });
  }

  return elements;
}

export async function docxToHtml(filename: string): Promise<string> {
  const buffer = readFileBuffer(filename);
  if (!buffer) throw new Error('Could not read file: ' + filename);
  const result = await mammoth.convertToHtml({ buffer });
  return result.value;
}

// ─── TEXT REPLACEMENT (Paragraph-level — THE CORE FIX) ───────────────────────
//
// Instead of matching within a single <w:t>, we process each <w:p> paragraph:
//   1. Extract FULL concatenated text of all runs
//   2. If search text found → rebuild paragraph with a clean single run
//   3. This handles ALL split-run scenarios

export async function replaceTextInDocx(
  filename: string, find: string, replace: string, caseSensitive: boolean = false
): Promise<{ replacements: number }> {
  const dx = await loadDocx(filename);
  let total = 0;

  const paths = ['word/document.xml'];
  for (let i = 1; i <= 20; i++) paths.push(`word/header${i}.xml`, `word/footer${i}.xml`);

  for (const path of paths) {
    const xml = await getXml(dx, path);
    if (!xml) continue;
    const { newContent, count } = replaceParagraphLevel(xml, find, replace, caseSensitive);
    if (count > 0) { setXml(dx, path, newContent); total += count; }
  }

  if (total > 0) await saveDocx(dx, filename);
  return { replacements: total };
}

function replaceParagraphLevel(xml: string, find: string, replace: string, caseSensitive: boolean): { newContent: string; count: number } {
  let total = 0;
  const flags = caseSensitive ? 'g' : 'gi';
  const re = new RegExp(escapeRegex(find), flags);

  const newContent = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    const text = extractParaText(paraXml);
    const matches = text.match(re);
    if (!matches) return paraXml;
    total += matches.length;
    const newText = text.replace(re, replace);
    const pPr = extractPPr(paraXml);
    const rPr = extractFirstRunRPr(paraXml);
    return `<w:p>${pPr}${buildRun(rPr, newText)}</w:p>`;
  });

  return { newContent, count: total };
}

// ─── ADD CONTENT ──────────────────────────────────────────────────────────────

export async function addToDocx(filename: string, contentXml: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const bodyClose = xml.lastIndexOf('</w:body>');
  if (bodyClose === -1) throw new Error('No </w:body>');
  xml = xml.slice(0, bodyClose) + contentXml + xml.slice(bodyClose);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

// ─── XML GENERATORS ───────────────────────────────────────────────────────────

export function headingXml(text: string, level: number = 2): string {
  return `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

export function coloredHeadingXml(text: string, level: number = 2, color: string = '1E40AF'): string {
  const sizes: Record<number, number> = { 1: 48, 2: 36, 3: 32, 4: 28, 5: 24, 6: 22 };
  const sz = sizes[level] || 32;
  return `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:color w:val="${color}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

export function paragraphXml(text: string): string {
  return text.split('\n').map(line => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`).join('');
}

export function styledParagraphXml(text: string, options: {
  bold?: boolean; italic?: boolean; underline?: boolean; strikethrough?: boolean;
  color?: string; font?: string; fontSize?: number;
  alignment?: 'left' | 'center' | 'right' | 'justify';
  highlight?: string;
} = {}): string {
  const rPr: string[] = [];
  if (options.bold) rPr.push('<w:b/><w:bCs/>');
  if (options.italic) rPr.push('<w:i/><w:iCs/>');
  if (options.underline) rPr.push('<w:u w:val="single"/>');
  if (options.strikethrough) rPr.push('<w:strike/>');
  if (options.color) rPr.push(`<w:color w:val="${options.color.replace('#', '')}"/>`);
  if (options.font) rPr.push(`<w:rFonts w:ascii="${options.font}" w:hAnsi="${options.font}" w:cs="${options.font}"/>`);
  if (options.fontSize) rPr.push(`<w:sz w:val="${options.fontSize * 2}"/><w:szCs w:val="${options.fontSize * 2}"/>`);
  if (options.highlight) rPr.push(`<w:highlight w:val="${options.highlight}"/>`);
  const pPr: string[] = [];
  if (options.alignment) {
    const jcMap: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' };
    pPr.push(`<w:jc w:val="${jcMap[options.alignment] || options.alignment}"/>`);
  }
  return `<w:p>${pPr.length ? '<w:pPr>' + pPr.join('') + '</w:pPr>' : ''}<w:r>${rPr.length ? '<w:rPr>' + rPr.join('') + '</w:rPr>' : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

export function bulletListXml(items: string[]): string {
  return items.map(item =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(item)}</w:t></w:r></w:p>`
  ).join('');
}

export function tableXml(headers: string[], rows: string[][]): string {
  const hRow = `<w:tr>${headers.map(h => `<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2D3748"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">${escapeXml(h)}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`;
  const dRows = rows.map((row, ri) => {
    const fill = ri % 2 === 0 ? 'F8FAFC' : 'FFFFFF';
    return `<w:tr>${row.map(cell => `<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${escapeXml(String(cell))}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`;
  }).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/><w:left w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/><w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/></w:tblBorders><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>${hRow}${dRows}</w:tbl>`;
}

export function styledTableXml(headers: string[], rows: string[][]): string { return tableXml(headers, rows); }

// ─── APPLY TEXT STYLE (paragraph-level fix) ───────────────────────────────────

export async function applyTextStyle(
  filename: string, find: string,
  options: { bold?: boolean; italic?: boolean; underline?: boolean; color?: string; font?: string; fontSize?: number; strikethrough?: boolean }
): Promise<number> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  let count = 0;

  const newXml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    const text = extractParaText(paraXml);
    if (!text.toLowerCase().includes(find.toLowerCase())) return paraXml;
    count++;
    const pPr = extractPPr(paraXml);
    const rPrParts: string[] = [];
    if (options.bold) rPrParts.push('<w:b/><w:bCs/>');
    if (options.italic) rPrParts.push('<w:i/><w:iCs/>');
    if (options.underline) rPrParts.push('<w:u w:val="single"/>');
    if (options.strikethrough) rPrParts.push('<w:strike/>');
    if (options.color) rPrParts.push(`<w:color w:val="${options.color.replace('#', '')}"/>`);
    if (options.font) rPrParts.push(`<w:rFonts w:ascii="${options.font}" w:hAnsi="${options.font}" w:cs="${options.font}"/>`);
    if (options.fontSize) rPrParts.push(`<w:sz w:val="${options.fontSize * 2}"/><w:szCs w:val="${options.fontSize * 2}"/>`);
    const rPr = rPrParts.length ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';

    const lower = text.toLowerCase();
    const lowerFind = find.toLowerCase();
    const idx = lower.indexOf(lowerFind);
    const before = text.slice(0, idx);
    const matched = text.slice(idx, idx + find.length);
    const after = text.slice(idx + find.length);

    let result = `<w:p>${pPr}`;
    if (before) result += buildRun('', before);
    result += `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(matched)}</w:t></w:r>`;
    if (after) result += buildRun('', after);
    return result + '</w:p>';
  });

  if (count > 0) { setXml(dx, 'word/document.xml', newXml); await saveDocx(dx, filename); }
  return count;
}

// ─── DELETE ELEMENT (paragraph-level fix) ─────────────────────────────────────

export async function deleteElementByContent(filename: string, search: string): Promise<number> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  let count = 0;
  const newXml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    if (extractParaText(paraXml).toLowerCase().includes(search.toLowerCase())) { count++; return ''; }
    return paraXml;
  });
  if (count > 0) { setXml(dx, 'word/document.xml', newXml); await saveDocx(dx, filename); }
  return count;
}

// ─── HEADER / FOOTER ──────────────────────────────────────────────────────────

export async function addHeaderToDocx(filename: string, text: string): Promise<void> {
  const dx = await loadDocx(filename);
  const rId = 'rIdHdr1';
  const hdrXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:hdr>`;
  dx.zip.file('word/header1.xml', hdrXml);

  let relsXml = await getXml(dx, 'word/_rels/document.xml.rels') || '';
  if (!relsXml.includes('header1.xml')) {
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`);
    setXml(dx, 'word/_rels/document.xml.rels', relsXml);
  }

  let docXml = await getXml(dx, 'word/document.xml');
  if (docXml && !docXml.includes(`r:id="${rId}"`)) {
    docXml = docXml.replace('<w:sectPr>', `<w:sectPr><w:headerReference w:type="default" r:id="${rId}"/>`);
    setXml(dx, 'word/document.xml', docXml);
  }

  let ctXml = await getXml(dx, '[Content_Types].xml');
  if (ctXml && !ctXml.includes('header1.xml')) {
    ctXml = ctXml.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
    setXml(dx, '[Content_Types].xml', ctXml);
  }

  dx.changed = true;
  await saveDocx(dx, filename);
}

export async function addFooterToDocx(filename: string, text: string): Promise<void> {
  const dx = await loadDocx(filename);
  const rId = 'rIdFtr1';
  const ftrXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:ftr>`;
  dx.zip.file('word/footer1.xml', ftrXml);

  let relsXml = await getXml(dx, 'word/_rels/document.xml.rels') || '';
  if (!relsXml.includes('footer1.xml')) {
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`);
    setXml(dx, 'word/_rels/document.xml.rels', relsXml);
  }

  let docXml = await getXml(dx, 'word/document.xml');
  if (docXml && !docXml.includes(`r:id="${rId}"`)) {
    docXml = docXml.replace('<w:sectPr>', `<w:sectPr><w:footerReference w:type="default" r:id="${rId}"/>`);
    setXml(dx, 'word/document.xml', docXml);
  }

  let ctXml = await getXml(dx, '[Content_Types].xml');
  if (ctXml && !ctXml.includes('footer1.xml')) {
    ctXml = ctXml.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
    setXml(dx, '[Content_Types].xml', ctXml);
  }

  dx.changed = true;
  await saveDocx(dx, filename);
}

export async function removeHeaderFromDocx(filename: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (xml) { xml = xml.replace(/<w:headerReference[^/]*\/>/g, ''); setXml(dx, 'word/document.xml', xml); await saveDocx(dx, filename); }
}

export async function removeFooterFromDocx(filename: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (xml) { xml = xml.replace(/<w:footerReference[^/]*\/>/g, ''); setXml(dx, 'word/document.xml', xml); await saveDocx(dx, filename); }
}

// ─── DOCUMENT LAYOUT ──────────────────────────────────────────────────────────

export async function setMargins(filename: string, top: number, bottom: number, left: number, right: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const pgMar = `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="720" w:footer="720" w:gutter="0"/>`;
  xml = xml.includes('<w:pgMar') ? xml.replace(/<w:pgMar[^/]*\/>/, pgMar) : xml.replace('</w:sectPr>', pgMar + '</w:sectPr>');
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function setOrientation(filename: string, orientation: 'portrait' | 'landscape'): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const w = orientation === 'landscape' ? 15840 : 12240;
  const h = orientation === 'landscape' ? 12240 : 15840;
  const orient = orientation === 'landscape' ? ' w:orient="landscape"' : '';
  const pgSz = `<w:pgSz w:w="${w}" w:h="${h}"${orient}/>`;
  xml = xml.includes('<w:pgSz') ? xml.replace(/<w:pgSz[^/]*\/>/, pgSz) : xml.replace('</w:sectPr>', pgSz + '</w:sectPr>');
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function setDocumentFont(filename: string, font: string): Promise<void> {
  const dx = await loadDocx(filename);
  let stylesXml = await getXml(dx, 'word/styles.xml');
  if (stylesXml) {
    const defaultRPr = `<w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/></w:rPr></w:rPrDefault>`;
    stylesXml = stylesXml.includes('<w:rPrDefault>')
      ? stylesXml.replace(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/, defaultRPr)
      : stylesXml.replace('<w:docDefaults>', '<w:docDefaults>' + defaultRPr);
    setXml(dx, 'word/styles.xml', stylesXml);
    await saveDocx(dx, filename);
  }
}

export async function setLineSpacing(filename: string, spacing: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const lineVal = Math.round(spacing * 240);
  const spacingXml = `<w:spacing w:line="${lineVal}" w:lineRule="auto"/>`;
  const newXml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    const pPr = extractPPr(paraXml);
    const newPPr = pPr
      ? (pPr.includes('<w:spacing') ? pPr.replace(/<w:spacing[^/]*\/>/, spacingXml) : pPr.replace('</w:pPr>', spacingXml + '</w:pPr>'))
      : `<w:pPr>${spacingXml}</w:pPr>`;
    return pPr ? paraXml.replace(pPr, newPPr) : paraXml.replace('<w:p>', '<w:p>' + newPPr);
  });
  setXml(dx, 'word/document.xml', newXml);
  await saveDocx(dx, filename);
}

export async function setParagraphSpacingToDocx(filename: string, before: number, after: number, lineSpacing?: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  let spacingXml = `<w:spacing w:before="${before}" w:after="${after}"`;
  if (lineSpacing) spacingXml += ` w:line="${lineSpacing}" w:lineRule="auto"`;
  spacingXml += '/>';
  const newXml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    const pPr = extractPPr(paraXml);
    const newPPr = pPr
      ? (pPr.includes('<w:spacing') ? pPr.replace(/<w:spacing[^/]*\/>/, spacingXml) : pPr.replace('</w:pPr>', spacingXml + '</w:pPr>'))
      : `<w:pPr>${spacingXml}</w:pPr>`;
    return pPr ? paraXml.replace(pPr, newPPr) : paraXml.replace('<w:p>', '<w:p>' + newPPr);
  });
  setXml(dx, 'word/document.xml', newXml);
  await saveDocx(dx, filename);
}

export async function setFirstLineIndentToDocx(filename: string, twips: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const indentXml = `<w:ind w:firstLine="${twips}"/>`;
  const newXml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    const pPr = extractPPr(paraXml);
    const newPPr = pPr
      ? (pPr.includes('<w:ind') ? pPr.replace(/<w:ind[^/]*\/>/, indentXml) : pPr.replace('</w:pPr>', indentXml + '</w:pPr>'))
      : `<w:pPr>${indentXml}</w:pPr>`;
    return pPr ? paraXml.replace(pPr, newPPr) : paraXml.replace('<w:p>', '<w:p>' + newPPr);
  });
  setXml(dx, 'word/document.xml', newXml);
  await saveDocx(dx, filename);
}

export async function setColumnsInDocx(filename: string, count: number, spacing: number = 720, separator: boolean = false): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const sepAttr = separator ? ' w:sep="1"' : '';
  const colsXml = `<w:cols w:num="${count}" w:space="${spacing}"${sepAttr}/>`;
  xml = xml.includes('<w:cols') ? xml.replace(/<w:cols[^/]*\/>/, colsXml) : xml.replace('</w:sectPr>', colsXml + '</w:sectPr>');
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function addSectionBreakToDocx(filename: string, type: string = 'nextPage'): Promise<void> {
  const types: Record<string, string> = { nextPage: 'nextPage', continuous: 'continuous', evenPage: 'evenPage', oddPage: 'oddPage' };
  const t = types[type] || 'nextPage';
  await addToDocx(filename, `<w:p><w:pPr><w:sectPr><w:type w:val="${t}"/></w:sectPr></w:pPr></w:p>`);
}

export async function addColumnBreakToDocx(filename: string): Promise<void> {
  await addToDocx(filename, `<w:p><w:r><w:br w:type="column"/></w:r></w:p>`);
}

// ─── TABLE OPERATIONS ─────────────────────────────────────────────────────────

export async function addTableRowToDocx(filename: string, tableIndex: number, data: string[]): Promise<number> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found (${tables.length} tables exist)`);
  const { start, end } = tables[tableIndex];
  const tbl = xml.slice(start, end);
  const ins = tbl.lastIndexOf('</w:tbl>');
  const rowXml = `<w:tr>${data.map(c => `<w:tc><w:p><w:r><w:t xml:space="preserve">${escapeXml(String(c))}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`;
  const newTbl = tbl.slice(0, ins) + rowXml + tbl.slice(ins);
  xml = xml.slice(0, start) + newTbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
  return data.length;
}

export async function updateTableCellInDocx(filename: string, tableIndex: number, rowIndex: number, colIndex: number, value: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const rows = findTableRows(tbl);
  if (rowIndex >= rows.length) throw new Error(`Row ${rowIndex} not found`);
  const row = rows[rowIndex];
  const cells = findTableCells(row.content);
  if (colIndex >= cells.length) throw new Error(`Col ${colIndex} not found`);
  const cell = cells[colIndex];
  const newCell = `<w:tc><w:p><w:r><w:t xml:space="preserve">${escapeXml(value)}</w:t></w:r></w:p></w:tc>`;
  const newRow = row.content.slice(0, cell.start) + newCell + row.content.slice(cell.end);
  tbl = tbl.slice(0, row.start) + newRow + tbl.slice(row.end);
  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function deleteTableRowFromDocx(filename: string, tableIndex: number, rowIndex: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const rows = findTableRows(tbl);
  if (rowIndex >= rows.length) throw new Error(`Row ${rowIndex} not found`);
  tbl = tbl.slice(0, rows[rowIndex].start) + tbl.slice(rows[rowIndex].end);
  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function deleteTableFromDocx(filename: string, tableIndex: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  xml = xml.slice(0, tables[tableIndex].start) + xml.slice(tables[tableIndex].end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function formatTableCellInDocx(
  filename: string, tableIndex: number, rowIndex: number, colIndex: number,
  options: { bg?: string; bold?: boolean; color?: string; font?: string; fontSize?: number; align?: string }
): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const rows = findTableRows(tbl);
  if (rowIndex >= rows.length) throw new Error(`Row ${rowIndex} not found`);
  const row = rows[rowIndex];
  const cells = findTableCells(row.content);
  if (colIndex >= cells.length) throw new Error(`Col ${colIndex} not found`);
  const cell = cells[colIndex];
  const cellText = extractParaText(cell.content);
  const tcPr = options.bg ? `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${options.bg.replace('#', '')}"/></w:tcPr>` : '';
  const rPrParts: string[] = [];
  if (options.bold) rPrParts.push('<w:b/><w:bCs/>');
  if (options.color) rPrParts.push(`<w:color w:val="${options.color.replace('#', '')}"/>`);
  if (options.font) rPrParts.push(`<w:rFonts w:ascii="${options.font}" w:hAnsi="${options.font}"/>`);
  if (options.fontSize) rPrParts.push(`<w:sz w:val="${options.fontSize * 2}"/>`);
  const rPr = rPrParts.length ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
  const pPrParts: string[] = [];
  if (options.align) { const m: Record<string,string> = {left:'left',center:'center',right:'right',justify:'both'}; pPrParts.push(`<w:jc w:val="${m[options.align] || options.align}"/>`); }
  const pPr = pPrParts.length ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : '';
  const newCell = `<w:tc>${tcPr}<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(cellText)}</w:t></w:r></w:p></w:tc>`;
  const newRow = row.content.slice(0, cell.start) + newCell + row.content.slice(cell.end);
  tbl = tbl.slice(0, row.start) + newRow + tbl.slice(row.end);
  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function addColumnToTable(filename: string, tableIndex: number, header: string, values: string[]): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const rows = findTableRows(tbl);
  let offset = 0;
  rows.forEach((row, ri) => {
    const val = ri === 0 ? header : (values[ri - 1] ?? '');
    const isH = ri === 0;
    const newCell = isH
      ? `<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2D3748"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">${escapeXml(val)}</w:t></w:r></w:p></w:tc>`
      : `<w:tc><w:p><w:r><w:t xml:space="preserve">${escapeXml(val)}</w:t></w:r></w:p></w:tc>`;
    const ins = row.start + offset + (tbl.slice(row.start + offset, row.end + offset)).lastIndexOf('</w:tr>');
    tbl = tbl.slice(0, ins) + newCell + tbl.slice(ins);
    offset += newCell.length;
  });
  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function deleteColumnFromTable(filename: string, tableIndex: number, colIndex: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const rows = findTableRows(tbl);
  let offset = 0;
  rows.forEach((row) => {
    const rowContent = tbl.slice(row.start + offset, row.end + offset);
    const cells = findTableCells(rowContent);
    if (colIndex < cells.length) {
      const cell = cells[colIndex];
      const newRow = rowContent.slice(0, cell.start) + rowContent.slice(cell.end);
      tbl = tbl.slice(0, row.start + offset) + newRow + tbl.slice(row.end + offset);
      offset += newRow.length - rowContent.length;
    }
  });
  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function setTableWidthToDocx(filename: string, tableIndex: number, width: number, widthType: string = 'dxa', alignment: string = 'center'): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const m = tbl.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/);
  if (m) {
    let newTblPr = m[0]
      .replace(/<w:tblW[^/]*\/>/, `<w:tblW w:w="${width}" w:type="${widthType}"/>`)
      .replace(/<w:jc[^/]*\/>/g, '');
    newTblPr = newTblPr.replace('</w:tblPr>', `<w:jc w:val="${alignment}"/></w:tblPr>`);
    tbl = tbl.replace(m[0], newTblPr);
  }
  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function setTableColumnWidthsToDocx(filename: string, tableIndex: number, widths: number[]): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);
  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const tblGrid = '<w:tblGrid>' + widths.map(w => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid>';
  tbl = tbl.includes('<w:tblGrid>') ? tbl.replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, tblGrid) : tbl.replace('</w:tblPr>', '</w:tblPr>' + tblGrid);
  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

// ─── COUNTS ───────────────────────────────────────────────────────────────────

export async function countTablesInDocx(filename: string): Promise<number> {
  const dx = await loadDocx(filename);
  const xml = await getXml(dx, 'word/document.xml');
  return xml ? findTables(xml).length : 0;
}

export async function countImagesInDocx(filename: string): Promise<number> {
  const dx = await loadDocx(filename);
  const xml = await getXml(dx, 'word/document.xml');
  return xml ? (xml.match(/<wp:inline|<wp:anchor/g) || []).length : 0;
}

// ─── GRAPHICS ─────────────────────────────────────────────────────────────────

export async function embedImageInDocx(
  filename: string, imageBuffer: Buffer, width: number = 400, height: number = 300,
  align: string = 'center', wrapStyle: string = 'inline', x?: number, y?: number
): Promise<{ success: boolean; message: string }> {
  const dx = await loadDocx(filename);
  let imageNum = 1;
  const files = Object.keys(dx.zip.files);
  for (const f of files) { const m = f.match(/image(\d+)\.(png|jpg)/); if (m) imageNum = Math.max(imageNum, parseInt(m[1]) + 1); }
  const imageName = `image${imageNum}.png`;
  const rId = `rIdImg${imageNum}`;
  dx.zip.file(`word/media/${imageName}`, imageBuffer);

  let ctXml = await getXml(dx, '[Content_Types].xml');
  if (ctXml && !ctXml.includes('Extension="png"')) {
    ctXml = ctXml.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n  <Default Extension="png" ContentType="image/png"/>');
    setXml(dx, '[Content_Types].xml', ctXml);
  }

  let relsXml = await getXml(dx, 'word/_rels/document.xml.rels');
  if (relsXml && !relsXml.includes(`Target="media/${imageName}"`)) {
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${imageName}"/></Relationships>`);
    setXml(dx, 'word/_rels/document.xml.rels', relsXml);
  }

  const cx = width * 9525, cy = height * 9525;
  const jc = align === 'right' ? 'right' : align === 'left' ? 'left' : 'center';
  const drawingXml = `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${100 + imageNum}" name="${imageName}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${200 + imageNum}" name="${imageName}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

  let docXml = await getXml(dx, 'word/document.xml');
  if (!docXml) throw new Error('No document.xml');
  const bc = docXml.lastIndexOf('</w:body>');
  docXml = docXml.slice(0, bc) + drawingXml + docXml.slice(bc);
  setXml(dx, 'word/document.xml', docXml);
  await saveDocx(dx, filename);
  return { success: true, message: 'Embedded image: ' + imageName };
}

export async function embedImagePositionedInDocx(filename: string, buf: Buffer, w: number, h: number, x: number, y: number, wrap: string = 'square'): Promise<{success:boolean;message:string}> {
  return embedImageInDocx(filename, buf, w, h, 'left', wrap, x, y);
}

export async function deleteImageFromDocx(filename: string, imageIndex: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  let idx = 0;
  const newXml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (para) => {
    if ((para.includes('<wp:inline') || para.includes('<wp:anchor')) && idx++ === imageIndex) return '';
    return para;
  });
  setXml(dx, 'word/document.xml', newXml);
  await saveDocx(dx, filename);
}

export async function embedChartInDocx(filename: string, chartConfig: any, width: number = 500, height: number = 350): Promise<{success:boolean;message:string}> {
  const { renderChart } = await import('./chart-engine');
  return embedImageInDocx(filename, await renderChart(chartConfig), width, height);
}

// ─── WATERMARK / PAGE BORDER / TEXT BOX ──────────────────────────────────────

export async function addWatermarkToDocx(filename: string, text: string, color: string = 'C0C0C0', fontSize: number = 72, font: string = 'Arial'): Promise<void> {
  const dx = await loadDocx(filename);
  const rId = 'rIdWMhdr';
  const hdrXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:452pt;height:113pt;rotation:315;z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="#${color.replace('#','')}" stroked="f"><v:fill on="t" type="solid"/><v:textpath on="t" string="${escapeXml(text)}" style="font-family:&quot;${font}&quot;;font-size:${fontSize}pt;font-weight:bold;font-style:italic"/></v:shape></w:pict></w:r></w:p></w:hdr>`;
  dx.zip.file('word/headerWM.xml', hdrXml);
  let relsXml = await getXml(dx, 'word/_rels/document.xml.rels') || '';
  if (!relsXml.includes('headerWM.xml')) {
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerWM.xml"/></Relationships>`);
    setXml(dx, 'word/_rels/document.xml.rels', relsXml);
  }
  let docXml = await getXml(dx, 'word/document.xml');
  if (docXml && !docXml.includes('headerWM')) {
    docXml = docXml.replace('<w:sectPr>', `<w:sectPr><w:headerReference w:type="default" r:id="${rId}"/>`);
    setXml(dx, 'word/document.xml', docXml);
  }
  let ctXml = await getXml(dx, '[Content_Types].xml');
  if (ctXml && !ctXml.includes('headerWM.xml')) { ctXml = ctXml.replace('</Types>', '<Override PartName="/word/headerWM.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'); setXml(dx, '[Content_Types].xml', ctXml); }
  dx.changed = true;
  await saveDocx(dx, filename);
}

export async function addPageBorderToDocx(filename: string, style: string = 'single', color: string = '000000', size: number = 4): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const c = color.replace('#', '');
  const borderXml = `<w:pgBorders w:offsetFrom="page"><w:top w:val="${style}" w:sz="${size}" w:space="24" w:color="${c}"/><w:left w:val="${style}" w:sz="${size}" w:space="24" w:color="${c}"/><w:bottom w:val="${style}" w:sz="${size}" w:space="24" w:color="${c}"/><w:right w:val="${style}" w:sz="${size}" w:space="24" w:color="${c}"/></w:pgBorders>`;
  xml = xml.includes('<w:pgBorders') ? xml.replace(/<w:pgBorders[\s\S]*?<\/w:pgBorders>/, borderXml) : xml.replace('</w:sectPr>', borderXml + '</w:sectPr>');
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function addTextBoxToDocx(filename: string, text: string, width: number = 200, height: number = 100, fillColor: string = 'FFFFFF', borderColor: string = '000000', fontSize: number = 12, bold: boolean = false, color: string = '000000', alignment: string = 'left', x: number = 0, y: number = 0): Promise<void> {
  const cx = width * 9525, cy = height * 9525, posX = x * 9525, posY = y * 9525;
  const tbXml = `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${posX}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${posY}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="300" name="TextBox"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fillColor.replace('#','')}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${borderColor.replace('#','')}"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="${alignment}"/></w:pPr><w:r><w:rPr>${bold?'<w:b/>':''}<w:sz w:val="${fontSize*2}"/><w:color w:val="${color.replace('#','')}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
  await addToDocx(filename, tbXml);
}

export async function addDropCapParagraphToDocx(filename: string, text: string, lines: number = 3, font: string = '', color: string = ''): Promise<void> {
  const dropCapXml = `<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="${lines}" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr><w:r><w:rPr>${font?`<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`:''}${color?`<w:color w:val="${color.replace('#','')}"/>`:''}${`<w:sz w:val="${lines*24}"/>`}</w:rPr><w:t>${escapeXml(text.charAt(0))}</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">${escapeXml(text.slice(1))}</w:t></w:r></w:p>`;
  await addToDocx(filename, dropCapXml);
}

export async function addTabStopParagraphToDocx(filename: string, text: string, tabStops: Array<{pos:number;align:string}>, fontSize: number = 12, bold: boolean = false, color: string = ''): Promise<void> {
  const tabs = tabStops.map(t => `<w:tab w:val="${t.align}" w:pos="${t.pos}"/>`).join('');
  await addToDocx(filename, `<w:p><w:pPr><w:tabs>${tabs}</w:tabs></w:pPr><w:r><w:rPr>${bold?'<w:b/>':''}<w:sz w:val="${fontSize*2}"/>${color?`<w:color w:val="${color.replace('#','')}"/>`:''}  </w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`);
}

export async function addFormattedPageNumbersToDocx(filename: string, format: string = 'Page {n} of {total}', alignment: string = 'center', showTotal: boolean = true, font: string = '', fontSize: number = 12, color: string = ''): Promise<void> {
  const rPr = `<w:rPr>${font?`<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`:''}${`<w:sz w:val="${fontSize*2}"/>`}${color?`<w:color w:val="${color.replace('#','')}"/>`:''}  </w:rPr>`;
  const ftrXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="${alignment}"/></w:pPr><w:r>${rPr}<w:t xml:space="preserve">Page </w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>${showTotal?`<w:r>${rPr}<w:t xml:space="preserve"> of </w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`:''}</w:p></w:ftr>`;
  const dx = await loadDocx(filename);
  const rId = 'rIdPNFtr';
  dx.zip.file('word/footerPN.xml', ftrXml);
  let relsXml = await getXml(dx, 'word/_rels/document.xml.rels') || '';
  if (!relsXml.includes('footerPN.xml')) { relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerPN.xml"/></Relationships>`); setXml(dx, 'word/_rels/document.xml.rels', relsXml); }
  let docXml = await getXml(dx, 'word/document.xml');
  if (docXml && !docXml.includes('footerPN')) { docXml = docXml.replace('<w:sectPr>', `<w:sectPr><w:footerReference w:type="default" r:id="${rId}"/>`); setXml(dx, 'word/document.xml', docXml); }
  let ctXml = await getXml(dx, '[Content_Types].xml');
  if (ctXml && !ctXml.includes('footerPN.xml')) { ctXml = ctXml.replace('</Types>', '<Override PartName="/word/footerPN.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'); setXml(dx, '[Content_Types].xml', ctXml); }
  dx.changed = true;
  await saveDocx(dx, filename);
}

export async function addHyperlinkToDocx(filename: string, text: string, url: string, color: string = '0563C1', underline: boolean = true): Promise<void> {
  const dx = await loadDocx(filename);
  const rId = `rIdLink${Date.now()}`;
  let relsXml = await getXml(dx, 'word/_rels/document.xml.rels') || '';
  relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml(url)}" TargetMode="External"/></Relationships>`);
  setXml(dx, 'word/_rels/document.xml.rels', relsXml);
  const hlXml = `<w:p><w:hyperlink r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:rPr><w:color w:val="${color.replace('#','')}"/>${underline?'<w:u w:val="single"/>':''}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:hyperlink></w:p>`;
  let docXml = await getXml(dx, 'word/document.xml');
  if (!docXml) throw new Error('No document.xml');
  const bc = docXml.lastIndexOf('</w:body>');
  docXml = docXml.slice(0, bc) + hlXml + docXml.slice(bc);
  setXml(dx, 'word/document.xml', docXml);
  await saveDocx(dx, filename);
}

export async function clearAllContentFromDocx(filename: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  const sectPrMatch = xml.match(/<w:sectPr>[\s\S]*?<\/w:sectPr>/);
  const sectPr = sectPrMatch ? sectPrMatch[0] : '<w:sectPr/>';
  const bodyStart = xml.indexOf('<w:body>');
  const header = xml.slice(0, bodyStart + '<w:body>'.length);
  xml = header + `<w:p/>${sectPr}</w:body></w:document>`;
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

export async function addHighlightParagraphToDocx(filename: string, text: string, highlight: string = 'yellow', bold: boolean = false, color: string = ''): Promise<void> {
  await addToDocx(filename, styledParagraphXml(text, { highlight, bold, color: color || undefined }));
}

// ═══════════════════════════════════════════════════════════════════════════════
// INDEXED PARAGRAPH ACCESS — gives AI precise surgical control over every element
// ═══════════════════════════════════════════════════════════════════════════════

export interface IndexedParagraph {
  index: number;       // 0-based position in document
  type: 'heading' | 'paragraph' | 'table' | 'image' | 'list' | 'other';
  level?: number;      // heading level 1-6
  text: string;        // full concatenated text
  isEmpty: boolean;
  xmlStart: number;    // byte offset in document.xml (internal)
  xmlEnd: number;
  rowCount?: number;   // for tables: number of rows
  colCount?: number;   // for tables: number of columns
}

/** Return every top-level block with its index for AI targeting */
export async function getIndexedParagraphs(filename: string): Promise<IndexedParagraph[]> {
  const dx = await loadDocx(filename);
  const xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const results: IndexedParagraph[] = [];
  let index = 0;

  // Match paragraphs and tables at the body level
  const blockRegex = /(<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:tbl>[\s\S]*?<\/w:tbl>)/g;
  let m: RegExpExecArray | null;

  while ((m = blockRegex.exec(xml)) !== null) {
    const block = m[0];
    const start = m.index;
    const end = m.index + block.length;

    if (block.startsWith('<w:tbl>') || block.startsWith('<w:tbl ')) {
      // Table - extract row count for better context
      const tableText = extractParaText(block);
      const rows = findTableRows(block);
      const cols = rows.length > 0 ? findTableCells(rows[0].content).length : 0;
      results.push({ 
        index, 
        type: 'table', 
        text: `[TABLE ${rows.length} rows x ${cols} cols] ${tableText.slice(0, 80)}`, 
        isEmpty: !tableText.trim(), 
        xmlStart: start, 
        xmlEnd: end,
        rowCount: rows.length,
        colCount: cols
      });
    } else {
      // Paragraph
      const text = extractParaText(block);
      const pPr = extractPPr(block);

      // Detect heading
      const headingMatch = pPr.match(/w:val="Heading(\d+)"/);
      const styleMatch = pPr.match(/w:val="[Hh]eading\s*(\d+)"/);
      const lvl = headingMatch ? parseInt(headingMatch[1]) : styleMatch ? parseInt(styleMatch[1]) : undefined;

      // Detect image
      const isImage = block.includes('<wp:inline') || block.includes('<wp:anchor');

      let type: IndexedParagraph['type'] = 'paragraph';
      if (lvl) type = 'heading';
      else if (isImage) type = 'image';

      results.push({ index, type, level: lvl, text: text.slice(0, 200), isEmpty: !text.trim() && !isImage, xmlStart: start, xmlEnd: end });
    }
    index++;
  }

  return results;
}

/** Insert content (XML string) BEFORE paragraph at given index */
export async function insertBeforeIndex(filename: string, targetIndex: number, contentXml: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const blocks = getAllBlocks(xml);
  if (targetIndex >= blocks.length) {
    // Append at end
    const bodyClose = xml.lastIndexOf('</w:body>');
    xml = xml.slice(0, bodyClose) + contentXml + xml.slice(bodyClose);
  } else {
    const { start } = blocks[targetIndex];
    xml = xml.slice(0, start) + contentXml + xml.slice(start);
  }

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Insert content AFTER paragraph at given index */
export async function insertAfterIndex(filename: string, targetIndex: number, contentXml: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const blocks = getAllBlocks(xml);
  if (targetIndex >= blocks.length) {
    const bodyClose = xml.lastIndexOf('</w:body>');
    xml = xml.slice(0, bodyClose) + contentXml + xml.slice(bodyClose);
  } else {
    const { end } = blocks[targetIndex];
    xml = xml.slice(0, end) + contentXml + xml.slice(end);
  }

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Replace paragraph/table at index with new XML */
export async function replaceAtIndex(filename: string, targetIndex: number, newContentXml: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const blocks = getAllBlocks(xml);
  if (targetIndex >= blocks.length) throw new Error(`Index ${targetIndex} out of range (${blocks.length} blocks)`);

  const { start, end } = blocks[targetIndex];
  xml = xml.slice(0, start) + newContentXml + xml.slice(end);

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Delete the block at index */
export async function deleteAtIndex(filename: string, targetIndex: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const blocks = getAllBlocks(xml);
  if (targetIndex >= blocks.length) throw new Error(`Index ${targetIndex} out of range`);

  const { start, end } = blocks[targetIndex];
  xml = xml.slice(0, start) + xml.slice(end);

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Format (restyle) a paragraph at index: change font, size, color, bold, italic, alignment, indent */
export async function formatAtIndex(
  filename: string,
  targetIndex: number,
  options: {
    bold?: boolean; italic?: boolean; underline?: boolean; strikethrough?: boolean;
    color?: string; font?: string; fontSize?: number;
    alignment?: 'left' | 'center' | 'right' | 'justify';
    headingLevel?: number;
    indent?: number;       // left indent in twips
    spaceAfter?: number;   // spacing after in twips
    spaceBefore?: number;
    highlight?: string;
    lineSpacing?: number;  // 240 = single, 480 = double
  }
): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const blocks = getAllBlocks(xml);
  if (targetIndex >= blocks.length) throw new Error(`Index ${targetIndex} out of range`);

  const { start, end } = blocks[targetIndex];
  const blockXml = xml.slice(start, end);
  const text = extractParaText(blockXml);
  const existingPPr = extractPPr(blockXml);

  // Build new pPr
  const pPrParts: string[] = [];
  if (options.headingLevel) pPrParts.push(`<w:pStyle w:val="Heading${options.headingLevel}"/>`);
  if (options.alignment) {
    const jc: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' };
    pPrParts.push(`<w:jc w:val="${jc[options.alignment]}"/>`);
  }
  if (options.indent !== undefined) pPrParts.push(`<w:ind w:left="${options.indent}"/>`);
  if (options.spaceBefore !== undefined || options.spaceAfter !== undefined || options.lineSpacing !== undefined) {
    let sp = '<w:spacing';
    if (options.spaceBefore !== undefined) sp += ` w:before="${options.spaceBefore}"`;
    if (options.spaceAfter !== undefined) sp += ` w:after="${options.spaceAfter}"`;
    if (options.lineSpacing !== undefined) sp += ` w:line="${options.lineSpacing}" w:lineRule="auto"`;
    sp += '/>';
    pPrParts.push(sp);
  }

  // Build new rPr
  const rPrParts: string[] = [];
  if (options.bold) rPrParts.push('<w:b/><w:bCs/>');
  if (options.italic) rPrParts.push('<w:i/><w:iCs/>');
  if (options.underline) rPrParts.push('<w:u w:val="single"/>');
  if (options.strikethrough) rPrParts.push('<w:strike/>');
  if (options.color) rPrParts.push(`<w:color w:val="${options.color.replace('#', '')}"/>`);
  if (options.font) rPrParts.push(`<w:rFonts w:ascii="${options.font}" w:hAnsi="${options.font}" w:cs="${options.font}"/>`);
  if (options.fontSize) rPrParts.push(`<w:sz w:val="${options.fontSize * 2}"/><w:szCs w:val="${options.fontSize * 2}"/>`);
  if (options.highlight) rPrParts.push(`<w:highlight w:val="${options.highlight}"/>`);

  const newPPr = pPrParts.length ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : existingPPr;
  const rPr = rPrParts.length ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
  const newBlock = `<w:p>${newPPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;

  xml = xml.slice(0, start) + newBlock + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Replace the text of a paragraph at index, preserving its formatting */
export async function replaceTextAtIndex(
  filename: string,
  targetIndex: number,
  newText: string
): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const blocks = getAllBlocks(xml);
  if (targetIndex >= blocks.length) throw new Error(`Index ${targetIndex} out of range`);

  const { start, end } = blocks[targetIndex];
  const blockXml = xml.slice(start, end);
  const pPr = extractPPr(blockXml);
  const rPr = extractFirstRunRPr(blockXml);

  const newBlock = `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(newText)}</w:t></w:r></w:p>`;
  xml = xml.slice(0, start) + newBlock + xml.slice(end);

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Duplicate the block at targetIndex, inserting the copy after it */
export async function duplicateAtIndex(filename: string, targetIndex: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const blocks = getAllBlocks(xml);
  if (targetIndex >= blocks.length) throw new Error(`Index ${targetIndex} out of range`);

  const { start, end } = blocks[targetIndex];
  const blockXml = xml.slice(start, end);
  xml = xml.slice(0, end) + blockXml + xml.slice(end);

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Move block from sourceIndex to just before destIndex */
export async function moveBlockToIndex(filename: string, sourceIndex: number, destIndex: number): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  let blocks = getAllBlocks(xml);
  if (sourceIndex >= blocks.length) throw new Error(`Source index ${sourceIndex} out of range`);
  if (destIndex > blocks.length) throw new Error(`Dest index ${destIndex} out of range`);

  const srcBlock = xml.slice(blocks[sourceIndex].start, blocks[sourceIndex].end);

  // Remove source
  xml = xml.slice(0, blocks[sourceIndex].start) + xml.slice(blocks[sourceIndex].end);

  // Recalculate after removal
  blocks = getAllBlocks(xml);
  const adjustedDest = destIndex > sourceIndex ? Math.min(destIndex - 1, blocks.length) : Math.min(destIndex, blocks.length);

  // Insert at destination
  if (adjustedDest >= blocks.length) {
    const bodyClose = xml.lastIndexOf('</w:body>');
    xml = xml.slice(0, bodyClose) + srcBlock + xml.slice(bodyClose);
  } else {
    const ins = blocks[adjustedDest].start;
    xml = xml.slice(0, ins) + srcBlock + xml.slice(ins);
  }

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Get raw document XML — AI can read and understand exact structure */
export async function getDocumentXml(filename: string): Promise<string> {
  const dx = await loadDocx(filename);
  const xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');
  // Return body content only (strip header boilerplate)
  const bodyStart = xml.indexOf('<w:body>');
  const bodyEnd = xml.lastIndexOf('</w:body>') + '</w:body>'.length;
  return bodyStart !== -1 ? xml.slice(bodyStart, bodyEnd) : xml;
}

/** Set raw document XML — full surgical control */
export async function setDocumentXml(filename: string, bodyXml: string): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  // Replace body content
  const bodyStart = xml.indexOf('<w:body>');
  const bodyEnd = xml.lastIndexOf('</w:body>') + '</w:body>'.length;

  if (bodyStart !== -1) {
    // Ensure bodyXml wraps in w:body if not already
    const newBody = bodyXml.startsWith('<w:body>') ? bodyXml : `<w:body>${bodyXml}</w:body>`;
    xml = xml.slice(0, bodyStart) + newBody + xml.slice(bodyEnd);
  }

  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Merge a range of table cells (colspan/rowspan) */
export async function mergeTableCellsInDocx(
  filename: string,
  tableIndex: number,
  startRow: number, startCol: number,
  endRow: number, endCol: number
): Promise<void> {
  const dx = await loadDocx(filename);
  let xml = await getXml(dx, 'word/document.xml');
  if (!xml) throw new Error('No document.xml');

  const tables = findTables(xml);
  if (tableIndex >= tables.length) throw new Error(`Table ${tableIndex} not found`);

  const { start, end } = tables[tableIndex];
  let tbl = xml.slice(start, end);
  const rows = findTableRows(tbl);

  // Apply horizontal merge (w:hMerge) within each row
  let offset = 0;
  for (let ri = startRow; ri <= endRow; ri++) {
    if (ri >= rows.length) break;
    const row = rows[ri];
    const rowContent = tbl.slice(row.start + offset, row.end + offset);
    const cells = findTableCells(rowContent);
    let rowOffset = 0;

    for (let ci = startCol; ci <= endCol; ci++) {
      if (ci >= cells.length) break;
      const cell = cells[ci];
      const cellContent = rowContent.slice(cell.start + rowOffset, cell.end + rowOffset);

      let mergeXml: string;
      if (ci === startCol) {
        // First cell: restart merge
        mergeXml = cellContent.replace(
          /<w:tcPr>/, '<w:tcPr><w:hMerge w:val="restart"/>'
        ).replace(
          /(<w:tc>)(?!<w:tcPr>)/, '$1<w:tcPr><w:hMerge w:val="restart"/></w:tcPr>'
        );
      } else {
        // Continuation cells: continue merge and empty content
        mergeXml = `<w:tc><w:tcPr><w:hMerge/></w:tcPr><w:p/></w:tc>`;
      }

      const newRowContent = rowContent.slice(0, cell.start + rowOffset) + mergeXml + rowContent.slice(cell.end + rowOffset);
      rowOffset += mergeXml.length - cellContent.length;
      const newRow = newRowContent;

      const prevRowLen = row.end + offset - (row.start + offset);
      tbl = tbl.slice(0, row.start + offset) + newRow + tbl.slice(row.end + offset);
      offset += newRow.length - prevRowLen;
    }
  }

  xml = xml.slice(0, start) + tbl + xml.slice(end);
  setXml(dx, 'word/document.xml', xml);
  await saveDocx(dx, filename);
}

/** Internal: get all top-level body blocks with start/end offsets */
function getAllBlocks(xml: string): Array<{ start: number; end: number }> {
  const results: Array<{ start: number; end: number }> = [];
  const bodyStart = xml.indexOf('<w:body>') + '<w:body>'.length;
  const bodyEnd = xml.lastIndexOf('</w:body>');
  const body = xml.slice(bodyStart, bodyEnd);

  const blockRe = /(<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:tbl>[\s\S]*?<\/w:tbl>)/g;
  let m: RegExpExecArray | null;

  while ((m = blockRe.exec(body)) !== null) {
    results.push({
      start: bodyStart + m.index,
      end: bodyStart + m.index + m[0].length,
    });
  }

  return results;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EXCEL FULL-READ — give AI complete visibility into spreadsheet state
// ═══════════════════════════════════════════════════════════════════════════════

export interface CellInfo {
  address: string;
  value: any;
  formula?: string;
  type: string;
  formatted?: string;
  style?: {
    bold?: boolean; italic?: boolean; color?: string;
    bgColor?: string; fontSize?: number; numberFormat?: string;
    alignment?: string; border?: string;
  };
}

export interface SheetReadResult {
  name: string;
  rowCount: number;
  colCount: number;
  cells: CellInfo[];        // all non-empty cells
  grid: string[][];         // human-readable 2D grid (values as strings)
  formulas: Record<string, string>;
  mergedCells: string[];
  headers: string[];
}

export async function readSpreadsheetFull(
  filename: string,
  sheetName?: string
): Promise<SheetReadResult[]> {
  const { readFileBuffer } = await import('./file-storage');
  const ExcelJS = (await import('exceljs')).default;

  const buffer = readFileBuffer(filename);
  if (!buffer) throw new Error('Could not read: ' + filename);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);

  const sheets = sheetName
    ? [wb.getWorksheet(sheetName)].filter(Boolean)
    : wb.worksheets;

  return (sheets as any[]).map((ws: any) => {
    const cells: CellInfo[] = [];
    const formulas: Record<string, string> = {};
    const grid: string[][] = [];

    ws.eachRow({ includeEmpty: false }, (row: any, rn: number) => {
      const gridRow: string[] = [];
      row.eachCell({ includeEmpty: false }, (cell: any, cn: number) => {
        const addr = `${colLetter(cn)}${rn}`;
        const isFormula = typeof cell.value === 'object' && cell.value && 'formula' in cell.value;
        const formula = isFormula ? '=' + (cell.value as any).formula : undefined;
        const rawVal = isFormula ? (cell.value as any).result ?? '' : cell.value;

        if (formula) formulas[addr] = formula;

        const style: CellInfo['style'] = {};
        try {
          if (cell.font?.bold) style.bold = true;
          if (cell.font?.italic) style.italic = true;
          if (cell.font?.color?.argb) style.color = cell.font.color.argb.slice(2); // strip alpha
          if (cell.fill?.fgColor?.argb) style.bgColor = cell.fill.fgColor.argb.slice(2);
          if (cell.font?.size) style.fontSize = cell.font.size;
          if (cell.numFmt) style.numberFormat = cell.numFmt;
          if (cell.alignment?.horizontal) style.alignment = cell.alignment.horizontal;
        } catch {}

        cells.push({
          address: addr,
          value: rawVal,
          formula,
          type: typeof rawVal,
          formatted: String(rawVal ?? ''),
          style: Object.keys(style).length ? style : undefined,
        });

        gridRow[cn] = String(rawVal ?? '');
      });
      grid[rn] = gridRow;
    });

    // Build clean 2D array
    const maxRow = ws.rowCount;
    const maxCol = ws.columnCount;
    const clean2D: string[][] = [];
    for (let r = 1; r <= maxRow; r++) {
      const row: string[] = [];
      for (let c = 1; c <= maxCol; c++) {
        const cell = ws.getCell(r, c);
        const v = typeof cell.value === 'object' && cell.value && 'formula' in cell.value
          ? (cell.value as any).result ?? '' : cell.value;
        row.push(String(v ?? ''));
      }
      clean2D.push(row);
    }

    // Merged cells
    const mergedCells: string[] = [];
    try {
      for (const [key] of Object.entries((ws as any)._merges || {})) {
        mergedCells.push(key as string);
      }
    } catch {}

    const headers = clean2D[0] || [];

    return {
      name: ws.name,
      rowCount: ws.rowCount - 1,
      colCount: ws.columnCount,
      cells,
      grid: clean2D,
      formulas,
      mergedCells,
      headers,
    };
  });
}

/** Bulk update multiple cells in one pass — maximally efficient */
export async function bulkUpdateCells(
  filename: string,
  sheetName: string,
  updates: Array<{
    cell: string;
    value?: any;
    formula?: string;
    style?: {
      bold?: boolean; italic?: boolean; color?: string; bgColor?: string;
      fontSize?: number; numberFormat?: string; alignment?: string;
      wrapText?: boolean; border?: 'thin' | 'medium' | 'thick';
    };
  }>
): Promise<{ updated: number }> {
  const { readFileBuffer, writeFileBuffer } = await import('./file-storage');
  const ExcelJS = (await import('exceljs')).default;

  const buffer = readFileBuffer(filename);
  if (!buffer) throw new Error('Could not read: ' + filename);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);

  const ws = wb.getWorksheet(sheetName) || wb.worksheets[0];
  if (!ws) throw new Error('Sheet not found: ' + sheetName);

  let updated = 0;
  for (const upd of updates) {
    const cell = ws.getCell(upd.cell);

    if (upd.formula !== undefined) {
      cell.value = { formula: upd.formula.replace(/^=/, '') };
    } else if (upd.value !== undefined) {
      // Try to preserve numeric types
      const num = typeof upd.value === 'string' && upd.value !== '' ? Number(upd.value) : NaN;
      cell.value = (!isNaN(num) && upd.value !== '') ? num : upd.value;
    }

    if (upd.style) {
      const s = upd.style;
      if (s.bold !== undefined || s.italic !== undefined || s.color !== undefined || s.fontSize !== undefined) {
        cell.font = {
          ...cell.font,
          ...(s.bold !== undefined && { bold: s.bold }),
          ...(s.italic !== undefined && { italic: s.italic }),
          ...(s.color && { color: { argb: 'FF' + s.color.replace('#', '') } }),
          ...(s.fontSize && { size: s.fontSize }),
        };
      }
      if (s.bgColor) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + s.bgColor.replace('#', '') } };
      }
      if (s.numberFormat) cell.numFmt = s.numberFormat;
      if (s.alignment) cell.alignment = { ...cell.alignment, horizontal: s.alignment as any };
      if (s.wrapText !== undefined) cell.alignment = { ...cell.alignment, wrapText: s.wrapText };
      if (s.border) {
        const bStyle = { style: s.border as any };
        cell.border = { top: bStyle, bottom: bStyle, left: bStyle, right: bStyle };
      }
    }

    updated++;
  }

  const out = await wb.xlsx.writeBuffer();
  writeFileBuffer(filename, Buffer.from(out));
  return { updated };
}

/** Helper: column number → letter (1 → A, 26 → Z, 27 → AA) */
function colLetter(n: number): string {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}
