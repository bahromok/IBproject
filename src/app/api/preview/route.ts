import { NextRequest, NextResponse } from 'next/server';
import { docxToHtml } from '@/lib/doc-reader';
import { readFileBuffer, fileExists } from '@/lib/file-storage';
import * as XLSX from 'xlsx';

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const filename = searchParams.get('filename');

    if (!filename) {
      return NextResponse.json({ error: 'Filename is required' }, { status: 400 });
    }

    if (!fileExists(filename)) {
      return NextResponse.json({ error: 'File not found' }, { status: 404 });
    }

    if (filename.endsWith('.docx')) {
      try {
        const html = await docxToHtml(filename);
        return NextResponse.json({
          type: 'docx',
          html: html || '<p><em>Empty document</em></p>',
          filename,
        });
      } catch (e: any) {
        return NextResponse.json({
          type: 'docx',
          html: '<p style="color:#888">Preview not available: ' + e.message + '</p>',
          filename,
        });
      }
    }

    if (filename.endsWith('.xlsx')) {
      const buffer = readFileBuffer(filename);
      if (!buffer) {
        return NextResponse.json({ error: 'Could not read file' }, { status: 500 });
      }

      const wb = XLSX.read(buffer, { type: 'buffer' });
      const sheets: Record<string, string> = {};

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        // Skip sheets that have no content (!ref is undefined)
        if (!ws['!ref']) {
          sheets[sheetName] = '<table><tr><td><em>Empty sheet</em></td></tr></table>';
          continue;
        }
        try {
          sheets[sheetName] = XLSX.utils.sheet_to_html(ws);
        } catch (e: any) {
          sheets[sheetName] = '<table><tr><td><em>Error rendering sheet: ' + e.message + '</em></td></tr></table>';
        }
      }

      return NextResponse.json({
        type: 'xlsx',
        sheets,
        activeSheet: wb.SheetNames[0],
        filename,
      });
    }

    return NextResponse.json({ error: 'Unsupported file type' }, { status: 400 });
  } catch (error: any) {
    console.error('Preview error:', error);
    return NextResponse.json(
      { error: 'Preview failed: ' + error.message },
      { status: 500 }
    );
  }
}
