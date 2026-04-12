import { NextRequest, NextResponse } from 'next/server';
import { listFiles, readFileBuffer, deleteFile } from '@/lib/file-storage';

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const action = searchParams.get('action');
    const filename = searchParams.get('filename');

    if (action === 'download' && filename) {
      const buffer = readFileBuffer(filename);
      if (!buffer) {
        return NextResponse.json({ error: 'File not found' }, { status: 404 });
      }

      const contentType = filename.endsWith('.docx')
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

      return new NextResponse(buffer as unknown as BodyInit, {
        headers: {
          'Content-Type': contentType,
          'Content-Disposition': `attachment; filename="${filename}"`,
          'Content-Length': buffer.length.toString(),
        },
      });
    }

    const files = listFiles();
    return NextResponse.json({ files });
  } catch (error) {
    console.error('Files API error:', error);
    return NextResponse.json(
      { error: 'Failed to fetch files' },
      { status: 500 }
    );
  }
}

export async function DELETE(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const filename = searchParams.get('filename');

    if (!filename) {
      return NextResponse.json({ error: 'Filename is required' }, { status: 400 });
    }

    const deleted = deleteFile(filename);
    if (deleted) {
      return NextResponse.json({ success: true, message: `Deleted ${filename}` });
    }

    return NextResponse.json({ error: 'File not found' }, { status: 404 });
  } catch (error) {
    console.error('Delete file error:', error);
    return NextResponse.json(
      { error: 'Failed to delete file' },
      { status: 500 }
    );
  }
}
