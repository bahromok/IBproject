import { NextRequest, NextResponse } from 'next/server';
import { ensureStorageDir, getFilePath } from '@/lib/file-storage';
import fs from 'fs';
import path from 'path';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File | null;

    if (!file) {
      return NextResponse.json({ error: 'No file provided' }, { status: 400 });
    }

    // Validate file type
    const ext = path.extname(file.name).toLowerCase();
    if (ext !== '.docx' && ext !== '.xlsx') {
      return NextResponse.json(
        { error: 'Only .docx and .xlsx files are supported' },
        { status: 400 }
      );
    }

    // Validate file size (max 10MB)
    if (file.size > 10 * 1024 * 1024) {
      return NextResponse.json(
        { error: 'File too large. Maximum size is 10MB.' },
        { status: 400 }
      );
    }

    ensureStorageDir();

    // Save file
    const buffer = Buffer.from(await file.arrayBuffer());
    const filename = file.name.replace(/[^a-zA-Z0-9._-]/g, '_'); // Sanitize filename
    const filePath = getFilePath(filename);
    fs.writeFileSync(filePath, buffer);

    return NextResponse.json({
      success: true,
      message: 'File uploaded: ' + filename,
      file: {
        name: filename,
        type: ext.slice(1),
        size: file.size,
      },
    });
  } catch (error: any) {
    console.error('Upload error:', error);
    return NextResponse.json(
      { error: 'Upload failed: ' + error.message },
      { status: 500 }
    );
  }
}
