import fs from 'fs';
import path from 'path';

const STORAGE_DIR = path.join(process.cwd(), 'storage');

export function ensureStorageDir() {
  if (!fs.existsSync(STORAGE_DIR)) {
    fs.mkdirSync(STORAGE_DIR, { recursive: true });
  }
}

export function getFilePath(filename: string): string {
  ensureStorageDir();
  return path.join(STORAGE_DIR, filename);
}

export function listFiles(): Array<{ name: string; type: string; size: number; createdAt: string; updatedAt: string }> {
  ensureStorageDir();
  const files = fs.readdirSync(STORAGE_DIR);
  return files
    .filter(f => f.endsWith('.docx') || f.endsWith('.xlsx'))
    .map(f => {
      const filePath = path.join(STORAGE_DIR, f);
      const stats = fs.statSync(filePath);
      const ext = path.extname(f).slice(1);
      return {
        name: f,
        type: ext,
        size: stats.size,
        createdAt: stats.birthtime.toISOString(),
        updatedAt: stats.mtime.toISOString(),
      };
    });
}

export function deleteFile(filename: string): boolean {
  const filePath = getFilePath(filename);
  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
    // Also delete metadata file if it exists
    const metaPath = filePath + '.meta.json';
    if (fs.existsSync(metaPath)) {
      fs.unlinkSync(metaPath);
    }
    return true;
  }
  return false;
}

export function fileExists(filename: string): boolean {
  return fs.existsSync(getFilePath(filename));
}

export function readFileBuffer(filename: string): Buffer | null {
  const filePath = getFilePath(filename);
  if (fs.existsSync(filePath)) {
    return fs.readFileSync(filePath);
  }
  return null;
}

export function writeFileBuffer(filename: string, buffer: Buffer): void {
  ensureStorageDir();
  const filePath = getFilePath(filename);
  fs.writeFileSync(filePath, buffer);
}
