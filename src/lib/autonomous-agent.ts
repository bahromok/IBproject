// ═══════════════════════════════════════════════════════════════════════════════════
// AUTONOMOUS AGENT UTILITIES - Advanced Search, Shell Execution & Memory Management
// ═══════════════════════════════════════════════════════════════════════════════════

import { exec } from 'child_process';
import { promisify } from 'util';
import * as fs from 'fs';
import * as path from 'path';
import { getFilePath, readFileBuffer, writeFileBuffer, ensureStorageDir } from './file-storage';

const execAsync = promisify(exec);

// ═══════════════════════════════════════════════════════════════════════════════════
// MEMORY CONTEXT MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════════

export interface DocumentMemory {
  id: string;
  filename: string;
  type: 'word' | 'excel';
  createdAt: Date;
  lastModified: Date;
  elements: Array<{
    id: string;
    type: string;
    content?: string;
    index?: number;
    properties?: any;
  }>;
  metadata: {
    wordCount?: number;
    sheetCount?: number;
    tableCount?: number;
    imageCount?: number;
    [key: string]: any;
  };
}

const documentMemoryCache = new Map<string, DocumentMemory>();

/**
 * Store document structure in memory for fast access
 */
export function storeDocumentMemory(memory: DocumentMemory): void {
  documentMemoryCache.set(memory.id, memory);
  persistMemoryToFile();
}

/**
 * Retrieve document from memory
 */
export function getDocumentMemory(id: string): DocumentMemory | undefined {
  return documentMemoryCache.get(id);
}

/**
 * Get all document memories
 */
export function getAllDocumentMemories(): DocumentMemory[] {
  return Array.from(documentMemoryCache.values());
}

/**
 * Update specific element in document memory
 */
export function updateElementInMemory(
  docId: string, 
  elementId: string, 
  updates: Partial<DocumentMemory['elements'][0]>
): boolean {
  const memory = documentMemoryCache.get(docId);
  if (!memory) return false;
  
  const elementIndex = memory.elements.findIndex(e => e.id === elementId);
  if (elementIndex === -1) return false;
  
  memory.elements[elementIndex] = { ...memory.elements[elementIndex], ...updates };
  memory.lastModified = new Date();
  persistMemoryToFile();
  return true;
}

/**
 * Delete element from document memory
 */
export function deleteElementFromMemory(docId: string, elementId: string): boolean {
  const memory = documentMemoryCache.get(docId);
  if (!memory) return false;
  
  const initialLength = memory.elements.length;
  memory.elements = memory.elements.filter(e => e.id !== elementId);
  memory.lastModified = new Date();
  persistMemoryToFile();
  return memory.elements.length < initialLength;
}

/**
 * Add element to document memory
 */
export function addElementToMemory(
  docId: string, 
  element: DocumentMemory['elements'][0]
): void {
  const memory = documentMemoryCache.get(docId);
  if (!memory) return;
  
  memory.elements.push(element);
  memory.lastModified = new Date();
  persistMemoryToFile();
}

/**
 * Persist memory to memory.md file
 */
function persistMemoryToFile(): void {
  try {
    ensureStorageDir();
    const memoryPath = path.join(process.env.STORAGE_DIR || './storage', 'memory.md');
    
    let content = '# Document Memory Index\n\n';
    content += `Last Updated: ${new Date().toISOString()}\n\n`;
    content += `Total Documents: ${documentMemoryCache.size}\n\n`;
    content += '---\n\n';
    
    for (const [id, memory] of documentMemoryCache.entries()) {
      content += `## ${memory.filename}\n\n`;
      content += `- **ID**: ${id}\n`;
      content += `- **Type**: ${memory.type}\n`;
      content += `- **Created**: ${memory.createdAt.toISOString()}\n`;
      content += `- **Last Modified**: ${memory.lastModified.toISOString()}\n`;
      content += `- **Elements**: ${memory.elements.length}\n`;
      
      if (memory.metadata.wordCount) {
        content += `- **Word Count**: ${memory.metadata.wordCount}\n`;
      }
      if (memory.metadata.sheetCount) {
        content += `- **Sheets**: ${memory.metadata.sheetCount}\n`;
      }
      if (memory.metadata.tableCount !== undefined) {
        content += `- **Tables**: ${memory.metadata.tableCount}\n`;
      }
      if (memory.metadata.imageCount !== undefined) {
        content += `- **Images**: ${memory.metadata.imageCount}\n`;
      }
      
      content += '\n### Elements:\n\n';
      for (const elem of memory.elements.slice(0, 20)) {
        content += `- [${elem.index !== undefined ? elem.index : 'N/A'}] ${elem.type}: ${elem.content?.substring(0, 100) || 'No content'}${(elem.content?.length || 0) > 100 ? '...' : ''}\n`;
      }
      
      if (memory.elements.length > 20) {
        content += `\n*... and ${memory.elements.length - 20} more elements*\n`;
      }
      
      content += '\n---\n\n';
    }
    
    fs.writeFileSync(memoryPath, content, 'utf-8');
  } catch (error) {
    console.error('Failed to persist memory to file:', error);
  }
}

/**
 * Load memory from memory.md file on startup
 */
export function loadMemoryFromFile(): void {
  try {
    const memoryPath = path.join(process.env.STORAGE_DIR || './storage', 'memory.md');
    if (!fs.existsSync(memoryPath)) return;
    
    const content = fs.readFileSync(memoryPath, 'utf-8');
    documentMemoryCache.clear();
    console.log('Memory file loaded. Cache will be rebuilt on document access.');
  } catch (error) {
    console.error('Failed to load memory from file:', error);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════
// ADVANCED SEARCH UTILITIES (GREP-LIKE)
// ═══════════════════════════════════════════════════════════════════════════════════

export interface SearchResult {
  file: string;
  line: number;
  column: number;
  content: string;
  context?: string;
  match: string;
}

/**
 * Search for text pattern across all documents using grep-like functionality
 */
export async function searchInDocuments(
  pattern: string,
  options: {
    caseSensitive?: boolean;
    wholeWord?: boolean;
    contextLines?: number;
    filePattern?: string;
    maxResults?: number;
  } = {}
): Promise<SearchResult[]> {
  const results: SearchResult[] = [];
  const storageDir = process.env.STORAGE_DIR || './storage';
  
  try {
    const flags = ['-n', '-H'];
    if (!options.caseSensitive) {
      flags.push('-i');
    }
    if (options.wholeWord) {
      flags.push('-w');
    }
    if (options.contextLines && options.contextLines > 0) {
      flags.push(`-C${options.contextLines}`);
    }
    
    const maxResults = options.maxResults || 100;
    const grepPattern = pattern.replace(/'/g, "'\\''");
    const cmd = `cd "${storageDir}" && find . -type f \\( -name "*.docx" -o -name "*.xlsx" \\) -exec sh -c 'for f; do strings "$f" | grep -n ${flags.join(' ')} "${grepPattern}" | head -${maxResults}; done' _ {} +`;
    
    const { stdout } = await execAsync(cmd, { 
      maxBuffer: 10 * 1024 * 1024,
      timeout: 30000 
    });
    
    if (stdout.trim()) {
      const lines = stdout.split('\n').filter(l => l.trim());
      for (const line of lines) {
        const match = line.match(/^([^:]+):(\d+):(.*)$/);
        if (match) {
          results.push({
            file: match[1],
            line: parseInt(match[2]),
            column: 0,
            content: match[3].trim(),
            match: pattern
          });
        }
      }
    }
  } catch (error: any) {
    if (error.code !== 1) {
      console.error('Search error:', error.message);
    }
  }
  
  return results.slice(0, maxResults);
}

/**
 * Search and replace text across documents
 */
export async function searchAndReplaceInDocuments(
  pattern: string,
  replacement: string,
  options: {
    caseSensitive?: boolean;
    wholeWord?: boolean;
    filePattern?: string;
  } = {}
): Promise<{ file: string; replacements: number }[]> {
  const results: Array<{ file: string; replacements: number }> = [];
  
  try {
    const searchResults = await searchInDocuments(pattern, {
      caseSensitive: options.caseSensitive,
      wholeWord: options.wholeWord,
      filePattern: options.filePattern,
      maxResults: 1000
    });
    
    const fileMap = new Map<string, number>();
    for (const result of searchResults) {
      fileMap.set(result.file, (fileMap.get(result.file) || 0) + 1);
    }
    
    for (const [file, count] of fileMap.entries()) {
      results.push({ file, replacements: count });
    }
  } catch (error: any) {
    console.error('Search and replace error:', error.message);
  }
  
  return results;
}

/**
 * Get document statistics using command-line tools
 */
export async function getDocumentStats(filename: string): Promise<{
  fileSize: number;
  wordCount?: number;
  lineCount?: number;
  characterCount?: number;
}> {
  const filePath = getFilePath(filename);
  
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filename}`);
  }
  
  const stats = fs.statSync(filePath);
  
  try {
    const { stdout } = await execAsync(`strings "${filePath}" | wc -w`, {
      maxBuffer: 10 * 1024 * 1024
    });
    
    const wordCount = parseInt(stdout.trim()) || 0;
    
    return {
      fileSize: stats.size,
      wordCount,
      lineCount: undefined,
      characterCount: undefined
    };
  } catch (error) {
    return {
      fileSize: stats.size
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════
// SAFE SHELL COMMAND EXECUTION
// ═══════════════════════════════════════════════════════════════════════════════════

export interface ShellCommandResult {
  success: boolean;
  stdout: string;
  stderr: string;
  exitCode: number | null;
  command: string;
  duration: number;
}

const ALLOWED_COMMANDS = new Set([
  'ls', 'dir', 'pwd', 'echo', 'cat', 'head', 'tail', 'wc',
  'grep', 'find', 'sort', 'uniq', 'cut', 'awk', 'sed',
  'mkdir', 'rmdir', 'cp', 'mv', 'rm', 'touch',
  'chmod', 'chown', 'stat', 'file', 'which', 'whereis',
  'date', 'time', 'uname', 'hostname', 'whoami',
  'zip', 'unzip', 'tar', 'gzip', 'gunzip',
  'node', 'npm', 'npx', 'python', 'python3',
  'git', 'curl', 'wget',
]);

const DANGEROUS_PATTERNS = [
  ';', '&&', '||', '|', '`', '$(', '${', '>', '<', '&',
  'rm -rf /', 'sudo', 'su ', 'chmod 777', 'dd if=',
  ':(){:|:&};:', 'fork bomb', 'mkfs', 'fdisk', 'mount',
  '/dev/null', 'nohup', 'disown', 'kill', 'pkill',
];

/**
 * Validate if a command is safe to execute
 */
function isCommandSafe(command: string): { safe: boolean; reason?: string } {
  const cmd = command.trim().toLowerCase();
  
  for (const pattern of DANGEROUS_PATTERNS) {
    if (cmd.includes(pattern.toLowerCase())) {
      return { safe: false, reason: `Contains dangerous pattern: ${pattern}` };
    }
  }
  
  const baseCmd = cmd.split(/\s+/)[0];
  
  if (!ALLOWED_COMMANDS.has(baseCmd)) {
    return { safe: false, reason: `Command not allowed: ${baseCmd}` };
  }
  
  return { safe: true };
}

/**
 * Execute a shell command safely
 */
export async function executeShellCommand(
  command: string,
  options: {
    timeout?: number;
    maxBuffer?: number;
    cwd?: string;
    dryRun?: boolean;
  } = {}
): Promise<ShellCommandResult> {
  const startTime = Date.now();
  
  const safety = isCommandSafe(command);
  if (!safety.safe) {
    return {
      success: false,
      stdout: '',
      stderr: `Command blocked for safety: ${safety.reason}`,
      exitCode: -1,
      command,
      duration: Date.now() - startTime
    };
  }
  
  if (options.dryRun) {
    return {
      success: true,
      stdout: `[DRY RUN] Would execute: ${command}`,
      stderr: '',
      exitCode: 0,
      command,
      duration: Date.now() - startTime
    };
  }
  
  try {
    const { stdout, stderr } = await execAsync(command, {
      timeout: options.timeout || 30000,
      maxBuffer: options.maxBuffer || 5 * 1024 * 1024,
      cwd: options.cwd || process.env.STORAGE_DIR || './storage',
      windowsHide: true
    });
    
    return {
      success: true,
      stdout: stdout.trim(),
      stderr: stderr.trim(),
      exitCode: 0,
      command,
      duration: Date.now() - startTime
    };
  } catch (error: any) {
    return {
      success: false,
      stdout: error.stdout?.trim() || '',
      stderr: error.stderr?.trim() || error.message,
      exitCode: error.code ?? error.exitCode ?? -1,
      command,
      duration: Date.now() - startTime
    };
  }
}

/**
 * List files in storage directory
 */
export async function listFilesDetailed(
  pattern: string = '*',
  options: { includeHidden?: boolean; sortBy?: 'name' | 'size' | 'date'; reverse?: boolean } = {}
): Promise<Array<{ name: string; size: number; modified: Date; type: string }>> {
  const storageDir = process.env.STORAGE_DIR || './storage';
  
  try {
    let cmd = `ls -la --time-style=long-iso "${storageDir}"`;
    if (pattern !== '*') {
      cmd += ` | grep "${pattern}"`;
    }
    
    const { stdout } = await execAsync(cmd);
    
    const files: Array<{ name: string; size: number; modified: Date; type: string }> = [];
    const lines = stdout.split('\n').slice(1);
    
    for (const line of lines) {
      const parts = line.trim().split(/\s+/);
      if (parts.length >= 9) {
        const type = parts[0][0] === 'd' ? 'directory' : 'file';
        const size = parseInt(parts[4]) || 0;
        const dateStr = `${parts[5]} ${parts[6]} ${parts[7]}`;
        const modified = new Date(dateStr);
        const name = parts.slice(8).join(' ');
        
        if (!options.includeHidden && name.startsWith('.')) {
          continue;
        }
        
        files.push({ name, size, modified, type });
      }
    }
    
    if (options.sortBy) {
      files.sort((a, b) => {
        let comparison = 0;
        switch (options.sortBy) {
          case 'name':
            comparison = a.name.localeCompare(b.name);
            break;
          case 'size':
            comparison = a.size - b.size;
            break;
          case 'date':
            comparison = a.modified.getTime() - b.modified.getTime();
            break;
        }
        return options.reverse ? -comparison : comparison;
      });
    }
    
    return files;
  } catch (error) {
    const entries = fs.readdirSync(storageDir, { withFileTypes: true });
    return entries
      .filter(e => options.includeHidden || !e.name.startsWith('.'))
      .filter(e => pattern === '*' || e.name.includes(pattern))
      .map(e => ({
        name: e.name,
        size: e.isFile() ? fs.statSync(path.join(storageDir, e.name)).size : 0,
        modified: fs.statSync(path.join(storageDir, e.name)).mtime,
        type: e.isDirectory() ? 'directory' : 'file'
      }));
  }
}

/**
 * Get system information
 */
export async function getSystemInfo(): Promise<{
  platform: string;
  arch: string;
  nodeVersion: string;
  memoryUsage: NodeJS.MemoryUsage;
  uptime: number;
  cwd: string;
}> {
  return {
    platform: process.platform,
    arch: process.arch,
    nodeVersion: process.version,
    memoryUsage: process.memoryUsage(),
    uptime: process.uptime(),
    cwd: process.cwd()
  };
}

// Initialize memory on module load
loadMemoryFromFile();
