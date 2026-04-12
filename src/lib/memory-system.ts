/**
 * Memory System - Context awareness and file history tracking
 * Includes MEMORY.md synchronization for persistent document structure
 */

import fs from 'fs';
import path from 'path';
import { getFilePath } from './file-storage';

// In-memory storage for file histories
const fileHistories = new Map<string, string[]>();
const conversationContext: any[] = [];
const documentStructures = new Map<string, DocumentStructure>();

interface DocumentStructure {
  filename: string;
  lastUpdated: string;
  elements: ElementInfo[];
  tables: TableInfo[];
  images: ImageInfo[];
}

interface ElementInfo {
  id: number;
  type: 'heading' | 'paragraph' | 'table' | 'image' | 'list' | 'chart';
  content: string;
  level?: number;
  tableRows?: number;
}

interface TableInfo {
  id: number;
  rows: number;
  columns: number;
  headers?: string[];
}

interface ImageInfo {
  id: number;
  type: 'image' | 'chart';
  caption?: string;
}

const MEMORY_FILE = 'MEMORY.md';

/**
 * Sync MEMORY.md file with current document structures
 */
export async function syncMemoryFile(): Promise<void> {
  try {
    let content = '# OfficeAI Document Memory\n\n';
    content += `Last updated: ${new Date().toISOString()}\n\n`;
    content += 'This file tracks the exact structure of all documents for precise AI editing.\n\n';
    
    for (const [filename, structure] of documentStructures.entries()) {
      content += `## ${filename}\n\n`;
      content += `**Last Updated:** ${structure.lastUpdated}\n\n`;
      
      if (structure.elements.length > 0) {
        content += '### Elements (by ID)\n\n';
        content += '| ID | Type | Content Preview |\n';
        content += '|-----|------|----------------|\n';
        
        for (const elem of structure.elements) {
          const preview = elem.content.substring(0, 50).replace(/\n/g, ' ') + (elem.content.length > 50 ? '...' : '');
          let typeLabel = elem.type;
          if (elem.type === 'heading') typeLabel += ` (H${elem.level})`;
          if (elem.type === 'table') typeLabel += ` (${elem.tableRows} rows)`;
          
          content += `| @${elem.id} | ${typeLabel} | ${preview} |\n`;
        }
        content += '\n';
      }
      
      if (structure.tables.length > 0) {
        content += '### Tables Detail\n\n';
        for (const table of structure.tables) {
          content += `**Table @${table.id}:** ${table.rows} rows × ${table.columns} columns\n`;
          if (table.headers) {
            content += `Headers: ${table.headers.join(' | ')}\n`;
          }
          content += '\n';
        }
      }
      
      if (structure.images.length > 0) {
        content += '### Images & Charts\n\n';
        for (const img of structure.images) {
          content += `**${img.type} @${img.id}:** ${img.caption || 'No caption'}\n`;
        }
        content += '\n';
      }
      
      // Add edit history
      const history = getFileHistory(filename);
      if (history.length > 0) {
        content += '### Recent Edits\n\n';
        history.slice(-5).forEach((entry, i) => {
          content += `${i + 1}. ${entry}\n`;
        });
        content += '\n';
      }
      
      content += '---\n\n';
    }
    
    if (documentStructures.size === 0) {
      content += '*No documents tracked yet. Create or read a document to start tracking.*\n';
    }
    
    const memoryPath = getFilePath(MEMORY_FILE);
    fs.writeFileSync(memoryPath, content, 'utf-8');
  } catch (error) {
    console.error('Failed to sync MEMORY.md:', error);
  }
}

/**
 * Update document structure from get_paragraph_index result
 */
export function updateDocumentStructure(filename: string, elements: any[]): void {
  const structure: DocumentStructure = {
    filename,
    lastUpdated: new Date().toISOString(),
    elements: [],
    tables: [],
    images: [],
  };
  
  let tableIdCounter = 0;
  let imageIdCounter = 0;
  
  for (let i = 0; i < elements.length; i++) {
    const elem = elements[i];
    const elementInfo: ElementInfo = {
      id: i,
      type: elem.type || 'paragraph',
      content: elem.content || elem.text || '',
      level: elem.level,
      tableRows: elem.tableRows,
    };
    
    structure.elements.push(elementInfo);
    
    if (elem.type === 'table' && elem.tableRows) {
      const tableInfo: TableInfo = {
        id: tableIdCounter++,
        rows: elem.tableRows,
        columns: elem.tableColumns || 0,
        headers: elem.headers,
      };
      structure.tables.push(tableInfo);
    }
    
    if (elem.type === 'image' || elem.type === 'chart') {
      structure.images.push({
        id: imageIdCounter++,
        type: elem.type,
        caption: elem.caption,
      });
    }
  }
  
  documentStructures.set(filename, structure);
  addFileHistoryEntry(filename, `Structure updated: ${elements.length} elements indexed`);
  
  // Async sync to file
  syncMemoryFile().catch(console.error);
}

/**
 * Get document structure by filename
 */
export function getDocumentStructure(filename: string): DocumentStructure | undefined {
  return documentStructures.get(filename);
}

/**
 * Get element by ID for precise editing
 */
export function getElementById(filename: string, elementId: number): ElementInfo | undefined {
  const structure = documentStructures.get(filename);
  if (!structure) return undefined;
  return structure.elements.find(e => e.id === elementId);
}

/**
 * Get table by ID for precise editing
 */
export function getTableById(filename: string, tableId: number): TableInfo | undefined {
  const structure = documentStructures.get(filename);
  if (!structure) return undefined;
  return structure.tables.find(t => t.id === tableId);
}

/**
 * Search memory for elements matching text
 */
export function searchMemory(filename: string, searchText: string): ElementInfo[] {
  const structure = documentStructures.get(filename);
  if (!structure) return [];
  
  const lowerSearch = searchText.toLowerCase();
  return structure.elements.filter(elem => 
    elem.content.toLowerCase().includes(lowerSearch)
  );
}

/**
 * Add an entry to file history
 */
export function addFileHistoryEntry(filename: string, entry: string): void {
  if (!fileHistories.has(filename)) {
    fileHistories.set(filename, []);
  }
  const history = fileHistories.get(filename)!;
  history.push(`[${new Date().toISOString()}] ${entry}`);
  
  // Keep last 50 entries
  if (history.length > 50) {
    fileHistories.set(filename, history.slice(-50));
  }
  
  // Sync memory file after history update
  syncMemoryFile().catch(console.error);
}

/**
 * Get file history
 */
export function getFileHistory(filename: string): string[] {
  return fileHistories.get(filename) || [];
}

/**
 * Get formatted file history for AI context
 */
export function getFileHistoryContext(filename: string): string {
  const history = getFileHistory(filename);
  if (history.length === 0) {
    return `No edit history for ${filename}.`;
  }
  
  return `Recent edits to ${filename}:\n` + 
    history.slice(-10).map((entry, i) => `  ${i + 1}. ${entry}`).join('\n');
}

/**
 * Add conversation context
 */
export function addConversationContext(entry: {
  role: 'user' | 'assistant';
  content: string;
  files?: string[];
}): void {
  conversationContext.push({
    ...entry,
    timestamp: new Date().toISOString()
  });
  
  // Keep last 20 messages
  if (conversationContext.length > 20) {
    conversationContext.shift();
  }
}

/**
 * Get conversation context for AI
 */
export function getConversationContext(): any[] {
  return conversationContext;
}

/**
 * Get referenced files from recent conversation
 */
export function getReferencedFiles(): string[] {
  const files = new Set<string>();
  for (const entry of conversationContext) {
    if (entry.files) {
      entry.files.forEach((f: string) => files.add(f));
    }
  }
  return Array.from(files);
}

/**
 * Clear all memory
 */
export function clearMemory(): void {
  fileHistories.clear();
  conversationContext.length = 0;
  documentStructures.clear();
  
  // Clear MEMORY.md file
  try {
    const memoryPath = getFilePath(MEMORY_FILE);
    if (fs.existsSync(memoryPath)) {
      fs.writeFileSync(memoryPath, '# OfficeAI Document Memory\n\n*Cleared*\n', 'utf-8');
    }
  } catch (error) {
    console.error('Failed to clear MEMORY.md:', error);
  }
}

/**
 * Get memory summary for debugging
 */
export function getMemorySummary(): {
  filesTracked: number;
  totalHistoryEntries: number;
  conversationMessages: number;
  totalElements: number;
} {
  let totalEntries = 0;
  let totalElements = 0;
  
  fileHistories.forEach(history => {
    totalEntries += history.length;
  });
  
  documentStructures.forEach(structure => {
    totalElements += structure.elements.length;
  });
  
  return {
    filesTracked: documentStructures.size,
    totalHistoryEntries: totalEntries,
    conversationMessages: conversationContext.length,
    totalElements,
  };
}

/**
 * Load MEMORY.md on startup
 */
export async function loadMemoryFile(): Promise<void> {
  try {
    const memoryPath = getFilePath(MEMORY_FILE);
    if (fs.existsSync(memoryPath)) {
      const content = fs.readFileSync(memoryPath, 'utf-8');
      console.log('MEMORY.md loaded successfully');
      // Parse and restore structures if needed
    } else {
      // Create initial MEMORY.md
      await syncMemoryFile();
    }
  } catch (error) {
    console.error('Failed to load MEMORY.md:', error);
  }
}

// Auto-load on module import
loadMemoryFile().catch(console.error);
