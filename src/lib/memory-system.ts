/**
 * Memory System - Context awareness and file history tracking
 */

// In-memory storage for file histories
const fileHistories = new Map<string, string[]>();
const conversationContext: any[] = [];

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
}

/**
 * Get memory summary for debugging
 */
export function getMemorySummary(): {
  filesTracked: number;
  totalHistoryEntries: number;
  conversationMessages: number;
} {
  let totalEntries = 0;
  fileHistories.forEach(history => {
    totalEntries += history.length;
  });
  
  return {
    filesTracked: fileHistories.size,
    totalHistoryEntries: totalEntries,
    conversationMessages: conversationContext.length
  };
}
