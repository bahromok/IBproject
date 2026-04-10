import { chatCompletion, type ChatMessage } from './groq-client';
import { TOOL_DEFINITIONS, executeTool, ToolResult } from './agent-tools';
import { listFiles as listStorageFiles } from './file-storage';
import { extractExcelStructure, extractWordStructure } from './file-structure';

// ─── INTERFACES ───────────────────────────────────────────────────────────────

export interface AgentResponse {
  thinking: string;
  plan: PlanStep[];
  toolCalls: any[];
  results: ToolResult[];
  finalMessage: string;
  corrections: number;
  isChat: boolean;
}

export interface PlanStep {
  step: number;
  action: string;
  tool: string;
  details: string;
}

export interface HistoryMessage {
  role: 'user' | 'assistant';
  content: string;
  filenames?: string[];
}

interface FileContext {
  filename: string;
  type: 'excel' | 'word';
  structure: string;
}

// ─── CONVERSATIONAL DETECTION ─────────────────────────────────────────────────

const CHAT_PATTERNS = [
  /^(hi|hello|hey|howdy|sup|yo|hiya|greetings)\b/i,
  /^(thanks|thank you|thx|ty|appreciate)\b/i,
  /^(bye|goodbye|see you|later|cya)\b/i,
  /^(good morning|good afternoon|good evening|good night)/i,
  /^(what can you do|what are you|who are you|help me|how do you work|what do you do)/i,
  /^(ok|okay|sure|yes|no|maybe|alright|got it|understood|right)\.?$/i,
  /^(how are you|how's it going|what's up|whats up)/i,
];

function isConversational(message: string): boolean {
  const trimmed = message.trim().toLowerCase();

  // Very short messages are likely conversational
  if (trimmed.length < 15 && !trimmed.includes('create') && !trimmed.includes('make') && !trimmed.includes('add')) {
    for (const pattern of CHAT_PATTERNS) {
      if (pattern.test(trimmed)) return true;
    }
  }

  // Check patterns
  for (const pattern of CHAT_PATTERNS) {
    if (pattern.test(trimmed)) return true;
  }

  // Questions without document intent
  if (trimmed.match(/^(what|who|how|why|when|where|can you|do you|are you|is there)/i) &&
      !trimmed.match(/\b(create|make|build|generate|write|edit|modify|add|delete|remove|fix|update|design|chart|document|excel|word|spreadsheet|table|report)\b/i)) {
    return true;
  }

  return false;
}

function generateChatResponse(message: string): string {
  const lower = message.trim().toLowerCase();

  if (/^(hi|hello|hey|howdy|sup|yo|hiya|greetings)/i.test(lower)) {
    return "Hey! I'm your document assistant. I can create and edit Word documents (.docx) and Excel spreadsheets (.xlsx) — including charts, tables, formulas, styling, and more.\n\nWhat would you like me to build?";
  }

  if (/^(thanks|thank you|thx|ty)/i.test(lower)) {
    return "You're welcome! Let me know if you need anything else.";
  }

  if (/^(bye|goodbye|see you|later|cya)/i.test(lower)) {
    return "Goodbye! Come back anytime you need documents created or edited.";
  }

  if (/^(good morning|good afternoon|good evening)/i.test(lower)) {
    return "Hello! Ready to help you with documents. What do you need?";
  }

  if (/(what can you do|what are you|who are you|help me|how do you work|what do you do)/i.test(lower)) {
    return `I'm OfficeAI — I create and edit professional documents. Here's what I can do:

**Word Documents (.docx)**
• Create formatted documents with headings, paragraphs, tables
• Add charts (pie, bar, line) as images
• Insert images, headers, footers, page numbers
• Style text (bold, italic, color, font, size)
• Set margins, orientation, line spacing
• Replace text, delete elements, add hyperlinks

**Excel Spreadsheets (.xlsx)**
• Create spreadsheets with data, headers, styling
• Add charts (pie, bar, line, doughnut)
• Formulas (SUM, AVERAGE, VLOOKUP, IF, etc.)
• Conditional formatting, data validation
• Sort, filter, freeze panes, merge cells
• Currency/percentage formatting

Just tell me what you want in plain English — like "create a sales report with a pie chart" or "add a Total column with SUM formulas to budget.xlsx".`;
  }

  if (/^(ok|okay|sure|yes|no|maybe|alright|got it|understood)/i.test(lower)) {
    return "Got it. What would you like me to do next?";
  }

  if (/^(how are you|how's it going|what's up)/i.test(lower)) {
    return "I'm ready to help! Tell me what document you'd like to create or edit.";
  }

  return "I can help you create or edit documents. Just describe what you need — like 'create a quarterly report' or 'add a chart to my spreadsheet'.";
}

// ─── MAIN AGENT LOOP ──────────────────────────────────────────────────────────

export async function runAutonomousAgent(
  userMessage: string,
  onProgress?: (status: string, progress: number, thinking?: string) => void,
  conversationHistory?: HistoryMessage[],
  abortSignal?: AbortSignal
): Promise<AgentResponse> {
  let corrections = 0;
  const maxCorrections = 2;

  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 0: CHECK FOR CONVERSATIONAL INPUT
  // ═══════════════════════════════════════════════════════════════════════════════

  if (isConversational(userMessage)) {
    onProgress?.('Responding...', 50, 'Conversational message');
    const chatMsg = generateChatResponse(userMessage);
    onProgress?.('Done!', 100);
    return {
      thinking: 'Conversational message - providing chat response',
      plan: [],
      toolCalls: [],
      results: [],
      finalMessage: chatMsg,
      corrections: 0,
      isChat: true,
    };
  }

  // Check abort
  if (abortSignal?.aborted) {
    return { thinking: 'Aborted', plan: [], toolCalls: [], results: [], finalMessage: 'Request cancelled.', corrections: 0, isChat: true };
  }

  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 1: READ - Gather file structures
  // ═══════════════════════════════════════════════════════════════════════════════

  onProgress?.('Reading files...', 5, 'Analyzing file structures...');

  const files = listStorageFiles();
  const fileContexts: FileContext[] = [];

  for (const f of files) {
    if (abortSignal?.aborted) {
      return { thinking: 'Aborted', plan: [], toolCalls: [], results: [], finalMessage: 'Request cancelled.', corrections: 0, isChat: true };
    }
    try {
      if (f.name.endsWith('.xlsx')) {
        const structure = await extractExcelStructure(f.name);
        fileContexts.push({
          filename: f.name,
          type: 'excel',
          structure: structure.summary,
        });
      } else if (f.name.endsWith('.docx')) {
        const structure = await extractWordStructure(f.name);
        fileContexts.push({
          filename: f.name,
          type: 'word',
          structure: structure.structure,
        });
      }
    } catch (e) { /* skip unreadable files */ }
  }

  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 2: PLAN - Generate execution plan
  // ═══════════════════════════════════════════════════════════════════════════════

  onProgress?.('Planning...', 15, 'Designing approach...');

  const fileList = files.length > 0
    ? files.map(f => `- ${f.name} (${f.type.toUpperCase()}, ${formatSize(f.size)})`).join('\n')
    : 'No files exist yet.';

  const structureContext = fileContexts.length > 0
    ? '\n\nFILE STRUCTURES:\n' + fileContexts.map(fc => `--- ${fc.filename} ---\n${fc.structure}`).join('\n\n')
    : '';

  const historyContext = buildHistoryContext(conversationHistory);

  // Detect active file from conversation history
  const activeFile = detectActiveFile(conversationHistory, files);
  const activeFileContext = activeFile
    ? `\nACTIVE FILE: ${activeFile}\nUse this as the target for edit operations unless user specifies another file.`
    : '';

  const systemPrompt = `You are OfficeAI, a document automation assistant. You create and edit Word (.docx) and Excel (.xlsx) files.

${TOOL_DEFINITIONS}

CURRENT STATE
FILES IN STORAGE:
${fileList}
${structureContext}
${activeFileContext}
${historyContext}`;

  const messages: ChatMessage[] = [
    { role: 'system', content: systemPrompt }
  ];

  if (conversationHistory && conversationHistory.length > 0) {
    const recent = conversationHistory.slice(-6);
    for (const msg of recent) {
      messages.push({ role: msg.role as 'user' | 'assistant', content: msg.content });
    }
  }

  messages.push({ role: 'user', content: userMessage });

  onProgress?.('Thinking...', 20, 'Analyzing your request...');

  if (abortSignal?.aborted) {
    return { thinking: 'Aborted', plan: [], toolCalls: [], results: [], finalMessage: 'Request cancelled.', corrections: 0, isChat: true };
  }

  let parsedResponse: {
    thinking: string;
    plan: PlanStep[];
    tool_calls: any[];
    message: string;
  } | null = null;

  for (let attempt = 0; attempt < 3; attempt++) {
    if (abortSignal?.aborted) {
      return { thinking: 'Aborted', plan: [], toolCalls: [], results: [], finalMessage: 'Request cancelled.', corrections: 0, isChat: true };
    }
    try {
      const aiResponse = await chatCompletion(messages, { temperature: 0.1, maxTokens: 4096 });
      parsedResponse = parseJsonResponse(aiResponse);
      if (parsedResponse && parsedResponse.tool_calls && parsedResponse.tool_calls.length > 0) break;
    } catch (e) {
      if (attempt === 2) {
        // Last resort: chat fallback instead of creating files
        return {
          thinking: 'Failed to parse AI response',
          plan: [],
          toolCalls: [],
          results: [],
          finalMessage: "I had trouble understanding that. Could you rephrase? For example:\n• 'Create a budget spreadsheet with a pie chart'\n• 'Add a row to report.xlsx'\n• 'Make a professional Word document about Q4 results'",
          corrections: 0,
          isChat: true,
        };
      }
    }
  }

  // If AI returned no tool_calls, treat as chat instead of creating files
  if (!parsedResponse || !parsedResponse.tool_calls || parsedResponse.tool_calls.length === 0) {
    return {
      thinking: 'No actionable document operations identified',
      plan: [],
      toolCalls: [],
      results: [],
      finalMessage: "I can help with that! Try asking me to:\n• Create a document or spreadsheet\n• Edit an existing file\n• Add charts, tables, or formulas\n\nJust describe what you need.",
      corrections: 0,
      isChat: true,
    };
  }

  onProgress?.('Executing plan...', 30, parsedResponse.thinking);

  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 3: EXECUTE - Run each tool call
  // ═══════════════════════════════════════════════════════════════════════════════

  const results: ToolResult[] = [];
  const toolCalls = parsedResponse.tool_calls;

  for (let i = 0; i < toolCalls.length; i++) {
    // Check abort between steps
    if (abortSignal?.aborted) {
      results.push({ success: false, message: 'Cancelled by user' });
      break;
    }

    const toolCall = toolCalls[i];
    const progress = 30 + (i / Math.max(toolCalls.length, 1)) * 60;

    onProgress?.(
      'Step ' + (i + 1) + '/' + toolCalls.length + ': ' + (toolCall.tool || 'unknown'),
      progress,
      parsedResponse!.thinking
    );

    const result = await executeTool(toolCall, (status, p) => {
      onProgress?.(status, progress + (p / 100) * (60 / Math.max(toolCalls.length, 1)), parsedResponse!.thinking);
    });

    // ═══════════════════════════════════════════════════════════════════════════════
    // PHASE 4: VERIFY & CORRECT
    // ═══════════════════════════════════════════════════════════════════════════════

    if (!result.success && corrections < maxCorrections) {
      onProgress?.('Correcting issue...', progress, result.message);

      const fixResult = await attemptFix(toolCall, result);
      if (fixResult.success) {
        results.push(fixResult);
        corrections++;
        onProgress?.('Fixed!', progress, 'Self-corrected');
      } else {
        results.push(result);
      }
    } else {
      results.push(result);
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 5: FINALIZE
  // ═══════════════════════════════════════════════════════════════════════════════

  onProgress?.('Done!', 100);

  let finalMessage = parsedResponse.message || '';
  const filenames: string[] = [];
  const modifiedFiles = new Set<string>();

  for (const r of results) {
    if (r.filename) {
      filenames.push(r.filename);
      if (r.success) modifiedFiles.add(r.filename);
    }
  }

  // Re-read structures of modified files for updated state
  const updatedStructures: Record<string, string> = {};
  for (const fname of modifiedFiles) {
    try {
      if (fname.endsWith('.xlsx')) {
        const s = await extractExcelStructure(fname);
        updatedStructures[fname] = s.summary;
      } else if (fname.endsWith('.docx')) {
        const s = await extractWordStructure(fname);
        updatedStructures[fname] = s.structure;
      }
    } catch { }
  }

  if (results.length > 0) {
    const resultLines = results.map(r => r.success ? r.message.split('\n')[0] : 'Failed: ' + r.message);
    finalMessage += '\n\n' + resultLines.join('\n');
  }

  return {
    thinking: parsedResponse.thinking,
    plan: parsedResponse.plan || [],
    toolCalls,
    results,
    finalMessage,
    corrections,
    isChat: false,
  };
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function parseJsonResponse(response: string): {
  thinking: string;
  plan: PlanStep[];
  tool_calls: any[];
  message: string;
} | null {
  try {
    return JSON.parse(response);
  } catch { }

  const jsonMatch = response.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try {
      const parsed = JSON.parse(jsonMatch[0]);
      if (parsed.tool_calls) {
        return {
          thinking: parsed.thinking || 'Processing...',
          plan: parsed.plan || [],
          tool_calls: parsed.tool_calls,
          message: parsed.message || 'Done.',
        };
      }
    } catch { }
  }

  const codeBlockMatch = response.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (codeBlockMatch) {
    try {
      const parsed = JSON.parse(codeBlockMatch[1].trim());
      return {
        thinking: parsed.thinking || 'Processing...',
        plan: parsed.plan || [],
        tool_calls: parsed.tool_calls || [],
        message: parsed.message || 'Done.',
      };
    } catch { }
  }

  return null;
}

async function attemptFix(toolCall: any, result: ToolResult): Promise<ToolResult> {
  try {
    if (toolCall.filename && !toolCall.filename.includes('.')) {
      const ext = toolCall.tool === 'create_spreadsheet' || toolCall.tool === 'edit_spreadsheet' ? '.xlsx' : '.docx';
      toolCall.filename += ext;
      return await executeTool(toolCall);
    }

    if (result.message.includes('not found') && toolCall.tool === 'edit_spreadsheet') {
      const createResult = await executeTool({
        tool: 'create_spreadsheet',
        filename: toolCall.filename,
        sheets: [{ name: 'Sheet1', headers: ['Column1'] }],
      });
      if (createResult.success) {
        return await executeTool(toolCall);
      }
    }

    if (result.message.includes('not found') && toolCall.tool === 'edit_document') {
      const createResult = await executeTool({
        tool: 'create_document',
        filename: toolCall.filename,
        title: 'Document',
        sections: [{ heading: 'Document', content: 'Content' }],
      });
      if (createResult.success) {
        return await executeTool(toolCall);
      }
    }
  } catch { }

  return result;
}

function detectActiveFile(
  history?: HistoryMessage[],
  existingFiles?: { name: string }[]
): string | null {
  if (!history || history.length === 0) return null;

  const allFilenames = new Set(existingFiles?.map(f => f.name) || []);

  // Scan history in reverse for the most recently mentioned file
  for (let i = history.length - 1; i >= 0; i--) {
    const msg = history[i];
    if (msg.filenames && msg.filenames.length > 0) {
      // Return the last filename from the most recent message that has filenames
      return msg.filenames[msg.filenames.length - 1];
    }
    // Also scan content for filename patterns
    const fileMatch = msg.content.match(/\b([\w-]+\.(docx|xlsx))\b/i);
    if (fileMatch && allFilenames.has(fileMatch[1])) {
      return fileMatch[1];
    }
  }

  // If files exist but no history mentions them, return the most recent file
  if (existingFiles && existingFiles.length > 0) {
    return existingFiles[0].name;
  }

  return null;
}

function buildHistoryContext(history?: HistoryMessage[]): string {
  if (!history || history.length === 0) return '';

  const recent = history.slice(-4);
  const lines: string[] = ['RECENT CONVERSATION:'];

  for (const msg of recent) {
    const role = msg.role === 'user' ? 'User' : 'AI';
    const content = msg.content.slice(0, 150).replace(/\n/g, ' ');
    lines.push(`${role}: ${content}`);
  }

  return lines.join('\n') + '\n';
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}
