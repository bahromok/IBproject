import { chatCompletion, type ChatMessage } from './groq-client';
import { executeTool, ToolResult } from './agent-tools';
import { extractExcelStructure, extractWordStructure } from './file-structure';
import { getStyleTemplate } from './style-engine';
import { getFileHistory, addFileHistoryEntry } from './memory-system';
import { validateFile, fixIssues } from './validation-layer';

export interface PlanStep {
  tool: string;
  params: any;
  description: string;
}

export interface AgentPlan {
  steps: PlanStep[];
  reasoning: string;
}

export interface AgentExecutionResult {
  success: boolean;
  message: string;
  plan: AgentPlan;
  results: ToolResult[];
  corrections: number;
}

/**
 * Autonomous Agent Loop - The core intelligence engine
 * 
 * Flow: UNDERSTAND → PLAN → EXECUTE → VERIFY → FIX → COMPLETE
 */
export async function runAgentLoop(
  task: string,
  context: {
    conversationHistory?: any[];
    currentFiles?: string[];
  },
  onProgress?: (status: string, progress: number, thinking?: string) => void
): Promise<AgentExecutionResult> {
  
  let corrections = 0;
  const maxCorrections = 3;
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 1: UNDERSTAND - Extract file structures and context
  // ═══════════════════════════════════════════════════════════════════════════════
  
  onProgress?.('Understanding context...', 5, 'Analyzing files and requirements...');
  
  const fileStructures: Record<string, any> = {};
  const fileHistories: Record<string, string[]> = {};
  
  // Extract structure for each relevant file
  for (const filename of context.currentFiles || []) {
    try {
      if (filename.endsWith('.xlsx')) {
        fileStructures[filename] = await extractExcelStructure(filename);
      } else if (filename.endsWith('.docx')) {
        fileStructures[filename] = await extractWordStructure(filename);
      }
      fileHistories[filename] = getFileHistory(filename);
    } catch (e) {
      // File might not exist yet
    }
  }
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 2: PLAN - Generate multi-step execution plan
  // ═══════════════════════════════════════════════════════════════════════════════
  
  onProgress?.('Creating execution plan...', 15, 'Designing multi-step approach...');
  
  const plan = await generatePlan(task, fileStructures, fileHistories);
  
  onProgress?.('Plan created: ' + plan.steps.length + ' steps', 20, plan.reasoning);
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 3: EXECUTE - Run each step with verification
  // ═══════════════════════════════════════════════════════════════════════════════
  
  const results: ToolResult[] = [];
  const executedSteps: string[] = [];
  
  for (let i = 0; i < plan.steps.length; i++) {
    const step = plan.steps[i];
    const baseProgress = 25 + (i / plan.steps.length) * 60;
    
    onProgress?.(
      'Step ' + (i + 1) + '/' + plan.steps.length + ': ' + step.description,
      baseProgress,
      'Executing...'
    );
    
    // Execute the step
    let result = await executeTool(step.params, (status, progress) => {
      onProgress?.(
        status,
        baseProgress + (progress / 100) * (60 / plan.steps.length),
        step.description
      );
    });
    
    // ═══════════════════════════════════════════════════════════════════════════════
    // PHASE 4: VERIFY & FIX - Self-correction loop
    // ═══════════════════════════════════════════════════════════════════════════════
    
    if (!result.success && corrections < maxCorrections) {
      onProgress?.('Detecting issue, attempting self-correction...', baseProgress, result.message);
      
      // Try to fix the issue
      const fixResult = await attemptFix(step, result, fileStructures);
      
      if (fixResult.success) {
        result = fixResult;
        corrections++;
        onProgress?.('Fixed! Continuing...', baseProgress, 'Self-corrected successfully');
      }
    }
    
    // Validate the result
    if (result.success && step.params.filename) {
      const issues = await validateFile(step.params.filename);
      if (issues.length > 0 && corrections < maxCorrections) {
        onProgress?.('Validating output...', baseProgress, 'Checking for issues...');
        await fixIssues(step.params.filename, issues);
        corrections++;
      }
    }
    
    results.push(result);
    executedSteps.push(step.description + ': ' + (result.success ? '✓' : '✗'));
    
    // Record in history
    if (step.params.filename) {
      addFileHistoryEntry(step.params.filename, step.description + ' - ' + (result.success ? 'success' : 'failed'));
    }
  }
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PHASE 5: FINALIZE - Generate summary
  // ═══════════════════════════════════════════════════════════════════════════════
  
  onProgress?.('Finalizing...', 95, 'Completing task...');
  
  const successCount = results.filter(r => r.success).length;
  const failCount = results.filter(r => !r.success).length;
  
  let message = '';
  if (failCount === 0) {
    message = '✅ All ' + results.length + ' steps completed successfully!';
  } else {
    message = '⚠️ ' + successCount + ' succeeded, ' + failCount + ' failed. Applied ' + corrections + ' corrections.';
  }
  
  message += '\n\nSteps executed:\n' + executedSteps.join('\n');
  
  onProgress?.('Done!', 100);
  
  return {
    success: failCount === 0,
    message,
    plan,
    results,
    corrections
  };
}

/**
 * Generate a multi-step execution plan using AI
 */
async function generatePlan(
  task: string,
  fileStructures: Record<string, any>,
  fileHistories: Record<string, string[]>
): Promise<AgentPlan> {
  
  const structureContext = Object.keys(fileStructures).length > 0
    ? 'Existing file structures:\n' + JSON.stringify(fileStructures, null, 2)
    : 'No existing files.';
  
  const historyContext = Object.keys(fileHistories).length > 0
    ? 'File edit history:\n' + JSON.stringify(fileHistories, null, 2)
    : '';
  
  const prompt = `You are an expert document automation agent. Create a detailed execution plan.

TASK: ${task}

${structureContext}

${historyContext}

AVAILABLE TOOLS (only these 9 tools can be used):
1. create_document - Create Word document with filename, title, sections
2. edit_document - Modify Word document (use "operations" array with types: replace_text, add_heading, add_paragraph, add_bullet_list, add_table, add_table_row, add_section, add_page_break, add_chart, add_image, add_header, add_footer)
3. create_spreadsheet - Create Excel workbook with filename, sheets
4. edit_spreadsheet - Modify Excel workbook (use "operations" array with types: add_row, update_cell, set_formula, add_column, set_cell_style, set_number_format, merge_cells, freeze_panes, add_chart)
5. read_document - Read document content
6. analyze_file - Analyze file structure
7. list_files - List all files
8. delete_file - Delete a file
9. rename_file - Rename a file

IMPORTANT: Operations like add_row, update_cell etc. are NOT tools - they go inside edit_document/edit_spreadsheet "operations" array.

PROFESSIONAL STYLES:
- Headers: bold, color 1E40AF (blue) or 2D3748 (dark)
- Tables: headerBg 2D3748, headerFont FFFFFF
- Financial: number format $#,##0.00
- Percentage: number format 0.00%

Return JSON with this exact structure:
{
  "reasoning": "Brief explanation of approach",
  "steps": [
    {
      "tool": "tool_name",
      "params": { ... },
      "description": "What this step does"
    }
  ]
}

Generate the complete plan. Be specific with file names, operations, and styling.`;

  const response = await chatCompletion([
    { role: 'system', content: 'You are an expert document automation planner. Always respond with valid JSON.' },
    { role: 'user', content: prompt }
  ], { temperature: 0.2 });
  
  try {
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      return {
        reasoning: parsed.reasoning || 'Executing task...',
        steps: parsed.steps || []
      };
    }
  } catch (e) {
    // Fallback plan
  }
  
  // Fallback: create simple plan
  return {
    reasoning: 'Creating a document based on your request.',
    steps: [{
      tool: 'create_document',
      params: {
        tool: 'create_document',
        filename: 'document.docx',
        title: 'Document',
        sections: [{ heading: 'Content', content: task }]
      },
      description: 'Create document with task content'
    }]
  };
}

/**
 * Attempt to fix a failed step using AI
 */
async function attemptFix(
  step: PlanStep,
  result: ToolResult,
  fileStructures: Record<string, any>
): Promise<ToolResult> {
  
  const fixPrompt = `A document operation failed. Fix it.

ORIGINAL STEP: ${JSON.stringify(step)}
ERROR: ${result.message}
FILE STRUCTURES: ${JSON.stringify(fileStructures)}

Analyze the error and create a corrected version of the operation.
Return JSON with the corrected tool call:
{
  "tool": "corrected_tool_name",
  "params": { ... },
  "reason": "Why this should fix the issue"
}`;

  try {
    const response = await chatCompletion([
      { role: 'system', content: 'You fix failed document operations. Always respond with valid JSON.' },
      { role: 'user', content: fixPrompt }
    ], { temperature: 0.2 });
    
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const fix = JSON.parse(jsonMatch[0]);
      return await executeTool(fix.params);
    }
  } catch (e) {
    // Fix failed
  }
  
  return result; // Return original error
}

export { generatePlan, attemptFix };
