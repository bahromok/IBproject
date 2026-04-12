import { chatCompletion, chatCompletionJSON, type ChatMessage } from './ai-client';
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
 * Generate a multi-step execution plan using AI with enhanced reasoning
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
  
  const prompt = `You are an expert document automation agent with advanced analytical capabilities. Create a detailed execution plan.

TASK: ${task}

${structureContext}

${historyContext}

AVAILABLE TOOLS (only these 14 tools can be used):
1. read_document - Read Word document content and structure
2. get_paragraph_index - Get indexed view of all document blocks (USE BEFORE EDITING)
3. create_document - Create Word document with filename, title, sections
4. edit_document - Modify Word document (use "operations" array)
5. read_spreadsheet_full - Read ALL Excel cells with values, formulas, styles (USE BEFORE EDITING)
6. create_spreadsheet - Create Excel workbook with filename, sheets
7. edit_spreadsheet - Modify Excel workbook (use "operations" array)
8. bulk_update_cells - Efficiently update many Excel cells at once
9. analyze_file - Analyze file structure and metadata
10. list_files - List all files in storage
11. delete_file - Delete a file
12. rename_file - Rename a file
13. get_document_xml - Get raw Word XML for advanced debugging
14. set_document_xml - Replace entire document body with XML (NUCLEAR OPTION)

IMPORTANT RULES:
- Operations like add_row, update_cell etc. go INSIDE edit_document/edit_spreadsheet "operations" array
- ALWAYS read before editing: use get_paragraph_index for Word, read_spreadsheet_full for Excel
- Use indexed operations (insert_before_index, replace_at_index, etc.) for precise Word editing
- Use bulk_update_cells for efficient Excel updates

PROFESSIONAL STYLES:
- Headers: bold, color 1E40AF (blue) or 2D3748 (dark gray)
- Tables: headerBg 2D3748, headerFont FFFFFF (white)
- Financial numbers: format $#,##0.00
- Percentages: format 0.00%
- Alternating rows: use light gray background F3F4F6

Return JSON with this exact structure:
{
  "reasoning": "Brief explanation of your approach and analysis",
  "steps": [
    {
      "tool": "tool_name",
      "params": { ... },
      "description": "What this step does in one sentence"
    }
  ]
}

Generate the complete plan. Be specific with file names, operations, data values, and styling.`;

  // Use enhanced JSON mode with reasoning capability
  try {
    const parsed = await chatCompletionJSON<{ reasoning: string; steps: PlanStep[] }>([
      { role: 'system', content: 'You are an expert document automation planner. Always respond with valid JSON only.' },
      { role: 'user', content: prompt }
    ], { 
      temperature: 0.2,
      taskType: 'reasoning' // Use best model for planning
    });
    
    return {
      reasoning: parsed.reasoning || 'Analyzing task and creating optimal execution plan...',
      steps: parsed.steps || []
    };
  } catch (e) {
    console.error('Plan generation failed, using fallback:', e);
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
 * Attempt to fix a failed step using AI with enhanced analysis
 */
async function attemptFix(
  step: PlanStep,
  result: ToolResult,
  fileStructures: Record<string, any>
): Promise<ToolResult> {
  
  const fixPrompt = `A document operation failed. Analyze the error and create a corrected version.

ORIGINAL STEP: ${JSON.stringify(step)}
ERROR: ${result.message}
FILE STRUCTURES: ${JSON.stringify(fileStructures)}

Analyze what went wrong and provide a corrected operation.
Return JSON with this structure:
{
  "tool": "corrected_tool_name",
  "params": { ... },
  "reason": "Why this fix should work"
}`;

  try {
    const fix = await chatCompletionJSON<{ tool: string; params: any; reason: string }>([
      { role: 'system', content: 'You fix failed document operations. Always respond with valid JSON only.' },
      { role: 'user', content: fixPrompt }
    ], { 
      temperature: 0.2,
      taskType: 'analysis' // Use analytical model for debugging
    });
    
    return await executeTool(fix.params);
  } catch (e) {
    console.error('Fix attempt failed:', e);
  }
  
  return result; // Return original error
}

export { generatePlan, attemptFix };
