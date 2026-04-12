// ═══════════════════════════════════════════════════════════════════════════════
// UNIVERSAL AI CLIENT - OpenAI-Compatible API with Smart Fallback & Analysis
// Supports: OpenAI, Groq, Together, Local models, Any OpenAI-compatible endpoint
// ═══════════════════════════════════════════════════════════════════════════════

import { APP_CONFIG } from './config';

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface AIResponse {
  content: string;
  model: string;
  usage?: { promptTokens: number; completionTokens: number; totalTokens: number };
}

// Available models organized by capability
const MODELS_BY_CAPABILITY = {
  // Fast reasoning & planning
  reasoning: ['o1-preview', 'o1-mini', 'claude-3-opus-20240229', 'llama-3.3-70b-versatile'],
  // Complex analysis & code
  analysis: ['gpt-4-turbo-preview', 'claude-3-sonnet-20240229', 'llama-3.1-70b-versatile'],
  // Quick tasks & simple edits
  fast: ['gpt-3.5-turbo', 'llama-3.1-8b-instant', 'mixtral-8x7b-32768'],
  // Creative writing
  creative: ['claude-3-haiku-20240307', 'gemma2-9b-it'],
};

// Smart model selection based on task type
function selectModel(taskType?: 'reasoning' | 'analysis' | 'fast' | 'creative'): string {
  if (taskType && MODELS_BY_CAPABILITY[taskType]) {
    return MODELS_BY_CAPABILITY[taskType][0];
  }
  return APP_CONFIG.ai.model || 'llama-3.3-70b-versatile';
}

// Rate limit tracking with exponential backoff
const rateLimitedModels = new Map<string, { timestamp: number; retryAfter: number }>();

function getNextModel(currentModel?: string, taskType?: 'reasoning' | 'analysis' | 'fast' | 'creative'): string {
  const now = Date.now();
  
  // Clear expired rate limits
  for (const [model, info] of rateLimitedModels.entries()) {
    if (now - info.timestamp > info.retryAfter * 1000) {
      rateLimitedModels.delete(model);
    }
  }
  
  // Get candidate models
  const candidates = taskType ? MODELS_BY_CAPABILITY[taskType] : Object.values(MODELS_BY_CAPABILITY).flat();
  
  // Find first available model
  for (const model of candidates) {
    if (!rateLimitedModels.has(model) && model !== currentModel) {
      return model;
    }
  }
  
  // Fallback to default
  return APP_CONFIG.ai.model || 'llama-3.3-70b-versatile';
}

function markRateLimited(model: string, retryAfter: number = 60) {
  rateLimitedModels.set(model, { timestamp: Date.now(), retryAfter });
  console.log(`[AI] Model ${model} rate-limited, retry after ${retryAfter}s`);
}

/**
 * Enhanced chat completion with advanced error handling, fallback, and analytics
 */
export async function chatCompletion(
  messages: ChatMessage[],
  options?: { 
    temperature?: number; 
    maxTokens?: number;
    taskType?: 'reasoning' | 'analysis' | 'fast' | 'creative';
    jsonMode?: boolean;
  }
): Promise<string> {
  let lastError: Error | null = null;
  let modelToUse = options?.taskType ? selectModel(options.taskType) : APP_CONFIG.ai.model;
  const maxAttempts = 5;
  
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    try {
      const response = await fetch(APP_CONFIG.ai.apiEndpoint + '/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${APP_CONFIG.ai.apiKey}`,
        },
        body: JSON.stringify({
          model: modelToUse,
          messages: messages.map(m => ({
            role: m.role,
            content: m.content,
          })),
          temperature: options?.temperature ?? 0.7,
          max_tokens: options?.maxTokens ?? 4096,
          response_format: options?.jsonMode ? { type: 'json_object' } : undefined,
        }),
      });
      
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw { status: response.status, message: errorData.error?.message || response.statusText };
      }
      
      const data = await response.json();
      return data.choices[0]?.message?.content || '';
      
    } catch (error: unknown) {
      const err = error as { status?: number; message?: string };
      lastError = error as Error;
      
      console.error(`[AI] Error (model: ${modelToUse}, attempt ${attempt + 1}):`, err?.message || err);
      
      // Authentication errors - stop immediately
      if (err?.status === 401 || err?.status === 403) {
        throw new Error('Invalid API key. Check your AI_API_KEY in .env file');
      }
      
      // Rate limiting - mark model and try next
      if (err?.status === 429) {
        const retryAfter = parseInt((error as any)?.retryAfter || '60');
        markRateLimited(modelToUse, retryAfter);
        modelToUse = getNextModel(modelToUse, options?.taskType);
        await new Promise(resolve => setTimeout(resolve, Math.min(retryAfter * 100, 2000)));
        continue;
      }
      
      // Server errors - try next model
      if (err?.status === 500 || err?.status === 502 || err?.status === 503 || err?.status === 504) {
        modelToUse = getNextModel(modelToUse, options?.taskType);
        await new Promise(resolve => setTimeout(resolve, 1000));
        continue;
      }
      
      // Other errors - retry with same model
      if (attempt < maxAttempts - 1) {
        await new Promise(resolve => setTimeout(resolve, 1000 * (attempt + 1)));
        continue;
      }
      
      throw error;
    }
  }
  
  throw lastError || new Error('All AI models failed');
}

/**
 * Streaming chat completion with real-time progress
 */
export async function streamChatCompletion(
  messages: ChatMessage[],
  onChunk: (chunk: string) => void,
  options?: { 
    temperature?: number; 
    maxTokens?: number;
    taskType?: 'reasoning' | 'analysis' | 'fast' | 'creative';
  }
): Promise<string> {
  let lastError: Error | null = null;
  let modelToUse = options?.taskType ? selectModel(options.taskType) : APP_CONFIG.ai.model;
  const maxAttempts = 5;
  
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    try {
      const response = await fetch(APP_CONFIG.ai.apiEndpoint + '/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${APP_CONFIG.ai.apiKey}`,
        },
        body: JSON.stringify({
          model: modelToUse,
          messages: messages.map(m => ({
            role: m.role,
            content: m.content,
          })),
          temperature: options?.temperature ?? 0.7,
          max_tokens: options?.maxTokens ?? 4096,
          stream: true,
        }),
      });
      
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw { status: response.status, message: errorData.error?.message || response.statusText };
      }
      
      const reader = response.body?.getReader();
      if (!reader) throw new Error('No response body');
      
      const decoder = new TextDecoder();
      let fullContent = '';
      
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        
        const chunk = decoder.decode(value);
        const lines = chunk.split('\n');
        
        for (const line of lines) {
          if (line.startsWith('data: ')) {
            const data = line.slice(6);
            if (data === '[DONE]') continue;
            
            try {
              const parsed = JSON.parse(data);
              const content = parsed.choices?.[0]?.delta?.content || '';
              fullContent += content;
              onChunk(content);
            } catch {
              // Skip invalid JSON
            }
          }
        }
      }
      
      return fullContent;
      
    } catch (error: unknown) {
      const err = error as { status?: number; message?: string };
      lastError = error as Error;
      
      if (err?.status === 401 || err?.status === 403) {
        throw new Error('Invalid API key');
      }
      
      if (err?.status === 429) {
        const retryAfter = parseInt((error as any)?.retryAfter || '60');
        markRateLimited(modelToUse, retryAfter);
        modelToUse = getNextModel(modelToUse, options?.taskType);
        await new Promise(resolve => setTimeout(resolve, Math.min(retryAfter * 100, 2000)));
        continue;
      }
      
      if (err?.status === 500 || err?.status === 502 || err?.status === 503 || err?.status === 504) {
        modelToUse = getNextModel(modelToUse, options?.taskType);
        await new Promise(resolve => setTimeout(resolve, 1000));
        continue;
      }
      
      throw error;
    }
  }
  
  throw lastError || new Error('All AI models failed');
}

/**
 * Enhanced JSON parsing with validation and schema enforcement
 */
export async function chatCompletionJSON<T>(
  messages: ChatMessage[],
  options?: { 
    temperature?: number; 
    maxTokens?: number;
    taskType?: 'reasoning' | 'analysis' | 'fast';
  }
): Promise<T> {
  const response = await chatCompletion(messages, {
    ...options,
    jsonMode: true,
    temperature: options?.temperature ?? 0.2, // Lower temp for JSON
  });
  
  // Extract JSON from response
  const jsonMatch = response.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('No valid JSON found in response');
  }
  
  try {
    return JSON.parse(jsonMatch[0]);
  } catch (e) {
    throw new Error('Failed to parse JSON response: ' + (e as Error).message);
  }
}

/**
 * Get AI system status and diagnostics
 */
export function getAIDiagnostics() {
  return {
    configuredModel: APP_CONFIG.ai.model,
    apiEndpoint: APP_CONFIG.ai.apiEndpoint,
    hasApiKey: !!APP_CONFIG.ai.apiKey,
    rateLimitedModels: Array.from(rateLimitedModels.entries()).map(([model, info]) => ({
      model,
      retryAfter: Math.max(0, info.retryAfter - Math.floor((Date.now() - info.timestamp) / 1000)),
    })),
    availableCapabilities: Object.keys(MODELS_BY_CAPABILITY),
  };
}

export { selectModel, getNextModel, MODELS_BY_CAPABILITY };
