import Groq from 'groq-sdk';
import { APP_CONFIG } from './config';

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

const AVAILABLE_MODELS = [
  'llama-3.1-8b-instant',      // Arzon, lekin parametrlari kam va nisbatan kuchsiz
  'llama-3.3-70b-versatile',   // Standart, muvozanatli
  'llama-3.1-70b-versatile',   // Alternativ
  'mixtral-8x7b-32768',        // Murakkab vazifalar uchun yaxshi
  'gemma2-9b-it',              // Zaxira
];

const groq = new Groq({
  apiKey: APP_CONFIG.ai.apiKey,
  dangerouslyAllowBrowser: false,
});

// Track rate-limited models
const rateLimitedModels = new Map<string, number>();

/**
 * Get the next available model (not rate-limited)
 */
function getNextModel(currentModel?: string): string {
  const now = Date.now();
  
  // Clear expired rate limits (60 seconds)
  for (const [model, timestamp] of rateLimitedModels.entries()) {
    if (now - timestamp > 60000) {
      rateLimitedModels.delete(model);
    }
  }
  
  // Find first model not rate-limited
  for (const model of AVAILABLE_MODELS) {
    if (!rateLimitedModels.has(model) && model !== currentModel) {
      return model;
    }
  }
  
  // If all are rate-limited, use the first one anyway
  return AVAILABLE_MODELS[0];
}

/**
 * Mark a model as rate-limited
 */
function markRateLimited(model: string) {
  rateLimitedModels.set(model, Date.now());
}

/**
 * Send a chat completion request with automatic model fallback
 */
export async function chatCompletion(
  messages: ChatMessage[],
  options?: { temperature?: number; maxTokens?: number }
): Promise<string> {
  let lastError: Error | null = null;
  let modelToUse = APP_CONFIG.ai.model;
  
  // Try up to 3 different models
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const response = await groq.chat.completions.create({
        model: modelToUse,
        messages: messages.map(m => ({
          role: m.role as 'system' | 'user' | 'assistant',
          content: m.content,
        })),
        temperature: options?.temperature ?? 0.7,
        max_tokens: options?.maxTokens ?? 4096,
      });

      return response.choices[0]?.message?.content || '';
    } catch (error: unknown) {
      const err = error as { status?: number; message?: string };
      lastError = error as Error;
      
      console.error(`Groq API Error (model: ${modelToUse}):`, err?.message || err);
      
      // Handle different error types
      if (err?.status === 401 || err?.status === 403) {
        throw new Error('API key invalid or expired. Please check your Groq API key in .env file');
      }
      
      if (err?.status === 429) {
        // Rate limited - mark this model and try next
        markRateLimited(modelToUse);
        modelToUse = getNextModel(modelToUse);
        console.log(`Rate limited on model, switching to: ${modelToUse}`);
        
        // Wait a bit before retry
        await new Promise(resolve => setTimeout(resolve, 1000));
        continue;
      }
      
      if (err?.status === 500 || err?.status === 502 || err?.status === 503) {
        // Server error - try next model
        modelToUse = getNextModel(modelToUse);
        console.log(`Server error, switching to: ${modelToUse}`);
        continue;
      }
      
      throw error;
    }
  }
  
  throw lastError || new Error('All models failed');
}

/**
 * Stream a chat completion request with model fallback
 */
export async function streamChatCompletion(
  messages: ChatMessage[],
  onChunk: (chunk: string) => void,
  options?: { temperature?: number; maxTokens?: number }
): Promise<string> {
  let lastError: Error | null = null;
  let modelToUse = APP_CONFIG.ai.model;
  
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const stream = await groq.chat.completions.create({
        model: modelToUse,
        messages: messages.map(m => ({
          role: m.role as 'system' | 'user' | 'assistant',
          content: m.content,
        })),
        temperature: options?.temperature ?? 0.7,
        max_tokens: options?.maxTokens ?? 4096,
        stream: true,
      });

      let fullContent = '';
      for await (const chunk of stream) {
        const content = chunk.choices[0]?.delta?.content || '';
        fullContent += content;
        onChunk(content);
      }

      return fullContent;
    } catch (error: unknown) {
      const err = error as { status?: number; message?: string };
      lastError = error as Error;
      
      if (err?.status === 401 || err?.status === 403) {
        throw new Error('API key invalid. Check .env file');
      }
      
      if (err?.status === 429) {
        markRateLimited(modelToUse);
        modelToUse = getNextModel(modelToUse);
        console.log(`Rate limited, switching to: ${modelToUse}`);
        await new Promise(resolve => setTimeout(resolve, 1000));
        continue;
      }
      
      if (err?.status === 500 || err?.status === 502 || err?.status === 503) {
        modelToUse = getNextModel(modelToUse);
        continue;
      }
      
      throw error;
    }
  }
  
  throw lastError || new Error('All models failed');
}

/**
 * Get current API configuration
 */
export function getApiConfig() {
  return {
    model: APP_CONFIG.ai.model,
    availableModels: AVAILABLE_MODELS,
    hasApiKey: !!APP_CONFIG.ai.apiKey,
    rateLimitedModels: Array.from(rateLimitedModels.keys()),
  };
}

/**
 * Get available models list
 */
export function getAvailableModels(): string[] {
  return AVAILABLE_MODELS;
}
