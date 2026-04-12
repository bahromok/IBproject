/**
 * Google Gemini Client - Support for Gemini models with vision capabilities
 */

import { GoogleGenerativeAI, GenerativeModel } from '@google/generative-ai';
import { APP_CONFIG } from './config';

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

const AVAILABLE_MODELS = [
  'gemini-1.5-flash',      // Fast, efficient
  'gemini-1.5-pro',        // Most capable
  'gemini-2.0-flash-exp',  // Experimental fast
];

let genAI: GoogleGenerativeAI | null = null;
let model: GenerativeModel | null = null;

// Track rate-limited models
const rateLimitedModels = new Map<string, number>();

/**
 * Initialize Gemini client
 */
function initGemini(): void {
  if (!genAI) {
    const apiKey = process.env.GEMINI_API_KEY || APP_CONFIG.ai.geminiApiKey;
    if (!apiKey) {
      throw new Error('GEMINI_API_KEY not found in environment variables');
    }
    genAI = new GoogleGenerativeAI(apiKey);
  }
}

/**
 * Get model instance
 */
function getModel(modelName?: string): GenerativeModel {
  initGemini();
  
  const modelToUse = modelName || APP_CONFIG.ai.geminiModel || 'gemini-1.5-flash';
  
  if (!genAI) {
    throw new Error('Gemini client not initialized');
  }
  
  return genAI.getGenerativeModel({ model: modelToUse });
}

/**
 * Mark a model as rate-limited
 */
function markRateLimited(modelName: string) {
  rateLimitedModels.set(modelName, Date.now());
}

/**
 * Convert chat messages to Gemini format
 */
function convertMessages(messages: ChatMessage[]): any[] {
  const geminiMessages: any[] = [];
  
  for (const msg of messages) {
    if (msg.role === 'system') {
      // System messages become user messages with context
      geminiMessages.push({
        role: 'user',
        parts: [{ text: `System instruction: ${msg.content}` }]
      });
      geminiMessages.push({
        role: 'model',
        parts: [{ text: 'Understood. I will follow these instructions.' }]
      });
    } else if (msg.role === 'user') {
      geminiMessages.push({
        role: 'user',
        parts: [{ text: msg.content }]
      });
    } else if (msg.role === 'assistant') {
      geminiMessages.push({
        role: 'model',
        parts: [{ text: msg.content }]
      });
    }
  }
  
  return geminiMessages;
}

/**
 * Send a chat completion request to Gemini
 */
export async function chatCompletion(
  messages: ChatMessage[],
  options?: { temperature?: number; maxTokens?: number; model?: string }
): Promise<string> {
  let lastError: Error | null = null;
  let modelToUse = options?.model || APP_CONFIG.ai.geminiModel || 'gemini-1.5-flash';
  
  // Try up to 3 different models
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      initGemini();
      const currentModel = getModel(modelToUse);
      
      const chat = currentModel.startChat({
        history: convertMessages(messages.slice(0, -1)),
        generationConfig: {
          temperature: options?.temperature ?? 0.7,
          maxOutputTokens: options?.maxTokens ?? 8192,
        },
      });
      
      const lastMessage = messages[messages.length - 1];
      const result = await chat.sendMessage(lastMessage.content);
      const response = await result.response;
      
      return response.text();
    } catch (error: unknown) {
      const err = error as { status?: number; message?: string };
      lastError = error as Error;
      
      console.error(`Gemini API Error (model: ${modelToUse}):`, err?.message || err);
      
      // Handle different error types
      if (err?.message?.includes('API key') || err?.message?.includes('invalid')) {
        throw new Error('Gemini API key invalid or expired. Please check your GEMINI_API_KEY in .env file');
      }
      
      if (err?.status === 429 || err?.message?.includes('quota') || err?.message?.includes('rate limit')) {
        // Rate limited - mark this model and try next
        markRateLimited(modelToUse);
        
        // Try next available model
        const nextModel = AVAILABLE_MODELS.find(m => m !== modelToUse && !rateLimitedModels.has(m));
        if (nextModel) {
          modelToUse = nextModel;
          console.log(`Rate limited on ${modelToUse}, switching to: ${modelToUse}`);
          await new Promise(resolve => setTimeout(resolve, 1000));
          continue;
        }
      }
      
      if (err?.status === 500 || err?.status === 502 || err?.status === 503) {
        // Server error - try next model
        const nextModel = AVAILABLE_MODELS.find(m => m !== modelToUse);
        if (nextModel) {
          modelToUse = nextModel;
          console.log(`Server error, switching to: ${modelToUse}`);
          continue;
        }
      }
      
      throw error;
    }
  }
  
  throw lastError || new Error('All Gemini models failed');
}

/**
 * Stream a chat completion request
 */
export async function streamChatCompletion(
  messages: ChatMessage[],
  onChunk: (chunk: string) => void,
  options?: { temperature?: number; maxTokens?: number; model?: string }
): Promise<string> {
  let lastError: Error | null = null;
  let modelToUse = options?.model || APP_CONFIG.ai.geminiModel || 'gemini-1.5-flash';
  
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      initGemini();
      const currentModel = getModel(modelToUse);
      
      const chat = currentModel.startChat({
        history: convertMessages(messages.slice(0, -1)),
        generationConfig: {
          temperature: options?.temperature ?? 0.7,
          maxOutputTokens: options?.maxTokens ?? 8192,
        },
      });
      
      const lastMessage = messages[messages.length - 1];
      const result = await chat.sendMessageStream(lastMessage.content);
      
      let fullContent = '';
      for await (const chunk of result.stream) {
        const text = chunk.text();
        fullContent += text;
        onChunk(text);
      }
      
      return fullContent;
    } catch (error: unknown) {
      const err = error as { status?: number; message?: string };
      lastError = error as Error;
      
      if (err?.message?.includes('API key') || err?.message?.includes('invalid')) {
        throw new Error('Gemini API key invalid. Check .env file');
      }
      
      if (err?.status === 429 || err?.message?.includes('quota') || err?.message?.includes('rate limit')) {
        markRateLimited(modelToUse);
        const nextModel = AVAILABLE_MODELS.find(m => m !== modelToUse && !rateLimitedModels.has(m));
        if (nextModel) {
          modelToUse = nextModel;
          console.log(`Rate limited, switching to: ${modelToUse}`);
          await new Promise(resolve => setTimeout(resolve, 1000));
          continue;
        }
      }
      
      if (err?.status === 500 || err?.status === 502 || err?.status === 503) {
        const nextModel = AVAILABLE_MODELS.find(m => m !== modelToUse);
        if (nextModel) {
          modelToUse = nextModel;
          continue;
        }
      }
      
      throw error;
    }
  }
  
  throw lastError || new Error('All Gemini models failed');
}

/**
 * Get current API configuration
 */
export function getApiConfig() {
  return {
    provider: 'gemini',
    model: APP_CONFIG.ai.geminiModel || 'gemini-1.5-flash',
    availableModels: AVAILABLE_MODELS,
    hasApiKey: !!process.env.GEMINI_API_KEY,
    rateLimitedModels: Array.from(rateLimitedModels.keys()),
  };
}

/**
 * Get available models list
 */
export function getAvailableModels(): string[] {
  return AVAILABLE_MODELS;
}

/**
 * Test Gemini connection
 */
export async function testConnection(): Promise<boolean> {
  try {
    initGemini();
    const testModel = getModel('gemini-1.5-flash');
    const result = await testModel.generateContent('Say "OK"');
    const response = await result.response;
    return response.text().toLowerCase().includes('ok');
  } catch (error) {
    console.error('Gemini connection test failed:', error);
    return false;
  }
}
