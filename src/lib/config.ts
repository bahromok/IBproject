export const APP_CONFIG = {
  ai: {
    apiKey: process.env.GROQ_API_KEY || '',
    model: 'llama-3.3-70b-versatile',
    temperature: 0.7,
    maxTokens: 4096,
    
    // Gemini configuration
    geminiApiKey: process.env.GEMINI_API_KEY || '',
    geminiModel: 'gemini-1.5-flash',
    
    // Available models for fallback
    models: [
      'llama-3.1-8b-instant',      // Fastest, cheapest
      'llama-3.3-70b-versatile',   // Default, balanced
      'llama-3.1-70b-versatile',   // Alternative
      'mixtral-8x7b-32768',        // Complex tasks
      'gemma2-9b-it',              // Backup
    ],
    
    // Gemini models
    geminiModels: [
      'gemini-1.5-flash',      // Fast, efficient
      'gemini-1.5-pro',        // Most capable
      'gemini-2.0-flash-exp',  // Experimental fast
    ],
  },
  storage: {
    directory: 'storage',
  },
};
