export const APP_CONFIG = {
  ai: {
    apiKey: process.env.GROQ_API_KEY || '',
    model: 'llama-3.3-70b-versatile',
    temperature: 0.7,
    maxTokens: 4096,
    // Available models for fallback
    models: [
      'llama-3.1-8b-instant',      // Fastest, cheapest
      'llama-3.3-70b-versatile',   // Default, balanced
      'llama-3.1-70b-versatile',   // Alternative
      'mixtral-8x7b-32768',        // Complex tasks
      'gemma2-9b-it',              // Backup
    ],
  },
  storage: {
    directory: 'storage',
  },
};
