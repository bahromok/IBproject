export const APP_CONFIG = {
  ai: {
    apiKey: process.env.AI_API_KEY || '',
    apiEndpoint: process.env.AI_API_ENDPOINT || 'https://api.openai.com/v1',
    model: process.env.AI_MODEL || 'gpt-4-turbo-preview',
    temperature: 0.7,
    maxTokens: 4096,
    // Smart fallback models organized by capability
    models: {
      reasoning: ['o1-preview', 'o1-mini', 'claude-3-opus-20240229', 'gpt-4-turbo-preview'],
      analysis: ['gpt-4-turbo-preview', 'claude-3-sonnet-20240229', 'gpt-4'],
      fast: ['gpt-3.5-turbo', 'gpt-4-turbo', 'mixtral-8x7b-32768'],
      creative: ['claude-3-haiku-20240307', 'gpt-4'],
    },
  },
  storage: {
    directory: 'storage',
  },
};
