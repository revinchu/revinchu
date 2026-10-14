// UI와 서버가 같은 공급자·기본 모델을 사용한다. API 키는 이 설정에 저장하지 않는다.
export const AI_PROVIDERS = Object.freeze({
  openai: { label: 'ChatGPT · OpenAI', model: 'gpt-4.1-mini', url: 'https://api.openai.com/v1/chat/completions' },
  anthropic: { label: 'Claude', model: 'claude-haiku-4-5', url: 'https://api.anthropic.com/v1/messages' },
  perplexity: { label: 'Perplexity', model: 'sonar', url: 'https://api.perplexity.ai/v1/sonar' },
  xai: { label: 'Grok · xAI', model: 'grok-4.7', url: 'https://api.x.ai/v1/chat/completions' },
  deepseek: { label: 'DeepSeek', model: 'deepseek-flash', url: 'https://api.deepseek.com/chat/completions' },
});
