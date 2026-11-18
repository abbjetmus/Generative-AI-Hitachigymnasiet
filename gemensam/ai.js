// Gemensam AI-klient för alla kodexempel i kursen.
// Inställningarna läses från filen .env i repots rotmapp (se .env.example).
// Byter du AI_BASE_URL och AI_MODEL kör samma kod mot en annan leverantör,
// till exempel en lokal modell i Ollama.
import OpenAI from 'openai';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const envFile = fileURLToPath(new URL('../.env', import.meta.url));
if (existsSync(envFile)) process.loadEnvFile(envFile);

if (!process.env.AI_API_KEY) {
  console.error('Hittar ingen AI_API_KEY. Kopiera .env.example till .env och fyll i din nyckel.');
  process.exit(1);
}

export const client = new OpenAI({
  baseURL: process.env.AI_BASE_URL ?? 'https://llm.aiqu.ai/v1',
  apiKey: process.env.AI_API_KEY,
});

export const MODEL = process.env.AI_MODEL ?? 'gpt-oss-120b';
export const EMBED_MODEL = process.env.AI_EMBED_MODEL ?? 'bge-m3';

// Skickar en lista med meddelanden och returnerar svarstexten.
export async function chatta(messages, options = {}) {
  const response = await client.chat.completions.create({ model: MODEL, messages, ...options });
  return response.choices[0].message.content;
}
