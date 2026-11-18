import OpenAI from 'openai'

// Klienten pratar med /ai på din egen utvecklingsserver, inte direkt med
// AI-tjänsten. Servern lägger till den riktiga nyckeln (se gemensam/ai-vite.js).
// Därför står det ingen nyckel här, och därför är det ofarligt att slå på
// dangerouslyAllowBrowser: det finns ingen hemlighet i webbläsaren att läcka.
export const client = new OpenAI({
  baseURL: `${window.location.origin}/ai`,
  apiKey: 'laggs-till-av-servern',
  dangerouslyAllowBrowser: true,
})

export const MODEL = __AI_MODEL__
export const EMBED_MODEL = __AI_EMBED_MODEL__

// Skickar en lista med meddelanden och returnerar svarstexten.
export async function chatta(messages, options = {}) {
  const response = await client.chat.completions.create({ model: MODEL, messages, ...options })
  return response.choices[0].message.content
}
