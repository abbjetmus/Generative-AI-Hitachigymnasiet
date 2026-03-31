// Sätta upp OpenAI-klienten
import OpenAI from "https://esm.sh/openai";

const client = new OpenAI({
    baseURL: "https://llm.aiqu.ai",
    apiKey: "sk-JKVPhL8DGwfdzxMstj1IJg",
    defaultHeaders: { "x-litellm-api-key": "sk-JKVPhL8DGwfdzxMstj1IJg" },
    dangerouslyAllowBrowser: true,
});

const MODEL = "gpt-oss-120b";

// Skapa prompten och skicka förfrågan till API:et
const messages = [
    {
        role: "system",
        content: "Du är en hjälpsam AI-assistent. Svara på svenska om användaren skriver på svenska.",
    },
];

// Skicka förfrågan till API:et och hämta svaret
const response = await client.chat.completions.create({
    model: MODEL,
    messages: messages,
});

