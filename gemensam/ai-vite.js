// Vite-plugin som kopplar ett Vue-projekt till AI-tjänsten utan att nyckeln
// hamnar i webbläsaren.
//
// Webbläsaren skickar sina anrop till /ai på den egna utvecklingsservern.
// Servern (Vite, som körs i Node på din dator) läser nyckeln från .env i
// repots rotmapp, lägger på den och skickar vidare till AI-tjänsten.
// Nyckeln syns alltså aldrig i webbläsarens kod eller nätverksflik.
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export function aiPlugin() {
  const envFile = fileURLToPath(new URL('../.env', import.meta.url));
  if (existsSync(envFile)) process.loadEnvFile(envFile);

  const baseURL = process.env.AI_BASE_URL ?? 'https://llm.aiqu.ai/v1';
  const apiKey = process.env.AI_API_KEY;
  if (!apiKey) {
    console.warn('\n⚠️  Hittar ingen AI_API_KEY. Kopiera .env.example till .env i repots rotmapp och fyll i din nyckel.\n');
  }

  return {
    name: 'ai-proxy',
    config: () => ({
      // Modellnamnen är inte hemliga, så de får följa med till webbläsaren.
      define: {
        __AI_MODEL__: JSON.stringify(process.env.AI_MODEL ?? 'gpt-oss-120b'),
        __AI_EMBED_MODEL__: JSON.stringify(process.env.AI_EMBED_MODEL ?? 'bge-m3'),
      },
      server: {
        proxy: {
          '/ai': {
            target: baseURL,
            changeOrigin: true,
            rewrite: (path) => path.replace(/^\/ai/, ''),
            configure: (proxy) => {
              proxy.on('proxyReq', (req) => {
                req.setHeader('Authorization', `Bearer ${apiKey}`);
              });
            },
          },
        },
      },
    }),
  };
}
