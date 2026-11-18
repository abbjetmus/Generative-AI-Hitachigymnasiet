import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
// Ändra sökvägen om projektet ligger på ett annat djup i repot.
import { aiPlugin } from '../../gemensam/ai-vite.js'

export default defineConfig({
  plugins: [vue(), aiPlugin()],
})
