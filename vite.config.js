import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// IMPORTANTE: cambia 'NOMBRE-DEL-REPO' por el nombre real de tu repositorio de GitHub.
// Si publicas en un dominio propio o en Vercel/Netlify, deja base: '/'.
export default defineConfig({
  plugins: [react()],
  base: '/NOMBRE-DEL-REPO/',
});
