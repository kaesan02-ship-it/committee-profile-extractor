import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const repoName = 'committee-profile-extractor';

export default defineConfig({
  base: `/${repoName}/`,
  plugins: [react()],
  server: {
    port: 5173,
    host: true,
  },
});
