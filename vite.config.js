import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const repoName = process.env.GITHUB_REPOSITORY?.split('/')[1] || 'committee-profile-extractor';
const basePath = process.env.VITE_BASE_PATH || `/${repoName}/`;

export default defineConfig({
  base: basePath,
  plugins: [react()],
  server: {
    port: 5173,
    host: true,
  },
});
