import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
	plugins: [react()],
	server: {
		port: 8080
	},
	css: {
		preprocessorOptions: {
			scss: {
				silenceDeprecations: ['mixed-decls', 'color-functions', 'global-builtin', 'import'],
				quietDeps: true, // Add this line to suppress warnings (above needed for bootstrap SCSS Dart messages)
				//api: 'modern',
			},
		}
	},
	base: '/PptxGenJS/demos/vite/'
})
