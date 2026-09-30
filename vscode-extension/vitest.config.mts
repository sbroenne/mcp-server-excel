import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		environment: 'node',
		include: ['tests/**/*.test.ts'],
		alias: { vscode: fileURLToPath(new URL('./tests/vscode.ts', import.meta.url)) },
		clearMocks: true,
		restoreMocks: true,
		unstubGlobals: true,
		testTimeout: 5000,
		hookTimeout: 5000
	}
});
