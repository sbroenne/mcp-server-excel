import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { listFiles, PackageManager } from '@vscode/vsce';
import { describe, expect, it, vi } from 'vitest';

const vsceRequire = createRequire(createRequire(import.meta.url).resolve('@vscode/vsce'));
const { Minimatch } = vsceRequire('minimatch') as typeof import('minimatch');

const included = [
	'package.json', 'README.md', 'LICENSE', 'icon.png',
	'bin/Sbroenne.ExcelMcp.McpServer.exe', 'out/extension.js', 'out/prerequisites.js'
];
const excluded = [
	'src/extension.ts', 'out/extension.js.map', 'out/tests/fixture.js',
	'tests/extension.test.ts', 'node_modules/minimatch/index.js',
	'.vscode/settings.json', 'tsconfig.json', 'vitest.config.mts', 'old.vsix',
	'skills/excel-mcp-report-formatting/SKILL.md',
	'skills/excel-mcp-report-formatting/VERSION',
	'skills/excel-mcp-report-formatting/references/formatting.md'
];

async function inspectFixture(extraIgnore = '', extraFiles: string[] = []) {
	const root = await mkdtemp(join(tmpdir(), 'excelmcp-vsce-files-'));
	try {
		await Promise.all([...included, ...excluded, ...extraFiles].map(async file => {
			const path = join(root, file);
			await mkdir(dirname(path), { recursive: true });
			await writeFile(path, file === 'package.json'
				? JSON.stringify({ name: 'synthetic-extension', version: '1.0.0', engines: { vscode: '^1.125.0' } })
				: 'synthetic packaging fixture');
		}));
		const ignore = await readFile(new URL('../.vscodeignore', import.meta.url), 'utf8');
		await writeFile(join(root, '.vscodeignore'), ignore + extraIgnore);
		return (await listFiles({ cwd: root, packageManager: PackageManager.None }))
			.map(file => file.replaceAll('\\', '/'));
	} finally {
		await rm(root, { recursive: true, force: true });
	}
}

describe('VSCE packaging file filters', () => {
	it('keeps the server and extension assets while excluding development files', async () => {
		const files = await inspectFixture();
		expect(files).toHaveLength(included.length);
		expect(files).toEqual(expect.arrayContaining(included));
		expect(files.some(file => file.startsWith('skills/'))).toBe(false);
	});

	it('preserves an explicit negation without including other source maps', async () => {
		const files = await inspectFixture('\n!out/retained.js.map\n', ['out/retained.js.map']);
		expect(files).toHaveLength(included.length + 1);
		expect(files).toEqual(expect.arrayContaining([...included, 'out/retained.js.map']));
		expect(files.some(file => file.startsWith('skills/'))).toBe(false);
	});
});

describe('VSCE minimatch preprocessing limits', () => {
	it('bounds generated glob alternatives during first-phase preprocessing', () => {
		const matcher = new Minimatch('safe', { optimizationLevel: 2, braceExpandMax: 64 });
		const parts = Array.from({ length: 10 }, () => '**/../a/b').join('/').split('/');
		const alternatives = matcher.firstPhasePreProcess([parts]);
		expect(alternatives.length).toBeGreaterThan(0);
		expect(alternatives.length).toBeLessThanOrEqual(64);
	});

	it('retains large alternative sets without quadratic pairwise processing', () => {
		const matcher = new Minimatch('safe', { optimizationLevel: 2 });
		const parts = Array.from({ length: 1001 }, (_, index) => ['x', String(index)]);
		const expected = parts.map(part => [...part]);
		const compare = vi.spyOn(matcher, 'partsMatch');
		expect(matcher.secondPhasePreProcess(parts)).toEqual(expected);
		expect(compare.mock.calls.length).toBe(0);
	});
});
