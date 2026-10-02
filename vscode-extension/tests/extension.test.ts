import type { ExecFileOptions } from 'node:child_process';
import { join } from 'node:path';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as vscode from 'vscode';
import { activate } from '../src/extension';
import { noCancellation, output, window as mockedWindow } from './vscode';

const probes = vi.hoisted(() => ({
	access: vi.fn<(path: string, mode?: number) => Promise<void>>(),
	query: vi.fn<(file: string, args: readonly string[], options: ExecFileOptions,
		callback: (error: Error | null, stdout: string, stderr: string) => void) => void>()
}));

vi.mock('node:fs/promises', () => ({ access: probes.access }));
vi.mock('node:child_process', () => ({ execFile: probes.query }));

function createContext() {
	const extensionPath = join('C:', 'fixtures', 'excel-extension');
	return {
		extensionPath,
		extension: {
			id: 'sbroenne.excel-mcp',
			extensionPath,
			extensionUri: vscode.Uri.file(extensionPath),
			isActive: true,
			packageJSON: { version: '2.1.0' },
			exports: undefined,
			extensionKind: vscode.ExtensionKind.UI,
			activate: async () => undefined
		},
		subscriptions: [] as vscode.Disposable[],
		globalState: {
			get: vi.fn().mockReturnValue(true),
			update: vi.fn().mockResolvedValue(undefined),
			keys: () => [],
			setKeysForSync: vi.fn()
		}
	} satisfies Parameters<typeof activate>[0];
}

async function registeredProvider(context = createContext()) {
	await activate(context);
	const registration = vi.mocked(vscode.lm.registerMcpServerDefinitionProvider).mock.calls.at(-1);
	expect(registration?.[0]).toBe('excel-mcp');
	if (!registration) {
		throw new Error('The extension did not register its MCP provider.');
	}
	return { context, provider: registration[1] };
}

async function serverDefinition(provider: vscode.McpServerDefinitionProvider) {
	const definitions = await provider.provideMcpServerDefinitions(noCancellation);
	expect(definitions).toHaveLength(1);
	if (!definitions?.[0]) {
		throw new Error('The provider did not return its bundled server.');
	}
	return definitions[0];
}

async function resolveServer(provider: vscode.McpServerDefinitionProvider,
	token: vscode.CancellationToken = noCancellation) {
	expect(provider.resolveMcpServerDefinition).toBeTypeOf('function');
	if (!provider.resolveMcpServerDefinition) {
		throw new Error('The provider has no launch-time resolution hook.');
	}
	return provider.resolveMcpServerDefinition(await serverDefinition(provider), token);
}

beforeEach(() => {
	probes.access.mockReset().mockResolvedValue(undefined);
	probes.query.mockReset().mockImplementation((_file, _args, _options, callback) => callback(null, '', ''));
});

describe('MCP registration and launch', () => {
	it('registers the packaged version and preserves the bundled command', async () => {
		const { context, provider } = await registeredProvider();
		expect(await serverDefinition(provider)).toMatchObject({
			label: 'excel-mcp',
			command: join(context.extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe'),
			args: [],
			env: {},
			version: '2.1.0'
		});
	});

	it('uses the new packaged version after an extension update', async () => {
		const context = createContext();
		context.extension.packageJSON.version = '2.2.0';
		const { provider } = await registeredProvider(context);
		expect((await serverDefinition(provider)).version).toBe('2.2.0');
	});

	it('discovers tools without prerequisite probes or prompts', async () => {
		const { provider } = await registeredProvider();
		await serverDefinition(provider);
		expect(probes.access).not.toHaveBeenCalled();
		expect(probes.query).not.toHaveBeenCalled();
		expect(vscode.window.showErrorMessage).not.toHaveBeenCalled();
		expect(vscode.window.showInformationMessage).not.toHaveBeenCalled();
	});

	it('checks the executable and registration before allowing launch', async () => {
		const { provider } = await registeredProvider();
		const server = await resolveServer(provider);
		expect(server?.label).toBe('excel-mcp');
		expect(probes.access).toHaveBeenCalledOnce();
		expect(probes.query).toHaveBeenCalledOnce();
		const [_command, args, options] = probes.query.mock.calls[0];
		expect(args).toContain('-NonInteractive');
		expect(args.join(' ')).not.toContain('New-Object -ComObject');
		expect(options.timeout).toBe(10000);
		expect(options.windowsHide).toBe(true);
	});

	it('rejects a missing server without exposing its private path', async () => {
		probes.access.mockRejectedValue(Object.assign(new Error('private-fixture-path'), { code: 'ENOENT' }));
		const { provider } = await registeredProvider();
		await expect(resolveServer(provider)).rejects.toThrow(/reinstall/i);
		expect(probes.query).not.toHaveBeenCalled();
		expect(JSON.stringify(output.appendLine.mock.calls)).not.toContain('private-fixture-path');
	});

	it('distinguishes missing Excel from an unsuccessful registration query', async () => {
		probes.query.mockImplementation((_file, _args, _options, callback) =>
			callback(Object.assign(new Error('private-fixture'), { code: 2 }), '', ''));
		const { provider } = await registeredProvider();
		await expect(resolveServer(provider)).rejects.toThrow(/install.*Excel/i);
	});

	it('does not label permission or probe errors as missing Excel', async () => {
		probes.query.mockImplementation((_file, _args, _options, callback) =>
			callback(Object.assign(new Error('private-fixture'), { code: 1 }), '', ''));
		const { provider } = await registeredProvider();
		await expect(resolveServer(provider)).rejects.toThrow(/could not check/i);
		expect(JSON.stringify(output.appendLine.mock.calls)).not.toContain('private-fixture');
	});

	it('reports a timed-out prerequisite query', async () => {
		probes.query.mockImplementation((_file, _args, _options, callback) =>
			callback(Object.assign(new Error('private-fixture'), { killed: true }), '', ''));
		const { provider } = await registeredProvider();
		await expect(resolveServer(provider)).rejects.toThrow(/timed out/i);
	});

	it('rejects a non-Windows host before checking local files', async () => {
		vi.stubGlobal('process', { ...process, platform: 'linux' });
		const { provider } = await registeredProvider();
		await expect(resolveServer(provider)).rejects.toThrow(/Windows desktop/i);
		expect(probes.access).not.toHaveBeenCalled();
		expect(probes.query).not.toHaveBeenCalled();
	});

	it('cancels before probing without showing a setup error', async () => {
		const { provider } = await registeredProvider();
		await expect(resolveServer(provider, {
			...noCancellation,
			isCancellationRequested: true
		})).rejects.toThrow(/cancel/i);
		expect(probes.access).not.toHaveBeenCalled();
		expect(probes.query).not.toHaveBeenCalled();
		expect(vscode.window.showErrorMessage).not.toHaveBeenCalled();
	});

	it('aborts an in-flight registration query and disposes the listener', async () => {
		let cancel: (() => void) | undefined;
		let canceled = false;
		const dispose = vi.fn();
		const token: vscode.CancellationToken = {
			get isCancellationRequested() { return canceled; },
			onCancellationRequested: listener => {
				cancel = () => { canceled = true; listener(undefined); };
				return { dispose };
			}
		};
		probes.query.mockImplementation((_file, _args, options, callback) => {
			options.signal?.addEventListener('abort', () =>
				callback(Object.assign(new Error('aborted'), { name: 'AbortError' }), '', ''), { once: true });
		});
		const { provider } = await registeredProvider();
		const launch = Promise.resolve(resolveServer(provider, token));
		const rejected = expect(launch).rejects.toThrow(/cancel/i);
		await vi.waitFor(() => expect(probes.query).toHaveBeenCalledOnce());
		expect(cancel).toBeTypeOf('function');
		cancel?.();
		await rejected;
		expect(probes.query.mock.calls[0][2].signal?.aborted).toBe(true);
		expect(dispose).toHaveBeenCalledOnce();
		expect(vscode.window.showErrorMessage).not.toHaveBeenCalled();
	});
});

describe('First-run help', () => {
	it('opens the user guides instead of installation instructions after installation', async () => {
		const context = createContext();
		context.globalState.get.mockReturnValue(false);
		mockedWindow.showInformationMessage.mockResolvedValueOnce('Getting Started');
		await registeredProvider(context);
		await vi.waitFor(() => expect(vscode.env.openExternal).toHaveBeenCalledOnce());
		const [uri] = vi.mocked(vscode.env.openExternal).mock.calls[0];
		expect(uri.toString()).toBe(
			'https://excelmcpserver.dev/guides/'
		);
	});

	it('explains automatic Copilot setup without claiming the server has started', async () => {
		const context = createContext();
		context.globalState.get.mockReturnValue(false);
		await registeredProvider(context);
		const message = vi.mocked(vscode.window.showInformationMessage).mock.calls[0]?.[0];
		expect(message).toMatch(/Copilot/);
		expect(message).toMatch(/excel-mcp/);
		expect(message).toMatch(/starts excel-mcp automatically/i);
		expect(message).toMatch(/send an Excel request/i);
		expect(message).toMatch(/when needed/i);
		expect(message).toMatch(/if prompted/i);
		expect(message).not.toContain('MCP: List Servers');
		expect(message).not.toMatch(/activated|connected|now available/i);
		expect(context.globalState.update).toHaveBeenCalledWith('excelmcp.hasShownWelcome', true);
	});

	it('awaits persistence of the welcome state', async () => {
		const context = createContext();
		context.globalState.get.mockReturnValue(false);
		let release: (() => void) | undefined;
		context.globalState.update.mockImplementation(() => new Promise<void>(resolve => { release = resolve; }));
		let completed = false;
		const activation = activate(context).then(() => { completed = true; });
		await Promise.resolve();
		expect(completed).toBe(false);
		expect(release).toBeTypeOf('function');
		release?.();
		await activation;
	});

	it('creates and owns the setup output channel', async () => {
		const { context } = await registeredProvider();
		expect(vscode.window.createOutputChannel).toHaveBeenCalledWith('ExcelMcp');
		expect(context.subscriptions).toContain(output);
	});

	it('keeps the provider usable and logs a failed welcome-state write', async () => {
		const context = createContext();
		context.globalState.get.mockReturnValue(false);
		context.globalState.update.mockRejectedValue(new Error('Cannot save the welcome preference.'));
		const { provider } = await registeredProvider(context);
		expect(await resolveServer(provider)).toMatchObject({ label: 'excel-mcp' });
		expect(output.appendLine).toHaveBeenCalledWith(
			'Could not save the welcome preference. Getting-started help may appear again. Cannot save the welcome preference.'
		);
	});
});
