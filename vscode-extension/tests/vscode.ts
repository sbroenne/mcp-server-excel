import { vi } from 'vitest';

export class McpStdioServerDefinition {
	constructor(
		public label: string,
		public command: string,
		public args: string[] = [],
		public env: Record<string, string | number | null> = {},
		public version?: string
	) {}
}

export class CancellationError extends Error {
	constructor() {
		super('Canceled');
		this.name = 'Canceled';
	}
}

export const noCancellation = {
	isCancellationRequested: false,
	onCancellationRequested: vi.fn(() => ({ dispose: vi.fn() }))
};

export const Uri = {
	file: (value: string) => ({ fsPath: value }),
	parse: (value: string) => ({ toString: () => value })
};

export const ExtensionKind = { UI: 1, Workspace: 2 };
export const output = { appendLine: vi.fn(), dispose: vi.fn(), show: vi.fn() };
export const lm = {
	registerMcpServerDefinitionProvider: vi.fn(() => ({ dispose: vi.fn() }))
};
export const window = {
	createOutputChannel: vi.fn(() => output),
	showInformationMessage: vi.fn(async (): Promise<string | undefined> => undefined),
	showErrorMessage: vi.fn(async () => undefined)
};
export const env = { openExternal: vi.fn(async () => true) };
