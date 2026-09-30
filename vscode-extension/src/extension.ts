import * as vscode from 'vscode';
import { join } from 'node:path';
import { checkLaunchPrerequisites, LaunchSetupError } from './prerequisites';

const userGuideUrl = 'https://excelmcpserver.dev/guides/';

export async function activate(context: Pick<vscode.ExtensionContext, 'extension' | 'extensionPath' | 'globalState' | 'subscriptions'>) {
	const output = vscode.window.createOutputChannel('ExcelMcp');
	context.subscriptions.push(output);
	const version: unknown = context.extension.packageJSON.version;
	if (typeof version !== 'string' || !/^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/.test(version)) {
		const message = 'ExcelMcp package version is invalid. Reinstall the extension.';
		output.appendLine(message);
		throw new Error(message);
	}
	const executable = join(context.extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe');

	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', {
			provideMcpServerDefinitions: async () => [
				new vscode.McpStdioServerDefinition('excel-mcp', executable, [], {}, version)
			],
			resolveMcpServerDefinition: async (server, token) => {
				const controller = new AbortController();
				const cancellation = token.onCancellationRequested(() => controller.abort());
				if (token.isCancellationRequested) {
					controller.abort();
				}
				try {
					await checkLaunchPrerequisites(executable, controller.signal);
					output.appendLine('Launch prerequisites verified. VS Code manages server startup and approvals.');
					return server;
				} catch (error) {
					if (controller.signal.aborted) {
						throw new vscode.CancellationError();
					}
					if (error instanceof LaunchSetupError) {
						output.appendLine(error.message);
						void showSetupError(error.message, output);
					}
					throw error;
				} finally {
					cancellation.dispose();
				}
			}
		})
	);
	output.appendLine(`Registered bundled MCP server version ${version}. Server logs: MCP: List Servers > excel-mcp > Show Output.`);

	const hasShownWelcome = context.globalState.get<boolean>('excelmcp.hasShownWelcome', false);
	if (!hasShownWelcome) {
		void showWelcomeMessage(output);
		try {
			await context.globalState.update('excelmcp.hasShownWelcome', true);
		} catch (error) {
			const detail = error instanceof Error ? error.message : String(error);
			output.appendLine(`Could not save the welcome preference. Getting-started help may appear again. ${detail}`);
		}
	}
}

async function showWelcomeMessage(output: vscode.OutputChannel) {
	try {
		const selection = await vscode.window.showInformationMessage(
			'ExcelMcp bundles real Excel automation. Open Copilot Chat and use MCP: List Servers to start excel-mcp and approve it when prompted.',
			'Getting Started'
		);
		if (selection === 'Getting Started') {
			if (!await vscode.env.openExternal(vscode.Uri.parse(userGuideUrl))) {
				output.appendLine(`Could not open the user guide. Visit ${userGuideUrl}`);
			}
		}
	} catch {
		output.appendLine(`Could not display getting-started help. Visit ${userGuideUrl}`);
	}
}

async function showSetupError(message: string, output: vscode.OutputChannel) {
	try {
		if (await vscode.window.showErrorMessage(message, 'Show Setup Output') === 'Show Setup Output') {
			output.show();
		}
	} catch {
		output.appendLine('Could not display the setup notification. See the setup error above.');
	}
}
