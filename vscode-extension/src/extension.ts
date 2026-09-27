import * as vscode from 'vscode';
import * as path from 'path';

/**
 * ExcelMcp VS Code Extension
 *
 * This extension provides MCP server definitions for the ExcelMcp MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft Excel
 * through the platform Excel automation backend.
 *
 * The extension bundles a self-contained MCP server executable, so no .NET SDK
 * or runtime installation is required.
 *
 * Agent Skills are registered via the chatSkills contribution point in package.json.
 */

export async function activate(context: vscode.ExtensionContext) {
	console.log('ExcelMcp extension is now active');

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', {
			provideMcpServerDefinitions: async () => {
				const runtime = resolveBundledRuntime(process.platform, process.arch);
				const mcpServerPath = path.join(context.extensionPath, 'bin', runtime.directory, runtime.executable);

				return [
					new vscode.McpStdioServerDefinition(
						'excel-mcp',
						mcpServerPath,
						[],
						{
							// Optional environment variables can be added here if needed
						}
					)
				];
			}
		})
	);

	// Show welcome message on first activation
	const hasShownWelcome = context.globalState.get<boolean>('excelmcp.hasShownWelcome', false);
	if (!hasShownWelcome) {
		showWelcomeMessage();
		context.globalState.update('excelmcp.hasShownWelcome', true);
	}
}

export function resolveBundledRuntime(
	platform: NodeJS.Platform,
	architecture: string
): { directory: string; executable: string } {
	if (platform === 'win32' && architecture === 'x64') {
		return {
			directory: 'win32-x64',
			executable: 'Sbroenne.ExcelMcp.McpServer.exe'
		};
	}

	if (platform === 'darwin' && architecture === 'arm64') {
		return {
			directory: 'darwin-arm64',
			executable: 'Sbroenne.ExcelMcp.McpServer'
		};
	}

	throw new Error(
		`Excel MCP Server does not include a runtime for ${platform}-${architecture}. ` +
		'Supported platforms are Windows x64 and Apple Silicon macOS. Intel macOS is not supported.'
	);
}

function showWelcomeMessage() {
	const message = 'ExcelMcp extension activated! The Excel MCP server is now available for AI assistants.';
	const learnMore = 'Learn More';

	vscode.window.showInformationMessage(message, learnMore).then(selection => {
		if (selection === learnMore) {
			vscode.env.openExternal(vscode.Uri.parse('https://github.com/sbroenne/mcp-server-excel'));
		}
	});
}

export function deactivate() {
	console.log('ExcelMcp extension is now deactivated');
}
