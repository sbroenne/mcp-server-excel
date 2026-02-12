import * as vscode from 'vscode';
import * as path from 'path';

/**
 * ExcelMcp VS Code Extension
 *
 * This extension provides MCP server definitions for the ExcelMcp MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft Excel
 * through native COM automation.
 *
 * The extension bundles self-contained executables for both the MCP server and CLI -
 * no .NET SDK or runtime installation required.
 *
 * Agent Skills are registered via the chatSkills contribution point in package.json.
 */

export async function activate(context: vscode.ExtensionContext) {
	console.log('ExcelMcp extension is now active');

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', {
			provideMcpServerDefinitions: async () => {
				// Return the MCP server definition for ExcelMcp
				const extensionPath = context.extensionPath;
				const mcpServerPath = path.join(extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe');

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
