import * as vscode from 'vscode';

/**
 * ExcelMcp VS Code Extension
 * 
 * This extension provides MCP server definitions for the ExcelMcp MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft Excel
 * through native COM automation.
 */

export function activate(context: vscode.ExtensionContext) {
	console.log('ExcelMcp extension is now active');

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('excelmcp', {
			provideMcpServerDefinitions: async () => {
				// Return the MCP server definition for ExcelMcp
				// This uses the dnx command to run the NuGet-hosted MCP server
				return [
					new vscode.McpStdioServerDefinition(
						'ExcelMcp - Excel Automation',
						'dnx',
						['Sbroenne.ExcelMcp.McpServer', '--yes'],
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
		showWelcomeMessage(context);
		context.globalState.update('excelmcp.hasShownWelcome', true);
	}
}

function showWelcomeMessage(context: vscode.ExtensionContext) {
	const message = 'ExcelMcp extension activated! The Excel MCP server is now available for AI assistants.';
	const learnMore = 'Learn More';
	const dontShowAgain = "Don't Show Again";

	vscode.window.showInformationMessage(message, learnMore, dontShowAgain).then(selection => {
		if (selection === learnMore) {
			vscode.env.openExternal(vscode.Uri.parse('https://github.com/sbroenne/mcp-server-excel'));
		}
	});
}

export function deactivate() {
	console.log('ExcelMcp extension is now deactivated');
}
