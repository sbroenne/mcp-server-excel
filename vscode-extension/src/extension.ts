import * as vscode from 'vscode';
import * as path from 'path';
import { StatusBarMcp, showSessionsQuickPick } from './statusBarMcp';

/**
 * ExcelMcp VS Code Extension
 *
 * This extension provides MCP server definitions for the ExcelMcp MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft Excel
 * through native COM automation.
 */

export async function activate(context: vscode.ExtensionContext) {
	// Register command FIRST to ensure it's always available
	// This must happen before any async operations that could fail
	const showCmd = vscode.commands.registerCommand('excelMcp.showSessions', async () => {
		await showSessionsQuickPick();
	});
	context.subscriptions.push(showCmd);

	// Initialize status bar
	const status = new StatusBarMcp();
	status.show();
	context.subscriptions.push(status);

	// Ensure .NET runtime is available (optional - uses ms-dotnettools.vscode-dotnet-runtime if installed)
	try {
		await ensureDotNetRuntime();
	} catch (error) {
		// .NET runtime extension is optional - the MCP server requires .NET 8 on the system
		const errorMessage = error instanceof Error ? error.message : String(error);
		if (process.env.VSCODE_EXTENSION_TEST !== 'true') {
			vscode.window.showErrorMessage(
				`ExcelMcp: Failed to setup .NET environment: ${errorMessage}. ` +
				`The extension may not work correctly.`
			);
		}
	}

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', {
			provideMcpServerDefinitions: async () => {
				const extensionPath = context.extensionPath;
				const mcpServerPath = path.join(extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe');

				return [
					new vscode.McpStdioServerDefinition(
						'Excel MCP Server',
						mcpServerPath,
						[],
						{}
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

async function ensureDotNetRuntime(): Promise<void> {
	const dotnetExtension = vscode.extensions.getExtension('ms-dotnettools.vscode-dotnet-runtime');

	if (!dotnetExtension) {
		// Extension not installed - this is fine, user needs .NET 8 installed on system
		return;
	}

	if (!dotnetExtension.isActive) {
		await dotnetExtension.activate();
	}

	// Request .NET 8 runtime using the command-based API
	const requestingExtensionId = 'sbroenne.excel-mcp';

	await vscode.commands.executeCommand('dotnet.showAcquisitionLog');
	await vscode.commands.executeCommand<{ dotnetPath: string }>('dotnet.acquire', {
		version: '8.0',
		requestingExtensionId
	});
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
	// Cleanup handled by disposables in context.subscriptions
}
