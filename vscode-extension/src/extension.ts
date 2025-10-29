import * as vscode from 'vscode';
import * as path from 'path';

/**
 * ExcelMcp VS Code Extension
 *
 * This extension provides MCP server definitions for the ExcelMcp MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft Excel
 * through native COM automation.
 */

export async function activate(context: vscode.ExtensionContext) {
	console.log('ExcelMcp extension is now active');

	// Ensure .NET runtime is available (still needed for the bundled executable)
	try {
		await ensureDotNetRuntime();
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		vscode.window.showErrorMessage(
			`ExcelMcp: Failed to setup .NET environment: ${errorMessage}. ` +
			`The extension may not work correctly.`
		);
	}

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('excelmcp', {
			provideMcpServerDefinitions: async () => {
				// Return the MCP server definition for ExcelMcp
				// Use the bundled executable path
				const extensionPath = context.extensionPath;
				const mcpServerPath = path.join(extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe');

				return [
					new vscode.McpStdioServerDefinition(
						'ExcelMcp - Excel Automation',
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

async function ensureDotNetRuntime(): Promise<void> {
	try {
		// Request .NET runtime acquisition via the .NET Install Tool extension
		const dotnetExtension = vscode.extensions.getExtension('ms-dotnettools.vscode-dotnet-runtime');

		if (!dotnetExtension) {
			throw new Error('.NET Install Tool extension not found. Please install ms-dotnettools.vscode-dotnet-runtime');
		}

		if (!dotnetExtension.isActive) {
			await dotnetExtension.activate();
		}

		// Request .NET 8 runtime using the command-based API
		// The extension uses commands, not direct exports
		const requestingExtensionId = 'sbroenne.excelmcp';

		await vscode.commands.executeCommand('dotnet.showAcquisitionLog');
		const result = await vscode.commands.executeCommand<{ dotnetPath: string }>('dotnet.acquire', {
			version: '8.0',
			requestingExtensionId
		});

		if (result?.dotnetPath) {
			console.log(`ExcelMcp: .NET runtime available at ${result.dotnetPath}`);
		}

		console.log('ExcelMcp: .NET runtime setup completed (MCP server is bundled with extension)');
	} catch (error) {
		console.error('ExcelMcp: Error during .NET runtime setup:', error);
		throw error;
	}
}function showWelcomeMessage() {
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
