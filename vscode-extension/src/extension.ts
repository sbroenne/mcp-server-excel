import * as vscode from 'vscode';

/**
 * ExcelMcp VS Code Extension
 * 
 * This extension provides MCP server definitions for the ExcelMcp MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft Excel
 * through native COM automation.
 */

export async function activate(context: vscode.ExtensionContext) {
	console.log('ExcelMcp extension is now active');

	// Ensure .NET runtime is available and tool is installed
	try {
		await ensureDotNetAndTool();
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
				// Uses dotnet tool run with the globally installed tool
				// This works with .NET 8 runtime (auto-installed by .NET Install Tool extension)
				return [
					new vscode.McpStdioServerDefinition(
						'ExcelMcp - Excel Automation',
						'dotnet',
						['tool', 'run', 'mcp-excel'],
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

async function ensureDotNetAndTool(): Promise<void> {
	try {
		// Request .NET runtime acquisition via the .NET Install Tool extension
		const dotnetExtension = vscode.extensions.getExtension('ms-dotnettools.vscode-dotnet-runtime');
		
		if (!dotnetExtension) {
			throw new Error('.NET Install Tool extension not found. Please install ms-dotnettools.vscode-dotnet-runtime');
		}

		if (!dotnetExtension.isActive) {
			await dotnetExtension.activate();
		}

		// Request .NET 8 runtime
		const dotnetApi = dotnetExtension.exports;
		const dotnetPath = await dotnetApi.acquireDotNet('8.0', 'runtime');
		
		console.log(`ExcelMcp: .NET runtime available at ${dotnetPath.dotnetPath}`);

		// Check if the MCP server tool is installed
		const terminal = vscode.window.createTerminal({
			name: 'ExcelMcp Setup',
			hideFromUser: true
		});

		// Install the tool if not already installed
		// Using --ignore-failed-sources to handle offline scenarios gracefully
		terminal.sendText('dotnet tool install --global Sbroenne.ExcelMcp.McpServer --ignore-failed-sources || dotnet tool update --global Sbroenne.ExcelMcp.McpServer --ignore-failed-sources');
		terminal.dispose();

		console.log('ExcelMcp: MCP server tool installation/update initiated');
	} catch (error) {
		console.error('ExcelMcp: Error during .NET/tool setup:', error);
		throw error;
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
