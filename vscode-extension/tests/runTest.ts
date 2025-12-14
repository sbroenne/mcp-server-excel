/**
 * VS Code Extension Integration Test Runner
 *
 * This script downloads VS Code, launches it with the extension installed,
 * and runs the integration test suite inside the real VS Code environment.
 */

import * as path from 'path';
import * as fs from 'fs';
import { execSync } from 'child_process';
import { runTests } from '@vscode/test-electron';

/**
 * Ensure the MCP server is built before running tests
 */
function ensureMcpServerBuilt(): void {
  const extensionRoot = path.resolve(__dirname, '../../');
  const mcpServerExe = path.join(extensionRoot, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe');

  // Check if MCP server executable exists
  if (!fs.existsSync(mcpServerExe)) {
    console.log('MCP server not found, building...');
    try {
      execSync('npm run build:mcp-server', {
        cwd: extensionRoot,
        stdio: 'inherit'
      });
      console.log('MCP server built successfully');
    } catch (error) {
      console.error('Failed to build MCP server:', error);
      throw error;
    }
  } else {
    console.log('MCP server found at:', mcpServerExe);
  }
}

/**
 * Setup VS Code test user data with MCP enabled
 */
function setupTestUserData(): string {
  const extensionRoot = path.resolve(__dirname, '../../');
  const userDataPath = path.join(extensionRoot, '.vscode-test', 'user-data');
  const userSettingsDir = path.join(userDataPath, 'User');
  const settingsPath = path.join(userSettingsDir, 'settings.json');

  // Ensure directory exists
  if (!fs.existsSync(userSettingsDir)) {
    fs.mkdirSync(userSettingsDir, { recursive: true });
  }

  // Create or update settings to enable MCP
  const settings = {
    'chat.mcp.enabled': true,
    // Disable telemetry and other noise during tests
    'telemetry.telemetryLevel': 'off',
    'update.mode': 'none',
    'extensions.autoUpdate': false,
  };

  fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2));
  console.log('Test settings configured at:', settingsPath);

  return userDataPath;
}

async function main() {
  try {
    // Ensure MCP server is built
    ensureMcpServerBuilt();

    // Setup test user data
    const userDataPath = setupTestUserData();

    // The folder containing the Extension Manifest package.json
    const extensionDevelopmentPath = path.resolve(__dirname, '../../');

    // The path to the extension test script
    const extensionTestsPath = path.resolve(__dirname, './suite/index');

    // Download VS Code, unzip it and run the integration tests
    await runTests({
      extensionDevelopmentPath,
      extensionTestsPath,
      launchArgs: [
        '--disable-gpu',        // Avoid GPU issues in CI
        `--user-data-dir=${userDataPath}`, // Use our configured user data
      ],
      // Set environment variable to indicate test mode
      extensionTestsEnv: {
        VSCODE_EXTENSION_TEST: 'true',
        // Signal that MCP server should be available
        MCP_SERVER_AVAILABLE: 'true',
      },
    });
  } catch (err) {
    console.error('Failed to run tests:', err);
    process.exit(1);
  }
}

main();
