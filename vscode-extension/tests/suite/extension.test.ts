/**
 * Extension Activation Integration Tests
 *
 * These tests verify that the Excel MCP extension activates correctly
 * and registers its components with VS Code.
 */

import * as assert from 'assert';
import * as vscode from 'vscode';

suite('Extension Activation', () => {

  test('extension should be present', () => {
    const extension = vscode.extensions.getExtension('sbroenne.excel-mcp');
    assert.ok(extension, 'Extension should be registered with VS Code');
  });

  test('extension should activate', async function() {
    this.timeout(30000);

    const extension = vscode.extensions.getExtension('sbroenne.excel-mcp');
    assert.ok(extension, 'Extension should exist');

    // Activate the extension if not already active
    if (!extension.isActive) {
      await extension.activate();
    }

    assert.strictEqual(extension.isActive, true, 'Extension should be active');
  });

  test('showSessions command should be registered', async function() {
    this.timeout(30000);

    const extension = vscode.extensions.getExtension('sbroenne.excel-mcp');
    assert.ok(extension, 'Extension should exist');

    // Activate if needed
    if (!extension.isActive) {
      await extension.activate();
    }

    // Give a moment for command registration
    await new Promise(resolve => setTimeout(resolve, 500));

    const commands = await vscode.commands.getCommands(true);
    const hasCommand = commands.includes('excelMcp.showSessions');

    if (!hasCommand) {
      console.log('Commands containing excel:',
        commands.filter(c => c.toLowerCase().includes('excel')));
    }

    assert.ok(hasCommand, 'excelMcp.showSessions command should be registered');
  });

  test('pollIntervalMs configuration should have correct default', () => {
    const config = vscode.workspace.getConfiguration('excelMcp');
    const pollInterval = config.get<number>('pollIntervalMs');

    // Default is 3000ms as specified in package.json
    assert.strictEqual(pollInterval, 3000, 'Default poll interval should be 3000ms');
  });

  test('pollIntervalMs configuration should be within valid range', () => {
    const config = vscode.workspace.getConfiguration('excelMcp');
    const pollInterval = config.get<number>('pollIntervalMs', 3000);

    assert.ok(pollInterval >= 1000, 'Poll interval should be at least 1000ms');
    assert.ok(pollInterval <= 60000, 'Poll interval should be at most 60000ms');
  });
});
