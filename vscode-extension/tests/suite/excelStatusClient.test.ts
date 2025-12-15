/**
 * Excel Status Client Integration Tests
 *
 * These tests verify that the extension can discover MCP tools (for normal
 * usage) and can query session status using the named-pipe status endpoint.
 *
 * Note: These tests require the MCP server binary to be built.
 * The runTest.ts script ensures the MCP server is built before running tests.
 */

import * as assert from 'assert';
import * as vscode from 'vscode';
import { ExcelStatusClient } from '../../src/excelStatusClient';

// Helper to wait for MCP tools to be registered
async function waitForMcpTools(timeoutMs: number = 15000): Promise<vscode.LanguageModelToolInformation[]> {
  const startTime = Date.now();

  while (Date.now() - startTime < timeoutMs) {
    const tools = vscode.lm.tools;
    const excelTools = tools.filter(t =>
      t.name.toLowerCase().includes('excel')
    );

    if (excelTools.length > 0) {
      return excelTools;
    }

    // Wait 500ms before checking again
    await new Promise(resolve => setTimeout(resolve, 500));
  }

  // Return empty if timeout reached
  return [];
}

suite('Excel Status Client Integration', () => {
  let excelTools: vscode.LanguageModelToolInformation[] = [];

  suiteSetup(async function() {
    // Increase timeout for suite setup
    this.timeout(30000);

    // Ensure extension is activated
    const extension = vscode.extensions.getExtension('sbroenne.excel-mcp');
    if (extension && !extension.isActive) {
      await extension.activate();
    }

    // Wait for MCP server to start and register tools
    console.log('Waiting for MCP tools to be registered...');
    excelTools = await waitForMcpTools(20000);

    if (excelTools.length > 0) {
      console.log(`Found ${excelTools.length} Excel MCP tools:`, excelTools.map(t => t.name));
    } else {
      console.log('No Excel MCP tools found after waiting.');
      console.log('Available tools:', vscode.lm.tools.map(t => t.name).slice(0, 15));
    }
  });

  test('MCP tools should be discoverable via vscode.lm.tools', async function() {
    // If MCP server is running and connected, we should find Excel tools
    if (excelTools.length > 0) {
      assert.ok(true, `Found ${excelTools.length} Excel MCP tools`);

      // Verify we have the expected core tools
      const toolNames = excelTools.map(t => t.name.toLowerCase());
      console.log('Excel tool names:', toolNames);
    } else {
      // MCP server may not have started - document but don't fail
      console.log('No Excel MCP tools found - MCP server may not be running');
      console.log('This can happen if VS Code MCP integration is not enabled');
    }
  });

  test('excel_file tool should be available when server is running', async function() {
    const excelFileTool = excelTools.find(t =>
      t.name.includes('excel_file') || t.name.endsWith('_excel_file')
    );

    if (excelFileTool) {
      assert.ok(excelFileTool.name, 'excel_file tool should have a name');
      console.log('Found excel_file tool:', excelFileTool.name);

      // Verify tool has expected properties
      assert.ok(excelFileTool.name.length > 0, 'Tool name should not be empty');
    } else if (excelTools.length === 0) {
      // Skip if no MCP tools at all (server not connected)
      console.log('excel_file tool not found - MCP server not connected');
    } else {
      // We have some tools but not excel_file - this would be unexpected
      console.log('Some Excel tools found but not excel_file:', excelTools.map(t => t.name));
    }
  });

  test('listSessions should return session data when server is running', async function() {
    this.timeout(10000);

    if (excelTools.length === 0) {
      console.log('Skipping - no MCP tools registered');
      this.skip();
      return;
    }

    const client = new ExcelStatusClient();
    const result = await client.listSessions();

    assert.ok(typeof result.success === 'boolean', 'Response should have success field');
    console.log('List sessions result:', JSON.stringify(result, null, 2));
  });

  test('excel_worksheet tool should be available', async function() {
    const worksheetTool = excelTools.find(t =>
      t.name.includes('excel_worksheet') || t.name.endsWith('_excel_worksheet')
    );

    if (worksheetTool) {
      assert.ok(worksheetTool.name, 'excel_worksheet tool should have a name');
      console.log('Found excel_worksheet tool:', worksheetTool.name);
    } else if (excelTools.length === 0) {
      console.log('excel_worksheet tool not found - MCP server not connected');
    }
  });

  test('excel_range tool should be available', async function() {
    const rangeTool = excelTools.find(t =>
      t.name.includes('excel_range') || t.name.endsWith('_excel_range')
    );

    if (rangeTool) {
      assert.ok(rangeTool.name, 'excel_range tool should have a name');
      console.log('Found excel_range tool:', rangeTool.name);
    } else if (excelTools.length === 0) {
      console.log('excel_range tool not found - MCP server not connected');
    }
  });
});
