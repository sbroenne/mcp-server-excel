/**
 * MCP Server Direct Integration Tests
 *
 * These tests communicate directly with the MCP server via JSON-RPC,
 * bypassing vscode.lm.tools. This enables true end-to-end testing
 * without requiring GitHub Copilot to be signed in.
 *
 * Note: These tests require Excel to be installed on the system.
 */

import * as assert from 'assert';
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { McpTestClient, getMcpServerPath } from './mcpTestClient';

suite('MCP Server Direct Integration', function() {
  // Increase timeout for all tests in this suite
  this.timeout(30000);

  let client: McpTestClient;
  let serverPath: string;

  suiteSetup(async function() {
    // Get extension path
    const extension = vscode.extensions.getExtension('sbroenne.excel-mcp');
    if (!extension) {
      console.log('Extension not found - skipping MCP Server tests');
      this.skip();
      return;
    }

    serverPath = getMcpServerPath(extension.extensionPath);
    console.log('MCP server path:', serverPath);

    // Check if server executable exists
    if (!fs.existsSync(serverPath)) {
      console.log('MCP server executable not found - skipping tests');
      this.skip();
      return;
    }

    // Start MCP server
    client = new McpTestClient(serverPath);
    try {
      await client.start();
      console.log('MCP server started');
    } catch (err) {
      console.error('Failed to start MCP server:', err);
      this.skip();
    }
  });

  suiteTeardown(async () => {
    if (client) {
      client.stop();
      console.log('MCP server stopped');
    }
  });

  test('server should respond to initialize request', async function() {
    const result = await client.initialize() as any;

    assert.ok(result, 'Initialize should return a result');
    assert.ok(result.protocolVersion, 'Result should have protocolVersion');
    assert.ok(result.serverInfo, 'Result should have serverInfo');

    console.log('Server info:', result.serverInfo);
  });

  test('server should list available tools', async function() {
    const result = await client.listTools() as any;

    assert.ok(result, 'listTools should return a result');
    assert.ok(result.tools, 'Result should have tools array');
    assert.ok(Array.isArray(result.tools), 'tools should be an array');
    assert.ok(result.tools.length > 0, 'Should have at least one tool');

    // Log tool names for visibility
    const toolNames = result.tools.map((t: any) => t.name);
    console.log(`Found ${toolNames.length} tools:`, toolNames);

    // Verify expected tools exist
    const expectedTools = ['excel_file', 'excel_worksheet', 'excel_range', 'excel_table'];
    for (const expected of expectedTools) {
      const found = toolNames.some((name: string) => name === expected);
      assert.ok(found, `Tool '${expected}' should be available`);
    }
  });

  test('excel_file List action should return empty sessions', async function() {
    const result = await client.callTool('excel_file', { action: 'List' }) as any;

    assert.ok(result, 'callTool should return a result');
    assert.ok(result.content, 'Result should have content array');
    assert.ok(Array.isArray(result.content), 'content should be an array');

    // Parse the JSON response from the text content
    const textContent = result.content.find((c: any) => c.type === 'text');
    assert.ok(textContent, 'Should have text content');

    const data = JSON.parse(textContent.text);
    console.log('List sessions response:', data);

    assert.ok('success' in data, 'Response should have success field');
    // No active sessions expected in test environment
    if (data.success) {
      assert.ok(Array.isArray(data.sessions) || data.sessions === undefined,
        'sessions should be an array or undefined');
    }
  });

  test('excel_file Test action should check Excel availability', async function() {
    const result = await client.callTool('excel_file', { action: 'Test' }) as any;

    assert.ok(result, 'callTool should return a result');
    assert.ok(result.content, 'Result should have content array');

    const textContent = result.content.find((c: any) => c.type === 'text');
    assert.ok(textContent, 'Should have text content');

    const data = JSON.parse(textContent.text);
    console.log('Test Excel response:', data);

    assert.ok('success' in data, 'Response should have success field');
    // This will be true if Excel is installed, false if not
    console.log('Excel available:', data.success);
  });

  test('tool schema should have proper input definitions', async function() {
    const result = await client.listTools() as any;

    // Find excel_file tool
    const excelFileTool = result.tools.find((t: any) => t.name === 'excel_file');
    assert.ok(excelFileTool, 'excel_file tool should exist');

    // Verify it has input schema
    assert.ok(excelFileTool.inputSchema, 'Tool should have inputSchema');
    assert.strictEqual(excelFileTool.inputSchema.type, 'object', 'Schema type should be object');

    // Verify required properties
    const properties = excelFileTool.inputSchema.properties;
    assert.ok(properties, 'Schema should have properties');
    assert.ok(properties.action, 'Should have action property');

    console.log('excel_file schema:', JSON.stringify(excelFileTool.inputSchema, null, 2));
  });
});

/**
 * Session Lifecycle Tests
 *
 * These tests verify the complete session handling workflow:
 * - Opening Excel files creates sessions
 * - Listing sessions returns the correct data
 * - Closing sessions properly cleans up
 *
 * These are the tests that verify the showSessionsQuickPick functionality
 * works correctly with the real MCP server.
 */
suite('MCP Session Lifecycle', function() {
  this.timeout(60000); // Sessions require Excel which is slow

  let client: McpTestClient;
  let serverPath: string;
  let testDir: string;
  let testFilePath: string;
  let excelAvailable = false;

  suiteSetup(async function() {
    // Get extension path
    const extension = vscode.extensions.getExtension('sbroenne.excel-mcp');
    if (!extension) {
      console.log('Extension not found - skipping session tests');
      this.skip();
      return;
    }

    serverPath = getMcpServerPath(extension.extensionPath);

    if (!fs.existsSync(serverPath)) {
      console.log('MCP server not found - skipping session tests');
      this.skip();
      return;
    }

    // Create test directory and file
    testDir = path.join(os.tmpdir(), `excel-mcp-session-test-${Date.now()}`);
    fs.mkdirSync(testDir, { recursive: true });
    testFilePath = path.join(testDir, 'test-session.xlsx');

    // Start MCP server
    client = new McpTestClient(serverPath);
    try {
      await client.start();
      await client.initialize();
      console.log('MCP server started for session tests');
    } catch (err) {
      console.error('Failed to start MCP server:', err);
      this.skip();
      return;
    }

    // Check if Excel is available by trying to create and open a file
    try {
      // First create the file
      const createResult = await client.callTool('excel_file', {
        action: 'CreateEmpty',
        excelPath: testFilePath,
        showExcel: false
      }) as any;

      const createContent = createResult.content?.find((c: any) => c.type === 'text');
      if (createContent) {
        const createData = JSON.parse(createContent.text);
        if (createData.success) {
          // Now open it as a session to verify sessions work
          const openResult = await client.callTool('excel_file', {
            action: 'Open',
            excelPath: testFilePath,
            showExcel: false
          }) as any;

          const openContent = openResult.content?.find((c: any) => c.type === 'text');
          if (openContent) {
            const openData = JSON.parse(openContent.text);
            excelAvailable = openData.success === true && !!openData.sessionId;

            if (excelAvailable) {
              // Clean up the session we just created
              await client.callTool('excel_file', {
                action: 'Close',
                sessionId: openData.sessionId,
                save: false
              });
            }
          }
        }
      }
      console.log('Excel available:', excelAvailable);
    } catch (err) {
      console.log('Excel not available:', err);
      excelAvailable = false;
    }
  });

  suiteTeardown(async () => {
    if (client) {
      client.stop();
    }
    // Clean up test directory
    if (testDir && fs.existsSync(testDir)) {
      try {
        fs.rmSync(testDir, { recursive: true, force: true });
      } catch {
        // Ignore cleanup errors
      }
    }
  });

  test('opening a file should create a session', async function() {
    if (!excelAvailable) {
      this.skip();
      return;
    }

    // First create the file
    await client.callTool('excel_file', {
      action: 'CreateEmpty',
      excelPath: testFilePath,
      showExcel: false
    });

    // Now open it as a session
    const result = await client.callTool('excel_file', {
      action: 'Open',
      excelPath: testFilePath,
      showExcel: false
    }) as any;

    const textContent = result.content?.find((c: any) => c.type === 'text');
    assert.ok(textContent, 'Should have text response');

    const data = JSON.parse(textContent.text);
    console.log('Open response:', data);

    assert.strictEqual(data.success, true, 'Open should succeed');
    assert.ok(data.sessionId, 'Should return a sessionId');
    assert.ok(data.filePath, 'Should return the filePath');

    // Store for later tests
    const sessionId = data.sessionId;

    // Clean up
    await client.callTool('excel_file', {
      action: 'Close',
      sessionId: sessionId,
      save: false
    });
  });

  test('listing sessions should return active sessions', async function() {
    if (!excelAvailable) {
      this.skip();
      return;
    }

    // Create the file first
    await client.callTool('excel_file', {
      action: 'CreateEmpty',
      excelPath: testFilePath,
      showExcel: false
    });

    // Open as session
    const createResult = await client.callTool('excel_file', {
      action: 'Open',
      excelPath: testFilePath,
      showExcel: false
    }) as any;

    const createData = JSON.parse(createResult.content?.find((c: any) => c.type === 'text')?.text || '{}');
    assert.strictEqual(createData.success, true, 'Should open session');
    const sessionId = createData.sessionId;

    try {
      // List sessions
      const listResult = await client.callTool('excel_file', {
        action: 'List'
      }) as any;

      const listData = JSON.parse(listResult.content?.find((c: any) => c.type === 'text')?.text || '{}');
      console.log('List sessions:', listData);

      assert.strictEqual(listData.success, true, 'List should succeed');
      assert.ok(Array.isArray(listData.sessions), 'sessions should be an array');
      assert.ok(listData.sessions.length >= 1, 'Should have at least one session');

      // Find our session
      const ourSession = listData.sessions.find((s: any) => s.sessionId === sessionId);
      assert.ok(ourSession, 'Our session should be in the list');
      assert.ok(ourSession.filePath, 'Session should have filePath');
      assert.strictEqual(typeof ourSession.canClose, 'boolean', 'Session should have canClose');
      assert.strictEqual(typeof ourSession.isExcelVisible, 'boolean', 'Session should have isExcelVisible');

      console.log('Found session:', ourSession);
    } finally {
      // Clean up
      await client.callTool('excel_file', {
        action: 'Close',
        sessionId: sessionId,
        save: false
      });
    }
  });

  test('closing a session should remove it from the list', async function() {
    if (!excelAvailable) {
      this.skip();
      return;
    }

    // Create the file first
    await client.callTool('excel_file', {
      action: 'CreateEmpty',
      excelPath: testFilePath,
      showExcel: false
    });

    // Open as session
    const createResult = await client.callTool('excel_file', {
      action: 'Open',
      excelPath: testFilePath,
      showExcel: false
    }) as any;

    const createData = JSON.parse(createResult.content?.find((c: any) => c.type === 'text')?.text || '{}');
    assert.strictEqual(createData.success, true, 'Should open session');
    const sessionId = createData.sessionId;

    // Verify session exists
    let listResult = await client.callTool('excel_file', { action: 'List' }) as any;
    let listData = JSON.parse(listResult.content?.find((c: any) => c.type === 'text')?.text || '{}');
    const sessionsBefore = listData.sessions?.length || 0;
    assert.ok(listData.sessions?.some((s: any) => s.sessionId === sessionId), 'Session should exist before close');

    // Close the session
    const closeResult = await client.callTool('excel_file', {
      action: 'Close',
      sessionId: sessionId,
      save: false
    }) as any;

    const closeData = JSON.parse(closeResult.content?.find((c: any) => c.type === 'text')?.text || '{}');
    console.log('Close response:', closeData);
    assert.strictEqual(closeData.success, true, 'Close should succeed');

    // Verify session is gone
    listResult = await client.callTool('excel_file', { action: 'List' }) as any;
    listData = JSON.parse(listResult.content?.find((c: any) => c.type === 'text')?.text || '{}');
    const sessionsAfter = listData.sessions?.length || 0;

    assert.ok(sessionsAfter < sessionsBefore, 'Session count should decrease after close');
    assert.ok(!listData.sessions?.some((s: any) => s.sessionId === sessionId), 'Session should not exist after close');

    console.log(`Sessions: ${sessionsBefore} before, ${sessionsAfter} after close`);
  });

  test('session info should include all required fields for QuickPick', async function() {
    if (!excelAvailable) {
      this.skip();
      return;
    }

    // Create the file first
    await client.callTool('excel_file', {
      action: 'CreateEmpty',
      excelPath: testFilePath,
      showExcel: false
    });

    // Open as session
    const createResult = await client.callTool('excel_file', {
      action: 'Open',
      excelPath: testFilePath,
      showExcel: false
    }) as any;

    const createData = JSON.parse(createResult.content?.find((c: any) => c.type === 'text')?.text || '{}');
    assert.strictEqual(createData.success, true, 'Should open session');
    const sessionId = createData.sessionId;

    try {
      // List sessions
      const listResult = await client.callTool('excel_file', { action: 'List' }) as any;
      const listData = JSON.parse(listResult.content?.find((c: any) => c.type === 'text')?.text || '{}');

      const session = listData.sessions?.find((s: any) => s.sessionId === sessionId);
      assert.ok(session, 'Should find the session');

      // Verify all fields needed by showSessionsQuickPick
      console.log('Session fields for QuickPick:', {
        sessionId: session.sessionId,
        filePath: session.filePath,
        canClose: session.canClose,
        isExcelVisible: session.isExcelVisible
      });

      // These are the exact fields used by showSessionsQuickPick
      assert.ok(typeof session.sessionId === 'string', 'sessionId should be string');
      assert.ok(typeof session.filePath === 'string', 'filePath should be string');
      assert.ok(typeof session.canClose === 'boolean', 'canClose should be boolean');
      assert.ok(typeof session.isExcelVisible === 'boolean', 'isExcelVisible should be boolean');

      // Verify filePath can be used with basename (showSessionsQuickPick uses this)
      const basename = session.filePath.split(/[/\\]/).pop();
      assert.ok(basename, 'Should be able to extract basename from filePath');
      assert.ok(basename.endsWith('.xlsx'), 'basename should end with .xlsx');

    } finally {
      await client.callTool('excel_file', {
        action: 'Close',
        sessionId: sessionId,
        save: false
      });
    }
  });
});
