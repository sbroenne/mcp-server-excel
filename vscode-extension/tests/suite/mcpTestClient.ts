/**
 * MCP Test Client - Direct Communication with MCP Server
 *
 * This module spawns the MCP server as a subprocess and communicates
 * via the JSON-RPC protocol over stdio, bypassing vscode.lm.tools.
 *
 * This enables true integration testing of the MCP server without
 * requiring GitHub Copilot to be signed in.
 */

import * as vscode from 'vscode';
import * as path from 'path';
import { spawn, ChildProcess } from 'child_process';

interface JsonRpcRequest {
  jsonrpc: '2.0';
  id: number;
  method: string;
  params?: Record<string, unknown>;
}

interface JsonRpcResponse {
  jsonrpc: '2.0';
  id: number;
  result?: unknown;
  error?: { code: number; message: string; data?: unknown };
}

/**
 * A test client that communicates directly with the MCP server via stdio.
 */
export class McpTestClient {
  private process: ChildProcess | null = null;
  private requestId = 0;
  private pending = new Map<number, { resolve: (r: unknown) => void; reject: (e: Error) => void }>();
  private buffer = '';

  constructor(private readonly serverDllPath: string) {}

  /**
   * Start the MCP server subprocess using dotnet to run the DLL
   */
  async start(): Promise<void> {
    return new Promise((resolve, reject) => {
      console.log(`Starting MCP server: dotnet ${this.serverDllPath}`);

      // Run via dotnet command since the server is framework-dependent (not self-contained)
      this.process = spawn('dotnet', [this.serverDllPath], {
        stdio: ['pipe', 'pipe', 'pipe'],
        shell: false,
      });

      if (!this.process.stdout || !this.process.stdin) {
        reject(new Error('Failed to create stdio streams'));
        return;
      }

      this.process.stdout.setEncoding('utf8');
      this.process.stdout.on('data', (data: string) => this.onData(data));

      this.process.stderr?.setEncoding('utf8');
      this.process.stderr?.on('data', (data: string) => {
        console.log('[MCP stderr]', data);
      });

      this.process.on('error', (err) => {
        console.error('[MCP process error]', err);
        reject(err);
      });

      this.process.on('exit', (code) => {
        console.log(`[MCP process exited] code=${code}`);
      });

      // Give the server time to start
      setTimeout(() => resolve(), 500);
    });
  }

  /**
   * Stop the MCP server subprocess
   */
  stop(): void {
    if (this.process) {
      this.process.kill();
      this.process = null;
    }
  }

  /**
   * Send a JSON-RPC request and wait for response
   */
  async request(method: string, params?: Record<string, unknown>): Promise<unknown> {
    if (!this.process?.stdin) {
      throw new Error('MCP server not started');
    }

    const id = ++this.requestId;
    const request: JsonRpcRequest = {
      jsonrpc: '2.0',
      id,
      method,
      params,
    };

    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });

      // MCP uses newline-delimited JSON (not LSP Content-Length headers)
      const msg = JSON.stringify(request) + '\n';
      this.process!.stdin!.write(msg);

      // Timeout after 10 seconds
      setTimeout(() => {
        if (this.pending.has(id)) {
          this.pending.delete(id);
          reject(new Error(`Request ${id} timed out`));
        }
      }, 10000);
    });
  }

  /**
   * Initialize the MCP session
   */
  async initialize(): Promise<unknown> {
    return this.request('initialize', {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'test-client', version: '1.0.0' },
    });
  }

  /**
   * List available tools
   */
  async listTools(): Promise<unknown> {
    return this.request('tools/list', {});
  }

  /**
   * Call a tool
   */
  async callTool(name: string, args: Record<string, unknown>): Promise<unknown> {
    return this.request('tools/call', { name, arguments: args });
  }

  private onData(data: string): void {
    this.buffer += data;

    // MCP uses newline-delimited JSON
    const lines = this.buffer.split('\n');

    // Keep the last incomplete line in the buffer
    this.buffer = lines.pop() || '';

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;

      try {
        const response: JsonRpcResponse = JSON.parse(trimmed);
        this.handleResponse(response);
      } catch (err) {
        console.error('Failed to parse response:', trimmed, err);
      }
    }
  }

  private handleResponse(response: JsonRpcResponse): void {
    const pending = this.pending.get(response.id);
    if (!pending) {
      console.log('Unexpected response:', response);
      return;
    }

    this.pending.delete(response.id);

    if (response.error) {
      pending.reject(new Error(response.error.message));
    } else {
      pending.resolve(response.result);
    }
  }
}

/**
 * Get the path to the MCP server DLL (not .exe since it's framework-dependent)
 */
export function getMcpServerPath(extensionPath: string): string {
  return path.join(extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.dll');
}
