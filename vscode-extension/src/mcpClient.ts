import * as vscode from 'vscode';

export type ExcelSessionInfo = {
  sessionId: string;
  filePath: string;
  canClose: boolean;
  isExcelVisible: boolean;
};

export type ListSessionsResult = {
  success: boolean;
  errorMessage?: string;
  sessions?: ExcelSessionInfo[];
};

export class McpClient {
  private readonly output = vscode.window.createOutputChannel('Excel MCP');

  /**
   * Resolve the fully-qualified tool name registered by the MCP server.
   * MCP tools from our provider appear in vscode.lm.tools with a prefix like `mcp_excel_mcp_ser_`.
   * We search for a tool matching the base name pattern.
   */
  private findToolName(baseName: string): string | undefined {
    const tools = vscode.lm.tools;
    const match = tools.find((t) => t.name.endsWith(`_${baseName}`) || t.name === baseName);
    return match?.name;
  }

  /**
   * Invoke a tool via vscode.lm.invokeTool and parse the JSON result.
   */
  private async invokeTool(baseName: string, params: Record<string, unknown>): Promise<any> {
    const toolName = this.findToolName(baseName);
    if (!toolName) {
      return { success: false, errorMessage: `Tool '${baseName}' not found. Ensure the MCP server is started.` };
    }
    const result = await vscode.lm.invokeTool(toolName, { toolInvocationToken: undefined, input: params });
    // result.content is Array<LanguageModelTextPart | ...>; we extract the first text part
    const textPart = result.content.find((p: any) => typeof p?.value === 'string') as { value: string } | undefined;
    if (!textPart) {
      return { success: false, errorMessage: 'No text content in tool result' };
    }
    try {
      return JSON.parse(textPart.value);
    } catch {
      return { success: false, errorMessage: 'Failed to parse tool result JSON' };
    }
  }

  async listSessions(): Promise<ListSessionsResult> {
    try {
      const result = await this.invokeTool('excel_file', { action: 'List' });
      return result as ListSessionsResult;
    } catch (err: any) {
      this.output.appendLine(`[listSessions] error: ${err?.message ?? String(err)}`);
      return { success: false, errorMessage: err?.message ?? String(err) };
    }
  }

  async closeSession(sessionId: string, save: boolean): Promise<{ success: boolean; errorMessage?: string }>{
    try {
      const result = await this.invokeTool('excel_file', { action: 'Close', sessionId, save });
      return result as { success: boolean; errorMessage?: string };
    } catch (err: any) {
      this.output.appendLine(`[closeSession] error: ${err?.message ?? String(err)}`);
      return { success: false, errorMessage: err?.message ?? String(err) };
    }
  }
}
