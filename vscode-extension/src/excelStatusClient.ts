import * as vscode from 'vscode';
import * as net from 'net';

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

/**
 * Lightweight client for communicating with the ExcelMcp server's status endpoint.
 *
 * This uses the per-window named pipe (status channel) to avoid VS Code's tool
 * invocation confirmation UI.
 */
export class ExcelStatusClient {
  private readonly output = vscode.window.createOutputChannel('Excel MCP');
  private readonly statusPipeName = getStatusPipeName();

  private async invokeStatusPipe(request: Record<string, unknown>): Promise<any> {
    const pipePath = `\\\\.\\pipe\\${this.statusPipeName}`;

    const attemptOnce = () => new Promise<any>((resolve) => {
      const socket = net.connect(pipePath);
      socket.setEncoding('utf8');

      let resolved = false;
      let buffer = '';

      const finish = (value: any) => {
        if (resolved) return;
        resolved = true;
        try { socket.end(); } catch { /* ignore */ }
        resolve(value);
      };

      const timeout = setTimeout(() => {
        try { socket.destroy(); } catch { /* ignore */ }
        finish({ success: false, errorMessage: 'Status pipe request timed out' });
      }, 1500);

      socket.on('connect', () => {
        try {
          socket.write(JSON.stringify(request) + '\n');
        } catch (err) {
          clearTimeout(timeout);
          finish({ success: false, errorMessage: (err as any)?.message ?? String(err) });
        }
      });

      socket.on('data', (data: string) => {
        buffer += data;
        const ix = buffer.indexOf('\n');
        if (ix < 0) return;

        const line = buffer.substring(0, ix).trim();
        clearTimeout(timeout);

        if (!line) {
          finish({ success: false, errorMessage: 'Empty status pipe response' });
          return;
        }

        try {
          finish(JSON.parse(line));
        } catch {
          finish({ success: false, errorMessage: 'Failed to parse status pipe JSON response' });
        }
      });

      socket.on('error', (err: NodeJS.ErrnoException) => {
        clearTimeout(timeout);
        finish({ success: false, errorMessage: err?.message ?? String(err) });
      });

      socket.on('close', () => {
        clearTimeout(timeout);
        if (!resolved) {
          finish({ success: false, errorMessage: 'Status pipe closed before response' });
        }
      });
    });

    // The MCP server process and pipe may not be ready yet; retry a couple times.
    let last: any;
    for (let i = 0; i < 3; i++) {
      last = await attemptOnce();
      if (last?.success === true || (last?.success === false && last?.errorMessage && !String(last.errorMessage).includes('ENOENT'))) {
        return last;
      }
      await new Promise((r) => setTimeout(r, 150));
    }
    return last;
  }

  async listSessions(): Promise<ListSessionsResult> {
    try {
      const result = await this.invokeStatusPipe({ action: 'list' });
      return result as ListSessionsResult;
    } catch (err: any) {
      this.output.appendLine(`[listSessions] error: ${err?.message ?? String(err)}`);
      return { success: false, errorMessage: err?.message ?? String(err) };
    }
  }

  async closeSession(sessionId: string, save: boolean): Promise<{ success: boolean; errorMessage?: string }>{
    try {
      const result = await this.invokeStatusPipe({ action: 'close', sessionId, save });
      return result as { success: boolean; errorMessage?: string };
    } catch (err: any) {
      this.output.appendLine(`[closeSession] error: ${err?.message ?? String(err)}`);
      return { success: false, errorMessage: err?.message ?? String(err) };
    }
  }
}

function getStatusPipeName(): string {
  const raw = vscode.env.sessionId || 'unknown';
  const safe = raw.replace(/[^a-zA-Z0-9._-]/g, '-');
  return `excelmcp-status-${safe}`;
}
