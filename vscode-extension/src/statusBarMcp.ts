import * as vscode from 'vscode';
import { ExcelStatusClient } from './excelStatusClient';
import { Poller } from './utils/polling';

export class StatusBarMcp {
  private readonly item: vscode.StatusBarItem;
  private readonly client: ExcelStatusClient;
  private poller?: Poller<any>;
  private _isVisible = false;

  /** Whether the status bar is currently visible (for testing) */
  get isVisible(): boolean {
    return this._isVisible;
  }

  constructor() {
    this.item = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 100);
    this.item.text = '$(graph) Excel MCP';
    this.item.tooltip = 'Excel MCP';
    this.item.command = 'excelMcp.showSessions';
    this.client = new ExcelStatusClient();
  }

  show() {
    // Start polling but don't show the item yet
    const intervalMs = vscode.workspace.getConfiguration('excelMcp').get<number>('pollIntervalMs', 3000);
    const backoffMs = Math.max(1000, Math.floor(intervalMs * 1.6));
    this.poller = new Poller(async () => await this.client.listSessions(), (r) => this.updateFromResult(r), (e) => this.updateDisconnected(e), intervalMs, backoffMs);
    this.poller.start();
  }

  dispose() {
    if (this.poller) this.poller.stop();
    this.item.dispose();
  }

  private updateFromResult(result: { success: boolean; errorMessage?: string; sessions?: any[] }) {
    if (!result.success) {
      this._isVisible = false;
      this.item.hide();
      return;
    }
    const count = (result.sessions ?? []).length;
    this.item.text = `$(check) Excel MCP (${count})`;
    this.item.tooltip = `Excel MCP: Connected\nActive sessions: ${count}`;
    this._isVisible = true;
    this.item.show();
  }

  private updateDisconnected(_err?: any) {
    this._isVisible = false;
    this.item.hide();
  }
}

export async function showSessionsQuickPick(client = new ExcelStatusClient()) {
  const res = await client.listSessions();
  if (!res.success) {
    void vscode.window.showErrorMessage(res.errorMessage ?? 'Failed to get sessions');
    return;
  }
  const sessions = res.sessions ?? [];
  if (sessions.length === 0) {
    void vscode.window.showInformationMessage('No active Excel MCP sessions');
    return;
  }
  const items = sessions.map((s) => ({
    label: `${basename(s.filePath)}`,
    description: s.filePath,
    detail: s.isExcelVisible ? 'Excel is visible' : 'Excel is hidden',
    session: s,
  }));
  const picked = await vscode.window.showQuickPick(items, { placeHolder: 'Select a session' });
  if (!picked) return;
  const action = await vscode.window.showQuickPick(['Close', 'Save & Close'], { placeHolder: 'Choose action' });
  if (!action) return;
  const save = action === 'Save & Close';
  const result = await client.closeSession(picked.session.sessionId, save);
  if (!result.success) {
    void vscode.window.showErrorMessage(result.errorMessage ?? 'Failed to close session');
    return;
  }
  void vscode.window.showInformationMessage('Session closed');
}

function basename(p: string) {
  const ix = Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/'));
  return ix >= 0 ? p.substring(ix + 1) : p;
}
