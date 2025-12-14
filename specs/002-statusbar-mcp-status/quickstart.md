# Quickstart: Status Bar MCP Server Monitor

## Prerequisites
- VS Code extension built and installed from `vscode-extension/`
- MCP Server available (extension starts it) and Excel closed when performing COM operations

## Usage
1. Open VS Code with the Excel MCP extension
2. Observe the status bar item labeled "Excel MCP":
   - Gray/disabled: server not running
   - Green/active: server running
3. Hover over the item to see server state and session count
4. Click the item to open a Quick Pick listing active sessions
5. Select a session and choose an action:
   - Close: terminates the session without saving
   - Save & Close: persists file changes then closes the session

## Notes
- Status updates occur within ~3–5 seconds
- If actions fail, a notification displays the server’s error message
- File paths may be truncated in labels; use details for full path
