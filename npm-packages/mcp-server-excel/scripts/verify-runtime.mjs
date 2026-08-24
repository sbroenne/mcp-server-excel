import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';

const launcherPath = process.argv[2];
if (!launcherPath) {
  throw new Error('Usage: node verify-runtime.mjs <launcher-path>');
}

const child = spawn(process.execPath, [launcherPath], {
  stdio: ['pipe', 'pipe', 'pipe'],
  windowsHide: true
});

let stdout = '';
let stderr = '';
let initialized;
let tools;

const timeout = setTimeout(() => {
  child.kill();
  throw new Error(`Timed out waiting for MCP responses.\nstdout: ${stdout}\nstderr: ${stderr}`);
}, 30_000);

child.stderr.on('data', chunk => {
  stderr += chunk.toString('utf8');
});

child.stdout.on('data', chunk => {
  stdout += chunk.toString('utf8');
  const lines = stdout.split(/\r?\n/);
  stdout = lines.pop() ?? '';

  for (const line of lines.filter(value => value.trim())) {
    const message = JSON.parse(line);
    if (message.id === 1) {
      initialized = message;
      child.stdin.write(
        `${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' })}\n`
      );
      child.stdin.write(
        `${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} })}\n`
      );
    } else if (message.id === 2) {
      tools = message;
      child.stdin.end();
    }
  }
});

child.stdin.write(
  `${JSON.stringify({
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      protocolVersion: '2025-06-18',
      capabilities: {},
      clientInfo: {
        name: 'excel-mcp-npm-smoke-test',
        version: '1.0.0'
      }
    }
  })}\n`
);

const exitCode = await new Promise((resolve, reject) => {
  child.once('error', reject);
  child.once('close', resolve);
});
clearTimeout(timeout);

assert.equal(exitCode, 0, `Launcher exited with code ${exitCode}.\nstderr: ${stderr}`);
assert.equal(initialized?.result?.serverInfo?.name, 'excel-mcp');
assert.ok(Array.isArray(tools?.result?.tools));
assert.ok(tools.result.tools.length > 0);

console.log(
  `MCP handshake succeeded with ${tools.result.tools.length} tools ` +
    `(server ${initialized.result.serverInfo.version}).`
);
