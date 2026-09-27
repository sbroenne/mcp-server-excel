import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';

const launcherPath = process.argv[2];
if (!launcherPath) {
  throw new Error('Usage: node verify-runtime.mjs <launcher-path>');
}

const runtimePackage =
  process.platform === 'darwin' && process.arch === 'arm64'
    ? '@sbroenne/excelcli-darwin-arm64'
    : '@sbroenne/excelcli-win32-x64';
const runtime = createRequire(launcherPath).resolve(runtimePackage);
for (const args of [['--version'], ['--help'], ['session', '--help'], ['--invalid-npm-smoke-option']]) {
  const options = { encoding: 'utf8', timeout: 30_000, windowsHide: true };
  const direct = spawnSync(runtime, args, options);
  const launched = spawnSync(process.execPath, [launcherPath, ...args], options);
  assert.ifError(direct.error);
  assert.ifError(launched.error);
  assert.equal(direct.signal, null);
  assert.equal(launched.signal, null);
  if (args[0] === '--invalid-npm-smoke-option') {
    assert.notEqual(direct.status, 0, 'Invalid arguments must fail.');
  } else {
    assert.equal(direct.status, 0, direct.stderr);
    assert.ok(direct.stdout.trim(), 'CLI help/version must produce output.');
  }
  assert.equal(launched.status, direct.status, 'Launcher must preserve exit status.');
  assert.equal(launched.stdout, direct.stdout, 'Launcher must preserve stdout.');
  assert.equal(launched.stderr, direct.stderr, 'Launcher must preserve stderr.');
}

console.log('CLI npm launcher preserves help, version, subcommands, and failure exit codes.');
