import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';

const launcherPath = process.argv[2];
if (!launcherPath) {
  throw new Error('Usage: node verify-runtime.mjs <launcher-path>');
}

const runtime = createRequire(launcherPath).resolve(`@sbroenne/excelcli-win32-${process.arch}`);

function withoutDuration(stdout) {
  const { durationMs } = JSON.parse(stdout);
  assert.ok(Number.isSafeInteger(durationMs) && durationMs >= 0,
    'Operation failures must report a non-negative durationMs.');
  return stdout.replace(/("durationMs"\s*:\s*)\d+/, '$1<duration>');
}

for (const args of [['--version'], ['--help'], ['session', '--help'], ['--invalid-npm-smoke-option']]) {
  const invalidArguments = args[0] === '--invalid-npm-smoke-option';
  const options = { encoding: 'utf8', timeout: 30_000, windowsHide: true };
  const direct = spawnSync(runtime, args, options);
  const launched = spawnSync(process.execPath, [launcherPath, ...args], options);
  assert.ifError(direct.error);
  assert.ifError(launched.error);
  assert.equal(direct.signal, null);
  assert.equal(launched.signal, null);
  if (invalidArguments) {
    assert.notEqual(direct.status, 0, 'Invalid arguments must fail.');
  } else {
    assert.equal(direct.status, 0, direct.stderr);
    assert.ok(direct.stdout.trim(), 'CLI help/version must produce output.');
  }
  assert.equal(launched.status, direct.status, 'Launcher must preserve exit status.');
  assert.equal(
    invalidArguments ? withoutDuration(launched.stdout) : launched.stdout,
    invalidArguments ? withoutDuration(direct.stdout) : direct.stdout,
    'Launcher must preserve stdout apart from per-execution durationMs.');
  assert.equal(launched.stderr, direct.stderr, 'Launcher must preserve stderr.');
}

console.log('CLI npm launcher preserves help, version, subcommands, and failure exit codes.');
