import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const extensionRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(resolve(extensionRoot, 'package.json'), 'utf8'));
const targets = ['win32-x64', 'darwin-arm64'];
const command = process.platform === 'win32' ? 'npx.cmd' : 'npx';

for (const target of targets) {
  const output = `excelmcp-${manifest.version}-${target}.vsix`;
  const result = spawnSync(
    command,
    ['vsce', 'package', '--target', target, '--out', output],
    { cwd: extensionRoot, encoding: 'utf8', stdio: 'inherit' }
  );

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}
