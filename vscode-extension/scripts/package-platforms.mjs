import { cpSync, existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const extensionRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(resolve(extensionRoot, 'package.json'), 'utf8'));
if (process.platform !== 'win32' && process.platform !== 'darwin') {
  throw new Error(`Extension packaging is unsupported on ${process.platform}.`);
}
const targets = process.platform === 'win32'
  ? ['win32-x64']
  : ['win32-x64', 'darwin-arm64'];
const command = process.platform === 'win32' ? 'npx.cmd' : 'npx';
const binRoot = resolve(extensionRoot, 'bin');
const savedBinRoot = mkdtempSync(resolve(tmpdir(), 'excelmcp-vsix-bin-'));

cpSync(binRoot, savedBinRoot, { recursive: true });

try {
  for (const target of targets) {
    rmSync(binRoot, { recursive: true, force: true });
    mkdirSync(binRoot, { recursive: true });
    const sourceRuntime = resolve(savedBinRoot, target);
    if (!existsSync(sourceRuntime)) {
      throw new Error(`Missing prepared runtime directory: ${sourceRuntime}`);
    }
    cpSync(sourceRuntime, resolve(binRoot, target), { recursive: true });

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
      throw new Error(`VSIX packaging failed for ${target} with exit code ${result.status ?? 1}.`);
    }
  }
} finally {
  rmSync(binRoot, { recursive: true, force: true });
  cpSync(savedBinRoot, binRoot, { recursive: true });
  rmSync(savedBinRoot, { recursive: true, force: true });
}
