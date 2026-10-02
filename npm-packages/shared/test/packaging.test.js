import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const repoRoot = fileURLToPath(new URL('../../../', import.meta.url));
const version = '9.8.7-test.1';

function executableFixture(arch) {
  const payload = Buffer.alloc(128);
  payload.write('MZ');
  payload.writeUInt32LE(64, 0x3c);
  payload.write('PE\0\0', 64);
  payload.writeUInt16LE(arch === 'arm64' ? 0xaa64 : 0x8664, 68);
  return payload;
}

const products = [
  ['McpServer', 'mcp-server-excel', 'mcp-excel'],
  ['Cli', 'excelcli', 'excelcli']
];

for (const [component, packageName, commandName, arch, platform] of products.flatMap(product =>
  [['x64', 'win32'], ['arm64', 'win32'], ['arm64', 'darwin']].map(([arch, platform]) => [...product, arch, platform])
)) {
  const isMac = platform === 'darwin';
  const runtimeIdentifier = isMac ? 'osx-arm64' : `win-${arch}`;
  const executableName = isMac ? commandName : `${commandName}.exe`;
  test(`${component} ${platform}/${arch} tarballs contain the matching runtime and shared launcher`, { timeout: 120_000 }, () => {
    const sandbox = mkdtempSync(join(tmpdir(), 'ExcelMcpNpmPack-'));
    const executable = join(sandbox, executableName);
    const payload = isMac ? Buffer.from('opaque native package fixture') : executableFixture(arch);
    const sourceManifest = join(repoRoot, 'npm-packages', packageName, 'package.json');
    const originalManifest = readFileSync(sourceManifest, 'utf8');
    try {
      writeFileSync(executable, payload);
      if (isMac) {
        mkdirSync(join(sandbox, 'helpers'));
        writeFileSync(join(sandbox, 'helpers', 'excelmcp-screencapture'), payload);
      }
      const result = spawnSync('pwsh', [
        '-NoProfile', '-File', join(repoRoot, 'scripts', 'Build-NpmPackages.ps1'),
        '-Component', component, '-Version', version,
        '-Architecture', arch,
        '-RuntimeIdentifier', runtimeIdentifier,
        '-RuntimeExecutable', executable, '-OutputDirectory', sandbox
      ], { encoding: 'utf8', timeout: 90_000 });
      assert.ifError(result.error);
      assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
      const packages = JSON.parse(result.stdout);

      for (const [kind, archive, name] of [
        ['launcher', packages.LauncherPackage, packageName],
        ['runtime', packages.RuntimePackage, `${packageName}-${platform}-${arch}`]
      ]) {
        const destination = join(sandbox, kind);
        mkdirSync(destination);
        const extraction = spawnSync('tar', ['-xf', archive, '-C', destination], { encoding: 'utf8' });
        assert.ifError(extraction.error);
        assert.equal(extraction.status, 0, extraction.stderr);
        const root = join(destination, 'package');
        const manifestBytes = readFileSync(join(root, 'package.json'));
        assert.notEqual(manifestBytes[0], 0xef, 'Manifest must not have a UTF-8 BOM.');
        assert.ok(!manifestBytes.includes(13), 'Manifest must use LF line endings.');
        const manifest = JSON.parse(manifestBytes.toString('utf8'));
        assert.equal(manifest.name, `@sbroenne/${name}`);
        assert.equal(manifest.version, version);
        assert.equal(readFileSync(join(root, 'LICENSE'), 'utf8'), readFileSync(join(repoRoot, 'LICENSE'), 'utf8'));
        if (kind === 'runtime') {
          assert.equal(manifest.main, executableName);
          assert.deepEqual(manifest.os, [platform]);
          assert.deepEqual(manifest.cpu, [arch]);
          assert.deepEqual(readFileSync(join(root, manifest.main)), payload);
          if (isMac) {
            assert.deepEqual(readFileSync(join(root, 'helpers', 'excelmcp-screencapture')), payload);
          }
        } else {
          assert.deepEqual(manifest.optionalDependencies, {
            [`@sbroenne/${packageName}-win32-x64`]: version,
            [`@sbroenne/${packageName}-darwin-arm64`]: version,
            [`@sbroenne/${packageName}-win32-arm64`]: version
          });
          assert.equal(manifest.bin[commandName], `bin/${commandName}.js`);
          assert.match(readFileSync(join(root, manifest.bin[commandName]), 'utf8'), new RegExp(`packageName: '@sbroenne/${packageName}'`));
          assert.equal(readFileSync(join(root, 'lib', 'launcher.js'), 'utf8'), readFileSync(join(repoRoot, 'npm-packages', 'shared', 'launcher.js'), 'utf8'));
          if (component === 'Cli') {
            assert.equal(manifest.mcpName, undefined, 'The CLI must not register as an MCP server.');
          }
        }
      }
      assert.equal(readFileSync(sourceManifest, 'utf8'), originalManifest, 'Packaging must not stamp the source tree.');
      if (!isMac && process.platform === 'win32' && process.arch !== arch) {
        const inspection = spawnSync('pwsh', [
          '-NoProfile', '-File', join(repoRoot, 'scripts', 'Test-NpmPackages.ps1'),
          '-Component', component, '-Architecture', arch,
          '-LauncherPackage', packages.LauncherPackage, '-RuntimePackage', packages.RuntimePackage
        ], { encoding: 'utf8', timeout: 25_000 });
        assert.ifError(inspection.error);
        assert.equal(inspection.status, 0, `${inspection.stdout}\n${inspection.stderr}`);
        assert.match(inspection.stdout, /archives validated/);
        assert.match(inspection.stdout, /execution NOT RUN/);
      }
    } finally {
      rmSync(sandbox, { recursive: true, force: true });
    }
  });
  if (!isMac) {
  test(`${component} rejects a mislabeled ${arch} runtime`, { timeout: 30_000 }, () => {
    const sandbox = mkdtempSync(join(tmpdir(), 'ExcelMcpNpmMismatch-'));
    try {
      const executable = join(sandbox, `${commandName}.exe`);
      writeFileSync(executable, executableFixture(arch === 'arm64' ? 'x64' : 'arm64'));
      const result = spawnSync('pwsh', [
        '-NoProfile', '-File', join(repoRoot, 'scripts', 'Build-NpmPackages.ps1'),
        '-Component', component, '-Version', version, '-Architecture', arch,
        '-RuntimeExecutable', executable, '-OutputDirectory', sandbox
      ], { encoding: 'utf8', timeout: 25_000 });
      assert.ifError(result.error);
      assert.notEqual(result.status, 0);
      assert.match(result.stderr, /machine type.*does not match/i);
    } finally {
      rmSync(sandbox, { recursive: true, force: true });
    }
  });
  }
}

for (const [name, payload] of [
  ['non-executable', Buffer.from('not an executable')],
  ['invalid offset', (() => {
    const payload = executableFixture('arm64');
    payload.writeUInt32LE(0xffffffff, 0x3c);
    return payload;
  })()],
  ['invalid PE signature', (() => {
    const payload = executableFixture('arm64');
    payload.writeUInt32LE(0, 64);
    return payload;
  })()]
]) {
  test(`packaging rejects ${name} before creating tarballs`, { timeout: 30_000 }, () => {
    const sandbox = mkdtempSync(join(tmpdir(), 'ExcelMcpNpmInvalid-'));
    try {
      const executable = join(sandbox, 'runtime.exe');
      writeFileSync(executable, payload);
      const result = spawnSync('pwsh', [
        '-NoProfile', '-File', join(repoRoot, 'scripts', 'Build-NpmPackages.ps1'),
        '-Version', version, '-Architecture', 'arm64',
        '-RuntimeExecutable', executable, '-OutputDirectory', sandbox
      ], { encoding: 'utf8', timeout: 25_000 });
      assert.ifError(result.error);
      assert.notEqual(result.status, 0);
      assert.match(result.stderr, /not a Windows executable|invalid executable header|invalid PE signature/);
    } finally {
      rmSync(sandbox, { recursive: true, force: true });
    }
  });
}
