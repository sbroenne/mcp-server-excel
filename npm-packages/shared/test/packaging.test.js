import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const repoRoot = fileURLToPath(new URL('../../../', import.meta.url));
const version = '9.8.7-test.1';

for (const [component, packageName, commandName] of [
  ['McpServer', 'mcp-server-excel', 'mcp-excel'],
  ['Cli', 'excelcli', 'excelcli']
]) {
  for (const [runtimeIdentifier, runtimeSuffix, executableName, expectedOs, expectedCpu] of [
    ['win-x64', 'win32-x64', `${commandName}.exe`, ['win32'], ['x64', 'arm64']],
    ['osx-arm64', 'darwin-arm64', commandName, ['darwin'], ['arm64']]
  ]) {
  test(`${component} ${runtimeIdentifier} tarballs contain the matching runtime and shared launcher`, { timeout: 120_000 }, () => {
    const sandbox = mkdtempSync(join(tmpdir(), 'ExcelMcpNpmPack-'));
    const executable = join(sandbox, executableName);
    const payload = Buffer.from('package fixture, not an executable');
    const sourceManifest = join(repoRoot, 'npm-packages', packageName, 'package.json');
    const originalManifest = readFileSync(sourceManifest, 'utf8');
    try {
      writeFileSync(executable, payload);
      const result = spawnSync('pwsh', [
        '-NoProfile', '-File', join(repoRoot, 'scripts', 'Build-NpmPackages.ps1'),
        '-Component', component, '-Version', version,
        '-RuntimeIdentifier', runtimeIdentifier,
        '-RuntimeExecutable', executable, '-OutputDirectory', sandbox
      ], { encoding: 'utf8', timeout: 90_000 });
      assert.ifError(result.error);
      assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
      const packages = JSON.parse(result.stdout);

      for (const [kind, archive, name] of [
        ['launcher', packages.LauncherPackage, packageName],
        ['runtime', packages.RuntimePackage, `${packageName}-${runtimeSuffix}`]
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
          assert.deepEqual(manifest.os, expectedOs);
          assert.deepEqual(manifest.cpu, expectedCpu);
          assert.deepEqual(readFileSync(join(root, manifest.main)), payload);
        } else {
          assert.deepEqual(manifest.optionalDependencies, {
            [`@sbroenne/${packageName}-win32-x64`]: version,
            [`@sbroenne/${packageName}-darwin-arm64`]: version
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
    } finally {
      rmSync(sandbox, { recursive: true, force: true });
    }
  });
  }
}
