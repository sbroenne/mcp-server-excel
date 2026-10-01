import { chmodSync, copyFileSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const extensionRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repositoryRoot = resolve(extensionRoot, '..');
const project = resolve(repositoryRoot, 'src', 'ExcelMcp.McpServer', 'ExcelMcp.McpServer.csproj');
const allTargets = [
  { runtime: 'win-x64', directory: 'win32-x64', executable: 'Sbroenne.ExcelMcp.McpServer.exe' },
  { runtime: 'osx-arm64', directory: 'darwin-arm64', executable: 'Sbroenne.ExcelMcp.McpServer' }
];
if (process.platform !== 'win32' && process.platform !== 'darwin') {
  throw new Error(`Extension runtime packaging is unsupported on ${process.platform}.`);
}
// Windows hosts package only the Windows runtime because the Darwin helper requires Swift and codesign.
const targets = process.platform === 'win32'
  ? allTargets.filter((target) => target.runtime === 'win-x64')
  : allTargets;

for (const target of targets) {
  const output = resolve(extensionRoot, 'bin', target.directory);
  const result = spawnSync(
    'dotnet',
    [
      'publish',
      project,
      '-c', 'Release',
      '-r', target.runtime,
      '--self-contained', 'true',
      '-p:PublishSingleFile=true',
      '-p:IncludeNativeLibrariesForSelfExtract=true',
      '-p:PublishTrimmed=false',
      '-p:PublishReadyToRun=false',
      '-p:NuGetAudit=false',
      '-nodeReuse:false',
      '-o', output,
      '--verbosity', 'minimal'
    ],
    { cwd: repositoryRoot, encoding: 'utf8', stdio: 'inherit' }
  );

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(`MCP runtime publish failed for ${target.runtime} with exit code ${result.status ?? 1}.`);
  }

  if (target.runtime.startsWith('osx-')) {
    const helperBuildRoot = resolve(extensionRoot, '.helper-build');
    const helperSource = resolve(helperBuildRoot, target.runtime, 'helpers', 'excelmcp-screencapture');
    const helperDirectory = resolve(output, 'helpers');
    const helperPath = resolve(helperDirectory, 'excelmcp-screencapture');
    rmSync(helperBuildRoot, { recursive: true, force: true });
    const helperBuild = spawnSync(
      'pwsh',
      [
        '-NoProfile', '-File', resolve(repositoryRoot, 'scripts', 'Build-MacScreenCaptureHelper.ps1'),
        '-RuntimeIdentifier', target.runtime,
        '-OutputRoot', helperBuildRoot
      ],
      { cwd: repositoryRoot, encoding: 'utf8', stdio: 'inherit' }
    );
    if (helperBuild.error) {
      throw helperBuild.error;
    }
    if (helperBuild.status !== 0) {
      throw new Error(`ScreenCaptureKit helper build failed with exit code ${helperBuild.status ?? 1}.`);
    }
    mkdirSync(helperDirectory, { recursive: true });
    copyFileSync(helperSource, helperPath);
    chmodSync(helperPath, 0o755);
    const helperSigning = spawnSync(
      'pwsh',
      ['-NoProfile', '-File', resolve(repositoryRoot, 'scripts', 'Sign-MacBinary.ps1'), '-Path', helperPath],
      { cwd: repositoryRoot, encoding: 'utf8', stdio: 'inherit' }
    );
    rmSync(helperBuildRoot, { recursive: true, force: true });
    if (helperSigning.error) {
      throw helperSigning.error;
    }
    if (helperSigning.status !== 0) {
      throw new Error(`ScreenCaptureKit helper signing failed with exit code ${helperSigning.status ?? 1}.`);
    }

    const executablePath = resolve(output, target.executable);
    chmodSync(executablePath, 0o755);
    const signing = spawnSync(
      'pwsh',
      [
        '-NoProfile', '-File', resolve(repositoryRoot, 'scripts', 'Sign-MacBinary.ps1'),
        '-Path', executablePath, '-AutomationClient'
      ],
      { cwd: repositoryRoot, encoding: 'utf8', stdio: 'inherit' }
    );
    if (signing.error) {
      throw signing.error;
    }
    if (signing.status !== 0) {
      throw new Error(`MCP runtime signing failed with exit code ${signing.status ?? 1}.`);
    }
  }
}
