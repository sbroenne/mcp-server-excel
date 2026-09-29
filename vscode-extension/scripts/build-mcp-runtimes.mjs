import { chmodSync, copyFileSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const extensionRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repositoryRoot = resolve(extensionRoot, '..');
const project = resolve(repositoryRoot, 'src', 'ExcelMcp.McpServer', 'ExcelMcp.McpServer.csproj');
const targets = [
  { runtime: 'win-x64', directory: 'win32-x64', executable: 'Sbroenne.ExcelMcp.McpServer.exe' },
  { runtime: 'osx-arm64', directory: 'darwin-arm64', executable: 'Sbroenne.ExcelMcp.McpServer' }
];

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
    process.exit(result.status ?? 1);
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
      process.exit(helperBuild.status ?? 1);
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
      process.exit(helperSigning.status ?? 1);
    }

    const executablePath = resolve(output, target.executable);
    chmodSync(executablePath, 0o755);
    const signing = spawnSync(
      'pwsh',
      ['-NoProfile', '-File', resolve(repositoryRoot, 'scripts', 'Sign-MacBinary.ps1'), '-Path', executablePath],
      { cwd: repositoryRoot, encoding: 'utf8', stdio: 'inherit' }
    );
    if (signing.error) {
      throw signing.error;
    }
    if (signing.status !== 0) {
      process.exit(signing.status ?? 1);
    }
  }
}
