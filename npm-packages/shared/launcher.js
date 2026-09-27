import { createRequire } from 'node:module';

import { foregroundChild } from 'foreground-child';

const require = createRequire(import.meta.url);

export function createLauncher({ packageName, commandName }) {
  const runtimePackageName = `${packageName}-win32-x64`;

  function resolveRuntime({
    platform = process.platform,
    arch = process.arch,
    resolvePackage = name => require.resolve(name)
  } = {}) {
    let runtimePackageName;
    if (platform === 'win32' && (arch === 'x64' || arch === 'arm64')) {
      runtimePackageName = `${packageName}-win32-x64`;
    } else if (platform === 'darwin' && arch === 'arm64') {
      runtimePackageName = `${packageName}-darwin-arm64`;
    } else if (platform === 'darwin' && arch === 'x64') {
      runtimePackageName = `${packageName}-darwin-x64`;
    } else {
      throw new Error(
        `ExcelMcp supports Windows x64/Arm64 and macOS x64/Arm64; this Node.js process is ${platform}-${arch}.`
      );
    }

    try {
      return resolvePackage(runtimePackageName);
    } catch (cause) {
      throw new Error(
        `Could not find ${runtimePackageName}. Reinstall ${packageName} ` +
          'with optional dependencies enabled; do not use --omit=optional.',
        { cause }
      );
    }
  }

  function launch({
    args = process.argv.slice(2),
    platform = process.platform,
    arch = process.arch,
    resolvePackage,
    foreground = foregroundChild
  } = {}) {
    const executable = resolveRuntime({ platform, arch, resolvePackage });
    return foreground(executable, args, {
      shell: false,
      stdio: 'inherit',
      windowsHide: true
    });
  }

  function main({ launchProcess = launch, stderr = process.stderr } = {}) {
    try {
      launchProcess();
      return undefined;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      stderr.write(`${commandName}: ${message}\n`);
      return 1;
    }
  }

  return { resolveRuntime, launch, main };
}
