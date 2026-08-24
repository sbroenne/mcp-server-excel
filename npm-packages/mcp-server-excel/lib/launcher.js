import { createRequire } from 'node:module';

import { foregroundChild } from 'foreground-child';

const runtimePackageName = '@sbroenne/mcp-server-excel-win32-x64';
const require = createRequire(import.meta.url);

export function resolveRuntime({
  platform = process.platform,
  arch = process.arch,
  resolvePackage = packageName => require.resolve(packageName)
} = {}) {
  if (platform !== 'win32') {
    throw new Error('ExcelMcp is Windows only.');
  }

  if (arch !== 'x64' && arch !== 'arm64') {
    throw new Error(`ExcelMcp requires Windows x64 or Arm64; this Node.js process is ${arch}.`);
  }

  try {
    return resolvePackage(runtimePackageName);
  } catch (cause) {
    throw new Error(
      `Could not find ${runtimePackageName}. Reinstall @sbroenne/mcp-server-excel ` +
        'with optional dependencies enabled; do not use --omit=optional.',
      { cause }
    );
  }
}

export function launch({
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

export function main({
  launchProcess = launch,
  stderr = process.stderr
} = {}) {
  try {
    launchProcess();
    return undefined;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    stderr.write(`excel-mcp: ${message}\n`);
    return 1;
  }
}
