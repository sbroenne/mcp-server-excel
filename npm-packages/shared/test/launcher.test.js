import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

import { createLauncher } from '../launcher.js';

const { launch, main, resolveRuntime } = createLauncher({
  packageName: '@sbroenne/mcp-server-excel',
  commandName: 'excel-mcp'
});

for (const [packageName, commandName] of [
  ['@sbroenne/excelcli', 'excelcli'],
  ['@sbroenne/mcp-server-excel', 'excel-mcp']
]) {
  for (const arch of ['x64', 'arm64']) {
    test(`${commandName} resolves only its matching ${arch} runtime`, () => {
      const launcher = createLauncher({ packageName, commandName });
      const expectedPackage = `${packageName}-win32-${arch}`;
      assert.equal(launcher.resolveRuntime({
        platform: 'win32',
        arch,
        resolvePackage: name => {
          assert.equal(name, expectedPackage);
          return 'C:\\runtime\\command.exe';
        }
      }), 'C:\\runtime\\command.exe');

      const attempts = [];
      assert.throws(() => launcher.resolveRuntime({
        platform: 'win32',
        arch,
        resolvePackage: name => {
          attempts.push(name);
          throw new Error('missing runtime');
        }
      }), error => error.message.includes(expectedPackage) &&
        error.message.includes('optional dependencies enabled'));
      assert.deepEqual(attempts, [expectedPackage], 'Missing runtimes must not fall back.');
    });
  }

  test(`${commandName} forwards real child I/O and nonzero exit status`, () => {
    const launcherUrl = new URL('../launcher.js', import.meta.url).href;
    const childCode = `
      const fs = require('node:fs');
      process.stdout.write(fs.readFileSync(0, 'utf8') + process.argv[1]);
      process.stderr.write('runtime diagnostic');
      process.exitCode = 23;
    `;
    const script = `
      import { createLauncher } from ${JSON.stringify(launcherUrl)};
      createLauncher(${JSON.stringify({ packageName, commandName })}).launch({
        platform: 'win32',
        arch: 'x64',
        resolvePackage: () => process.execPath,
        args: ${JSON.stringify(['-e', childCode, 'path with spaces'])}
      });
    `;
    const result = spawnSync(process.execPath, ['--input-type=module', '-e', script], {
      input: 'stdin:',
      encoding: 'utf8',
      timeout: 10_000
    });
    assert.ifError(result.error);
    assert.equal(result.status, 23, result.stderr);
    assert.equal(result.stdout, 'stdin:path with spaces');
    assert.equal(result.stderr, 'runtime diagnostic');
  });
}

test('CLI resolves its own runtime and forwards command arguments unchanged', () => {
  const cli = createLauncher({
    packageName: '@sbroenne/excelcli',
    commandName: 'excelcli'
  });
  let invocation;
  cli.launch({
    platform: 'win32',
    arch: 'arm64',
    args: ['-q', 'session', 'open', 'C:\\Data\\Book with spaces.xlsx'],
    resolvePackage: name => {
      assert.equal(name, '@sbroenne/excelcli-win32-arm64');
      return 'C:\\runtime\\excelcli.exe';
    },
    foreground: (...parameters) => { invocation = parameters; }
  });
  assert.deepEqual(invocation, [
    'C:\\runtime\\excelcli.exe',
    ['-q', 'session', 'open', 'C:\\Data\\Book with spaces.xlsx'],
    { shell: false, stdio: 'inherit', windowsHide: true }
  ]);
});

test('CLI missing runtime errors identify the CLI package and command', () => {
  const cli = createLauncher({
    packageName: '@sbroenne/excelcli',
    commandName: 'excelcli'
  });
  let stderr = '';
  const exitCode = cli.main({
    launchProcess: () => cli.resolveRuntime({
      platform: 'win32',
      arch: 'x64',
      resolvePackage: () => { throw new Error('package not found'); }
    }),
    stderr: { write: value => { stderr += value; } }
  });
  assert.equal(exitCode, 1);
  assert.match(stderr, /^excelcli: Could not find @sbroenne\/excelcli-win32-x64/);
  assert.match(stderr, /Reinstall @sbroenne\/excelcli with optional dependencies enabled/);
});

test('resolveRuntime rejects unsupported operating systems', () => {
  assert.throws(
    () => resolveRuntime({ platform: 'linux', arch: 'x64' }),
    /Windows only/
  );
});

test('resolveRuntime supports native Windows Arm64', () => {
  assert.equal(
    resolveRuntime({
      platform: 'win32',
      arch: 'arm64',
      resolvePackage: () => 'C:\\runtime\\mcp-excel.exe'
    }),
    'C:\\runtime\\mcp-excel.exe'
  );
});

test('resolveRuntime rejects unsupported Windows architectures', () => {
  assert.throws(
    () => resolveRuntime({ platform: 'win32', arch: 'ia32' }),
    /x64 or Arm64/
  );
});

test('resolveRuntime explains how to restore an omitted binary package', () => {
  assert.throws(
    () =>
      resolveRuntime({
        platform: 'win32',
        arch: 'x64',
        resolvePackage: () => {
          throw new Error('package not found');
        }
      }),
    /optional dependencies enabled/
  );
});

test('launch forwards arguments and foreground process options', () => {
  const expectedChild = {};
  let invocation;

  const actualChild = launch({
    platform: 'win32',
    arch: 'x64',
    args: ['--version'],
    resolvePackage: () => 'C:\\runtime\\mcp-excel.exe',
    foreground: (...parameters) => {
      invocation = parameters;
      return expectedChild;
    }
  });

  assert.equal(actualChild, expectedChild);
  assert.deepEqual(invocation, [
    'C:\\runtime\\mcp-excel.exe',
    ['--version'],
    {
      shell: false,
      stdio: 'inherit',
      windowsHide: true
    }
  ]);
});

test('main reports launcher failures only on stderr', () => {
  let stderr = '';

  const exitCode = main({
    launchProcess: () => {
      throw new Error('runtime unavailable');
    },
    stderr: {
      write: value => {
        stderr += value;
      }
    }
  });

  assert.equal(exitCode, 1);
  assert.equal(stderr, 'excel-mcp: runtime unavailable\n');
});

test('main leaves process lifetime to foreground-child after launch', () => {
  const child = {};

  const exitCode = main({
    launchProcess: () => child,
    stderr: {
      write: () => assert.fail('stderr should not be written')
    }
  });

  assert.equal(exitCode, undefined);
});
