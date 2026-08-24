import assert from 'node:assert/strict';
import test from 'node:test';

import { launch, main, resolveRuntime } from '../lib/launcher.js';

test('resolveRuntime rejects unsupported operating systems', () => {
  assert.throws(
    () => resolveRuntime({ platform: 'linux', arch: 'x64' }),
    /Windows only/
  );
});

test('resolveRuntime supports Windows Arm64 through x64 emulation', () => {
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
