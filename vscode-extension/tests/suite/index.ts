/**
 * Integration Test Suite Entry Point
 *
 * This module runs inside VS Code and discovers/executes all integration tests.
 * Uses Mocha for test framework (standard for VS Code extension tests).
 */

import * as path from 'path';
import Mocha from 'mocha';
import { glob } from 'glob';

export async function run(): Promise<void> {
  // Create the mocha test runner
  const mocha = new Mocha({
    ui: 'tdd',      // Use TDD interface (suite/test) instead of BDD (describe/it)
    color: true,
    timeout: 60000, // 60s timeout for integration tests
  });

  const testsRoot = path.resolve(__dirname, '.');

  // Find all test files
  const files = await glob('**/*.test.js', { cwd: testsRoot });

  // Add files to the test suite
  for (const file of files) {
    mocha.addFile(path.resolve(testsRoot, file));
  }

  return new Promise((resolve, reject) => {
    try {
      // Run the mocha tests
      mocha.run((failures) => {
        if (failures > 0) {
          reject(new Error(`${failures} tests failed.`));
        } else {
          resolve();
        }
      });
    } catch (err) {
      console.error(err);
      reject(err);
    }
  });
}
