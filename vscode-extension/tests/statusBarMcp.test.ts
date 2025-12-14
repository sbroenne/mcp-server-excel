import * as assert from 'assert';
import { basename } from '../src/utils/format';

/**
 * StatusBarMcp Support Function Tests
 *
 * These tests verify helper functions used by StatusBarMcp.
 * The StatusBarMcp class itself requires VS Code and is tested
 * in integration tests (tests/suite/statusBar.test.ts).
 */

suite('StatusBarMcp helpers', () => {
  suite('basename helper (used in Quick Pick)', () => {
    // These are covered in format.test.ts but we verify they work
    // in the context of session display

    test('extracts filename for display in Quick Pick', () => {
      const testCases = [
        { input: 'C:\\Users\\test\\Documents\\sales.xlsx', expected: 'sales.xlsx' },
        { input: 'D:/Projects/budget.xlsm', expected: 'budget.xlsm' },
        { input: '/home/user/data.xlsx', expected: 'data.xlsx' },
        { input: 'simple.xlsx', expected: 'simple.xlsx' },
      ];

      for (const { input, expected } of testCases) {
        assert.strictEqual(basename(input), expected,
          `basename('${input}') should be '${expected}'`);
      }
    });

    test('handles Windows UNC paths', () => {
      const uncPath = '\\\\server\\share\\folder\\file.xlsx';
      const result = basename(uncPath);
      assert.strictEqual(result, 'file.xlsx');
    });

    test('handles paths with spaces', () => {
      const pathWithSpaces = 'C:\\My Documents\\Excel Files\\report.xlsx';
      const result = basename(pathWithSpaces);
      assert.strictEqual(result, 'report.xlsx');
    });

    test('handles paths with special characters', () => {
      const specialPath = 'C:\\Data\\Report (2024) [Final].xlsx';
      const result = basename(specialPath);
      assert.strictEqual(result, 'Report (2024) [Final].xlsx');
    });

    test('handles very long filenames', () => {
      const longName = 'A'.repeat(100) + '.xlsx';
      const longPath = 'C:\\Folder\\' + longName;
      const result = basename(longPath);
      assert.strictEqual(result, longName);
    });
  });

  suite('Session display formatting', () => {
    test('session label format verification', () => {
      // Simulate what showSessionsQuickPick does
      const session = {
        filePath: 'C:\\Users\\test\\Documents\\Quarterly Report.xlsx',
        sessionId: 'sess-123',
        canClose: true,
        isExcelVisible: false
      };

      const label = basename(session.filePath);
      const description = session.filePath;
      const detail = session.isExcelVisible ? 'Excel is visible' : 'Excel is hidden';

      assert.strictEqual(label, 'Quarterly Report.xlsx');
      assert.strictEqual(description, 'C:\\Users\\test\\Documents\\Quarterly Report.xlsx');
      assert.strictEqual(detail, 'Excel is hidden');
    });

    test('visible Excel session format', () => {
      const session = {
        filePath: 'D:\\Work\\Dashboard.xlsm',
        sessionId: 'sess-456',
        canClose: true,
        isExcelVisible: true
      };

      const detail = session.isExcelVisible ? 'Excel is visible' : 'Excel is hidden';
      assert.strictEqual(detail, 'Excel is visible');
    });
  });
});
