import * as assert from 'assert';

/**
 * ExcelStatusClient Type Contract Tests
 *
 * These tests verify the expected type contracts for status client responses.
 * The actual ExcelStatusClient class requires VS Code and is tested in integration tests.
 *
 * These tests document the expected response shapes from the MCP server.
 */

// Type definitions (mirrored from excelStatusClient.ts for unit testing)
type ExcelSessionInfo = {
  sessionId: string;
  filePath: string;
  canClose: boolean;
  isExcelVisible: boolean;
};

type ListSessionsResult = {
  success: boolean;
  errorMessage?: string;
  sessions?: ExcelSessionInfo[];
};

suite('ExcelStatusClient Type Contracts', () => {
  suite('ListSessionsResult', () => {
    test('success result has correct shape', () => {
      const result: ListSessionsResult = {
        success: true,
        sessions: []
      };

      assert.strictEqual(result.success, true);
      assert.ok(Array.isArray(result.sessions));
      assert.strictEqual(result.errorMessage, undefined);
    });

    test('error result has correct shape', () => {
      const result: ListSessionsResult = {
        success: false,
        errorMessage: 'Connection failed'
      };

      assert.strictEqual(result.success, false);
      assert.strictEqual(result.errorMessage, 'Connection failed');
      assert.strictEqual(result.sessions, undefined);
    });

    test('success with sessions has correct shape', () => {
      const result: ListSessionsResult = {
        success: true,
        sessions: [
          {
            sessionId: 'sess-001',
            filePath: 'C:\\Temp\\test.xlsx',
            canClose: true,
            isExcelVisible: false
          },
          {
            sessionId: 'sess-002',
            filePath: 'D:\\Data\\report.xlsm',
            canClose: false,
            isExcelVisible: true
          }
        ]
      };

      assert.strictEqual(result.success, true);
      assert.strictEqual(result.sessions?.length, 2);
      assert.strictEqual(result.sessions?.[0].sessionId, 'sess-001');
      assert.strictEqual(result.sessions?.[1].isExcelVisible, true);
    });
  });

  suite('ExcelSessionInfo', () => {
    test('session info has all required fields', () => {
      const session: ExcelSessionInfo = {
        sessionId: 'test-123',
        filePath: 'C:\\Temp\\test.xlsx',
        canClose: true,
        isExcelVisible: false
      };

      assert.strictEqual(session.sessionId, 'test-123');
      assert.strictEqual(session.filePath, 'C:\\Temp\\test.xlsx');
      assert.strictEqual(session.canClose, true);
      assert.strictEqual(session.isExcelVisible, false);
    });

    test('session with visible Excel', () => {
      const session: ExcelSessionInfo = {
        sessionId: 'visible-session',
        filePath: 'C:\\Documents\\visible.xlsx',
        canClose: true,
        isExcelVisible: true
      };

      assert.strictEqual(session.isExcelVisible, true);
    });

    test('session that cannot be closed', () => {
      const session: ExcelSessionInfo = {
        sessionId: 'busy-session',
        filePath: 'C:\\Work\\processing.xlsx',
        canClose: false,
        isExcelVisible: false
      };

      assert.strictEqual(session.canClose, false);
    });
  });

  suite('CloseSessionResult', () => {
    test('successful close', () => {
      const result = {
        success: true
      };

      assert.strictEqual(result.success, true);
    });

    test('failed close with error message', () => {
      const result = {
        success: false,
        errorMessage: 'Session has active operations'
      };

      assert.strictEqual(result.success, false);
      assert.strictEqual(result.errorMessage, 'Session has active operations');
    });
  });
});
