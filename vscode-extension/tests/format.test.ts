import * as assert from 'assert';
import { basename, truncate } from '../src/utils/format';

suite('format utilities', () => {
  suite('basename', () => {
    test('extracts filename from forward-slash path', () => {
      assert.strictEqual(basename('C:/Temp/file.xlsx'), 'file.xlsx');
    });

    test('extracts filename from backslash path', () => {
      assert.strictEqual(basename('C:\\Temp\\file.xlsx'), 'file.xlsx');
    });

    test('returns filename unchanged if no path separator', () => {
      assert.strictEqual(basename('file.xlsx'), 'file.xlsx');
    });

    test('handles nested paths', () => {
      assert.strictEqual(basename('C:/Users/test/Documents/data.xlsx'), 'data.xlsx');
    });

    test('handles empty string', () => {
      assert.strictEqual(basename(''), '');
    });
  });

  suite('truncate', () => {
    test('returns string unchanged when shorter than max', () => {
      assert.strictEqual(truncate('hello', 10), 'hello');
    });

    test('returns string unchanged when equal to max', () => {
      assert.strictEqual(truncate('hello', 5), 'hello');
    });

    test('truncates and adds ellipsis when longer than max', () => {
      assert.strictEqual(truncate('1234567890', 5), '1234…');
    });

    test('handles edge case of max=1', () => {
      assert.strictEqual(truncate('hello', 1), '…');
    });

    test('handles empty string', () => {
      assert.strictEqual(truncate('', 5), '');
    });
  });
});
