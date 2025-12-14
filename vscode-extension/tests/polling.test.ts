import * as assert from 'assert';
import { Poller } from '../src/utils/polling';

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

suite('Poller', () => {
  test('calls onResult with fetched data', async () => {
    const results: number[] = [];
    let fetchCount = 0;

    const poller = new Poller<number>(
      async () => {
        fetchCount++;
        return fetchCount;
      },
      (r) => results.push(r),
      () => {},
      100, // intervalMs
      150  // backoffMs
    );

    poller.start();
    await delay(350);
    poller.stop();

    assert.ok(results.length >= 2, `Expected at least 2 results, got ${results.length}`);
    assert.strictEqual(results[0], 1);
  });

  test('applies backoff after error and recovers', async () => {
    const results: number[] = [];
    const errors: Error[] = [];
    let fetchCount = 0;

    const poller = new Poller<number>(
      async () => {
        fetchCount++;
        if (fetchCount === 1) {
          throw new Error('fail once');
        }
        return fetchCount;
      },
      (r) => results.push(r),
      (e) => errors.push(e),
      100,
      150
    );

    poller.start();
    await delay(400);
    poller.stop();

    // First call errors, subsequent calls succeed
    assert.strictEqual(errors.length, 1);
    assert.strictEqual(errors[0].message, 'fail once');
    assert.ok(results.length >= 1, `Expected at least 1 result, got ${results.length}`);
    assert.ok(results[results.length - 1] >= 2, `Expected last result >= 2, got ${results[results.length - 1]}`);
  });

  test('stops polling when stop() is called', async () => {
    let fetchCount = 0;

    const poller = new Poller<number>(
      async () => {
        fetchCount++;
        return fetchCount;
      },
      () => {},
      () => {},
      50,
      100
    );

    poller.start();
    await delay(75);
    poller.stop();
    const countAtStop = fetchCount;
    await delay(150);

    // Should not have increased after stop
    assert.strictEqual(fetchCount, countAtStop);
  });

  test('is safe to call stop() multiple times', () => {
    const poller = new Poller<number>(
      async () => 1,
      () => {},
      () => {},
      100,
      150
    );

    poller.start();
    poller.stop();
    // Should not throw
    poller.stop();
    assert.ok(true, 'Multiple stop() calls did not throw');
  });
});
