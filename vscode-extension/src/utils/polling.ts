export type PollFn<T> = () => Promise<T>;

export class Poller<T> {
  private timer?: ReturnType<typeof setInterval>;
  constructor(private readonly fn: PollFn<T>, private readonly onResult: (r: T) => void, private readonly onError: (e: any) => void, private intervalMs: number, private backoffMs: number) {}

  start() {
    this.stop();
    this.timer = setInterval(() => {
      this.fn().then(this.onResult).catch((e) => {
        this.onError(e);
        this.backoff();
      });
    }, this.intervalMs);
  }

  stop() {
    if (this.timer) clearInterval(this.timer);
    this.timer = undefined;
  }

  private backoff() {
    if (this.timer) clearInterval(this.timer);
    this.timer = setInterval(() => {
      this.fn().then((r) => {
        this.onResult(r);
        // restore normal interval after a successful call
        this.stop();
        this.start();
      }).catch(this.onError);
    }, this.backoffMs);
  }
}
