/**
 * A simple concurrency pool that limits the number of
 * simultaneously running async tasks.
 */
export class ConcurrencyPool {
  private running = 0;
  private queue: Array<() => void> = [];

  constructor(private limit: number) {}

  /**
   * Run `fn` respecting the concurrency limit.
   * If the pool is full, waits until a slot opens.
   */
  async run<T>(fn: () => Promise<T>): Promise<T> {
    if (this.running >= this.limit) {
      await new Promise<void>((resolve) => this.queue.push(resolve));
    }
    this.running++;
    try {
      return await fn();
    } finally {
      this.running--;
      const next = this.queue.shift();
      if (next) next();
    }
  }
}
