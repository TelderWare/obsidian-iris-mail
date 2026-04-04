import { ConcurrencyPool } from "../concurrency";

describe("ConcurrencyPool", () => {
  it("limits concurrent executions", async () => {
    const pool = new ConcurrencyPool(2);
    let running = 0;
    let maxRunning = 0;

    const task = () =>
      pool.run(async () => {
        running++;
        maxRunning = Math.max(maxRunning, running);
        await new Promise((r) => setTimeout(r, 10));
        running--;
      });

    await Promise.all([task(), task(), task(), task(), task()]);
    expect(maxRunning).toBeLessThanOrEqual(2);
  });

  it("returns results from tasks", async () => {
    const pool = new ConcurrencyPool(3);
    const results = await Promise.all([
      pool.run(async () => 1),
      pool.run(async () => 2),
      pool.run(async () => 3),
    ]);
    expect(results).toEqual([1, 2, 3]);
  });

  it("propagates errors", async () => {
    const pool = new ConcurrencyPool(1);
    await expect(
      pool.run(async () => {
        throw new Error("test error");
      }),
    ).rejects.toThrow("test error");
  });

  it("releases slot on error so subsequent tasks can run", async () => {
    const pool = new ConcurrencyPool(1);
    try {
      await pool.run(async () => {
        throw new Error("fail");
      });
    } catch { /* expected */ }

    const result = await pool.run(async () => "recovered");
    expect(result).toBe("recovered");
  });
});
