import { EmailStore } from "../EmailStore";

describe("EmailStore.hashPrompt", () => {
  it("returns a consistent hash for the same input", () => {
    const hash1 = EmailStore.hashPrompt("test prompt");
    const hash2 = EmailStore.hashPrompt("test prompt");
    expect(hash1).toBe(hash2);
  });

  it("returns different hashes for different inputs", () => {
    const hash1 = EmailStore.hashPrompt("prompt A");
    const hash2 = EmailStore.hashPrompt("prompt B");
    expect(hash1).not.toBe(hash2);
  });

  it("never contains dashes", () => {
    // Test with many inputs to check the old bug with negative hashes
    for (let i = 0; i < 100; i++) {
      const hash = EmailStore.hashPrompt(`test-${i}-prompt-${Math.random()}`);
      expect(hash).not.toContain("-");
    }
  });

  it("returns at least 5 characters", () => {
    const hash = EmailStore.hashPrompt("short");
    expect(hash.length).toBeGreaterThanOrEqual(5);
  });
});
