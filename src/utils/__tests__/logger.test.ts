import { logger, setDebugEnabled, isDebugEnabled } from "../logger";

describe("logger", () => {
  it("respects debug enabled flag", () => {
    const spy = jest.spyOn(console, "debug").mockImplementation();

    setDebugEnabled(false);
    logger.debug("Test", "should not appear");
    expect(spy).not.toHaveBeenCalled();

    setDebugEnabled(true);
    expect(isDebugEnabled()).toBe(true);
    logger.debug("Test", "should appear");
    expect(spy).toHaveBeenCalledTimes(1);

    spy.mockRestore();
    setDebugEnabled(false);
  });

  it("always logs warnings regardless of debug flag", () => {
    const spy = jest.spyOn(console, "warn").mockImplementation();

    setDebugEnabled(false);
    logger.warn("Test", "warning message");
    expect(spy).toHaveBeenCalledTimes(1);

    spy.mockRestore();
  });
});
