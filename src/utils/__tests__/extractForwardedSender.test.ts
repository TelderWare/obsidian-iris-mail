import { extractForwardedSender } from "../extractForwardedSender";

describe("extractForwardedSender", () => {
  it("extracts Outlook-style forwarded sender", () => {
    const html = `
      <div id="divRplyFwdMsg">
        <b>From:</b> John Smith &lt;john@example.com&gt;<br>
        <b>Sent:</b> Monday, January 1, 2025<br>
      </div>
    `;
    const result = extractForwardedSender(html);
    expect(result).toEqual({ name: "John Smith", address: "john@example.com" });
  });

  it("extracts Gmail wrote: style sender", () => {
    const html = `
      <div class="gmail_quote">
        On Mon, Jan 1, 2025 at 10:00 AM Jane Doe &lt;jane@example.com&gt; wrote:
        <blockquote>Original message</blockquote>
      </div>
    `;
    const result = extractForwardedSender(html);
    expect(result).not.toBeNull();
    expect(result!.address).toBe("jane@example.com");
  });

  it("extracts email-only From line", () => {
    const html = `
      <p>From: user@domain.com</p>
      <p>Some content</p>
    `;
    const result = extractForwardedSender(html);
    expect(result).toEqual({ address: "user@domain.com" });
  });

  it("returns null for non-forwarded email", () => {
    const html = "<p>Hello, this is a regular email.</p>";
    const result = extractForwardedSender(html);
    expect(result).toBeNull();
  });
});
