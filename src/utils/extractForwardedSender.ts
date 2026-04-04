/**
 * Attempts to extract the original sender from a forwarded email's HTML body.
 * Uses DOMParser for robust HTML handling, with regex fallback for edge cases.
 * Returns null if no forwarded sender information is found.
 */
export function extractForwardedSender(
  html: string,
): { name?: string; address?: string } | null {
  // Try DOM-based extraction first (more robust than regex)
  const domResult = extractViaDom(html);
  if (domResult) return domResult;

  // Fallback: regex for plain-text "On ... wrote:" patterns not caught by DOM
  const wrotePattern =
    /On\s.{10,120}?([\w\s.'-]+?)\s*(?:&lt;|<)([\w.+-]+@[\w.-]+\.\w+)(?:&gt;|>)\s*wrote:/i;
  const wroteMatch = wrotePattern.exec(html);
  if (wroteMatch) {
    return parseMatch(wroteMatch[1], wroteMatch[2]);
  }

  // Fallback: plain email-only "From:" line
  const emailOnlyFrom =
    /(?:^|>|\n)\s*From:\s*([\w.+-]+@[\w.-]+\.\w+)/im;
  const emailOnlyMatch = emailOnlyFrom.exec(html);
  if (emailOnlyMatch) {
    return { address: emailOnlyMatch[1].trim().toLowerCase() };
  }

  return null;
}

function extractViaDom(html: string): { name?: string; address?: string } | null {
  let doc: Document;
  try {
    const parser = new DOMParser();
    doc = parser.parseFromString(html, "text/html");
  } catch {
    return null;
  }

  // Strategy 1: Outlook-style forwarded header — look for a <b>From:</b> element
  const bolds = doc.querySelectorAll("b");
  for (const bold of Array.from(bolds)) {
    if (/^\s*From:\s*$/i.test(bold.textContent || "")) {
      // Get the text content AFTER the <b>From:</b> element
      let afterText = "";
      let sibling = bold.nextSibling;
      while (sibling) {
        afterText += sibling.textContent || "";
        if (afterText.length > 200) break;
        sibling = sibling.nextSibling;
      }
      // Fallback to parent text minus the label
      if (!afterText.trim()) {
        const parent = bold.parentElement;
        if (parent) {
          afterText = (parent.textContent || "").replace(/^.*From:\s*/i, "");
        }
      }
      const result = extractEmailFromText(afterText.trim());
      if (result) return result;
    }
  }

  // Strategy 2: Look in divRplyFwdMsg (Outlook forwarded block)
  const fwdBlock = doc.querySelector('[id="divRplyFwdMsg"]');
  if (fwdBlock) {
    const text = fwdBlock.textContent || "";
    const fromLine = text.match(/From:\s*([^\n]+)/i);
    if (fromLine) {
      const result = extractEmailFromText(fromLine[1]);
      if (result) return result;
    }
  }

  // Strategy 3: Gmail quote block
  const gmailQuote = doc.querySelector(".gmail_quote");
  if (gmailQuote) {
    const text = gmailQuote.textContent || "";
    // "On ... <email> wrote:" pattern
    const wroteMatch = text.match(/([\w\s.'-]+?)\s*<([\w.+-]+@[\w.-]+\.\w+)>\s*wrote:/i);
    if (wroteMatch) {
      return parseMatch(wroteMatch[1], wroteMatch[2]);
    }
  }

  return null;
}

function extractEmailFromText(text: string): { name?: string; address?: string } | null {
  // "Name <email>" or "Name [mailto:email]"
  const full = text.match(/([^<[\n]*?)(?:<|\[mailto:)([\w.+-]+@[\w.-]+\.\w+)/i);
  if (full) {
    return parseMatch(full[1], full[2]);
  }
  // Bare email address
  const bare = text.match(/([\w.+-]+@[\w.-]+\.\w+)/);
  if (bare) {
    return { address: bare[1].trim().toLowerCase() };
  }
  return null;
}

function parseMatch(
  rawName: string,
  rawAddress: string,
): { name?: string; address?: string } {
  const address = rawAddress.trim().toLowerCase();
  const name = rawName
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (name) {
    return { name, address };
  }
  return { address };
}
