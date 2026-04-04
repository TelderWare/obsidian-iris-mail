import TurndownService from "turndown";

let instance: TurndownService | null = null;

function getTurndown(): TurndownService {
  if (!instance) {
    instance = new TurndownService({
      headingStyle: "atx",
      codeBlockStyle: "fenced",
      bulletListMarker: "-",
    });

    // Remove tracking pixels (1x1 or 0x0 images)
    instance.addRule("removeTrackingPixels", {
      filter: (node: HTMLElement) => {
        if (node.nodeName !== "IMG") return false;
        const w = node.getAttribute("width");
        const h = node.getAttribute("height");
        return (w === "1" && h === "1") || (w === "0" && h === "0");
      },
      replacement: () => "",
    });

    // Remove style tags
    instance.addRule("removeStyleTags", {
      filter: "style",
      replacement: () => "",
    });

    // Remove hidden elements
    instance.addRule("removeHidden", {
      filter: (node: HTMLElement) => {
        const style = node.getAttribute("style") || "";
        return (
          style.includes("display:none") ||
          style.includes("display: none") ||
          style.includes("visibility:hidden")
        );
      },
      replacement: () => "",
    });
  }
  return instance;
}

export function htmlToMarkdown(html: string): string {
  return getTurndown().turndown(html);
}
