import { setIcon } from "obsidian";

interface ToolbarCallbacks {
  onRefresh: () => void;
}

export class Toolbar {
  constructor(containerEl: HTMLElement, callbacks: ToolbarCallbacks) {
    const refreshBtn = containerEl.createEl("button", {
      cls: "iris-toolbar-btn clickable-icon",
      attr: { "aria-label": "Refresh" },
    });
    setIcon(refreshBtn, "refresh-cw");
    refreshBtn.addEventListener("click", () => callbacks.onRefresh());
  }
}
