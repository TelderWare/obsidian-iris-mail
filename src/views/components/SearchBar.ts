import { setIcon } from "obsidian";

interface SearchBarCallbacks {
  onSearch: (query: string) => void;
}

export class SearchBar {
  private wrapperEl: HTMLElement;
  private inputEl: HTMLInputElement;

  constructor(containerEl: HTMLElement, callbacks: SearchBarCallbacks) {
    this.wrapperEl = containerEl.createDiv({
      cls: "iris-search-collapsible",
    });

    const iconBtn = this.wrapperEl.createEl("button", {
      cls: "iris-search-icon clickable-icon",
      attr: { "aria-label": "Search" },
    });
    setIcon(iconBtn, "search");

    this.inputEl = this.wrapperEl.createEl("input", {
      type: "text",
      placeholder: "Search…",
      cls: "iris-search-input",
    });

    // Expand on icon click
    iconBtn.addEventListener("click", () => {
      this.expand();
      this.inputEl.focus();
    });

    // Expand on hover
    this.wrapperEl.addEventListener("mouseenter", () => this.expand());

    // Collapse on mouse leave if empty and not focused
    this.wrapperEl.addEventListener("mouseleave", () => {
      if (!this.inputEl.value && document.activeElement !== this.inputEl) {
        this.collapse();
      }
    });

    // Collapse on blur if empty
    this.inputEl.addEventListener("blur", () => {
      if (!this.inputEl.value) {
        this.collapse();
      }
    });

    // Debounced search
    let debounceTimer: number | null = null;
    this.inputEl.addEventListener("input", () => {
      if (debounceTimer) window.clearTimeout(debounceTimer);
      debounceTimer = window.setTimeout(() => {
        callbacks.onSearch(this.inputEl.value);
      }, 500);
    });

    this.inputEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        if (debounceTimer) window.clearTimeout(debounceTimer);
        callbacks.onSearch(this.inputEl.value);
      }
      if (e.key === "Escape") {
        this.inputEl.value = "";
        callbacks.onSearch("");
        this.inputEl.blur();
      }
    });
  }

  clear(): void {
    this.inputEl.value = "";
    this.collapse();
  }

  private expand(): void {
    this.wrapperEl.addClass("is-expanded");
  }

  private collapse(): void {
    this.wrapperEl.removeClass("is-expanded");
  }
}
