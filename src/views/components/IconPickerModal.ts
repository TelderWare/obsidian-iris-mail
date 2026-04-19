import { App, Modal, getIconIds, setIcon } from "obsidian";

export class IconPickerModal extends Modal {
  private currentIcon: string;
  private onPick: (icon: string) => void;

  constructor(app: App, currentIcon: string, onPick: (icon: string) => void) {
    super(app);
    this.currentIcon = currentIcon;
    this.onPick = onPick;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass("iris-icon-picker-modal");
    contentEl.createEl("h3", { text: "Pick an icon" });

    const searchInput = contentEl.createEl("input", {
      cls: "iris-icon-picker-search",
      attr: { type: "text", placeholder: "Search icons…" },
    });

    const gridEl = contentEl.createDiv({ cls: "iris-icon-picker-grid" });

    const allIcons = getIconIds().sort();
    const render = (filter: string) => {
      gridEl.empty();
      const q = filter.trim().toLowerCase();
      const matches = q
        ? allIcons.filter((id) => id.toLowerCase().includes(q))
        : allIcons;
      for (const id of matches.slice(0, 400)) {
        const btn = gridEl.createEl("button", {
          cls: "iris-icon-picker-cell clickable-icon" +
            (id === this.currentIcon ? " is-active" : ""),
          attr: { "aria-label": id, title: id },
        });
        setIcon(btn, id);
        btn.addEventListener("click", () => {
          this.onPick(id);
          this.close();
        });
      }
      if (matches.length > 400) {
        gridEl.createDiv({
          cls: "iris-icon-picker-overflow",
          text: `Showing 400 of ${matches.length} — refine your search for more.`,
        });
      }
      if (matches.length === 0) {
        gridEl.createDiv({ cls: "iris-icon-picker-empty", text: "No matches." });
      }
    };

    render("");
    searchInput.addEventListener("input", () => render(searchInput.value));
    setTimeout(() => searchInput.focus(), 10);
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
