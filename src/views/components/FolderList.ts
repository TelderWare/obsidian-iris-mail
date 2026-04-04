import { setIcon } from "obsidian";
import type { MailFolder } from "../../types";
import { WELL_KNOWN_FOLDERS } from "../../constants";

interface FolderListCallbacks {
  onFolderSelect: (folder: MailFolder) => void;
}

const FOLDER_ICON_MAP: Record<string, string> = {
  Inbox: "inbox",
  "Sent Items": "send",
  Drafts: "file-edit",
  "Deleted Items": "trash-2",
  "Junk Email": "shield-x",
  Archive: "archive",
};

export class FolderList {
  private containerEl: HTMLElement;
  private callbacks: FolderListCallbacks;
  private selectedId: string | null = null;
  private folders: MailFolder[] = [];

  constructor(containerEl: HTMLElement, callbacks: FolderListCallbacks) {
    this.containerEl = containerEl;
    this.callbacks = callbacks;
  }

  render(folders: MailFolder[]): void {
    this.folders = folders;
    this.containerEl.empty();

    const sorted = this.sortFolders(
      folders.filter((f) => f.displayName && f.displayName in FOLDER_ICON_MAP)
    );

    for (const folder of sorted) {
      const name = folder.displayName || "";
      const iconName = FOLDER_ICON_MAP[name] || "folder";

      const btn = this.containerEl.createEl("button", {
        cls:
          "iris-folder-icon clickable-icon" +
          (folder.id === this.selectedId ? " is-selected" : ""),
        attr: { "aria-label": name },
      });

      setIcon(btn, iconName);

      if (folder.unreadItemCount && folder.unreadItemCount > 0) {
        btn.createSpan({
          cls: "iris-folder-badge",
          text: String(folder.unreadItemCount),
        });
      }

      btn.addEventListener("click", () => {
        this.selectedId = folder.id || null;
        this.render(this.folders);
        this.callbacks.onFolderSelect(folder);
      });
    }
  }

  private sortFolders(folders: MailFolder[]): MailFolder[] {
    const order = new Map(WELL_KNOWN_FOLDERS.map((name, i) => [name, i]));
    return [...folders].sort((a, b) => {
      const aOrder = order.get(a.displayName || "") ?? 999;
      const bOrder = order.get(b.displayName || "") ?? 999;
      if (aOrder !== bOrder) return aOrder - bOrder;
      return (a.displayName || "").localeCompare(b.displayName || "");
    });
  }
}
