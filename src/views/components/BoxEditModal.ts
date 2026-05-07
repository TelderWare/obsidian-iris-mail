import { App, Modal, Notice, setIcon } from "obsidian";
import { IconPickerModal } from "./IconPickerModal";
import type { Box } from "../../types";
import { TAG_COLOR_PALETTE } from "./CreateTagModal";

export interface BoxEditModalOptions {
  /** Present when editing an existing box. Absent when creating a new one. */
  initial?: Box;
  /** IDs already in use — used only when creating, to prevent collisions. */
  existingIds: Set<string>;
  /** All user-defined tag names, shown as chips to select into the box predicate. */
  availableTags: string[];
  onSubmit: (draft: Box) => void;
}

/**
 * Modal for creating / editing a Box (a named view over the message list).
 * For built-in boxes the name/icon/color/tag-predicate are editable but the
 * `builtin` marker is preserved so the view's predicate logic still applies.
 */
export class BoxEditModal extends Modal {
  private opts: BoxEditModalOptions;

  constructor(app: App, opts: BoxEditModalOptions) {
    super(app);
    this.opts = opts;
  }

  onOpen(): void {
    const { contentEl } = this;
    const isEdit = !!this.opts.initial;
    contentEl.empty();
    contentEl.addClass("iris-create-tag-modal");
    contentEl.createEl("h3", { text: isEdit ? "Edit box" : "New box" });

    let currentIcon = this.opts.initial?.icon || "inbox";
    let currentColor = this.opts.initial?.color || "";

    const nameRow = contentEl.createDiv({ cls: "iris-create-tag-field iris-create-tag-name-row" });
    const iconBtn = nameRow.createEl("button", {
      cls: "iris-create-tag-icon clickable-icon",
      attr: { "aria-label": "Pick icon" },
    });
    setIcon(iconBtn, currentIcon);
    const applyIconColor = () => {
      iconBtn.style.color = currentColor || "";
    };
    applyIconColor();
    iconBtn.addEventListener("click", () => {
      new IconPickerModal(this.app, currentIcon, (picked) => {
        currentIcon = picked;
        iconBtn.empty();
        setIcon(iconBtn, currentIcon);
        applyIconColor();
      }).open();
    });

    const nameInput = nameRow.createEl("input", {
      cls: "iris-create-tag-input",
      attr: { type: "text", placeholder: "Name" },
    });
    if (this.opts.initial) nameInput.value = this.opts.initial.name;

    const colorField = contentEl.createDiv({ cls: "iris-create-tag-field" });
    colorField.createEl("label", { text: "Color", cls: "iris-create-tag-label" });
    const swatchRow = colorField.createDiv({ cls: "iris-create-tag-swatches" });
    const swatches: HTMLElement[] = [];
    const markActive = () => {
      for (const s of swatches) {
        const hex = s.dataset.color || "";
        s.toggleClass("is-active", hex.toLowerCase() === currentColor.toLowerCase());
      }
    };
    const addSwatch = (hex: string, label: string) => {
      const el = swatchRow.createEl("button", {
        cls: "iris-create-tag-swatch",
        attr: { type: "button", "aria-label": label, title: label },
      });
      el.dataset.color = hex;
      if (hex) {
        el.style.background = hex;
      } else {
        el.addClass("is-default");
        setIcon(el, "circle-slash");
      }
      el.addEventListener("click", () => {
        currentColor = hex;
        markActive();
        applyIconColor();
      });
      swatches.push(el);
    };
    addSwatch("", "Default");
    for (const [name, hex] of TAG_COLOR_PALETTE) addSwatch(hex, name);
    markActive();

    // Secretary's predicate is purely in-flight state (messages currently
    // being classified/summarized) — tags on it are ignored, so don't offer
    // a tag selector that would silently do nothing.
    const usesTagPredicate = this.opts.initial?.builtin !== "secretary";

    // "Save" only makes sense for boxes whose predicate is driven by
    // locally-tracked state we can evaluate without a fresh server fetch.
    const supportsSaved =
      !this.opts.initial?.builtin ||
      this.opts.initial.builtin === "todo" ||
      this.opts.initial.builtin === "junk";
    let saved = !!this.opts.initial?.saved;
    const selectedTags = new Set(this.opts.initial?.tags || []);
    if (usesTagPredicate && this.opts.availableTags.length > 0) {
      const tagField = contentEl.createDiv({ cls: "iris-create-tag-field" });
      tagField.createEl("label", {
        text: "Tags",
        cls: "iris-create-tag-label",
        attr: { title: "Messages carrying any selected tag are included in this box." },
      });
      const chipRow = tagField.createDiv({ cls: "iris-create-tag-chips" });
      for (const tag of this.opts.availableTags) {
        const chip = chipRow.createEl("button", {
          cls: "iris-create-tag-chip" + (selectedTags.has(tag) ? " is-active" : ""),
          text: tag,
          attr: { type: "button" },
        });
        chip.addEventListener("click", () => {
          if (selectedTags.has(tag)) {
            selectedTags.delete(tag);
            chip.removeClass("is-active");
          } else {
            selectedTags.add(tag);
            chip.addClass("is-active");
          }
        });
      }
    }

    if (supportsSaved) {
      const savedField = contentEl.createDiv({ cls: "iris-create-tag-field iris-box-saved-field" });
      const savedLabel = savedField.createEl("label", { cls: "iris-box-saved-label" });
      const savedInput = savedLabel.createEl("input", { attr: { type: "checkbox" } });
      savedInput.checked = saved;
      savedLabel.createSpan({ text: "Save messages past sync window" });
      savedField.createEl("div", {
        cls: "iris-box-saved-help",
        text: "Keep these messages visible even after they age out of the server's sync window.",
      });
      savedInput.addEventListener("change", () => { saved = savedInput.checked; });
    }

    const footer = contentEl.createDiv({ cls: "iris-create-tag-footer" });
    const submitBtn = footer.createEl("button", {
      cls: "mod-cta",
      text: isEdit ? "Save" : "Create",
    });
    const cancelBtn = footer.createEl("button", { text: "Cancel" });

    submitBtn.addEventListener("click", () => {
      const name = nameInput.value.trim();
      if (!name) {
        new Notice("Box name is required");
        nameInput.focus();
        return;
      }
      const draft: Box = {
        id: this.opts.initial?.id || this.generateId(name),
        name,
        icon: currentIcon,
        color: currentColor || undefined,
        builtin: this.opts.initial?.builtin,
        tags: usesTagPredicate && selectedTags.size > 0 ? Array.from(selectedTags) : this.opts.initial?.tags,
        saved: supportsSaved ? saved : this.opts.initial?.saved,
      };
      this.opts.onSubmit(draft);
      this.close();
    });

    cancelBtn.addEventListener("click", () => this.close());

    setTimeout(() => nameInput.focus(), 10);
  }

  onClose(): void {
    this.contentEl.empty();
  }

  private generateId(name: string): string {
    const base = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "box";
    let id = base;
    let i = 2;
    while (this.opts.existingIds.has(id)) id = `${base}-${i++}`;
    return id;
  }
}
