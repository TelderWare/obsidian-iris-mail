import { App, Modal, Notice, setIcon } from "obsidian";
import { IconPickerModal } from "./IconPickerModal";

export const TAG_COLOR_PALETTE: Array<[string, string]> = [
  ["Red", "#e03131"],
  ["Orange", "#f76707"],
  ["Yellow", "#f59f00"],
  ["Green", "#2f9e44"],
  ["Teal", "#0ca678"],
  ["Blue", "#1971c2"],
  ["Purple", "#7950f2"],
  ["Pink", "#d6336c"],
];

export interface TagModalInitial {
  name: string;
  criteria: string;
  icon: string;
  /** Current color hex, or empty string for default. */
  color?: string;
  /** Current contradictions (other tag names mutually exclusive with this one). */
  contradicts?: string[];
  /** Tags this one precludes (directional: if this fires, skip those). */
  precludes?: string[];
  /** Tags that preclude this one (directional: if any of those fires, skip this). */
  precludedBy?: string[];
}

export interface TagModalOptions {
  /** Existing tag names (used for duplicate check in create mode only). */
  existingTags: string[];
  /** If present, the modal opens in edit mode with these initial values; the name is locked. */
  initial?: TagModalInitial;
  /** Called on Save. `iconExplicit` is true when the user changed the icon manually. */
  onSubmit: (
    name: string,
    criteria: string,
    icon: string,
    iconExplicit: boolean,
    color: string,
    contradicts: string[],
    precludes: string[],
    precludedBy: string[],
  ) => void;
  /** If provided, enables the criteria "Auto-generate" button. */
  onGenerate?: (name: string) => Promise<string>;
}

/**
 * Modal for creating or editing a tag. Fields: icon, name, color, criteria.
 * In edit mode the name is read-only.
 */
export class CreateTagModal extends Modal {
  private opts: TagModalOptions;

  constructor(app: App, opts: TagModalOptions) {
    super(app);
    this.opts = opts;
  }

  onOpen(): void {
    const { contentEl } = this;
    const isEdit = !!this.opts.initial;
    contentEl.empty();
    contentEl.addClass("iris-create-tag-modal");
    contentEl.createEl("h3", { text: isEdit ? "Edit tag" : "New tag" });

    let currentIcon = this.opts.initial?.icon || "tag";
    let iconExplicit = false;
    let currentColor = this.opts.initial?.color || "";

    // Icon + Name row
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
        iconExplicit = true;
        iconBtn.empty();
        setIcon(iconBtn, currentIcon);
        applyIconColor();
      }).open();
    });

    const nameInput = nameRow.createEl("input", {
      cls: "iris-create-tag-input",
      attr: { type: "text", placeholder: "Name" },
    });
    if (isEdit) {
      nameInput.value = this.opts.initial!.name;
      nameInput.readOnly = true;
      nameInput.addClass("is-locked");
    }

    // Color swatches
    const colorField = contentEl.createDiv({ cls: "iris-create-tag-field" });
    colorField.createEl("label", { text: "Color", cls: "iris-create-tag-label" });
    const swatchRow = colorField.createDiv({ cls: "iris-create-tag-swatches" });

    const swatches: HTMLElement[] = [];
    const markActive = () => {
      for (const s of swatches) {
        const hex = s.dataset.color || "";
        s.toggleClass(
          "is-active",
          hex.toLowerCase() === currentColor.toLowerCase(),
        );
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

    // Contradictions
    const otherTags = this.opts.existingTags.filter((t) => t !== this.opts.initial?.name);
    const contradicts = new Set(this.opts.initial?.contradicts || []);
    if (otherTags.length > 0) {
      const contraField = contentEl.createDiv({ cls: "iris-create-tag-field" });
      contraField.createEl("label", {
        text: "Contradicts",
        cls: "iris-create-tag-label",
        attr: { title: "Selected tags are mutually exclusive with this one — the classifier skips their API calls once either is confirmed." },
      });
      const chipRow = contraField.createDiv({ cls: "iris-create-tag-chips" });
      for (const other of otherTags) {
        const chip = chipRow.createEl("button", {
          cls: "iris-create-tag-chip" + (contradicts.has(other) ? " is-active" : ""),
          text: other,
          attr: { type: "button" },
        });
        chip.addEventListener("click", () => {
          if (contradicts.has(other)) {
            contradicts.delete(other);
            chip.removeClass("is-active");
          } else {
            contradicts.add(other);
            chip.addClass("is-active");
          }
        });
      }
    }

    // Criteria
    const critField = contentEl.createDiv({ cls: "iris-create-tag-field" });
    const critHeader = critField.createDiv({ cls: "iris-create-tag-crit-header" });
    critHeader.createEl("label", { text: "Criteria", cls: "iris-create-tag-label" });

    let genBtn: HTMLButtonElement | null = null;
    if (this.opts.onGenerate) {
      genBtn = critHeader.createEl("button", {
        cls: "iris-create-tag-generate mod-muted",
        text: "Auto-generate",
      });
    }

    const critArea = critField.createEl("textarea", {
      cls: "iris-create-tag-textarea",
      attr: {
        rows: "4",
        placeholder: "Criteria",
      },
    });
    if (this.opts.initial) critArea.value = this.opts.initial.criteria;

    if (genBtn) {
      genBtn.addEventListener("click", async () => {
        const trimmed = nameInput.value.trim();
        if (!trimmed) {
          new Notice("Enter a tag name first");
          nameInput.focus();
          return;
        }
        const prev = genBtn!.textContent;
        genBtn!.disabled = true;
        genBtn!.textContent = "Generating…";
        try {
          critArea.value = await this.opts.onGenerate!(trimmed);
        } catch (err) {
          new Notice(`Generation failed: ${err instanceof Error ? err.message : String(err)}`);
        } finally {
          genBtn!.disabled = false;
          genBtn!.textContent = prev || "Auto-generate";
        }
      });
    }

    // Precludes / Precluded by (directional skip relationships)
    const precludes = new Set(this.opts.initial?.precludes || []);
    const precludedBy = new Set(this.opts.initial?.precludedBy || []);
    if (otherTags.length > 0) {
      this.renderChipField(
        contentEl,
        "Precludes",
        "If this tag is confirmed on a message, the classifier skips these tags.",
        otherTags,
        precludes,
      );
      this.renderChipField(
        contentEl,
        "Precluded by",
        "If any of these tags is confirmed on a message, the classifier skips this one.",
        otherTags,
        precludedBy,
      );
    }

    // Footer
    const footer = contentEl.createDiv({ cls: "iris-create-tag-footer" });
    const submitBtn = footer.createEl("button", {
      cls: "mod-cta",
      text: isEdit ? "Save" : "Create",
    });
    const cancelBtn = footer.createEl("button", { text: "Cancel" });

    submitBtn.addEventListener("click", () => {
      const name = nameInput.value.trim();
      const criteria = critArea.value.trim();
      if (!name) {
        new Notice("Tag name is required");
        nameInput.focus();
        return;
      }
      if (!isEdit && this.opts.existingTags.includes(name)) {
        new Notice(`Tag "${name}" already exists`);
        nameInput.focus();
        return;
      }
      if (!criteria) {
        new Notice("Criteria is required");
        critArea.focus();
        return;
      }
      this.opts.onSubmit(
        name,
        criteria,
        currentIcon,
        iconExplicit,
        currentColor,
        Array.from(contradicts),
        Array.from(precludes),
        Array.from(precludedBy),
      );
      this.close();
    });

    cancelBtn.addEventListener("click", () => this.close());

    setTimeout(() => (isEdit ? critArea : nameInput).focus(), 10);
  }

  onClose(): void {
    this.contentEl.empty();
  }

  private renderChipField(
    parent: HTMLElement,
    label: string,
    tooltip: string,
    options: string[],
    selected: Set<string>,
  ): void {
    const field = parent.createDiv({ cls: "iris-create-tag-field" });
    field.createEl("label", {
      text: label,
      cls: "iris-create-tag-label",
      attr: { title: tooltip },
    });
    const row = field.createDiv({ cls: "iris-create-tag-chips" });
    for (const opt of options) {
      const chip = row.createEl("button", {
        cls: "iris-create-tag-chip" + (selected.has(opt) ? " is-active" : ""),
        text: opt,
        attr: { type: "button" },
      });
      chip.addEventListener("click", () => {
        if (selected.has(opt)) {
          selected.delete(opt);
          chip.removeClass("is-active");
        } else {
          selected.add(opt);
          chip.addClass("is-active");
        }
      });
    }
  }
}
