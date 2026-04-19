import { App, Modal, Notice, setIcon } from "obsidian";
import { IconPickerModal } from "./IconPickerModal";

export interface TagModalInitial {
  name: string;
  criteria: string;
  icon: string;
}

export interface TagModalOptions {
  /** Existing tag names (used for duplicate check in create mode only). */
  existingTags: string[];
  /** If present, the modal opens in edit mode with these initial values; the name is locked. */
  initial?: TagModalInitial;
  /** Called on Save. `iconExplicit` is true when the user changed the icon manually. */
  onSubmit: (name: string, criteria: string, icon: string, iconExplicit: boolean) => void;
  /** If provided, enables the criteria "Auto-generate" button. */
  onGenerate?: (name: string) => Promise<string>;
}

/**
 * Modal for creating or editing a tag. Fields: icon, name, criteria.
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

    // Icon + Name row
    const nameRow = contentEl.createDiv({ cls: "iris-create-tag-field iris-create-tag-name-row" });
    const iconBtn = nameRow.createEl("button", {
      cls: "iris-create-tag-icon clickable-icon",
      attr: { "aria-label": "Pick icon" },
    });
    setIcon(iconBtn, currentIcon);
    iconBtn.addEventListener("click", () => {
      new IconPickerModal(this.app, currentIcon, (picked) => {
        currentIcon = picked;
        iconExplicit = true;
        iconBtn.empty();
        setIcon(iconBtn, currentIcon);
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
      this.opts.onSubmit(name, criteria, currentIcon, iconExplicit);
      this.close();
    });

    cancelBtn.addEventListener("click", () => this.close());

    setTimeout(() => (isEdit ? critArea : nameInput).focus(), 10);
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
