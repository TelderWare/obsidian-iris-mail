import { App, Modal, setIcon } from "obsidian";
import type IrisMailPlugin from "../../main";
import {
  parseTagCategories,
  setTagContradictions,
  removeTagFromContradictions,
  setTagPrecludesList,
  setPrecludedByFor,
  getPrecludedBy,
  removeTagFromPrecludes,
  TAG_ICON_POOL,
} from "../../constants";
import { CreateTagModal } from "./CreateTagModal";
import { generateTagDescription, hasClaudeAccess, pickTagIcon } from "../../utils/claudeApi";

export interface TagsModalOptions {
  onChange?: () => void;
}

/**
 * Quick-access tag manager. Lists all user-defined tags with icon/color
 * swatches, supports click-to-edit (via CreateTagModal), delete, and create.
 * Mirrors the Settings tab's tag CRUD without routing through the settings
 * page rebuild.
 */
export class TagsModal extends Modal {
  private plugin: IrisMailPlugin;
  private opts: TagsModalOptions;
  private listEl!: HTMLElement;

  constructor(app: App, plugin: IrisMailPlugin, opts: TagsModalOptions = {}) {
    super(app);
    this.plugin = plugin;
    this.opts = opts;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass("iris-tags-modal");

    const header = contentEl.createDiv({ cls: "iris-tags-modal-header" });
    header.createEl("h3", { text: "Tags" });
    const addBtn = header.createEl("button", {
      cls: "mod-cta",
      text: "New tag",
    });
    addBtn.addEventListener("click", () => this.openCreate());

    this.listEl = contentEl.createDiv({ cls: "iris-tags-modal-list" });
    this.renderList();
  }

  onClose(): void {
    this.contentEl.empty();
  }

  private renderList(): void {
    this.listEl.empty();
    const s = this.plugin.settings;
    if (!s.tagIcons) s.tagIcons = {};
    if (!s.tagDescriptions) s.tagDescriptions = {};

    const categories = parseTagCategories(s.tagCategories);
    if (categories.length === 0) {
      this.listEl.createDiv({
        cls: "iris-tags-modal-empty",
        text: "No tags yet. Click New tag to add one.",
      });
      return;
    }

    for (const cat of categories) {
      const icon = s.tagIcons[cat] || "tag";
      const color = s.tagColors?.[cat] || "";
      const needsCriteria = !(s.tagDescriptions[cat] || "").trim();

      const row = this.listEl.createDiv({
        cls: "iris-tags-modal-item is-clickable" + (needsCriteria ? " needs-description" : ""),
        attr: { role: "button", tabindex: "0", "aria-label": `Edit tag "${cat}"` },
      });

      const iconEl = row.createSpan({ cls: "iris-tags-modal-icon" });
      setIcon(iconEl, icon);
      if (color) iconEl.style.color = color;

      row.createSpan({ cls: "iris-tags-modal-name", text: cat });

      if (needsCriteria) {
        const warn = row.createSpan({
          cls: "iris-tags-modal-warn",
          attr: { "aria-label": "Missing criteria — classifier accuracy will suffer" },
        });
        setIcon(warn, "alert-triangle");
      }

      const deleteBtn = row.createEl("button", {
        cls: "iris-tags-modal-delete clickable-icon",
        attr: { "aria-label": `Delete tag "${cat}"` },
      });
      setIcon(deleteBtn, "trash-2");
      deleteBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.deleteTag(cat);
      });

      row.addEventListener("click", () => this.openEdit(cat));
      row.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          this.openEdit(cat);
        }
      });
    }
  }

  /** Write all per-tag fields to settings (shared between create and edit). */
  private writeTagFields(
    cat: string,
    criteria: string,
    icon: string,
    color: string,
    contradicts: string[],
    precludes: string[],
    precludedBy: string[],
  ): void {
    const s = this.plugin.settings;
    if (!s.tagDescriptions) s.tagDescriptions = {};
    if (!s.tagIcons) s.tagIcons = {};
    if (!s.tagColors) s.tagColors = {};
    if (!s.tagContradictions) s.tagContradictions = {};
    if (!s.tagPrecludes) s.tagPrecludes = {};
    s.tagDescriptions[cat] = criteria;
    s.tagIcons[cat] = icon;
    if (color) s.tagColors[cat] = color;
    else delete s.tagColors[cat];
    setTagContradictions(s.tagContradictions, cat, contradicts);
    setTagPrecludesList(s.tagPrecludes, cat, precludes);
    setPrecludedByFor(s.tagPrecludes, cat, precludedBy);
  }

  /** Build the `onGenerate` callback used by CreateTagModal, or undefined if Claude isn't available. */
  private buildGenerate(excludeName?: string): ((name: string) => Promise<string>) | undefined {
    const s = this.plugin.settings;
    if (!s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return undefined;
    return (name) => {
      const others = parseTagCategories(s.tagCategories)
        .filter((n) => n !== (excludeName ?? name))
        .map((n) => ({ name: n, description: s.tagDescriptions?.[n] || "" }));
      return generateTagDescription(s.anthropicApiKey, s.claudeModel, name, others);
    };
  }

  private openCreate(): void {
    const s = this.plugin.settings;
    new CreateTagModal(this.app, {
      existingTags: parseTagCategories(s.tagCategories),
      onGenerate: this.buildGenerate(),
      onSubmit: async (name, criteria, icon, iconExplicit, color, contradicts, precludes, precludedBy) => {
        s.tagCategories = [...parseTagCategories(s.tagCategories), name].join(", ");
        const usedIcons = Object.values(s.tagIcons || {});
        this.writeTagFields(name, criteria, icon, color, contradicts, precludes, precludedBy);
        this.persist();
        this.renderList();

        // When the user didn't pick an icon explicitly, ask Claude to pick one
        // that doesn't collide with existing tags. Falls back silently.
        if (!iconExplicit && s.enableClaudeProcessing && hasClaudeAccess(s.anthropicApiKey)) {
          try {
            const picked = await pickTagIcon(
              s.anthropicApiKey, s.claudeModel, name, criteria,
              TAG_ICON_POOL, usedIcons,
            );
            if (picked) {
              s.tagIcons![name] = picked;
              this.persist();
              this.renderList();
            }
          } catch {
            // Keep fallback icon.
          }
        }
      },
    }).open();
  }

  private openEdit(cat: string): void {
    const s = this.plugin.settings;
    new CreateTagModal(this.app, {
      existingTags: parseTagCategories(s.tagCategories),
      initial: {
        name: cat,
        criteria: s.tagDescriptions?.[cat] || "",
        icon: s.tagIcons?.[cat] || "tag",
        color: s.tagColors?.[cat] || "",
        contradicts: s.tagContradictions?.[cat] || [],
        precludes: s.tagPrecludes?.[cat] || [],
        precludedBy: getPrecludedBy(s.tagPrecludes || {}, cat),
      },
      onGenerate: this.buildGenerate(cat),
      onSubmit: (_name, criteria, icon, _iconExplicit, color, contradicts, precludes, precludedBy) => {
        this.writeTagFields(cat, criteria, icon, color, contradicts, precludes, precludedBy);
        this.persist();
        this.renderList();
      },
    }).open();
  }

  private deleteTag(cat: string): void {
    const s = this.plugin.settings;
    const remaining = parseTagCategories(s.tagCategories).filter((n) => n !== cat);
    s.tagCategories = remaining.join(", ");
    if (s.tagIcons) delete s.tagIcons[cat];
    if (s.tagDescriptions) delete s.tagDescriptions[cat];
    if (s.tagColors) delete s.tagColors[cat];
    if (s.tagPromptVersions) delete s.tagPromptVersions[cat];
    if (s.tagContradictions) removeTagFromContradictions(s.tagContradictions, cat);
    if (s.tagPrecludes) removeTagFromPrecludes(s.tagPrecludes, cat);
    this.persist();
    this.renderList();
  }

  private persist(): void {
    this.plugin.scheduleSaveSettings();
    this.opts.onChange?.();
  }
}
