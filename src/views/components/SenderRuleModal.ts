import { App, Modal, Setting } from "obsidian";
import type { SenderRule } from "../../types";

export class SenderRuleModal extends Modal {
  private readonly address: string;
  private readonly displayName: string;
  private readonly tagCategories: string[];
  private readonly initial: SenderRule;
  private readonly onSave: (address: string, rule: SenderRule) => void;
  private readonly onRemove: (address: string) => void;

  constructor(
    app: App,
    address: string,
    displayName: string,
    tagCategories: string[],
    current: SenderRule | undefined,
    onSave: (address: string, rule: SenderRule) => void,
    onRemove: (address: string) => void,
  ) {
    super(app);
    this.address = address;
    this.displayName = displayName;
    this.tagCategories = tagCategories;
    this.initial = current ?? {};
    this.onSave = onSave;
    this.onRemove = onRemove;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();

    contentEl.createEl("h3", { text: "Sender rule" });
    contentEl.createEl("p", {
      text: this.displayName && this.displayName !== this.address
        ? `${this.displayName} <${this.address}>`
        : this.address,
      cls: "iris-nickname-modal-address",
    });

    let autoBin = !!this.initial.autoBin;
    let autoTag = this.initial.autoTag ?? "";

    new Setting(contentEl)
      .setName("Always bin messages from this sender")
      .setDesc("New messages from this sender are moved to the provider's trash folder on arrival.")
      .addToggle((toggle) =>
        toggle
          .setValue(autoBin)
          .onChange((v) => { autoBin = v; }),
      );

    const tagSetting = new Setting(contentEl)
      .setName("Always apply tag")
      .setDesc(
        this.tagCategories.length === 0
          ? "No tags defined yet. Add tags in settings first."
          : "Apply this tag to new messages from this sender.",
      );

    if (this.tagCategories.length > 0) {
      tagSetting.addDropdown((drop) => {
        drop.addOption("", "(none)");
        for (const tag of this.tagCategories) drop.addOption(tag, tag);
        drop.setValue(autoTag);
        drop.onChange((v) => { autoTag = v; });
      });
    }

    const buttons = new Setting(contentEl);

    buttons.addButton((btn) =>
      btn
        .setButtonText("Save")
        .setCta()
        .onClick(() => {
          const rule: SenderRule = {};
          if (autoBin) rule.autoBin = true;
          if (autoTag) rule.autoTag = autoTag;

          if (!rule.autoBin && !rule.autoTag) {
            // Empty rule → remove any existing rule.
            this.onRemove(this.address);
          } else {
            this.onSave(this.address, rule);
          }
          this.close();
        }),
    );

    if (this.initial.autoBin || this.initial.autoTag) {
      buttons.addButton((btn) =>
        btn
          .setButtonText("Remove rule")
          .setWarning()
          .onClick(() => {
            this.onRemove(this.address);
            this.close();
          }),
      );
    }

    buttons.addButton((btn) =>
      btn.setButtonText("Cancel").onClick(() => this.close()),
    );
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
