import { App, Modal, Setting } from "obsidian";

export class NicknameModal extends Modal {
  private address: string;
  private rawName: string;
  private currentNickname: string;
  private onSave: (address: string, nickname: string) => void;
  private onDelete: (address: string) => void;
  private onRegenerate?: () => Promise<string>;

  constructor(
    app: App,
    address: string,
    rawName: string,
    currentNickname: string,
    onSave: (address: string, nickname: string) => void,
    onDelete: (address: string) => void,
    onRegenerate?: () => Promise<string>,
  ) {
    super(app);
    this.address = address;
    this.rawName = rawName;
    this.currentNickname = currentNickname;
    this.onSave = onSave;
    this.onDelete = onDelete;
    this.onRegenerate = onRegenerate;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h3", { text: "Edit Nickname" });
    contentEl.createEl("p", {
      text: this.address,
      cls: "iris-nickname-modal-address",
    });

    let value = this.currentNickname;
    let textInput: HTMLInputElement | null = null;

    new Setting(contentEl)
      .setName("Nickname")
      .addText((text) => {
        textInput = text.inputEl;
        text
          .setPlaceholder(this.rawName)
          .setValue(this.currentNickname)
          .onChange((v) => { value = v; });
        text.inputEl.addEventListener("keydown", (e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            this.onSave(this.address, value.trim());
            this.close();
          }
        });
        // Auto-focus the input
        setTimeout(() => text.inputEl.focus(), 10);
      });

    const buttonRow = new Setting(contentEl)
      .addButton((btn) =>
        btn
          .setButtonText("Save")
          .setCta()
          .onClick(() => {
            this.onSave(this.address, value.trim());
            this.close();
          }),
      );

    if (this.onRegenerate) {
      buttonRow.addButton((btn) => {
        btn.setButtonText("Regenerate").onClick(async () => {
          if (!this.onRegenerate || !textInput) return;
          const prev = btn.buttonEl.textContent;
          btn.setDisabled(true).setButtonText("…");
          try {
            const next = await this.onRegenerate();
            value = next;
            textInput.value = next;
          } catch {
            // Keep modal open; user can retry or cancel.
          } finally {
            btn.setDisabled(false).setButtonText(prev || "Regenerate");
          }
        });
      });
    }

    buttonRow
      .addButton((btn) =>
        btn
          .setButtonText("Delete")
          .setWarning()
          .onClick(() => {
            this.onDelete(this.address);
            this.close();
          }),
      )
      .addButton((btn) =>
        btn.setButtonText("Cancel").onClick(() => this.close()),
      );
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
