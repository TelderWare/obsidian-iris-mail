import { App, PluginSettingTab, Setting, setIcon } from "obsidian";
import type IrisMailPlugin from "../main";
import { DEFAULT_CLAUDE_PROMPT, TAG_CLASSIFY_PROMPT, TAG_ICON_POOL, ITEM_DETECTION_PROMPT, parseTagCategories, bumpTagVersion } from "../constants";
import { CreateTagModal } from "../views/components/CreateTagModal";
import { generateTagDescription, hasClaudeAccess, pickTagIcon } from "../utils/claudeApi";
import { setDebugEnabled } from "../utils/logger";

export class IrisMailSettingsTab extends PluginSettingTab {
  plugin: IrisMailPlugin;

  constructor(app: App, plugin: IrisMailPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    containerEl.createEl("h2", { text: "Iris Mail Settings" });

    new Setting(containerEl)
      .setName("Azure Client ID")
      .setDesc(
        "Application (client) ID from your Azure Entra app registration. " +
          "Register at portal.azure.com > App registrations > New registration. " +
          "Set platform to 'Public client' with redirect URI http://localhost:{port}/redirect.",
      )
      .addText((text) =>
        text
          .setPlaceholder("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
          .setValue(this.plugin.settings.clientId)
          .onChange(async (value) => {
            this.plugin.settings.clientId = value.trim();
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Authority")
      .setDesc(
        "Azure authority URL. Use 'common' for any account, or your tenant ID for org-only.",
      )
      .addText((text) =>
        text
          .setPlaceholder("https://login.microsoftonline.com/common")
          .setValue(this.plugin.settings.authority)
          .onChange(async (value) => {
            this.plugin.settings.authority = value.trim();
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Sign-in method")
      .setDesc(
        "Browser redirect opens a browser window. Device code lets you sign in on any device.",
      )
      .addDropdown((dropdown) =>
        dropdown
          .addOptions({
            "auth-code": "Browser redirect",
            "device-code": "Device code",
          })
          .setValue(this.plugin.settings.authMethod)
          .onChange(async (value) => {
            this.plugin.settings.authMethod = value as "auth-code" | "device-code";
            this.plugin.scheduleSaveSettings();
            this.display(); // re-render to show/hide redirect port
          }),
      );

    if (this.plugin.settings.authMethod === "auth-code") {
      new Setting(containerEl)
        .setName("Redirect port")
        .setDesc(
          "Localhost port for OAuth redirect. Must match the redirect URI in your Azure app.",
        )
        .addText((text) =>
          text
            .setPlaceholder("3847")
            .setValue(String(this.plugin.settings.redirectPort))
            .onChange(async (value) => {
              const port = parseInt(value, 10);
              if (!isNaN(port) && port > 0 && port <= 65535) {
                this.plugin.settings.redirectPort = port;
                this.plugin.scheduleSaveSettings();
              }
            }),
        );
    }

    new Setting(containerEl)
      .setName("Auto-refresh interval (minutes)")
      .setDesc("How often to check for new emails. Set to 0 to disable.")
      .addText((text) =>
        text
          .setPlaceholder("5")
          .setValue(String(this.plugin.settings.refreshIntervalMinutes))
          .onChange(async (value) => {
            const mins = parseInt(value, 10);
            if (!isNaN(mins) && mins >= 0) {
              this.plugin.settings.refreshIntervalMinutes = mins;
              this.plugin.scheduleSaveSettings();
            }
          }),
      );

    new Setting(containerEl)
      .setName("Messages per page")
      .addDropdown((dropdown) =>
        dropdown
          .addOptions({ "10": "10", "25": "25", "50": "50" })
          .setValue(String(this.plugin.settings.pageSize))
          .onChange(async (value) => {
            this.plugin.settings.pageSize = parseInt(value, 10);
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Show read emails")
      .setDesc("If disabled, only unread messages are shown.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showReadEmails)
          .onChange(async (value) => {
            this.plugin.settings.showReadEmails = value;
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Resolve forwarded sender")
      .setDesc(
        "Show the original sender of forwarded emails instead of the forwarder.",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.resolveForwardedSender)
          .onChange(async (value) => {
            this.plugin.settings.resolveForwardedSender = value;
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Ribbon badge")
      .setDesc("What to show on the ribbon icon badge.")
      .addDropdown((dropdown) =>
        dropdown
          .addOptions({
            off: "Off",
            unread: "Unread count",
            total: "Total messages",
          })
          .setValue(this.plugin.settings.badgeCount)
          .onChange(async (value) => {
            this.plugin.settings.badgeCount = value as "off" | "unread" | "total";
            this.plugin.scheduleSaveSettings();
            this.plugin.updateBadge(-1); // signal a re-sync
          }),
      );

    new Setting(containerEl)
      .setName("Badge position")
      .setDesc("Where to show the badge on the ribbon icon.")
      .addDropdown((dropdown) =>
        dropdown
          .addOptions({
            "top-right": "Top right",
            "top-left": "Top left",
            "bottom-right": "Bottom right",
            "bottom-left": "Bottom left",
            off: "Disabled",
          })
          .setValue(this.plugin.settings.badgePosition)
          .onChange(async (value) => {
            this.plugin.settings.badgePosition = value as any;
            this.plugin.scheduleSaveSettings();
            this.plugin.updateBadge(-1);
          }),
      );

    // AI Processing section
    containerEl.createEl("h3", { text: "AI Processing" });

    new Setting(containerEl)
      .setName("Enable Claude processing")
      .setDesc("Use Claude AI to convert emails into clean, readable markdown.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.enableClaudeProcessing)
          .onChange(async (value) => {
            this.plugin.settings.enableClaudeProcessing = value;
            this.plugin.scheduleSaveSettings();
            this.display();
          }),
      );

    if (this.plugin.settings.enableClaudeProcessing) {
      new Setting(containerEl)
        .setName("Anthropic API key")
        .setDesc("Your Anthropic API key for Claude.")
        .addText((text) => {
          text.inputEl.type = "password";
          text
            .setPlaceholder("sk-ant-...")
            .setValue(this.plugin.settings.anthropicApiKey)
            .onChange(async (value) => {
              this.plugin.settings.anthropicApiKey = value.trim();
              this.plugin.scheduleSaveSettings();
            });
        });

      new Setting(containerEl)
        .setName("Model")
        .setDesc("Claude model to use for processing.")
        .addDropdown((dropdown) =>
          dropdown
            .addOptions({
              "claude-haiku-4-5-20251001": "Haiku 4.5 (fastest, cheapest)",
              "claude-sonnet-4-6": "Sonnet 4.6",
              "claude-opus-4-6": "Opus 4.6 (most capable)",
            })
            .setValue(this.plugin.settings.claudeModel)
            .onChange(async (value) => {
              this.plugin.settings.claudeModel = value;
              this.plugin.scheduleSaveSettings();
            }),
        );

      new Setting(containerEl)
        .setName("System prompt")
        .setDesc(
          "Instructions for how Claude should process emails. Leave blank to use the default.",
        )
        .addTextArea((textArea) => {
          textArea.inputEl.rows = 5;
          textArea.inputEl.style.width = "100%";
          textArea.inputEl.placeholder = DEFAULT_CLAUDE_PROMPT;
          textArea
            .setValue(this.plugin.settings.claudeSystemPrompt)
            .onChange(async (value) => {
            this.plugin.settings.claudeSystemPrompt = value.trim();
            this.plugin.scheduleSaveSettings();
          });
        });

      new Setting(containerEl)
        .setName("Prefetch limit")
        .setDesc(
          "How many messages to pre-process with Claude in the background when the inbox loads. " +
          "Set to 0 to disable, or -1 for all messages.",
        )
        .addText((text) =>
          text
            .setPlaceholder("10")
            .setValue(String(this.plugin.settings.prefetchLimit))
            .onChange(async (value) => {
              const n = parseInt(value, 10);
              if (!isNaN(n) && n >= -1) {
                this.plugin.settings.prefetchLimit = n;
                this.plugin.scheduleSaveSettings();
              }
            }),
        );
    }

    // Event & Task Detection section
    containerEl.createEl("h3", { text: "Event & Task Detection" });

    new Setting(containerEl)
      .setName("Auto-detect events & tasks")
      .setDesc("Automatically scan emails for calendar events and actionable tasks.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.enableAutoItemDetection)
          .onChange(async (value) => {
            this.plugin.settings.enableAutoItemDetection = value;
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Event note folder")
      .setDesc("Vault folder for accepted event notes.")
      .addText((text) =>
        text
          .setPlaceholder("Events")
          .setValue(this.plugin.settings.eventNoteFolderPath)
          .onChange(async (value) => {
            this.plugin.settings.eventNoteFolderPath = value;
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Task note folder")
      .setDesc("Vault folder for accepted task notes.")
      .addText((text) =>
        text
          .setPlaceholder("Tasks")
          .setValue(this.plugin.settings.taskNoteFolderPath)
          .onChange(async (value) => {
            this.plugin.settings.taskNoteFolderPath = value;
            this.plugin.scheduleSaveSettings();
          }),
      );

    {
      const promptSetting = new Setting(containerEl)
        .setName("Detection prompt")
        .setDesc("Custom prompt for extracting events and tasks. Leave blank for default.");

      promptSetting.addTextArea((area) =>
        area
          .setPlaceholder(ITEM_DETECTION_PROMPT.slice(0, 120) + "…")
          .setValue(this.plugin.settings.itemDetectionPrompt)
          .onChange(async (value) => {
            this.plugin.settings.itemDetectionPrompt = value;
            this.plugin.scheduleSaveSettings();
          }),
      );

      promptSetting.addButton((btn) =>
        btn.setButtonText("Reset").onClick(async () => {
          this.plugin.settings.itemDetectionPrompt = "";
          this.plugin.scheduleSaveSettings();
          this.display();
        }),
      );
    }

    // Tag Classification section
    containerEl.createEl("h3", { text: "Tag Classification" });

    // Show each defined tag with its icon and definition, plus an "Add tag" button
    const categories = parseTagCategories(this.plugin.settings.tagCategories);
    if (!this.plugin.settings.tagIcons) this.plugin.settings.tagIcons = {};
    if (!this.plugin.settings.tagDescriptions) this.plugin.settings.tagDescriptions = {};

    new Setting(containerEl)
      .setName("Tags")
      .setDesc("Each tag has a name, icon, and description used by the yes/no classifier.")
      .addButton((btn) =>
        btn
          .setButtonText("Add tag")
          .setCta()
          .onClick(() => {
            const s = this.plugin.settings;
            const canGenerate = s.enableClaudeProcessing && hasClaudeAccess(s.anthropicApiKey);
            new CreateTagModal(this.app, {
              existingTags: parseTagCategories(s.tagCategories),
              onGenerate: canGenerate
                ? (name) => generateTagDescription(
                    s.anthropicApiKey,
                    s.claudeModel,
                    name,
                    parseTagCategories(s.tagCategories).map((n) => ({
                      name: n,
                      description: s.tagDescriptions?.[n] || "",
                    })),
                  )
                : undefined,
              onSubmit: async (name, criteria, icon, iconExplicit) => {
                const updated = [...parseTagCategories(s.tagCategories), name];
                s.tagCategories = updated.join(", ");
                s.tagDescriptions[name] = criteria;

                const usedIcons = Object.values(s.tagIcons);
                s.tagIcons[name] = icon;
                this.plugin.scheduleSaveSettings();
                this.display();

                if (!iconExplicit && canGenerate) {
                  try {
                    const picked = await pickTagIcon(
                      s.anthropicApiKey, s.claudeModel, name, criteria,
                      TAG_ICON_POOL, usedIcons,
                    );
                    if (picked) {
                      s.tagIcons[name] = picked;
                      this.plugin.scheduleSaveSettings();
                      this.display();
                    }
                  } catch {
                    // Keep fallback.
                  }
                }
              },
            }).open();
          }),
      );

    if (categories.length > 0) {
      const tagListEl = containerEl.createDiv({ cls: "iris-settings-tag-list" });
      for (const cat of categories) {
        if (!this.plugin.settings.tagIcons[cat]) {
          this.plugin.settings.tagIcons[cat] = "tag";
        }

        const needsCriteria = !(this.plugin.settings.tagDescriptions[cat] || "").trim();
        const tagEl = tagListEl.createDiv({
          cls: "iris-settings-tag-item is-clickable" + (needsCriteria ? " needs-description" : ""),
          attr: { role: "button", tabindex: "0", "aria-label": `Edit tag "${cat}"` },
        });

        const iconEl = tagEl.createSpan({ cls: "iris-settings-tag-icon" });
        setIcon(iconEl, this.plugin.settings.tagIcons[cat]);
        tagEl.createSpan({ text: cat, cls: "iris-settings-tag-name" });

        if (needsCriteria) {
          const warn = tagEl.createSpan({
            cls: "iris-settings-tag-warn",
            attr: { "aria-label": "Missing criteria — classifier accuracy will suffer" },
          });
          setIcon(warn, "alert-triangle");
        }

        const deleteBtn = tagEl.createEl("button", {
          cls: "iris-settings-tag-delete clickable-icon",
          attr: { "aria-label": `Delete tag "${cat}"` },
        });
        setIcon(deleteBtn, "trash-2");
        deleteBtn.addEventListener("click", (e) => {
          e.stopPropagation();
          const remaining = parseTagCategories(this.plugin.settings.tagCategories)
            .filter((n) => n !== cat);
          this.plugin.settings.tagCategories = remaining.join(", ");
          delete this.plugin.settings.tagIcons[cat];
          delete this.plugin.settings.tagDescriptions[cat];
          if (this.plugin.settings.tagPromptVersions) {
            delete this.plugin.settings.tagPromptVersions[cat];
          }
          this.plugin.scheduleSaveSettings();
          this.display();
        });

        const openEdit = () => this.openTagEditModal(cat);
        tagEl.addEventListener("click", openEdit);
        tagEl.addEventListener("keydown", (e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            openEdit();
          }
        });
      }
    }

    new Setting(containerEl)
      .setName("Auto-tag new emails")
      .setDesc(
        "Use Claude to automatically predict tags for untagged emails.",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.enableAutoTagging)
          .onChange(async (value) => {
            this.plugin.settings.enableAutoTagging = value;
            this.plugin.scheduleSaveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Tag classification prompt")
      .setDesc(
        "Meta-instructions for the yes/no classifier. Each tag is evaluated separately against its own definition. Refined automatically when you deny tags. Leave blank to use the default.",
      )
      .addTextArea((textArea) => {
        textArea.inputEl.rows = 5;
        textArea.inputEl.style.width = "100%";
        textArea.inputEl.placeholder = TAG_CLASSIFY_PROMPT;
        textArea
          .setValue(this.plugin.settings.tagClassifyPrompt)
          .onChange(async (value) => {
            this.plugin.settings.tagClassifyPrompt = value.trim();
            this.plugin.scheduleSaveSettings();
          });
      })
      .addButton((btn) =>
        btn.setButtonText("Reset").onClick(async () => {
          this.plugin.settings.tagClassifyPrompt = "";
          this.plugin.scheduleSaveSettings();
          this.display();
        }),
      );

    new Setting(containerEl)
      .setName("Clear auto-tags")
      .setDesc("Remove all automatically assigned tags.")
      .addButton((btn) =>
        btn.setButtonText("Clear auto-tags").onClick(async () => {
          const allTags = this.plugin.store.getAllTags();
          for (const [msgId, entries] of allTags) {
            for (const entry of entries) {
              if (entry.source === "auto") {
                this.plugin.store.removeTag(msgId, entry.tag);
              }
            }
          }
          this.display();
        }),
      );

    // Account section
    containerEl.createEl("h3", { text: "Account" });

    if (this.plugin.authProvider?.isSignedIn()) {
      const account = this.plugin.authProvider.getAccount();
      new Setting(containerEl)
        .setName("Signed in as")
        .setDesc(account?.username || "Unknown")
        .addButton((btn) =>
          btn.setButtonText("Sign out").onClick(async () => {
            await this.plugin.handleLogout();
            this.display();
          }),
        );
    } else {
      new Setting(containerEl)
        .setName("Not signed in")
        .setDesc("Sign in to access your Outlook inbox.")
        .addButton((btn) =>
          btn
            .setButtonText("Sign in")
            .setCta()
            .onClick(async () => {
              await this.plugin.handleLogin();
              this.display();
            }),
        );
    }

    // Advanced section
    containerEl.createEl("h3", { text: "Advanced" });

    new Setting(containerEl)
      .setName("Debug logging")
      .setDesc("Log detailed debug information to the developer console.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.debugLogging)
          .onChange(async (value) => {
            this.plugin.settings.debugLogging = value;
            setDebugEnabled(value);
            this.plugin.scheduleSaveSettings();
          }),
      );
  }

  private openTagEditModal(cat: string): void {
    const s = this.plugin.settings;
    const canGenerate = s.enableClaudeProcessing && hasClaudeAccess(s.anthropicApiKey);
    new CreateTagModal(this.app, {
      existingTags: parseTagCategories(s.tagCategories),
      initial: {
        name: cat,
        criteria: s.tagDescriptions?.[cat] || "",
        icon: s.tagIcons?.[cat] || "tag",
      },
      onGenerate: canGenerate
        ? (name) => generateTagDescription(
            s.anthropicApiKey, s.claudeModel, name,
            parseTagCategories(s.tagCategories)
              .filter((n) => n !== cat)
              .map((n) => ({ name: n, description: s.tagDescriptions?.[n] || "" })),
          )
        : undefined,
      onSubmit: (_name, criteria, icon) => {
        const criteriaChanged = (s.tagDescriptions[cat] || "") !== criteria;
        s.tagDescriptions[cat] = criteria;
        s.tagIcons[cat] = icon;
        if (criteriaChanged) {
          if (!s.tagPromptVersions) s.tagPromptVersions = {};
          bumpTagVersion(s.tagPromptVersions, cat);
        }
        this.plugin.scheduleSaveSettings();
        this.display();
      },
    }).open();
  }
}
