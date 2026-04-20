import { App, PluginSettingTab, Setting, setIcon } from "obsidian";
import type IrisMailPlugin from "../main";
import { DEFAULT_CLAUDE_PROMPT, TAG_CLASSIFY_PROMPT, TAG_ICON_POOL, ITEM_DETECTION_PROMPT, parseTagCategories, bumpTagVersion, MSAL_AUTHORITY_DEFAULT } from "../constants";
import { CreateTagModal } from "../views/components/CreateTagModal";
import { generateTagDescription, hasClaudeAccess, pickTagIcon } from "../utils/claudeApi";
import { setDebugEnabled } from "../utils/logger";
import { ImapAuthProvider } from "../auth/ImapAuthProvider";
import { IMAP_PRESETS, getImapPreset } from "../mail/imapPresets";
import type { Account, MailProvider, AuthMethod } from "../types";

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

    // ── Accounts ────────────────────────────────────────
    containerEl.createEl("h3", { text: "Accounts" });

    for (const account of this.plugin.settings.accounts) {
      this.renderAccount(containerEl, account);
    }

    new Setting(containerEl)
      .setName("Add account")
      .setDesc("Connect via Azure (OAuth) for Outlook, or via IMAP for any provider that supports app passwords.")
      .addButton((btn) =>
        btn.setButtonText("Via Azure").onClick(async () => {
          await this.plugin.createAccount({ label: "Outlook", provider: "outlook" });
          this.display();
        }),
      )
      .addButton((btn) =>
        btn.setButtonText("Via IMAP").onClick(async () => {
          await this.plugin.createAccount({ label: "IMAP", provider: "imap" });
          this.display();
        }),
      );

    // Shared OAuth port for both providers' loopback redirect.
    new Setting(containerEl)
      .setName("Redirect port")
      .setDesc(
        "Localhost port for OAuth redirect. Must match the redirect URI registered in your Azure app.",
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
    containerEl.createEl("p", {
      text:
        "Right-click selected text inside an email to create an event or task note. " +
        "Use the refresh button on a message to scan the full email at once.",
      cls: "setting-item-description",
    });

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

  private renderAccount(containerEl: HTMLElement, account: Account): void {
    const wrap = containerEl.createDiv({ cls: "iris-account-card" });

    // Header: label + provider + sign-in/out
    const entry = this.plugin.accounts.get(account.id);
    const signedIn = !!entry?.auth.isSignedIn();
    const signedInAs = entry?.auth.getAccount()?.username;
    const providerLabel =
      account.provider === "imap" ? "IMAP" : "Outlook (Azure)";

    new Setting(wrap)
      .setName(account.label)
      .setDesc(
        providerLabel +
          (signedIn ? ` — signed in as ${signedInAs ?? "(unknown)"}` : " — not signed in") +
          (account.enabled ? "" : " — disabled"),
      )
      .addToggle((toggle) =>
        toggle
          .setTooltip("Include in unified inbox")
          .setValue(account.enabled)
          .onChange(async (value) => {
            await this.plugin.updateAccount({ ...account, enabled: value });
            this.display();
          }),
      )
      .addButton((btn) => {
        if (signedIn) {
          btn.setButtonText("Sign out").onClick(async () => {
            await this.plugin.logoutAccount(account.id);
            this.display();
          });
        } else {
          btn.setButtonText("Sign in").setCta().onClick(async () => {
            await this.plugin.loginAccount(account.id);
            this.display();
          });
        }
      })
      .addButton((btn) =>
        btn
          .setButtonText("Remove")
          .setWarning()
          .onClick(async () => {
            await this.plugin.removeAccount(account.id);
            this.display();
          }),
      );

    // Label edit
    new Setting(wrap)
      .setName("Label")
      .addText((text) =>
        text
          .setPlaceholder(providerLabel)
          .setValue(account.label)
          .onChange(async (value) => {
            await this.plugin.updateAccount({ ...account, label: value.trim() || account.label });
          }),
      );

    // Per-provider credentials
    if (account.provider === "outlook") {
      new Setting(wrap)
        .setName("Azure Client ID")
        .setDesc("Application (client) ID from your Azure Entra app registration.")
        .addText((text) =>
          text
            .setPlaceholder("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
            .setValue(account.clientId ?? "")
            .onChange(async (value) => {
              await this.plugin.updateAccount({ ...account, clientId: value.trim() });
            }),
        );

      new Setting(wrap)
        .setName("Authority")
        .setDesc("Use 'common' for any account, or your tenant ID for org-only.")
        .addText((text) =>
          text
            .setPlaceholder(MSAL_AUTHORITY_DEFAULT)
            .setValue(account.authority ?? "")
            .onChange(async (value) => {
              await this.plugin.updateAccount({ ...account, authority: value.trim() });
            }),
        );

      new Setting(wrap)
        .setName("Sign-in method")
        .addDropdown((dropdown) =>
          dropdown
            .addOptions({ "auth-code": "Browser redirect", "device-code": "Device code" })
            .setValue(account.authMethod ?? "auth-code")
            .onChange(async (value) => {
              await this.plugin.updateAccount({ ...account, authMethod: value as AuthMethod });
            }),
        );
    } else {
      // IMAP
      const preset = getImapPreset(account.imapPreset);
      const presetSetting = new Setting(wrap)
        .setName("Provider preset")
        .addDropdown((dd) => {
          for (const p of IMAP_PRESETS) dd.addOption(p.key, p.label);
          dd.setValue(account.imapPreset ?? "other");
          dd.onChange(async (value) => {
            const p = getImapPreset(value);
            const patch: Account = { ...account, imapPreset: value };
            if (p && p.host) {
              patch.imapHost = p.host;
              patch.imapPort = p.port;
              patch.imapSecure = p.secure;
            }
            await this.plugin.updateAccount(patch);
            this.display();
          });
        });

      if (preset?.hint) {
        const descFrag = document.createDocumentFragment();
        descFrag.appendText(preset.hint);
        if (preset.appPasswordUrl) {
          descFrag.appendText(" ");
          descFrag.createEl("a", {
            text: "Open app password page",
            href: preset.appPasswordUrl,
          });
        }
        presetSetting.setDesc(descFrag);
      } else {
        presetSetting.setDesc("Pick your provider, or 'Other' for a custom server.");
      }

      // Always read the current account by id when patching — the closure's
      // `account` snapshot goes stale as soon as any previous onChange fires.
      const patch = async (delta: Partial<Account>): Promise<void> => {
        const current = this.plugin.settings.accounts.find((a) => a.id === account.id);
        if (!current) return;
        await this.plugin.updateAccount({ ...current, ...delta });
      };

      new Setting(wrap)
        .setName("Email address")
        .setDesc("Your IMAP login — usually the full email address.")
        .addText((text) =>
          text
            .setPlaceholder("you@example.com")
            .setValue(account.imapEmail ?? "")
            .onChange(async (value) => {
              await patch({ imapEmail: value.trim() });
            }),
        );

      new Setting(wrap)
        .setName("IMAP host")
        .addText((text) =>
          text
            .setPlaceholder("imap.example.com")
            .setValue(account.imapHost ?? "")
            .onChange(async (value) => {
              await patch({ imapHost: value.trim() });
            }),
        );

      new Setting(wrap)
        .setName("IMAP port")
        .addText((text) =>
          text
            .setPlaceholder("993")
            .setValue(account.imapPort ? String(account.imapPort) : "")
            .onChange(async (value) => {
              const port = parseInt(value, 10);
              if (!isNaN(port) && port > 0 && port <= 65535) {
                await patch({ imapPort: port });
              }
            }),
        );

      new Setting(wrap)
        .setName("Use TLS")
        .setDesc("Required for port 993. Disable only for STARTTLS on port 143.")
        .addToggle((toggle) =>
          toggle
            .setValue(account.imapSecure ?? true)
            .onChange(async (value) => {
              await patch({ imapSecure: value });
            }),
        );

      {
        const imapAuth = entry?.auth instanceof ImapAuthProvider ? entry.auth : null;
        const hasSaved = !!imapAuth?.hasStoredPassword();
        const SENTINEL = "••••••••••••••••";
        new Setting(wrap)
          .setName("App password")
          .setDesc(
            hasSaved
              ? "A password is saved. Click and type to replace it."
              : "Saved encrypted as you type. Use an app-specific password — not your account password.",
          )
          .addText((text) => {
            text.inputEl.type = "password";
            if (hasSaved) {
              text.setValue(SENTINEL);
            } else {
              text.setPlaceholder("xxxx xxxx xxxx xxxx");
            }
            // Clear the sentinel on focus so the user can type without
            // appending to it. Restore on blur if they didn't type anything.
            text.inputEl.addEventListener("focus", () => {
              if (text.inputEl.value === SENTINEL) text.setValue("");
            });
            text.inputEl.addEventListener("blur", () => {
              if (!text.inputEl.value && hasSaved) text.setValue(SENTINEL);
            });
            text.onChange((value) => {
              if (value === SENTINEL) return;
              // Empty while something's saved = leave the stored password alone.
              if (!value && hasSaved) return;
              imapAuth?.setPassword(value);
            });
          });
      }
    }
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
