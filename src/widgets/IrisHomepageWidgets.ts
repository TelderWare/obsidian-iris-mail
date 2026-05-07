import type IrisMailPlugin from "../main";
import type { Box, Message } from "../types";
import { MessageList } from "../views/components/MessageList";
import { makeNameResolver } from "../utils/nameResolve";
import { makeEffectiveSenderResolver } from "../utils/effectiveSender";
import { InboxView } from "../views/InboxView";
import { VIEW_TYPE_IRIS_MAIL } from "../constants";

/**
 * Contract consumed by the Iris Homepage plugin. Each widget is created inside
 * an `IrisHomepageContext.containerEl`, and returns a handle that Iris Homepage
 * uses to tear it down or react to resize / config changes.
 */
export interface IrisHomepageContext {
  containerEl: HTMLElement;
  config?: unknown;
}

export interface IrisHomepageWidgetHandle {
  destroy(): void;
  onResize?(): void;
  onConfigChange?(config: unknown): void;
}

export interface IrisHomepageWidgetDescriptor {
  type: string;
  label: string;
  icon: string;
  defaultSizePx: { width: number; height: number };
  create(ctx: IrisHomepageContext): IrisHomepageWidgetHandle;
}

/** Inputs needed to apply a box's predicate to a message. */
export interface BoxMatchContext {
  junkIds: Set<string>;
  todoIds: Set<string>;
  pinnedIds: Set<string>;
  tagsByMessage: Map<string, Set<string>>;
  boxes: readonly Box[];
  /** Message IDs the classifier or summarizer is currently working on. */
  inFlight: Set<string>;
}

/**
 * Pure predicate mirroring InboxView.messageMatchesBox so widgets can filter
 * the shared message snapshot without owning any UI state of their own.
 */
export function messageMatchesBox(msg: Message, box: Box, ctx: BoxMatchContext): boolean {
  const id = msg.id;
  if (!id) return false;

  const msgTags = ctx.tagsByMessage.get(id) ?? new Set<string>();
  const boxes = ctx.boxes;

  const hasBuiltinTagMatch = (builtin: "junk" | "todo"): boolean => {
    if (msgTags.size === 0) return false;
    const wired = boxes.find((b) => b.builtin === builtin);
    if (!wired || !wired.tags || wired.tags.length === 0) return false;
    for (const t of wired.tags) if (msgTags.has(t)) return true;
    return false;
  };

  const junkSignal = ctx.junkIds.has(id) || hasBuiltinTagMatch("junk");
  const todoSignal = ctx.todoIds.has(id) || hasBuiltinTagMatch("todo");

  if (box.builtin === "secretary") {
    if (junkSignal) return false;
    return ctx.inFlight.has(id);
  }

  const isUnread = !msg.isRead;
  const effectiveTodo = !isUnread && todoSignal;
  const effectiveJunk = !isUnread && !effectiveTodo && junkSignal;

  switch (box.builtin) {
    case "in":
      return isUnread;
    case "todo":
      return effectiveTodo;
    case "junk":
      return effectiveJunk;
    case "read":
      return !isUnread && !effectiveTodo && !effectiveJunk;
    default: {
      if (effectiveTodo || effectiveJunk) return false;
      const boxTags = box.tags || [];
      if (boxTags.length === 0) return false;
      for (const t of boxTags) if (msgTags.has(t)) return true;
      return false;
    }
  }
}

function buildMatchContext(plugin: IrisMailPlugin): BoxMatchContext {
  const tagsByMessage = new Map<string, Set<string>>();
  for (const [id, entries] of plugin.store.getAllTags()) {
    tagsByMessage.set(id, new Set(entries.map((e) => e.tag)));
  }
  return {
    junkIds: plugin.store.getAllJunkIds(),
    todoIds: plugin.store.getAllTodoIds(),
    pinnedIds: plugin.store.getAllPinnedIds(),
    tagsByMessage,
    boxes: plugin.settings.boxes || [],
    inFlight: new Set(),
  };
}

/**
 * Per-widget UI state. Reuses the real InboxView MessageList so rows look
 * byte-identical to the main inbox (tags, attachment icons, account labels,
 * hover/selected/unread styling, IntersectionObserver-backed pagination).
 */
class BoxWidget {
  private root: HTMLElement;
  private listEl: HTMLElement;
  private messageList: MessageList;

  constructor(
    private readonly plugin: IrisMailPlugin,
    private box: Box,
    container: HTMLElement,
  ) {
    this.root = container.createDiv({
      cls: "iris-mail-container iris-homepage-widget-root",
    });

    // Match Iris Homepage's own "Recent notes" widget header.
    this.root.createEl("h6", { cls: "iris-hp-widget-title", text: box.name });

    // Matches the class the main view uses so message-list CSS applies as-is.
    this.listEl = this.root.createDiv({ cls: "iris-message-list" });

    const nameResolver = makeNameResolver(plugin.store.getAllNicknames());
    const effectiveSenderResolver = makeEffectiveSenderResolver(plugin);
    // Every interaction in the widget routes back to the full inbox view —
    // selection, pagination, nickname/rule edits all need the real UI.
    const openBox = () => void this.openInMail();
    this.messageList = new MessageList(
      this.listEl,
      {
        onMessageSelect: openBox,
        onLoadMore: openBox,
        onMultiSelect: () => { /* no-op — selection lives in the real view */ },
        onEditNickname: openBox,
        onEditSenderRule: openBox,
        onMarkSenderAsJunk: openBox,
      },
      nameResolver,
      effectiveSenderResolver,
    );

    this.render();
  }

  /** Re-seed the MessageList's tag caches from current store/settings. */
  private applyTagContext(): void {
    const s = this.plugin.settings;
    this.messageList.setTagCache(this.plugin.store.getAllTags());
    this.messageList.setTagIcons(new Map(Object.entries(s.tagIcons || {})));
    this.messageList.setTagColors(new Map(Object.entries(s.tagColors || {})));
    this.messageList.setTagDescriptions(
      new Map(
        Object.entries(s.tagDescriptions || {}).filter(
          ([, v]) => typeof v === "string" && v.trim().length > 0,
        ),
      ),
    );
    const hidden = s.tagHiddenInList || {};
    this.messageList.setHiddenListTags(
      new Set(Object.keys(hidden).filter((k) => hidden[k])),
    );
    this.messageList.setPinnedIds(this.plugin.store.getAllPinnedIds());
    // The user can toggle `resolveForwardedSender` between renders.
    this.messageList.setEffectiveSenderResolver(makeEffectiveSenderResolver(this.plugin));
  }

  /** Always re-read the current box definition — the user may have renamed
   *  or retagged it through the Iris Mail UI. */
  private refreshBoxRef(): void {
    const latest = (this.plugin.settings.boxes || []).find((b) => b.id === this.box.id);
    if (latest) this.box = latest;
  }

  render(): void {
    this.refreshBoxRef();
    this.applyTagContext();

    const ctx = buildMatchContext(this.plugin);
    const messages = this.plugin.getInboxMessages();
    const filtered = messages.filter((m) => messageMatchesBox(m, this.box, ctx));

    const dir = this.plugin.settings.sortNewestFirst ? -1 : 1;
    const byDate = (a: Message, b: Message) =>
      dir *
      (new Date(a.receivedDateTime || 0).getTime() -
        new Date(b.receivedDateTime || 0).getTime());
    // Pinned messages float to the top of every box, matching InboxView.
    const sorted = [...filtered].sort((a, b) => {
      const ap = a.id ? ctx.pinnedIds.has(a.id) : false;
      const bp = b.id ? ctx.pinnedIds.has(b.id) : false;
      if (ap !== bp) return ap ? -1 : 1;
      return byDate(a, b);
    });

    this.root.style.display = sorted.length === 0 ? "none" : "";

    const empty =
      messages.length === 0
        ? { icon: this.box.icon || "inbox", title: "No messages loaded yet" }
        : { icon: this.box.icon || "inbox", title: `No messages in ${this.box.name}` };

    void this.messageList.renderFlatMessages(sorted, false, empty);
  }

  /** Open the full Iris Mail view focused on this widget's box. */
  private async openInMail(): Promise<void> {
    this.plugin.settings.selectedBoxId = this.box.id;
    await this.plugin.saveSettings();
    await this.plugin.activateView();

    // activateView only reveals an existing leaf — it doesn't re-read
    // selectedBoxId — so push the selection straight onto the live view.
    for (const leaf of this.plugin.app.workspace.getLeavesOfType(VIEW_TYPE_IRIS_MAIL)) {
      if (leaf.view instanceof InboxView) leaf.view.showBox(this.box.id);
    }
  }

  destroy(): void {
    this.root.remove();
  }
}

/**
 * Build the widget descriptors Iris Homepage consumes. One widget per visible
 * box. Widgets auto-refresh whenever the plugin's inbox snapshot changes.
 */
export function buildIrisHomepageWidgets(
  plugin: IrisMailPlugin,
): IrisHomepageWidgetDescriptor[] {
  const boxes = (plugin.settings.boxes || []).filter((b) => !b.hidden);
  return boxes.map<IrisHomepageWidgetDescriptor>((box) => ({
    type: `iris-mail:box-${box.id}`,
    label: `${box.name} (Iris Mail)`,
    icon: box.icon || "inbox",
    defaultSizePx: { width: 380, height: 460 },
    create(ctx) {
      const widget = new BoxWidget(plugin, box, ctx.containerEl);
      const unsubscribe = plugin.onMessagesChanged(() => widget.render());

      if (plugin.getInboxMessages().length === 0 && plugin.accounts?.anySignedIn()) {
        void plugin.backgroundRefresh();
      }

      return {
        destroy: () => {
          unsubscribe();
          widget.destroy();
        },
      };
    },
  }));
}
