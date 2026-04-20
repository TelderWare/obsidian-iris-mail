import { Client } from "@microsoft/microsoft-graph-client";
import type { OutlookAuthProvider } from "../auth/OutlookAuthProvider";
import type { MailFolder, Message } from "../types";
import { MESSAGE_LIST_SELECT } from "../constants";
import type { MailApi, MailListResponse, ListMessagesOptions } from "./MailApi";

interface GraphPagedResponse<T> {
  "@odata.nextLink"?: string;
  value: T[];
}

export class OutlookMailApi implements MailApi {
  private client: Client | null = null;

  constructor(private authProvider: OutlookAuthProvider) {}

  private getClient(): Client {
    if (!this.client) {
      this.client = Client.initWithMiddleware({
        authProvider: this.authProvider,
      });
    }
    return this.client;
  }

  async listFolders(): Promise<MailFolder[]> {
    const response: GraphPagedResponse<MailFolder> = await this.getClient()
      .api("/me/mailFolders")
      .filter("isHidden eq false")
      .top(50)
      .get();
    return response.value;
  }

  async listMessages(
    folderId: string,
    options: ListMessagesOptions = {},
  ): Promise<MailListResponse<Message>> {
    let response: GraphPagedResponse<Message>;

    if (options.nextLink) {
      response = await this.getClient().api(options.nextLink).get();
    } else {
      let request = this.getClient()
        .api(`/me/mailFolders/${folderId}/messages`)
        .select(MESSAGE_LIST_SELECT)
        .orderby("receivedDateTime desc")
        .top(options.top ?? 25);

      if (options.search) {
        request = request.search(`"${options.search}"`);
      }
      if (options.unreadOnly) {
        request = request.filter("isRead eq false");
      }

      response = await request.get();
    }

    return {
      value: response.value,
      nextLink: response["@odata.nextLink"] ?? null,
    };
  }

  async getMessage(messageId: string): Promise<Message> {
    return await this.getClient().api(`/me/messages/${messageId}`).get();
  }

  async getMessageBody(messageId: string): Promise<Message> {
    return await this.getClient()
      .api(`/me/messages/${messageId}`)
      .select(
        "id,subject,body,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments",
      )
      .expand("attachments($select=id,name,size,contentType,isInline)")
      .get();
  }

  async markAsRead(messageId: string): Promise<void> {
    await this.getClient()
      .api(`/me/messages/${messageId}`)
      .update({ isRead: true });
  }

  async markAsUnread(messageId: string): Promise<void> {
    await this.getClient()
      .api(`/me/messages/${messageId}`)
      .update({ isRead: false });
  }

  async deleteMessage(messageId: string): Promise<void> {
    await this.getClient()
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: "deletedItems" });
  }
}
