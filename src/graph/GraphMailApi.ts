import { Client } from "@microsoft/microsoft-graph-client";
import type { AuthProvider } from "../auth/AuthProvider";
import type { MailFolder, Message, GraphPagedResponse } from "../types";
import { MESSAGE_LIST_SELECT } from "../constants";

export class GraphMailApi {
  private client: Client | null = null;

  constructor(private authProvider: AuthProvider) {}

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
    options: {
      top?: number;
      search?: string;
      filter?: string;
      nextLink?: string;
    } = {},
  ): Promise<GraphPagedResponse<Message>> {
    if (options.nextLink) {
      // nextLink is a full URL — pass it directly
      return await this.getClient().api(options.nextLink).get();
    }

    let request = this.getClient()
      .api(`/me/mailFolders/${folderId}/messages`)
      .select(MESSAGE_LIST_SELECT)
      .orderby("receivedDateTime desc")
      .top(options.top ?? 25);

    if (options.search) {
      request = request.search(`"${options.search}"`);
    }
    if (options.filter) {
      request = request.filter(options.filter);
    }

    return await request.get();
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

}
