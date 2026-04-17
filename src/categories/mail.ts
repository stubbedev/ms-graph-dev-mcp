import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

function buildRecipients(emails: string[]): Array<{ emailAddress: { address: string } }> {
  return emails.map((e) => ({ emailAddress: { address: e } }));
}

const READ_PERMISSIONS = {
  delegated: ["Mail.Read", "Mail.ReadBasic"],
  application: ["Mail.Read"],
};

const SEND_PERMISSIONS = {
  delegated: ["Mail.Send"],
  application: ["Mail.Send"],
};

const READWRITE_PERMISSIONS = {
  delegated: ["Mail.ReadWrite"],
  application: ["Mail.ReadWrite"],
};

export const mailTools: ToolDefinition[] = [
  {
    name: "graph_mail_list_messages",
    description: "List messages in a user mailbox",
    category: "mail",
    zodShape: {
      userId: z.string(),
      filter: z.string().optional(),
      select: z.array(z.string()).optional(),
      top: z.number().optional(),
    },
    handler: (args: { userId: string; filter?: string; select?: string[]; top?: number }) => {
      const params: string[] = [];
      if (args.filter) params.push(`$filter=${encodeURIComponent(args.filter)}`);
      if (args.select?.length) params.push(`$select=${args.select.join(",")}`);
      if (args.top) params.push(`$top=${args.top}`);
      const qs = params.length ? `?${params.join("&")}` : "";
      const endpoint = `${BASE}/users/${args.userId}/messages${qs}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `List messages for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-messages",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_get_message",
    description: "Get a specific mail message",
    category: "mail",
    zodShape: {
      userId: z.string(),
      messageId: z.string(),
    },
    handler: (args: { userId: string; messageId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/messages/${args.messageId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, messageId: args.messageId },
        queryParams: {},
        body: null,
        description: `Get message ${args.messageId} for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst message = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_send",
    description: "Send an email message",
    category: "mail",
    zodShape: {
      userId: z.string(),
      toRecipients: z.array(z.string()),
      subject: z.string(),
      body: z.string(),
      bodyContentType: z.enum(["Text", "HTML"]).optional(),
      ccRecipients: z.array(z.string()).optional(),
    },
    handler: (args: {
      userId: string;
      toRecipients: string[];
      subject: string;
      body: string;
      bodyContentType?: "Text" | "HTML";
      ccRecipients?: string[];
    }) => {
      const endpoint = `${BASE}/users/${args.userId}/sendMail`;
      const requestBody: Record<string, unknown> = {
        message: {
          subject: args.subject,
          body: { contentType: args.bodyContentType ?? "Text", content: args.body },
          toRecipients: buildRecipients(args.toRecipients),
          ...(args.ccRecipients?.length ? { ccRecipients: buildRecipients(args.ccRecipients) } : {}),
        },
        saveToSentItems: true,
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: requestBody,
        description: `Send an email on behalf of ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-sendmail",
        codeExample: `await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(requestBody, null, 2)})\n});`,
        requiredPermissions: SEND_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_create_draft",
    description: "Create a draft mail message",
    category: "mail",
    zodShape: {
      userId: z.string(),
      toRecipients: z.array(z.string()),
      subject: z.string(),
      body: z.string(),
      bodyContentType: z.enum(["Text", "HTML"]).optional(),
      ccRecipients: z.array(z.string()).optional(),
    },
    handler: (args: {
      userId: string;
      toRecipients: string[];
      subject: string;
      body: string;
      bodyContentType?: "Text" | "HTML";
      ccRecipients?: string[];
    }) => {
      const endpoint = `${BASE}/users/${args.userId}/messages`;
      const requestBody: Record<string, unknown> = {
        subject: args.subject,
        body: { contentType: args.bodyContentType ?? "Text", content: args.body },
        toRecipients: buildRecipients(args.toRecipients),
        ...(args.ccRecipients?.length ? { ccRecipients: buildRecipients(args.ccRecipients) } : {}),
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: requestBody,
        description: `Create a draft message for ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-messages",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(requestBody, null, 2)})\n});\nconst draft = await response.json();`,
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_reply",
    description: "Reply to a mail message",
    category: "mail",
    zodShape: {
      userId: z.string(),
      messageId: z.string(),
      comment: z.string(),
    },
    handler: (args: { userId: string; messageId: string; comment: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/messages/${args.messageId}/reply`;
      const body = { comment: args.comment };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, messageId: args.messageId },
        queryParams: {},
        body,
        description: `Reply to message ${args.messageId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-reply",
        codeExample: `await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});`,
        requiredPermissions: SEND_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_delete_message",
    description: "Delete a mail message",
    category: "mail",
    zodShape: {
      userId: z.string(),
      messageId: z.string(),
    },
    handler: (args: { userId: string; messageId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/messages/${args.messageId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, messageId: args.messageId },
        queryParams: {},
        body: null,
        description: `Delete message ${args.messageId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_list_folders",
    description: "List mail folders for a user",
    category: "mail",
    zodShape: {
      userId: z.string(),
    },
    handler: (args: { userId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/mailFolders`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `List mail folders for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_move_message",
    description: "Move a message to a different folder",
    category: "mail",
    zodShape: {
      userId: z.string(),
      messageId: z.string(),
      destinationFolderId: z.string(),
    },
    handler: (args: { userId: string; messageId: string; destinationFolderId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/messages/${args.messageId}/move`;
      const body = { destinationId: args.destinationFolderId };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, messageId: args.messageId },
        queryParams: {},
        body,
        description: `Move message ${args.messageId} to folder ${args.destinationFolderId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-move",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});\nconst movedMessage = await response.json();`,
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_mail_get_delta",
    description: "Get changes to messages in a mail folder since a previous delta token — returns only new, updated, or deleted messages.",
    category: "mail",
    zodShape: {
      userId: z.string(),
      folderId: z.string().optional().describe("Mail folder ID or well-known name like 'inbox'. Defaults to inbox."),
      deltaToken: z.string().optional(),
    },
    handler: (args: { userId: string; folderId?: string; deltaToken?: string }) => {
      let endpoint: string;
      if (args.deltaToken) {
        endpoint = args.deltaToken;
      } else {
        endpoint = `${BASE}/users/${args.userId}/mailFolders/${args.folderId ?? "inbox"}/messages/delta`;
      }
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `Get delta changes to messages in folder ${args.folderId ?? "inbox"} for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-delta",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();\n// Store data['@odata.deltaLink'] for next call`,
        requiredPermissions: READ_PERMISSIONS,
        notes: "Store @odata.deltaLink between syncs. Pass it back as deltaToken to receive only changes since last sync. @removed items have been deleted.",
      };
    },
  },
];
