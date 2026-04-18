import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

const SUBSCRIPTION_PERMISSIONS = {
  delegated: ["same as the subscribed resource"],
  application: ["same as the subscribed resource"],
};

export const subscriptionsTools: ToolDefinition[] = [
  {
    name: "graph_subscriptions_create",
    description: "Create a webhook subscription to receive change notifications for a Microsoft Graph resource.",
    category: "subscriptions",
    zodShape: {
      resource: z.string().describe("Resource path to watch (e.g. 'users', 'me/mailFolders/inbox/messages', 'sites/{siteId}/lists/{listId}', '/teams/{teamId}/channels')"),
      changeTypes: z.array(z.enum(["created", "updated", "deleted"])).describe("Event types to subscribe to — any combination of 'created', 'updated', 'deleted'"),
      notificationUrl: z.string().describe("HTTPS endpoint that will receive notifications — must be publicly reachable"),
      expirationMinutes: z.number().optional().describe("Subscription lifetime in minutes from now. Max varies by resource (mail/calendar/contacts: 4230 min, others: up to 4320 min). Defaults to 4230."),
      clientState: z.string().optional().describe("Secret value included in every notification so you can verify it came from Graph"),
    },
    handler: (args: {
      resource: string;
      changeTypes: ("created" | "updated" | "deleted")[];
      notificationUrl: string;
      expirationMinutes?: number;
      clientState?: string;
    }) => {
      const endpoint = `${BASE}/subscriptions`;
      const body = {
        changeType: args.changeTypes.join(","),
        notificationUrl: args.notificationUrl,
        resource: args.resource,
        expirationDateTime: new Date(Date.now() + (args.expirationMinutes ?? 4230) * 60000).toISOString(),
        clientState: args.clientState ?? "secretClientValue",
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body,
        description: `Create a webhook subscription for resource '${args.resource}'.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/subscription-post-subscriptions",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst subscription = await response.json();`,
        requiredPermissions: {
          delegated: ["varies by resource — use the same permission required to read that resource"],
          application: ["varies by resource — use the same permission required to read that resource"],
        },
        notes: "notificationUrl must be HTTPS and publicly reachable. Graph will send a validation request to the URL before creating the subscription. Subscriptions expire and must be renewed. Max expiry varies: mail/calendar/contacts = 4230 min, other resources = up to 4320 min (check docs for exact limits).",
      };
    },
  },
  {
    name: "graph_subscriptions_list",
    description: "List all active webhook subscriptions for the current app.",
    category: "subscriptions",
    zodShape: {},
    handler: (_args: Record<string, never>) => {
      const endpoint = `${BASE}/subscriptions`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body: null,
        description: "List all active webhook subscriptions.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/subscription-list",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: SUBSCRIPTION_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_subscriptions_get",
    description: "Get a specific webhook subscription by ID.",
    category: "subscriptions",
    zodShape: {
      subscriptionId: z.string().describe("Webhook subscription ID"),
    },
    handler: (args: { subscriptionId: string }) => {
      const endpoint = `${BASE}/subscriptions/${args.subscriptionId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { subscriptionId: args.subscriptionId },
        queryParams: {},
        body: null,
        description: `Get subscription ${args.subscriptionId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/subscription-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst subscription = await response.json();`,
        requiredPermissions: SUBSCRIPTION_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_subscriptions_delete",
    description: "Delete a webhook subscription to stop receiving notifications.",
    category: "subscriptions",
    zodShape: {
      subscriptionId: z.string().describe("Webhook subscription ID to delete"),
    },
    handler: (args: { subscriptionId: string }) => {
      const endpoint = `${BASE}/subscriptions/${args.subscriptionId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { subscriptionId: args.subscriptionId },
        queryParams: {},
        body: null,
        description: `Delete subscription ${args.subscriptionId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/subscription-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
        requiredPermissions: SUBSCRIPTION_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_subscriptions_renew",
    description: "Renew a webhook subscription before it expires to continue receiving notifications.",
    category: "subscriptions",
    zodShape: {
      subscriptionId: z.string().describe("Webhook subscription ID to renew"),
      expirationMinutes: z.number().optional().describe("New lifetime in minutes from now. Defaults to 4230. Check resource-specific limits in the docs."),
    },
    handler: (args: { subscriptionId: string; expirationMinutes?: number }) => {
      const endpoint = `${BASE}/subscriptions/${args.subscriptionId}`;
      const body = {
        expirationDateTime: new Date(Date.now() + (args.expirationMinutes ?? 4230) * 60000).toISOString(),
      };
      return {
        endpoint,
        method: "PATCH",
        headers: buildHeaders(),
        pathParams: { subscriptionId: args.subscriptionId },
        queryParams: {},
        body,
        description: `Renew subscription ${args.subscriptionId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/subscription-update",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'PATCH',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});\nconst updated = await response.json();`,
        requiredPermissions: SUBSCRIPTION_PERMISSIONS,
        notes: "Renew subscriptions before they expire — Graph does not auto-renew. Consider renewing when less than 15 minutes remain. Set up a scheduled job to handle renewal.",
      };
    },
  },
];
