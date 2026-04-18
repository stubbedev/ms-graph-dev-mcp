import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

const READ_PERMISSIONS = {
  delegated: ["Sites.Read.All"],
  application: ["Sites.Read.All"],
};

const WRITE_PERMISSIONS = {
  delegated: ["Sites.ReadWrite.All"],
  application: ["Sites.ReadWrite.All"],
};

export const sitesTools: ToolDefinition[] = [
  {
    name: "graph_sites_list",
    description: "List SharePoint sites",
    category: "sites",
    zodShape: {
      search: z.string().optional().describe("Search query to filter sites by name or URL (e.g. 'marketing')"),
    },
    handler: (args: { search?: string }) => {
      const qs = args.search ? `?search=${encodeURIComponent(args.search)}` : "";
      const endpoint = `${BASE}/sites${qs}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: args.search ? { search: args.search } : {},
        body: null,
        description: "List SharePoint sites.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/site-list",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_get",
    description: "Get a SharePoint site by ID",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID (e.g. contoso.sharepoint.com,guid,guid)"),
    },
    handler: (args: { siteId: string }) => {
      const endpoint = `${BASE}/sites/${args.siteId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId },
        queryParams: {},
        body: null,
        description: `Get site ${args.siteId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/site-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst site = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_get_by_url",
    description: "Get a SharePoint site by hostname and relative path",
    category: "sites",
    zodShape: {
      hostname: z.string().describe("SharePoint hostname (e.g. contoso.sharepoint.com)"),
      siteRelativePath: z.string().describe("Site-relative path (e.g. /sites/MySite)"),
    },
    handler: (args: { hostname: string; siteRelativePath: string }) => {
      const endpoint = `${BASE}/sites/${args.hostname}:${args.siteRelativePath}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { hostname: args.hostname, siteRelativePath: args.siteRelativePath },
        queryParams: {},
        body: null,
        description: `Get site at ${args.hostname}${args.siteRelativePath}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/site-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst site = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_list_lists",
    description: "List SharePoint lists in a site",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
    },
    handler: (args: { siteId: string }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId },
        queryParams: {},
        body: null,
        description: `List all lists in site ${args.siteId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/list-list",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_get_list",
    description: "Get a SharePoint list",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
      listId: z.string().describe("SharePoint list ID"),
    },
    handler: (args: { siteId: string; listId: string }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists/${args.listId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId, listId: args.listId },
        queryParams: {},
        body: null,
        description: `Get list ${args.listId} from site ${args.siteId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/list-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst list = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_create_list",
    description: "Create a SharePoint list",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID to create the list in"),
      displayName: z.string().describe("Display name for the new list"),
      template: z.string().optional().describe("List template (e.g. 'genericList', 'documentLibrary'). Defaults to genericList."),
    },
    handler: (args: { siteId: string; displayName: string; template?: string }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists`;
      const body: Record<string, unknown> = {
        displayName: args.displayName,
        list: { template: args.template ?? "genericList" },
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId },
        queryParams: {},
        body,
        description: `Create list '${args.displayName}' in site ${args.siteId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/list-create",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst list = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_list_items",
    description: "List items in a SharePoint list",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
      listId: z.string().describe("SharePoint list ID"),
      filter: z.string().optional().describe("OData $filter expression (e.g. \"fields/Status eq 'Active'\")"),
      select: z.array(z.string()).optional().describe("Fields to return (e.g. ['id', 'fields/Title', 'fields/Status'])"),
    },
    handler: (args: { siteId: string; listId: string; filter?: string; select?: string[] }) => {
      const params: string[] = ["expand=fields"];
      if (args.filter) params.push(`$filter=${encodeURIComponent(args.filter)}`);
      if (args.select?.length) params.push(`$select=${args.select.join(",")}`);
      const endpoint = `${BASE}/sites/${args.siteId}/lists/${args.listId}/items?${params.join("&")}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId, listId: args.listId },
        queryParams: { expand: "fields" },
        body: null,
        description: `List items in list ${args.listId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-list",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: "Column values are returned under the 'fields' property. This request already includes ?expand=fields. Without it the items array would contain only metadata, not column data.",
      };
    },
  },
  {
    name: "graph_sites_get_item",
    description: "Get a SharePoint list item",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
      listId: z.string().describe("SharePoint list ID"),
      itemId: z.string().describe("List item ID"),
    },
    handler: (args: { siteId: string; listId: string; itemId: string }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists/${args.listId}/items/${args.itemId}?expand=fields`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId, listId: args.listId, itemId: args.itemId },
        queryParams: { expand: "fields" },
        body: null,
        description: `Get item ${args.itemId} from list ${args.listId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst item = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: "Column values are returned under the 'fields' property. This request already includes ?expand=fields.",
      };
    },
  },
  {
    name: "graph_sites_create_item",
    description: "Create a SharePoint list item",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
      listId: z.string().describe("SharePoint list ID to create the item in"),
      fields: z.record(z.unknown()).describe("Column values for the new item — key-value pairs matching column internal names (e.g. { \"Title\": \"My item\", \"Status\": \"Active\" })"),
    },
    handler: (args: { siteId: string; listId: string; fields: Record<string, unknown> }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists/${args.listId}/items`;
      const body = { fields: args.fields };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId, listId: args.listId },
        queryParams: {},
        body,
        description: `Create item in list ${args.listId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-create",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst item = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_update_item",
    description: "Update a SharePoint list item",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
      listId: z.string().describe("SharePoint list ID"),
      itemId: z.string().describe("List item ID to update"),
      fields: z.record(z.unknown()).describe("Column values to update — key-value pairs matching column internal names (e.g. { \"Status\": \"Completed\" })"),
    },
    handler: (args: { siteId: string; listId: string; itemId: string; fields: Record<string, unknown> }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists/${args.listId}/items/${args.itemId}/fields`;
      return {
        endpoint,
        method: "PATCH",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId, listId: args.listId, itemId: args.itemId },
        queryParams: {},
        body: args.fields,
        description: `Update fields of item ${args.itemId} in list ${args.listId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-update",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'PATCH',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(args.fields, null, 2)})\n});\nconst updated = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_delete_item",
    description: "Delete a SharePoint list item",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
      listId: z.string().describe("SharePoint list ID"),
      itemId: z.string().describe("List item ID to delete"),
    },
    handler: (args: { siteId: string; listId: string; itemId: string }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists/${args.listId}/items/${args.itemId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId, listId: args.listId, itemId: args.itemId },
        queryParams: {},
        body: null,
        description: `Delete item ${args.itemId} from list ${args.listId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_sites_list_columns",
    description: "List columns in a SharePoint list",
    category: "sites",
    zodShape: {
      siteId: z.string().describe("SharePoint site ID"),
      listId: z.string().describe("SharePoint list ID"),
    },
    handler: (args: { siteId: string; listId: string }) => {
      const endpoint = `${BASE}/sites/${args.siteId}/lists/${args.listId}/columns`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { siteId: args.siteId, listId: args.listId },
        queryParams: {},
        body: null,
        description: `List columns in list ${args.listId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/list-list-columns",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
];
