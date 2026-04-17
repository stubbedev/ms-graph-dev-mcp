import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

export const usersTools: ToolDefinition[] = [
  {
    name: "graph_users_get",
    description: "Get a user by ID or UPN",
    category: "users",
    zodShape: {
      userId: z.string().describe("User ID or UPN"),
      select: z.array(z.string()).optional().describe("Fields to select"),
    },
    handler: (args: { userId: string; select?: string[] }) => {
      const selectParam = args.select?.length ? `?$select=${args.select.join(",")}` : "";
      const endpoint = `${BASE}/users/${args.userId}${selectParam}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: args.select?.length ? { $select: args.select.join(",") } : {},
        body: null,
        description: `Get properties for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst user = await response.json();`,
        requiredPermissions: {
          delegated: ["User.Read", "User.Read.All", "Directory.Read.All"],
          application: ["User.Read.All", "Directory.Read.All"],
        },
        notes: null,
      };
    },
  },
  {
    name: "graph_users_list",
    description: "List all users",
    category: "users",
    zodShape: {
      filter: z.string().optional().describe("OData filter expression"),
      select: z.array(z.string()).optional().describe("Fields to select"),
      top: z.number().optional().describe("Max number of results"),
      skip: z.number().optional().describe("Number of results to skip"),
    },
    handler: (args: { filter?: string; select?: string[]; top?: number; skip?: number }) => {
      const params: string[] = [];
      if (args.filter) params.push(`$filter=${encodeURIComponent(args.filter)}`);
      if (args.select?.length) params.push(`$select=${args.select.join(",")}`);
      if (args.top) params.push(`$top=${args.top}`);
      if (args.skip) params.push(`$skip=${args.skip}`);
      const qs = params.length ? `?${params.join("&")}` : "";
      const endpoint = `${BASE}/users${qs}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: Object.fromEntries(params.map((p) => p.split("="))),
        body: null,
        description: "List users in the tenant.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: {
          delegated: ["User.Read.All", "Directory.Read.All"],
          application: ["User.Read.All", "Directory.Read.All"],
        },
        notes: null,
      };
    },
  },
  {
    name: "graph_users_create",
    description: "Create a new user",
    category: "users",
    zodShape: {
      displayName: z.string(),
      mailNickname: z.string(),
      userPrincipalName: z.string(),
      password: z.string(),
      accountEnabled: z.boolean().optional(),
    },
    handler: (args: {
      displayName: string;
      mailNickname: string;
      userPrincipalName: string;
      password: string;
      accountEnabled?: boolean;
    }) => {
      const endpoint = `${BASE}/users`;
      const body = {
        accountEnabled: args.accountEnabled ?? true,
        displayName: args.displayName,
        mailNickname: args.mailNickname,
        userPrincipalName: args.userPrincipalName,
        passwordProfile: {
          forceChangePasswordNextSignIn: true,
          password: args.password,
        },
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body,
        description: "Create a new user.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-users",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst user = await response.json();`,
        requiredPermissions: {
          delegated: ["User.ReadWrite.All"],
          application: ["User.ReadWrite.All"],
        },
        notes: null,
      };
    },
  },
  {
    name: "graph_users_update",
    description: "Update user properties",
    category: "users",
    zodShape: {
      userId: z.string(),
      displayName: z.string().optional(),
      jobTitle: z.string().optional(),
      department: z.string().optional(),
      mobilePhone: z.string().optional(),
      officeLocation: z.string().optional(),
      preferredLanguage: z.string().optional(),
    },
    handler: (args: {
      userId: string;
      displayName?: string;
      jobTitle?: string;
      department?: string;
      mobilePhone?: string;
      officeLocation?: string;
      preferredLanguage?: string;
    }) => {
      const { userId, ...rest } = args;
      const endpoint = `${BASE}/users/${userId}`;
      const body = Object.fromEntries(Object.entries(rest).filter(([, v]) => v !== undefined));
      return {
        endpoint,
        method: "PATCH",
        headers: buildHeaders(),
        pathParams: { userId },
        queryParams: {},
        body,
        description: `Update properties of the user ${userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-update",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'PATCH',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});`,
        requiredPermissions: {
          delegated: ["User.ReadWrite.All"],
          application: ["User.ReadWrite.All"],
        },
        notes: null,
      };
    },
  },
  {
    name: "graph_users_delete",
    description: "Delete a user",
    category: "users",
    zodShape: {
      userId: z.string(),
    },
    handler: (args: { userId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `Delete user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
        requiredPermissions: {
          delegated: ["User.ReadWrite.All"],
          application: ["User.ReadWrite.All"],
        },
        notes: null,
      };
    },
  },
  {
    name: "graph_users_get_manager",
    description: "Get the manager of a user",
    category: "users",
    zodShape: {
      userId: z.string(),
    },
    handler: (args: { userId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/manager`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `Get the manager of user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-manager",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst manager = await response.json();`,
        requiredPermissions: {
          delegated: ["User.Read.All", "Directory.Read.All"],
          application: ["User.Read.All", "Directory.Read.All"],
        },
        notes: null,
      };
    },
  },
  {
    name: "graph_users_list_direct_reports",
    description: "List direct reports of a user",
    category: "users",
    zodShape: {
      userId: z.string(),
    },
    handler: (args: { userId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/directReports`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `List the direct reports of user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-directreports",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: {
          delegated: ["User.Read.All", "Directory.Read.All"],
          application: ["User.Read.All", "Directory.Read.All"],
        },
        notes: null,
      };
    },
  },
  {
    name: "graph_users_get_delta",
    description: "Get changes to users in Azure AD since a previous delta token — returns only added, updated, or deleted users.",
    category: "users",
    zodShape: {
      deltaToken: z.string().optional().describe("Token from previous delta response. Omit for initial sync."),
      select: z.array(z.string()).optional(),
    },
    handler: (args: { deltaToken?: string; select?: string[] }) => {
      let endpoint: string;
      if (args.deltaToken) {
        endpoint = args.deltaToken;
      } else {
        const params: string[] = [];
        if (args.select?.length) params.push(`$select=${args.select.join(",")}`);
        const qs = params.length ? `?${params.join("&")}` : "";
        endpoint = `${BASE}/users/delta${qs}`;
      }
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: args.select?.length ? { $select: args.select.join(",") } : {},
        body: null,
        description: "Get delta changes to users since the last sync.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-delta",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();\n// Store data['@odata.deltaLink'] for next call`,
        requiredPermissions: {
          delegated: ["User.Read.All", "Directory.Read.All"],
          application: ["User.Read.All", "Directory.Read.All"],
        },
        notes: "On initial call (no deltaToken), response includes all users and ends with @odata.deltaLink. Store that token and pass it on subsequent calls to get only changes. @removed in the response marks deleted users.",
      };
    },
  },
];
