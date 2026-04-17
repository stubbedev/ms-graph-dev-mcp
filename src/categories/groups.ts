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
  delegated: ["Group.Read.All", "Directory.Read.All"],
  application: ["Group.Read.All", "Directory.Read.All"],
};

const READWRITE_PERMISSIONS = {
  delegated: ["Group.ReadWrite.All"],
  application: ["Group.ReadWrite.All"],
};

const MEMBER_READ_PERMISSIONS = {
  delegated: ["GroupMember.Read.All", "Group.Read.All"],
  application: ["GroupMember.Read.All", "Group.Read.All"],
};

const MEMBER_READWRITE_PERMISSIONS = {
  delegated: ["GroupMember.ReadWrite.All", "Group.ReadWrite.All"],
  application: ["GroupMember.ReadWrite.All", "Group.ReadWrite.All"],
};

export const groupsTools: ToolDefinition[] = [
  {
    name: "graph_groups_list",
    description: "List all groups",
    category: "groups",
    zodShape: {
      filter: z.string().optional(),
      select: z.array(z.string()).optional(),
      top: z.number().optional(),
    },
    handler: (args: { filter?: string; select?: string[]; top?: number }) => {
      const params: string[] = [];
      if (args.filter) params.push(`$filter=${encodeURIComponent(args.filter)}`);
      if (args.select?.length) params.push(`$select=${args.select.join(",")}`);
      if (args.top) params.push(`$top=${args.top}`);
      const qs = params.length ? `?${params.join("&")}` : "";
      const endpoint = `${BASE}/groups${qs}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body: null,
        description: "List all groups in the tenant.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_groups_get",
    description: "Get a group by ID",
    category: "groups",
    zodShape: {
      groupId: z.string(),
      select: z.array(z.string()).optional(),
    },
    handler: (args: { groupId: string; select?: string[] }) => {
      const selectParam = args.select?.length ? `?$select=${args.select.join(",")}` : "";
      const endpoint = `${BASE}/groups/${args.groupId}${selectParam}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { groupId: args.groupId },
        queryParams: args.select?.length ? { $select: args.select.join(",") } : {},
        body: null,
        description: `Get group ${args.groupId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst group = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_groups_create",
    description: "Create a new group",
    category: "groups",
    zodShape: {
      displayName: z.string(),
      mailNickname: z.string(),
      description: z.string().optional(),
      groupTypes: z.array(z.string()).optional(),
      mailEnabled: z.boolean(),
      securityEnabled: z.boolean(),
    },
    handler: (args: {
      displayName: string;
      mailNickname: string;
      description?: string;
      groupTypes?: string[];
      mailEnabled: boolean;
      securityEnabled: boolean;
    }) => {
      const endpoint = `${BASE}/groups`;
      const body: Record<string, unknown> = {
        displayName: args.displayName,
        mailNickname: args.mailNickname,
        mailEnabled: args.mailEnabled,
        securityEnabled: args.securityEnabled,
        groupTypes: args.groupTypes ?? [],
      };
      if (args.description) body.description = args.description;
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body,
        description: `Create group '${args.displayName}'.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-post-groups",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst group = await response.json();`,
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_groups_delete",
    description: "Delete a group",
    category: "groups",
    zodShape: {
      groupId: z.string(),
    },
    handler: (args: { groupId: string }) => {
      const endpoint = `${BASE}/groups/${args.groupId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { groupId: args.groupId },
        queryParams: {},
        body: null,
        description: `Delete group ${args.groupId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
        requiredPermissions: READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_groups_list_members",
    description: "List members of a group",
    category: "groups",
    zodShape: {
      groupId: z.string(),
    },
    handler: (args: { groupId: string }) => {
      const endpoint = `${BASE}/groups/${args.groupId}/members`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { groupId: args.groupId },
        queryParams: {},
        body: null,
        description: `List members of group ${args.groupId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list-members",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: MEMBER_READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_groups_add_member",
    description: "Add a member to a group",
    category: "groups",
    zodShape: {
      groupId: z.string(),
      userId: z.string(),
    },
    handler: (args: { groupId: string; userId: string }) => {
      const endpoint = `${BASE}/groups/${args.groupId}/members/$ref`;
      const body = { "@odata.id": `${BASE}/directoryObjects/${args.userId}` };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { groupId: args.groupId },
        queryParams: {},
        body,
        description: `Add user ${args.userId} to group ${args.groupId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-post-members",
        codeExample: `await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});`,
        requiredPermissions: MEMBER_READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_groups_remove_member",
    description: "Remove a member from a group",
    category: "groups",
    zodShape: {
      groupId: z.string(),
      userId: z.string(),
    },
    handler: (args: { groupId: string; userId: string }) => {
      const endpoint = `${BASE}/groups/${args.groupId}/members/${args.userId}/$ref`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { groupId: args.groupId, userId: args.userId },
        queryParams: {},
        body: null,
        description: `Remove user ${args.userId} from group ${args.groupId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-delete-members",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
        requiredPermissions: MEMBER_READWRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_groups_list_owners",
    description: "List owners of a group",
    category: "groups",
    zodShape: {
      groupId: z.string(),
    },
    handler: (args: { groupId: string }) => {
      const endpoint = `${BASE}/groups/${args.groupId}/owners`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { groupId: args.groupId },
        queryParams: {},
        body: null,
        description: `List owners of group ${args.groupId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list-owners",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: MEMBER_READ_PERMISSIONS,
        notes: null,
      };
    },
  },
];
