import { z } from "zod";
import { ToolDefinition } from "../registry.js";

const BASE = "https://graph.microsoft.com/v1.0";

function buildHeaders(): Record<string, string> {
  return {
    Authorization: "Bearer {token}",
    "Content-Type": "application/json",
  };
}

function driveBasePath(args: { driveId?: string; userId?: string }): string {
  if (args.driveId) return `/drives/${args.driveId}`;
  if (args.userId) return `/users/${args.userId}/drive`;
  return "/me/drive";
}

const READ_PERMISSIONS = {
  delegated: ["Files.Read", "Files.Read.All", "Sites.Read.All"],
  application: ["Files.Read.All", "Sites.Read.All"],
};

const WRITE_PERMISSIONS = {
  delegated: ["Files.ReadWrite", "Files.ReadWrite.All", "Sites.ReadWrite.All"],
  application: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
};

export const filesTools: ToolDefinition[] = [
  {
    name: "graph_files_list_children",
    description: "List children of a drive item (folder)",
    category: "files",
    zodShape: {
      driveId: z.string().optional(),
      itemId: z.string().optional().describe("Item ID or 'root'"),
      userId: z.string().optional(),
    },
    handler: (args: { driveId?: string; itemId?: string; userId?: string }) => {
      const base = driveBasePath(args);
      const item = args.itemId ?? "root";
      const endpoint = `${BASE}${base}/items/${item}/children`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { driveId: args.driveId, itemId: item },
        queryParams: {},
        body: null,
        description: "List items inside a drive folder.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-list-children",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_get_item",
    description: "Get a drive item by ID",
    category: "files",
    zodShape: {
      driveId: z.string().optional(),
      itemId: z.string(),
      userId: z.string().optional(),
    },
    handler: (args: { driveId?: string; itemId: string; userId?: string }) => {
      const base = driveBasePath(args);
      const endpoint = `${BASE}${base}/items/${args.itemId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { itemId: args.itemId },
        queryParams: {},
        body: null,
        description: `Get drive item ${args.itemId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst item = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_upload_small",
    description: "Upload a small file (<4MB) using PUT",
    category: "files",
    zodShape: {
      fileName: z.string(),
      parentId: z.string().optional(),
      userId: z.string().optional(),
      contentType: z.string().optional(),
    },
    handler: (args: { fileName: string; parentId?: string; userId?: string; contentType?: string }) => {
      const base = driveBasePath(args);
      const parent = args.parentId ?? "root";
      const endpoint = `${BASE}${base}/items/${parent}:/${args.fileName}:/content`;
      const headers: Record<string, string> = {
        Authorization: "Bearer {token}",
        "Content-Type": args.contentType ?? "application/octet-stream",
      };
      return {
        endpoint,
        method: "PUT",
        headers,
        pathParams: { parentId: parent, fileName: args.fileName },
        queryParams: {},
        body: "{file binary content}",
        description: `Upload file '${args.fileName}' to parent item ${parent}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-put-content",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'PUT',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': '${args.contentType ?? "application/octet-stream"}' },\n  body: fileBuffer\n});\nconst item = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: "Only supports files under 4 MB. For larger files use graph_files_create_upload_session to get a resumable upload URL, then PUT each chunk to that URL.",
      };
    },
  },
  {
    name: "graph_files_create_upload_session",
    description: "Create an upload session for large files (>4MB)",
    category: "files",
    zodShape: {
      fileName: z.string(),
      parentId: z.string().optional(),
      userId: z.string().optional(),
    },
    handler: (args: { fileName: string; parentId?: string; userId?: string }) => {
      const base = driveBasePath(args);
      const parent = args.parentId ?? "root";
      const endpoint = `${BASE}${base}/items/${parent}:/${args.fileName}:/createUploadSession`;
      const body = {
        item: {
          "@microsoft.graph.conflictBehavior": "rename",
          name: args.fileName,
        },
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { parentId: parent, fileName: args.fileName },
        queryParams: {},
        body,
        description: `Create an upload session for large file '${args.fileName}'.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst session = await response.json();\n// Then upload chunks to session.uploadUrl`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_create_folder",
    description: "Create a new folder in a drive",
    category: "files",
    zodShape: {
      folderName: z.string(),
      parentId: z.string().optional(),
      userId: z.string().optional(),
    },
    handler: (args: { folderName: string; parentId?: string; userId?: string }) => {
      const base = driveBasePath(args);
      const parent = args.parentId ?? "root";
      const endpoint = `${BASE}${base}/items/${parent}/children`;
      const body = {
        name: args.folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { parentId: parent },
        queryParams: {},
        body,
        description: `Create folder '${args.folderName}'.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-post-children",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body, null, 2)})\n});\nconst folder = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_delete_item",
    description: "Delete a drive item",
    category: "files",
    zodShape: {
      itemId: z.string(),
      userId: z.string().optional(),
    },
    handler: (args: { itemId: string; userId?: string }) => {
      const base = driveBasePath(args);
      const endpoint = `${BASE}${base}/items/${args.itemId}`;
      return {
        endpoint,
        method: "DELETE",
        headers: buildHeaders(),
        pathParams: { itemId: args.itemId },
        queryParams: {},
        body: null,
        description: `Delete drive item ${args.itemId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-delete",
        codeExample: `await fetch('${endpoint}', {\n  method: 'DELETE',\n  headers: { 'Authorization': 'Bearer {token}' }\n});`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_move_item",
    description: "Move a drive item to a new parent",
    category: "files",
    zodShape: {
      itemId: z.string(),
      destinationParentId: z.string(),
      userId: z.string().optional(),
    },
    handler: (args: { itemId: string; destinationParentId: string; userId?: string }) => {
      const base = driveBasePath(args);
      const endpoint = `${BASE}${base}/items/${args.itemId}`;
      const body = { parentReference: { id: args.destinationParentId } };
      return {
        endpoint,
        method: "PATCH",
        headers: buildHeaders(),
        pathParams: { itemId: args.itemId },
        queryParams: {},
        body,
        description: `Move item ${args.itemId} to parent ${args.destinationParentId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-move",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'PATCH',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});\nconst item = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_copy_item",
    description: "Copy a drive item to a new location",
    category: "files",
    zodShape: {
      itemId: z.string(),
      destinationParentId: z.string(),
      newName: z.string().optional(),
      userId: z.string().optional(),
    },
    handler: (args: { itemId: string; destinationParentId: string; newName?: string; userId?: string }) => {
      const base = driveBasePath(args);
      const endpoint = `${BASE}${base}/items/${args.itemId}/copy`;
      const body: Record<string, unknown> = { parentReference: { id: args.destinationParentId } };
      if (args.newName) body.name = args.newName;
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { itemId: args.itemId },
        queryParams: {},
        body,
        description: `Copy item ${args.itemId} to parent ${args.destinationParentId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-copy",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});\n// Returns 202 Accepted with monitor URL`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: "Returns 202 Accepted immediately. Poll the URL in the Location response header to track copy completion — the copy may take time for large files.",
      };
    },
  },
  {
    name: "graph_files_search",
    description: "Search for files in OneDrive",
    category: "files",
    zodShape: {
      query: z.string(),
      userId: z.string().optional(),
    },
    handler: (args: { query: string; userId?: string }) => {
      const base = driveBasePath(args);
      const endpoint = `${BASE}${base}/root/search(q='${encodeURIComponent(args.query)}')`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: { q: args.query },
        body: null,
        description: `Search for files matching '${args.query}'.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-search",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_get_download_url",
    description: "Get a download URL for a drive item",
    category: "files",
    zodShape: {
      itemId: z.string(),
      userId: z.string().optional(),
    },
    handler: (args: { itemId: string; userId?: string }) => {
      const base = driveBasePath(args);
      const endpoint = `${BASE}${base}/items/${args.itemId}?$select=id,@microsoft.graph.downloadUrl`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { itemId: args.itemId },
        queryParams: { "$select": "id,@microsoft.graph.downloadUrl" },
        body: null,
        description: `Get the download URL for item ${args.itemId}. Use the '@microsoft.graph.downloadUrl' field from the response.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst item = await response.json();\nconst downloadUrl = item['@microsoft.graph.downloadUrl'];`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_files_get_delta",
    description: "Get changes to a drive (OneDrive or SharePoint document library) since a previous delta token — returns only added, modified, or deleted items.",
    category: "files",
    zodShape: {
      deltaToken: z.string().optional(),
      driveId: z.string().optional(),
      userId: z.string().optional(),
    },
    handler: (args: { deltaToken?: string; driveId?: string; userId?: string }) => {
      let endpoint: string;
      if (args.deltaToken && args.deltaToken.startsWith("http")) {
        endpoint = args.deltaToken;
      } else {
        const base = driveBasePath(args);
        endpoint = `${BASE}${base}/root/delta`;
        if (args.deltaToken) {
          endpoint += `?$deltatoken=${args.deltaToken}`;
        }
      }
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: {},
        queryParams: {},
        body: null,
        description: "Get delta changes to drive items since the last sync.",
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-delta",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();\n// Store data['@odata.deltaLink'] for next call`,
        requiredPermissions: READ_PERMISSIONS,
        notes: "On first call, returns all drive items plus @odata.deltaLink at the end. Use that token for subsequent calls to get only changes. Items with @removed annotation have been deleted.",
      };
    },
  },
];
