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
  delegated: ["Notes.Read", "Notes.ReadWrite"],
  application: ["Notes.Read.All", "Notes.ReadWrite.All"],
};

const WRITE_PERMISSIONS = {
  delegated: ["Notes.ReadWrite"],
  application: ["Notes.ReadWrite.All"],
};

export const notesTools: ToolDefinition[] = [
  {
    name: "graph_notes_list_notebooks",
    description: "List OneNote notebooks for a user",
    category: "notes",
    zodShape: {
      userId: z.string(),
    },
    handler: (args: { userId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/notebooks`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body: null,
        description: `List notebooks for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/onenote-list-notebooks",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_notes_get_notebook",
    description: "Get a specific OneNote notebook",
    category: "notes",
    zodShape: {
      userId: z.string(),
      notebookId: z.string(),
    },
    handler: (args: { userId: string; notebookId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/notebooks/${args.notebookId}`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, notebookId: args.notebookId },
        queryParams: {},
        body: null,
        description: `Get notebook ${args.notebookId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/notebook-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst notebook = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_notes_create_notebook",
    description: "Create a new OneNote notebook",
    category: "notes",
    zodShape: {
      userId: z.string(),
      displayName: z.string(),
    },
    handler: (args: { userId: string; displayName: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/notebooks`;
      const body = { displayName: args.displayName };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId },
        queryParams: {},
        body,
        description: `Create notebook '${args.displayName}' for user ${args.userId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/onenote-post-notebooks",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});\nconst notebook = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_notes_list_sections",
    description: "List sections in a notebook",
    category: "notes",
    zodShape: {
      userId: z.string(),
      notebookId: z.string(),
    },
    handler: (args: { userId: string; notebookId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/notebooks/${args.notebookId}/sections`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, notebookId: args.notebookId },
        queryParams: {},
        body: null,
        description: `List sections in notebook ${args.notebookId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/notebook-list-sections",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_notes_create_section",
    description: "Create a section in a notebook",
    category: "notes",
    zodShape: {
      userId: z.string(),
      notebookId: z.string(),
      displayName: z.string(),
    },
    handler: (args: { userId: string; notebookId: string; displayName: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/notebooks/${args.notebookId}/sections`;
      const body = { displayName: args.displayName };
      return {
        endpoint,
        method: "POST",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, notebookId: args.notebookId },
        queryParams: {},
        body,
        description: `Create section '${args.displayName}' in notebook ${args.notebookId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/notebook-post-sections",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'application/json' },\n  body: JSON.stringify(${JSON.stringify(body)})\n});\nconst section = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_notes_list_pages",
    description: "List pages in a OneNote section",
    category: "notes",
    zodShape: {
      userId: z.string(),
      sectionId: z.string(),
    },
    handler: (args: { userId: string; sectionId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/sections/${args.sectionId}/pages`;
      return {
        endpoint,
        method: "GET",
        headers: buildHeaders(),
        pathParams: { userId: args.userId, sectionId: args.sectionId },
        queryParams: {},
        body: null,
        description: `List pages in section ${args.sectionId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/section-list-pages",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}' }\n});\nconst data = await response.json();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_notes_create_page",
    description: "Create a page in a OneNote section",
    category: "notes",
    zodShape: {
      userId: z.string(),
      sectionId: z.string(),
      title: z.string(),
      htmlContent: z.string(),
    },
    handler: (args: { userId: string; sectionId: string; title: string; htmlContent: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/sections/${args.sectionId}/pages`;
      const htmlBody = `<!DOCTYPE html><html><head><title>${args.title}</title></head><body>${args.htmlContent}</body></html>`;
      return {
        endpoint,
        method: "POST",
        headers: {
          Authorization: "Bearer {token}",
          "Content-Type": "text/html",
        },
        pathParams: { userId: args.userId, sectionId: args.sectionId },
        queryParams: {},
        body: htmlBody,
        description: `Create page '${args.title}' in section ${args.sectionId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/section-post-pages",
        codeExample: `const html = \`${htmlBody}\`;\nconst response = await fetch('${endpoint}', {\n  method: 'POST',\n  headers: { 'Authorization': 'Bearer {token}', 'Content-Type': 'text/html' },\n  body: html\n});\nconst page = await response.json();`,
        requiredPermissions: WRITE_PERMISSIONS,
        notes: null,
      };
    },
  },
  {
    name: "graph_notes_get_page_content",
    description: "Get the HTML content of a OneNote page",
    category: "notes",
    zodShape: {
      userId: z.string(),
      pageId: z.string(),
    },
    handler: (args: { userId: string; pageId: string }) => {
      const endpoint = `${BASE}/users/${args.userId}/onenote/pages/${args.pageId}/content`;
      return {
        endpoint,
        method: "GET",
        headers: {
          Authorization: "Bearer {token}",
          Accept: "text/html",
        },
        pathParams: { userId: args.userId, pageId: args.pageId },
        queryParams: {},
        body: null,
        description: `Get HTML content of page ${args.pageId}.`,
        docsUrl: "https://learn.microsoft.com/en-us/graph/api/page-get",
        codeExample: `const response = await fetch('${endpoint}', {\n  method: 'GET',\n  headers: { 'Authorization': 'Bearer {token}', 'Accept': 'text/html' }\n});\nconst html = await response.text();`,
        requiredPermissions: READ_PERMISSIONS,
        notes: null,
      };
    },
  },
];
