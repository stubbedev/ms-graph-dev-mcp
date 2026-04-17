import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { ToolRegistry } from "./registry.js";
import {
  usersTools,
  filesTools,
  mailTools,
  calendarTools,
  groupsTools,
  notesTools,
  tasksTools,
  sitesTools,
} from "./categories/index.js";
import type { ToolDefinition } from "./registry.js";

// Static search lookup for search_graph_api bootstrap tool
const GRAPH_API_SEARCH_MAP: Array<{ keywords: string[]; endpoint: string; method: string; description: string; docsUrl: string }> = [
  { keywords: ["user", "users", "get user", "list users", "find user"], endpoint: "/users or /users/{id}", method: "GET", description: "Get or list users in the tenant", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-get" },
  { keywords: ["me", "current user", "signed in", "profile"], endpoint: "/me", method: "GET", description: "Get the signed-in user's profile", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-get" },
  { keywords: ["create user", "new user", "register user"], endpoint: "/users", method: "POST", description: "Create a new user", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-users" },
  { keywords: ["update user", "patch user", "modify user"], endpoint: "/users/{id}", method: "PATCH", description: "Update user properties", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-update" },
  { keywords: ["delete user", "remove user"], endpoint: "/users/{id}", method: "DELETE", description: "Delete a user", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-delete" },
  { keywords: ["send mail", "send email", "email", "sendmail"], endpoint: "/users/{id}/sendMail", method: "POST", description: "Send an email message", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-sendmail" },
  { keywords: ["list messages", "inbox", "emails", "messages"], endpoint: "/users/{id}/messages", method: "GET", description: "List messages in a mailbox", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-messages" },
  { keywords: ["calendar", "events", "list events", "meetings"], endpoint: "/users/{id}/events", method: "GET", description: "List calendar events", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-events" },
  { keywords: ["create event", "new event", "schedule meeting", "book meeting"], endpoint: "/users/{id}/events", method: "POST", description: "Create a calendar event", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-events" },
  { keywords: ["group", "groups", "list groups", "find group"], endpoint: "/groups", method: "GET", description: "List groups in the tenant", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list" },
  { keywords: ["create group", "new group", "microsoft 365 group"], endpoint: "/groups", method: "POST", description: "Create a new group", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-post-groups" },
  { keywords: ["group members", "members", "list members", "member of"], endpoint: "/groups/{id}/members", method: "GET", description: "List group members", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list-members" },
  { keywords: ["add member", "join group"], endpoint: "/groups/{id}/members/$ref", method: "POST", description: "Add a member to a group", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-post-members" },
  { keywords: ["files", "onedrive", "documents", "drive"], endpoint: "/me/drive/root/children", method: "GET", description: "List files in OneDrive root", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-list-children" },
  { keywords: ["upload file", "upload", "put file"], endpoint: "/me/drive/items/{parentId}:/{filename}:/content", method: "PUT", description: "Upload a file to OneDrive", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-put-content" },
  { keywords: ["download file", "get file", "download url"], endpoint: "/me/drive/items/{id}?$select=id,@microsoft.graph.downloadUrl", method: "GET", description: "Get file download URL", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-get" },
  { keywords: ["sharepoint", "site", "sites", "sharepoint site"], endpoint: "/sites/{siteId}", method: "GET", description: "Get a SharePoint site", docsUrl: "https://learn.microsoft.com/en-us/graph/api/site-get" },
  { keywords: ["list items", "sharepoint list", "sp list"], endpoint: "/sites/{siteId}/lists/{listId}/items", method: "GET", description: "List items in a SharePoint list", docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-list" },
  { keywords: ["create list item", "add item", "new item"], endpoint: "/sites/{siteId}/lists/{listId}/items", method: "POST", description: "Create a SharePoint list item", docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-create" },
  { keywords: ["planner", "tasks", "todo", "task list"], endpoint: "/planner/plans/{planId}/tasks", method: "GET", description: "List tasks in a Planner plan", docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks" },
  { keywords: ["create task", "new task", "add task"], endpoint: "/planner/tasks", method: "POST", description: "Create a Planner task", docsUrl: "https://learn.microsoft.com/en-us/graph/api/planner-post-tasks" },
  { keywords: ["onenote", "notebook", "notes"], endpoint: "/users/{id}/onenote/notebooks", method: "GET", description: "List OneNote notebooks", docsUrl: "https://learn.microsoft.com/en-us/graph/api/onenote-list-notebooks" },
  { keywords: ["manager", "reports", "org chart", "hierarchy"], endpoint: "/users/{id}/manager", method: "GET", description: "Get a user's manager", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-manager" },
  { keywords: ["free busy", "schedule", "availability", "find meeting"], endpoint: "/users/{id}/calendar/getSchedule", method: "POST", description: "Get free/busy schedule", docsUrl: "https://learn.microsoft.com/en-us/graph/api/calendar-getschedule" },
  { keywords: ["folder", "create folder", "new folder", "mkdir"], endpoint: "/me/drive/items/{parentId}/children", method: "POST", description: "Create a folder in OneDrive", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-post-children" },
  { keywords: ["mail folder", "inbox", "sent items", "drafts", "junk"], endpoint: "/users/{id}/mailFolders", method: "GET", description: "List mail folders", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders" },
  { keywords: ["reply", "reply to email", "reply to message"], endpoint: "/users/{id}/messages/{messageId}/reply", method: "POST", description: "Reply to a message", docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-reply" },
  { keywords: ["large file", "upload session", "resumable upload"], endpoint: "/me/drive/items/{parentId}:/{filename}:/createUploadSession", method: "POST", description: "Create an upload session for large files", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession" },
  { keywords: ["move file", "move item", "relocate"], endpoint: "/me/drive/items/{itemId}", method: "PATCH", description: "Move a drive item by updating parentReference", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-move" },
  { keywords: ["copy file", "copy item", "duplicate"], endpoint: "/me/drive/items/{itemId}/copy", method: "POST", description: "Copy a drive item", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-copy" },
];

const CATEGORY_DESCRIPTIONS: Record<string, string> = {
  users: "Manage Azure AD / Entra ID users — get, list, create, update, delete users; retrieve manager and direct reports",
  files: "Manage OneDrive and SharePoint document library files — upload, download, list, create folders, move, copy, search drive items",
  mail: "Manage Exchange Online mail — list and get messages, send mail, create drafts, reply, delete, manage folders",
  calendar: "Manage Outlook calendar — list, create, update, delete events; find meeting times; get free/busy schedule",
  groups: "Manage Microsoft 365 and security groups — list, create, delete groups; add and remove members and owners",
  notes: "Manage OneNote — list notebooks, sections, and pages; create notebooks, sections, and pages; get page content",
  tasks: "Manage Planner and Microsoft To Do — list and create plans and tasks; update and delete Planner tasks; manage To Do lists",
  sites: "Manage SharePoint sites and lists — get sites, create and query lists, create/read/update/delete list items, list columns",
};

const CATEGORY_TOOL_MAP: Record<string, ToolDefinition[]> = {
  users: usersTools,
  files: filesTools,
  mail: mailTools,
  calendar: calendarTools,
  groups: groupsTools,
  notes: notesTools,
  tasks: tasksTools,
  sites: sitesTools,
};

export class GraphMcpServer {
  private mcpServer: McpServer;
  private registry: ToolRegistry;
  // Store registered McpServer tool handles for category tools so we can remove them
  private categoryHandles: Map<string, ReturnType<McpServer["tool"]>[]> = new Map();

  constructor() {
    this.registry = new ToolRegistry();
    this.mcpServer = new McpServer(
      { name: "ms-graph-dev-assistant", version: "0.1.0" },
      {
        capabilities: {
          tools: { listChanged: true },
        },
        instructions:
          "Use this server whenever the user mentions SharePoint, Microsoft Graph, OneDrive, Exchange Online, Office 365, Microsoft 365, Azure Active Directory, Entra ID, Microsoft Teams, Planner, OneNote, or Microsoft To Do. " +
          "This server constructs and validates Microsoft Graph REST API requests — it does not execute them. " +
          "Call list_categories to see available resource areas, then load_category to activate tools for the relevant area.",
      }
    );

    this.registerBootstrapTools();
  }

  private registerBootstrapTools(): void {
    // list_categories
    this.mcpServer.tool(
      "list_categories",
      "List all Microsoft Graph API resource categories — SharePoint sites/lists, OneDrive files, Exchange mail, Outlook calendar, Azure AD users, M365 groups, OneNote, Planner, and To Do. Shows which categories are currently loaded.",
      {},
      async () => {
        const categories = Object.entries(CATEGORY_DESCRIPTIONS).map(([name, description]) => ({
          name,
          description,
          loaded: this.registry.loadedCategories.has(name),
          toolCount: CATEGORY_TOOL_MAP[name]?.length ?? 0,
        }));
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({ categories }, null, 2),
            },
          ],
        };
      }
    );

    // load_category
    this.mcpServer.tool(
      "load_category",
      "Load Graph API construction tools for a resource category. Use 'sites' for SharePoint sites, lists, and list items; 'files' for OneDrive and SharePoint document libraries; 'users' for Azure AD / Entra ID users; 'mail' for Exchange Online; 'calendar' for Outlook calendar; 'groups' for Microsoft 365 groups; 'notes' for OneNote; 'tasks' for Planner and Microsoft To Do. Sends tools/list_changed after loading.",
      { category: z.string().describe("Category to load: users, files, mail, calendar, groups, notes, tasks, or sites") },
      async ({ category }) => {
        const normalizedCategory = category.toLowerCase().trim();

        if (!CATEGORY_TOOL_MAP[normalizedCategory]) {
          const available = Object.keys(CATEGORY_TOOL_MAP).join(", ");
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify({
                  error: `Unknown category '${normalizedCategory}'. Available: ${available}`,
                }),
              },
            ],
          };
        }

        if (this.registry.loadedCategories.has(normalizedCategory)) {
          const toolNames = this.registry.getAll()
            .filter((t) => t.category === normalizedCategory)
            .map((t) => t.name);
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify({
                  message: `Category '${normalizedCategory}' is already loaded.`,
                  tools: toolNames,
                }),
              },
            ],
          };
        }

        const tools = CATEGORY_TOOL_MAP[normalizedCategory];
        const newToolNames = this.registry.registerCategory(normalizedCategory, tools);
        const handles = this.registerCategoryTools(normalizedCategory, tools);
        this.categoryHandles.set(normalizedCategory, handles);

        // Notify clients that tool list changed
        this.mcpServer.sendToolListChanged();

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({
                message: `Loaded ${newToolNames.length} tools for category '${normalizedCategory}'.`,
                tools: newToolNames,
              }),
            },
          ],
        };
      }
    );

    // search_graph_api
    this.mcpServer.tool(
      "search_graph_api",
      "Search Microsoft Graph REST API endpoints by keyword. Covers SharePoint sites and lists, OneDrive files, Exchange Online mail, Outlook calendar, Azure AD users, Microsoft 365 groups, Teams, Planner, OneNote, and To Do.",
      { query: z.string().describe("Search terms, e.g. 'sharepoint list items', 'upload file onedrive', 'send email', 'find meeting times'") },
      async ({ query }) => {
        const lowerQuery = query.toLowerCase();
        const queryWords = lowerQuery.split(/\s+/);

        const results = GRAPH_API_SEARCH_MAP
          .map((entry) => {
            const score = entry.keywords.reduce((acc, keyword) => {
              if (lowerQuery.includes(keyword)) return acc + 2;
              const keywordWords = keyword.split(/\s+/);
              const matchCount = keywordWords.filter((kw) => queryWords.some((qw) => qw.includes(kw) || kw.includes(qw))).length;
              return acc + matchCount;
            }, 0);
            return { ...entry, score };
          })
          .filter((r) => r.score > 0)
          .sort((a, b) => b.score - a.score)
          .slice(0, 5)
          .map(({ score: _score, ...rest }) => rest);

        if (results.length === 0) {
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify({
                  message: "No results found. Try keywords like: user, email, calendar, files, groups, sharepoint, tasks, onenote",
                  results: [],
                }),
              },
            ],
          };
        }

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({ query, results }, null, 2),
            },
          ],
        };
      }
    );
  }

  private registerCategoryTools(
    _categoryName: string,
    tools: ToolDefinition[]
  ): ReturnType<McpServer["tool"]>[] {
    const handles: ReturnType<McpServer["tool"]>[] = [];

    for (const tool of tools) {
      const handle = this.mcpServer.tool(
        tool.name,
        tool.description,
        tool.zodShape,
        async (args: Record<string, unknown>) => {
          try {
            const result = tool.handler(args);
            return {
              content: [
                {
                  type: "text" as const,
                  text: JSON.stringify(result, null, 2),
                },
              ],
            };
          } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return {
              content: [
                {
                  type: "text" as const,
                  text: JSON.stringify({ error: message }),
                },
              ],
              isError: true,
            };
          }
        }
      );
      handles.push(handle);
    }

    return handles;
  }

  async start(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.mcpServer.connect(transport);
  }
}
