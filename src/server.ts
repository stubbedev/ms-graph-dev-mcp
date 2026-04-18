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
  subscriptionsTools,
} from "./categories/index.js";
import type { ToolDefinition } from "./registry.js";

// Static search lookup for search_graph_api bootstrap tool
const GRAPH_API_SEARCH_MAP: Array<{ keywords: string[]; endpoint: string; method: string; description: string; docsUrl: string; category: string }> = [
  // users
  { keywords: ["user", "users", "get user", "list users", "find user", "lookup user", "user profile", "directory user", "entra user", "aad user"], endpoint: "/users or /users/{id}", method: "GET", description: "Get or list users in the tenant", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-get", category: "users" },
  { keywords: ["me", "current user", "signed in", "my profile", "self", "whoami"], endpoint: "/me", method: "GET", description: "Get the signed-in user's profile", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-get", category: "users" },
  { keywords: ["create user", "new user", "register user", "provision user", "add user"], endpoint: "/users", method: "POST", description: "Create a new user", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-users", category: "users" },
  { keywords: ["update user", "patch user", "modify user", "change user", "edit user"], endpoint: "/users/{id}", method: "PATCH", description: "Update user properties", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-update", category: "users" },
  { keywords: ["delete user", "remove user", "deactivate user"], endpoint: "/users/{id}", method: "DELETE", description: "Delete a user", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-delete", category: "users" },
  { keywords: ["manager", "reports", "direct reports", "org chart", "hierarchy", "reporting line"], endpoint: "/users/{id}/manager", method: "GET", description: "Get a user's manager", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-manager", category: "users" },
  { keywords: ["user delta", "user changes", "sync users", "track user changes"], endpoint: "/users/delta", method: "GET", description: "Get only changed users since last sync", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-delta", category: "users" },
  // mail
  { keywords: ["send mail", "send email", "send message", "email user", "sendmail", "compose email"], endpoint: "/users/{id}/sendMail", method: "POST", description: "Send an email message", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-sendmail", category: "mail" },
  { keywords: ["list messages", "inbox", "emails", "messages", "read email", "get emails", "check email", "unread"], endpoint: "/users/{id}/messages", method: "GET", description: "List messages in a mailbox", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-messages", category: "mail" },
  { keywords: ["reply", "reply to email", "reply to message", "respond to email"], endpoint: "/users/{id}/messages/{messageId}/reply", method: "POST", description: "Reply to a message", docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-reply", category: "mail" },
  { keywords: ["draft", "create draft", "save draft", "draft email"], endpoint: "/users/{id}/messages", method: "POST", description: "Create a draft email message", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-messages", category: "mail" },
  { keywords: ["mail folder", "inbox folder", "sent items", "drafts folder", "junk", "archive", "mailbox folder"], endpoint: "/users/{id}/mailFolders", method: "GET", description: "List mail folders", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders", category: "mail" },
  { keywords: ["move email", "move message", "move to folder"], endpoint: "/users/{id}/messages/{id}/move", method: "POST", description: "Move a message to a different folder", docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-move", category: "mail" },
  { keywords: ["mail delta", "email changes", "sync mail", "track mail changes"], endpoint: "/users/{id}/mailFolders/{id}/messages/delta", method: "GET", description: "Get only changed messages since last sync", docsUrl: "https://learn.microsoft.com/en-us/graph/api/message-delta", category: "mail" },
  // calendar
  { keywords: ["calendar", "events", "list events", "meetings", "appointments", "outlook calendar"], endpoint: "/users/{id}/events", method: "GET", description: "List calendar events", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-list-events", category: "calendar" },
  { keywords: ["create event", "new event", "schedule meeting", "book meeting", "add appointment", "invite"], endpoint: "/users/{id}/events", method: "POST", description: "Create a calendar event", docsUrl: "https://learn.microsoft.com/en-us/graph/api/user-post-events", category: "calendar" },
  { keywords: ["free busy", "availability", "find meeting", "meeting times", "when is free", "schedule finder"], endpoint: "/users/{id}/calendar/getSchedule", method: "POST", description: "Get free/busy schedule", docsUrl: "https://learn.microsoft.com/en-us/graph/api/calendar-getschedule", category: "calendar" },
  { keywords: ["calendar delta", "calendar changes", "sync calendar", "track event changes"], endpoint: "/me/calendarView/delta", method: "GET", description: "Get only changed calendar events since last sync", docsUrl: "https://learn.microsoft.com/en-us/graph/api/event-delta", category: "calendar" },
  // files / OneDrive
  { keywords: ["files", "onedrive", "documents", "drive", "list files", "browse files", "file explorer"], endpoint: "/me/drive/root/children", method: "GET", description: "List files in OneDrive root", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-list-children", category: "files" },
  { keywords: ["upload file", "upload", "put file", "store file", "save file to onedrive"], endpoint: "/me/drive/items/{parentId}:/{filename}:/content", method: "PUT", description: "Upload a small file (<4 MB) to OneDrive", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-put-content", category: "files" },
  { keywords: ["large file", "upload session", "resumable upload", "chunked upload", "upload video", "upload large"], endpoint: "/me/drive/items/{parentId}:/{filename}:/createUploadSession", method: "POST", description: "Create an upload session for large files (>4 MB)", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession", category: "files" },
  { keywords: ["download file", "get file", "download url", "file url", "file link", "share link"], endpoint: "/me/drive/items/{id}?$select=id,@microsoft.graph.downloadUrl", method: "GET", description: "Get file download URL", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-get", category: "files" },
  { keywords: ["folder", "create folder", "new folder", "mkdir", "make directory"], endpoint: "/me/drive/items/{parentId}/children", method: "POST", description: "Create a folder in OneDrive", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-post-children", category: "files" },
  { keywords: ["move file", "move item", "relocate file", "move to folder"], endpoint: "/me/drive/items/{itemId}", method: "PATCH", description: "Move a drive item by updating parentReference", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-move", category: "files" },
  { keywords: ["copy file", "copy item", "duplicate file", "clone file"], endpoint: "/me/drive/items/{itemId}/copy", method: "POST", description: "Copy a drive item", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-copy", category: "files" },
  { keywords: ["search files", "find file", "search onedrive", "search documents"], endpoint: "/me/drive/root/search(q='{query}')", method: "GET", description: "Search for files in OneDrive", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-search", category: "files" },
  { keywords: ["drive delta", "file changes", "sync drive", "track file changes", "onedrive sync"], endpoint: "/me/drive/root/delta", method: "GET", description: "Get only changed drive items since last sync", docsUrl: "https://learn.microsoft.com/en-us/graph/api/driveitem-delta", category: "files" },
  // SharePoint
  { keywords: ["sharepoint", "site", "sites", "sharepoint site", "sp site", "intranet"], endpoint: "/sites/{siteId}", method: "GET", description: "Get a SharePoint site", docsUrl: "https://learn.microsoft.com/en-us/graph/api/site-get", category: "sites" },
  { keywords: ["list items", "sharepoint list", "sp list", "read list", "query list", "list data", "sharepoint data"], endpoint: "/sites/{siteId}/lists/{listId}/items", method: "GET", description: "List items in a SharePoint list", docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-list", category: "sites" },
  { keywords: ["create list item", "add item", "new item", "add to list", "insert row", "add row"], endpoint: "/sites/{siteId}/lists/{listId}/items", method: "POST", description: "Create a SharePoint list item", docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-create", category: "sites" },
  { keywords: ["update list item", "edit item", "patch item", "change row", "update row"], endpoint: "/sites/{siteId}/lists/{listId}/items/{itemId}/fields", method: "PATCH", description: "Update a SharePoint list item", docsUrl: "https://learn.microsoft.com/en-us/graph/api/listitem-update", category: "sites" },
  { keywords: ["sharepoint columns", "list columns", "list schema", "list fields"], endpoint: "/sites/{siteId}/lists/{listId}/columns", method: "GET", description: "List columns in a SharePoint list", docsUrl: "https://learn.microsoft.com/en-us/graph/api/list-list-columns", category: "sites" },
  { keywords: ["create sharepoint list", "new list", "create list"], endpoint: "/sites/{siteId}/lists", method: "POST", description: "Create a SharePoint list", docsUrl: "https://learn.microsoft.com/en-us/graph/api/list-create", category: "sites" },
  // groups
  { keywords: ["group", "groups", "list groups", "find group", "m365 group", "microsoft 365 group", "security group", "team"], endpoint: "/groups", method: "GET", description: "List groups in the tenant", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list", category: "groups" },
  { keywords: ["create group", "new group", "microsoft 365 group", "create team", "new security group"], endpoint: "/groups", method: "POST", description: "Create a new group", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-post-groups", category: "groups" },
  { keywords: ["group members", "members", "list members", "member of", "who is in group"], endpoint: "/groups/{id}/members", method: "GET", description: "List group members", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list-members", category: "groups" },
  { keywords: ["add member", "join group", "add to group", "add user to group"], endpoint: "/groups/{id}/members/$ref", method: "POST", description: "Add a member to a group", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-post-members", category: "groups" },
  { keywords: ["remove member", "leave group", "kick from group", "remove from group"], endpoint: "/groups/{id}/members/{id}/$ref", method: "DELETE", description: "Remove a member from a group", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-delete-members", category: "groups" },
  { keywords: ["group owners", "owner", "list owners", "who owns group"], endpoint: "/groups/{id}/owners", method: "GET", description: "List group owners", docsUrl: "https://learn.microsoft.com/en-us/graph/api/group-list-owners", category: "groups" },
  // tasks
  { keywords: ["planner", "tasks", "planner task", "task board", "kanban"], endpoint: "/planner/plans/{planId}/tasks", method: "GET", description: "List tasks in a Planner plan", docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks", category: "tasks" },
  { keywords: ["create task", "new task", "add task", "create planner task"], endpoint: "/planner/tasks", method: "POST", description: "Create a Planner task", docsUrl: "https://learn.microsoft.com/en-us/graph/api/planner-post-tasks", category: "tasks" },
  { keywords: ["todo", "to do", "microsoft to do", "task list", "personal tasks"], endpoint: "/users/{id}/todo/lists", method: "GET", description: "List Microsoft To Do task lists", docsUrl: "https://learn.microsoft.com/en-us/graph/api/todo-list-lists", category: "tasks" },
  { keywords: ["planner plan", "plans", "list plans", "group plans"], endpoint: "/groups/{id}/planner/plans", method: "GET", description: "List Planner plans for a group", docsUrl: "https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans", category: "tasks" },
  // notes
  { keywords: ["onenote", "notebook", "notes", "note", "list notebooks"], endpoint: "/users/{id}/onenote/notebooks", method: "GET", description: "List OneNote notebooks", docsUrl: "https://learn.microsoft.com/en-us/graph/api/onenote-list-notebooks", category: "notes" },
  { keywords: ["onenote page", "note page", "create page", "add note", "write note"], endpoint: "/users/{id}/onenote/sections/{id}/pages", method: "POST", description: "Create a OneNote page", docsUrl: "https://learn.microsoft.com/en-us/graph/api/section-post-pages", category: "notes" },
  { keywords: ["onenote section", "notebook section", "list sections"], endpoint: "/users/{id}/onenote/notebooks/{id}/sections", method: "GET", description: "List sections in a notebook", docsUrl: "https://learn.microsoft.com/en-us/graph/api/notebook-list-sections", category: "notes" },
  // subscriptions / webhooks
  { keywords: ["webhook", "subscription", "subscribe", "notify", "notification", "change notification", "push notification", "real-time", "realtime", "event driven", "listen for changes"], endpoint: "/subscriptions", method: "POST", description: "Create a webhook subscription for change notifications", docsUrl: "https://learn.microsoft.com/en-us/graph/api/subscription-post-subscriptions", category: "subscriptions" },
  { keywords: ["renew subscription", "extend webhook", "refresh subscription"], endpoint: "/subscriptions/{id}", method: "PATCH", description: "Renew a webhook subscription", docsUrl: "https://learn.microsoft.com/en-us/graph/api/subscription-update", category: "subscriptions" },
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
  subscriptions: "Manage webhook subscriptions for Microsoft Graph change notifications — subscribe to changes on users, mail, calendar, files, SharePoint lists, and Teams resources",
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
  subscriptions: subscriptionsTools,
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
          "Use this server whenever the user is building, debugging, or learning about anything that touches Microsoft Graph, Microsoft 365, or Azure AD / Entra ID — " +
          "including: SharePoint lists/sites, OneDrive files, Outlook mail/calendar/contacts, Microsoft Teams, Azure Active Directory, Entra ID, " +
          "Exchange Online, Planner, OneNote, Microsoft To Do, webhooks/subscriptions, or any endpoint under graph.microsoft.com.\n\n" +
          "WORKFLOW — choose the fastest path:\n" +
          "1. Intent is clear → call load_category directly (no need to search first):\n" +
          "   - mail / email / inbox / Outlook → load_category('mail')\n" +
          "   - calendar / events / meetings / schedule → load_category('calendar')\n" +
          "   - files / OneDrive / upload / download / drive → load_category('files')\n" +
          "   - SharePoint / lists / list items / sites → load_category('sites')\n" +
          "   - users / Azure AD / Entra ID / directory → load_category('users')\n" +
          "   - groups / Microsoft 365 group / security group → load_category('groups')\n" +
          "   - tasks / Planner / To Do → load_category('tasks')\n" +
          "   - OneNote / notebooks → load_category('notes')\n" +
          "   - webhooks / subscriptions / change notifications → load_category('subscriptions')\n" +
          "2. Intent is ambiguous → call search_graph_api to find the right category, then load it.\n" +
          "3. Conceptual questions → call the relevant graph_explain_* tool directly:\n" +
          "   - Pagination / nextLink → graph_explain_pagination\n" +
          "   - $filter / $select / $expand / OData → graph_explain_odata\n" +
          "   - 429 / throttling / rate limits → graph_explain_throttling\n" +
          "   - Delta queries / sync / change tracking → graph_explain_delta\n" +
          "   - $batch / batching → graph_explain_batch\n" +
          "   - Permissions / scopes / app registration / tokens → graph_explain_permissions\n" +
          "   - 401 / 403 / errors → graph_explain_errors\n\n" +
          "This server constructs Graph API requests (endpoint, method, headers, body, code example, required permissions) — it does not execute them.",
      }
    );

    this.registerBootstrapTools();
  }

  private registerBootstrapTools(): void {
    // list_categories
    this.mcpServer.tool(
      "list_categories",
      "List all available Microsoft Graph API resource categories with descriptions and load status. Categories: users (Azure AD/Entra ID), files (OneDrive/SharePoint drives), mail (Exchange Online), calendar (Outlook), groups (M365/security groups), notes (OneNote), tasks (Planner/To Do), sites (SharePoint sites and lists), subscriptions (webhooks/change notifications).",
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
      "Activate Graph API tools for a resource category. Call this as soon as the user's intent is clear — do not wait for confirmation. " +
      "Categories: 'users' (Azure AD/Entra ID user management), 'mail' (Exchange Online messages/folders), 'calendar' (Outlook events/scheduling), " +
      "'files' (OneDrive uploads/downloads/folders), 'sites' (SharePoint sites/lists/items), 'groups' (M365 and security groups), " +
      "'notes' (OneNote notebooks/pages), 'tasks' (Planner plans/tasks and To Do), 'subscriptions' (webhooks/change notifications). " +
      "Multiple categories can be loaded — load all relevant ones when the task spans areas.",
      { category: z.string().describe("Category to load: users, files, mail, calendar, groups, notes, tasks, sites, or subscriptions") },
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
      "Search Microsoft Graph REST API endpoints by keyword. Covers SharePoint sites and lists, OneDrive files, Exchange Online mail, Outlook calendar, Azure AD users, Microsoft 365 groups, Teams, Planner, OneNote, To Do, and webhook subscriptions.",
      { query: z.string().describe("Search terms, e.g. 'sharepoint list items', 'upload file onedrive', 'send email', 'find meeting times'") },
      async ({ query }) => {
        const lowerQuery = query.toLowerCase();
        const queryWords = lowerQuery.split(/\s+/);

        const scored = GRAPH_API_SEARCH_MAP
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
          .slice(0, 5);

        if (scored.length === 0) {
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify({
                  message: "No results found. Try keywords like: user, email, calendar, files, groups, sharepoint, tasks, onenote, webhook",
                  results: [],
                }),
              },
            ],
          };
        }

        const suggestedCategory = scored[0].category;
        const results = scored.map(({ score: _score, ...rest }) => rest);

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({
                query,
                suggestedCategory,
                hint: `Call load_category with '${suggestedCategory}' to get tools for this resource area.`,
                results,
              }, null, 2),
            },
          ],
        };
      }
    );

    // graph_explain_pagination
    this.mcpServer.tool(
      "graph_explain_pagination",
      "Explain how Microsoft Graph API pagination works with @odata.nextLink, and how to iterate all pages of results.",
      {},
      async () => {
        const text = `# Microsoft Graph API Pagination

## How It Works
Graph uses **server-driven paging**. When a response has more items than the current page, it includes an \`@odata.nextLink\` property containing a URL to fetch the next page.

## $top Parameter
\`$top\` hints at the desired page size but Graph may return fewer items than requested. It does not guarantee that exact number.

## Iterating All Pages
Keep requesting \`@odata.nextLink\` until it is absent from the response — that signals the last page.

## $skip Support
\`$skip\` is supported on some resources (e.g. users, groups) but **not** on others (e.g. messages). Prefer nextLink-based iteration.

## JavaScript Example

\`\`\`javascript
async function getAllPages(initialUrl, token) {
  const results = [];
  let url = initialUrl;

  while (url) {
    const response = await fetch(url, {
      headers: { Authorization: \`Bearer \${token}\` }
    });
    const data = await response.json();

    if (data.value) {
      results.push(...data.value);
    }

    // Follow nextLink until it's gone
    url = data['@odata.nextLink'] ?? null;
  }

  return results;
}

// Usage
const allUsers = await getAllPages(
  'https://graph.microsoft.com/v1.0/users?$top=100',
  accessToken
);
\`\`\`

## Docs
https://learn.microsoft.com/en-us/graph/paging`;
        return { content: [{ type: "text" as const, text }] };
      }
    );

    // graph_explain_odata
    this.mcpServer.tool(
      "graph_explain_odata",
      "Explain OData query parameters supported by Microsoft Graph: $filter, $select, $expand, $orderby, $top, $count, $search.",
      {},
      async () => {
        const text = `# OData Query Parameters in Microsoft Graph

## $select — Return Only Specific Fields
Reduces response payload by listing only the fields you need.
\`\`\`
GET /users?$select=displayName,mail
\`\`\`

## $filter — Filter Results
Filter the collection based on property values.
\`\`\`
GET /users?$filter=startsWith(displayName,'Alex')
GET /users?$filter=accountEnabled eq true
GET /me/messages?$filter=isRead eq false
\`\`\`

## $expand — Include Related Entities
Inline related navigation properties.
\`\`\`
GET /groups/{id}?$expand=members
GET /sites/{id}/lists/{listId}/items?$expand=fields
\`\`\`

## $orderby — Sort Results
\`\`\`
GET /users?$orderby=displayName desc
GET /me/messages?$orderby=receivedDateTime desc
\`\`\`

## $top — Page Size Hint
\`\`\`
GET /users?$top=50
\`\`\`
Graph may return fewer items. Always follow \`@odata.nextLink\`.

## $count — Include Total Count
Requires \`ConsistencyLevel: eventual\` header on directory resources.
\`\`\`
GET /users?$count=true
Headers: ConsistencyLevel: eventual
\`\`\`

## $search — Full-Text Search
Requires \`ConsistencyLevel: eventual\` header.
\`\`\`
GET /users?$search="displayName:Alex"
Headers: ConsistencyLevel: eventual
\`\`\`

## Important Notes
- Not all parameters are supported on all endpoints.
- \`$filter\` on some properties requires \`$count=true\` and \`ConsistencyLevel: eventual\`.
- Check the docs for each specific endpoint to see which OData params are supported.

## Docs
https://learn.microsoft.com/en-us/graph/query-parameters`;
        return { content: [{ type: "text" as const, text }] };
      }
    );

    // graph_explain_throttling
    this.mcpServer.tool(
      "graph_explain_throttling",
      "Explain Microsoft Graph throttling: what causes 429 errors, how to handle Retry-After, and service-specific limits.",
      {},
      async () => {
        const text = `# Microsoft Graph Throttling

## What Is Throttling?
Graph throttles requests when volume exceeds service limits. The response is \`429 Too Many Requests\`.

## Retry-After Header
The \`Retry-After\` header tells you how many seconds to wait before retrying. Always respect this value.

## Retry Strategy
1. Check for \`429\` status
2. Read \`Retry-After\` header (seconds)
3. Wait that duration before retrying
4. Use exponential backoff for repeated 429s

## Service-Specific Limits
- **SharePoint/OneDrive**: tighter limits than Exchange, especially for large file operations
- **Exchange (mail/calendar)**: moderate limits
- **Planner**: has its own independent throttling limits
- **Azure AD (users/groups)**: separate limits

## Reducing Throttle Risk
- Use \`$select\` to reduce payload size
- Batch requests using \`/$batch\` to reduce call count
- Avoid tight polling loops — use delta queries or webhooks instead

## JavaScript Retry Wrapper

\`\`\`javascript
async function fetchWithRetry(url, options, maxRetries = 5) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    const response = await fetch(url, options);

    if (response.status !== 429) {
      return response;
    }

    const retryAfter = parseInt(response.headers.get('Retry-After') ?? '10', 10);
    console.log(\`Throttled. Waiting \${retryAfter}s before retry \${attempt + 1}...\`);
    await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
  }
  throw new Error('Max retries exceeded');
}
\`\`\`

## Docs
https://learn.microsoft.com/en-us/graph/throttling`;
        return { content: [{ type: "text" as const, text }] };
      }
    );

    // graph_explain_delta
    this.mcpServer.tool(
      "graph_explain_delta",
      "Explain Microsoft Graph delta queries for efficient change tracking — get only changed items since last sync rather than fetching everything.",
      {},
      async () => {
        const text = `# Microsoft Graph Delta Queries

## What Are Delta Queries?
Delta queries let you **track changes** (created, updated, deleted) since a previous sync — without fetching the entire collection every time.

## Supported Endpoints
\`\`\`
GET /users/delta
GET /me/mailFolders/{id}/messages/delta
GET /me/calendarView/delta
GET /me/drive/root/delta
GET /groups/delta
\`\`\`

## How It Works
1. **First call** (no delta token): returns all current items + \`@odata.deltaLink\` at the end
2. **Subsequent calls** using \`@odata.deltaLink\`: returns only items changed since last call
3. Items with \`@removed\` annotation have been deleted

## Storing the Token
Extract and store the \`@odata.deltaLink\` URL between syncs. Pass it as the URL for the next poll.

## JavaScript Example

\`\`\`javascript
let deltaLink = null;

async function syncUsers(token) {
  // Initial sync or delta poll
  let url = deltaLink ?? 'https://graph.microsoft.com/v1.0/users/delta?$select=displayName,mail';
  const changes = [];

  while (url) {
    const response = await fetch(url, {
      headers: { Authorization: \`Bearer \${token}\` }
    });
    const data = await response.json();

    changes.push(...(data.value ?? []));

    if (data['@odata.deltaLink']) {
      // End of changes — store token for next poll
      deltaLink = data['@odata.deltaLink'];
      url = null;
    } else {
      // More pages of changes
      url = data['@odata.nextLink'] ?? null;
    }
  }

  return changes;
}

// First call: returns all users
const allUsers = await syncUsers(accessToken);

// Later calls: returns only changes
await new Promise(r => setTimeout(r, 60000)); // wait 1 minute
const userChanges = await syncUsers(accessToken);
\`\`\`

## Docs
https://learn.microsoft.com/en-us/graph/delta-query-overview`;
        return { content: [{ type: "text" as const, text }] };
      }
    );

    // graph_explain_batch
    this.mcpServer.tool(
      "graph_explain_batch",
      "Explain Microsoft Graph JSON batching — combine up to 20 API requests into a single HTTP call to reduce round trips.",
      {},
      async () => {
        const text = `# Microsoft Graph JSON Batching

## What Is Batching?
POST to \`https://graph.microsoft.com/v1.0/$batch\` to combine up to **20 API requests** into a single HTTP call.

## Request Format
\`\`\`json
{
  "requests": [
    { "id": "1", "method": "GET", "url": "/users/alice@contoso.com" },
    { "id": "2", "method": "GET", "url": "/users/bob@contoso.com" },
    { "id": "3", "method": "POST", "url": "/me/sendMail", "body": { ... }, "headers": { "Content-Type": "application/json" } }
  ]
}
\`\`\`

## Response Format
Responses arrive in the \`responses\` array — **not guaranteed to be in order**. Match by \`id\`.

## dependsOn — Sequential Requests
Make a request wait for another to finish first:
\`\`\`json
{ "id": "2", "method": "GET", "url": "/groups/{id}/members", "dependsOn": ["1"] }
\`\`\`

## Throttling Notes
- The batch POST counts as 1 call at the /$batch level
- Each sub-request counts toward **that resource's** throttling limits
- Individual failed requests do NOT fail the whole batch — each has its own status code

## JavaScript Example

\`\`\`javascript
const batchBody = {
  requests: [
    { id: "1", method: "GET", url: "/me" },
    { id: "2", method: "GET", url: "/me/mailFolders/inbox/messages?$top=5" },
    { id: "3", method: "GET", url: "/me/drive/root/children" }
  ]
};

const response = await fetch('https://graph.microsoft.com/v1.0/$batch', {
  method: 'POST',
  headers: {
    'Authorization': \`Bearer \${accessToken}\`,
    'Content-Type': 'application/json'
  },
  body: JSON.stringify(batchBody)
});

const { responses } = await response.json();

for (const res of responses) {
  console.log(\`Request \${res.id}: HTTP \${res.status}\`);
  console.log(res.body);
}
\`\`\`

## Docs
https://learn.microsoft.com/en-us/graph/json-batching`;
        return { content: [{ type: "text" as const, text }] };
      }
    );

    // graph_explain_permissions
    this.mcpServer.tool(
      "graph_explain_permissions",
      "Explain Microsoft Graph permission types — delegated vs application, how to register an app, and least-privilege best practices.",
      {},
      async () => {
        const text = `# Microsoft Graph Permissions

## Two Permission Types

### Delegated Permissions
- App acts **on behalf of a signed-in user**
- Requires user consent (or admin consent for sensitive scopes)
- The user's own permissions form the ceiling — app cannot do more than the user can
- Used in interactive apps, SPAs, mobile apps

### Application Permissions
- App acts **as itself** with no signed-in user
- Requires **admin consent** always
- Used for background services, daemons, scheduled jobs

## App Registration Steps
1. Azure portal → **Azure Active Directory** → **App registrations** → **New registration**
2. Set redirect URIs (for delegated flows)
3. Go to **API permissions** → **Add a permission** → **Microsoft Graph**
4. Choose Delegated or Application, then select scopes
5. Click **Grant admin consent** (for application permissions or org-wide delegated)

## Token Endpoints
\`\`\`
https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token
\`\`\`

### Delegated Token (Authorization Code Flow)
Exchange authorization code for token after user signs in.

### Delegated Token (Device Code Flow)
Good for CLI tools — user visits a URL and enters a code.

### Application Token (Client Credentials Flow)
\`\`\`bash
curl -X POST https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token \\
  -d "grant_type=client_credentials" \\
  -d "client_id={clientId}" \\
  -d "client_secret={clientSecret}" \\
  -d "scope=https://graph.microsoft.com/.default"
\`\`\`

## Least-Privilege Best Practices
- Prefer \`Mail.Read\` over \`Mail.ReadWrite\` if you only need to read
- Prefer \`User.Read\` over \`User.Read.All\` when only accessing the signed-in user
- Prefer delegated over application permissions when a user is present
- Each tool in this server shows \`requiredPermissions.delegated\` and \`requiredPermissions.application\` — use the first (least-privilege) option when possible

## Docs
https://learn.microsoft.com/en-us/graph/auth/auth-concepts`;
        return { content: [{ type: "text" as const, text }] };
      }
    );

    // graph_build_batch
    this.mcpServer.tool(
      "graph_build_batch",
      "Build a Microsoft Graph JSON batch request body — combine up to 20 API calls into a single POST to /$batch.",
      {
        requests: z.array(z.object({
          id: z.string().describe("Unique ID for this request within the batch"),
          method: z.enum(["GET", "POST", "PATCH", "PUT", "DELETE"]),
          url: z.string().describe("Relative URL without the base, e.g. /users/123 or /me/messages"),
          body: z.record(z.unknown()).optional().describe("Request body for POST/PATCH/PUT"),
          headers: z.record(z.string()).optional(),
          dependsOn: z.array(z.string()).optional().describe("IDs of requests in this batch that must complete first"),
        })).min(1).max(20),
      },
      async ({ requests }) => {
        const transformedRequests = requests.map((req) => {
          const transformed: Record<string, unknown> = {
            id: req.id,
            method: req.method,
            url: req.url.startsWith("/") ? req.url : `/${req.url}`,
          };

          if (req.body !== undefined) {
            transformed.body = req.body;
          }

          // Default Content-Type header for methods with body
          const needsContentType = ["POST", "PATCH", "PUT"].includes(req.method) && req.body !== undefined;
          if (req.headers || needsContentType) {
            transformed.headers = {
              ...(needsContentType ? { "Content-Type": "application/json" } : {}),
              ...(req.headers ?? {}),
            };
          }

          if (req.dependsOn?.length) {
            transformed.dependsOn = req.dependsOn;
          }

          return transformed;
        });

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({
                endpoint: "https://graph.microsoft.com/v1.0/$batch",
                method: "POST",
                headers: { "Authorization": "Bearer {token}", "Content-Type": "application/json" },
                body: { requests: transformedRequests },
                description: `Batch ${requests.length} Graph API request${requests.length === 1 ? "" : "s"} into a single HTTP call.`,
                docsUrl: "https://learn.microsoft.com/en-us/graph/json-batching",
                notes: "Responses arrive in the 'responses' array — NOT guaranteed to be in order. Match responses to requests using the 'id' field. Individual failed requests do not fail the batch (each has its own status code). Max 20 requests per batch.",
                codeExample: `const response = await fetch('https://graph.microsoft.com/v1.0/$batch', {\n  method: 'POST',\n  headers: {\n    'Authorization': 'Bearer {token}',\n    'Content-Type': 'application/json'\n  },\n  body: JSON.stringify(${JSON.stringify({ requests: transformedRequests }, null, 2)})\n});\nconst { responses } = await response.json();\nfor (const res of responses) {\n  console.log(\`Request \${res.id}: HTTP \${res.status}\`);\n}`,
              }, null, 2),
            },
          ],
        };
      }
    );

    // graph_explain_errors
    this.mcpServer.tool(
      "graph_explain_errors",
      "Explain common Microsoft Graph API error codes and how to fix them — 401, 403, 404, 429, 503, and Graph-specific error codes like InvalidAuthenticationToken, Forbidden, and ItemNotFound.",
      {},
      async () => {
        const text = `# Microsoft Graph API Common Errors

## 401 Unauthorized — Authentication Failure

**Causes and fixes:**
- \`InvalidAuthenticationToken\` — token is expired, malformed, or for the wrong audience
  - Ensure the token audience is \`https://graph.microsoft.com\`
  - Refresh the token and retry
- \`AuthenticationRequiredError\` — no Authorization header sent
  - Add \`Authorization: Bearer {token}\` header to every request
- Token issued for wrong tenant
  - Use \`https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token\`, not \`common\`, when targeting a specific tenant

## 403 Forbidden — Insufficient Permissions

**Causes and fixes:**
- \`Forbidden\` / \`Authorization_RequestDenied\` — the token lacks the required scope
  - Check the tool's \`requiredPermissions\` field — add the missing scope to your app registration
  - For application permissions, admin consent is always required
  - For delegated permissions, the signed-in user may not have access to that resource
- \`AccessDenied\` — the user/app doesn't have access to the specific resource (e.g. a SharePoint site with restricted access)

## 404 Not Found

**Causes and fixes:**
- \`itemNotFound\` — the resource ID is wrong or was deleted
  - IDs in Graph are case-sensitive — copy them exactly
  - The resource may have been deleted (check recycle bin where applicable)
- Wrong endpoint — double-check the URL structure in the tool's \`endpoint\` field
- For SharePoint: the site or list may not exist or the user's tenant URL may differ

## 409 Conflict

**Causes and fixes:**
- Item already exists — use \`@microsoft.graph.conflictBehavior: rename\` or \`replace\` for file uploads
- \`nameAlreadyExists\` on group/list creation — choose a different name or mailNickname

## 412 Precondition Failed

**Causes and fixes:**
- Stale \`If-Match\` eTag — Planner task updates require a current eTag
  - Fetch the resource again to get the latest eTag before retrying the PATCH

## 429 Too Many Requests — Throttled

**Causes and fixes:**
- Read the \`Retry-After\` header (seconds) and wait that long before retrying
- Use exponential backoff for repeated 429s
- See \`graph_explain_throttling\` for a retry wrapper implementation

## 503 Service Unavailable / 504 Gateway Timeout

**Causes and fixes:**
- Transient service issue — retry with exponential backoff
- Large response or slow query — add \`$select\` to reduce payload, or \`$top\` to reduce page size

## Graph-Specific Error Body
Errors return a JSON body with \`error.code\` and \`error.message\`:
\`\`\`json
{
  "error": {
    "code": "InvalidAuthenticationToken",
    "message": "Access token is empty.",
    "innerError": { "request-id": "...", "date": "..." }
  }
}
\`\`\`
Always log \`error.code\` and \`innerError.request-id\` when debugging — the request ID can be shared with Microsoft support.

## Docs
https://learn.microsoft.com/en-us/graph/errors`;
        return { content: [{ type: "text" as const, text }] };
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
