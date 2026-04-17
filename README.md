# Microsoft Graph Dev MCP

An MCP server that helps you construct and validate [Microsoft Graph REST API](https://learn.microsoft.com/en-us/graph/api/overview) calls — no authentication required in the server itself.

Tools are loaded on demand by resource category. Ask about SharePoint and the sites tools appear. Ask about files and the OneDrive tools appear. The server starts lean and grows with your needs, and always knows the required permissions for every operation.

## What it does

- Constructs valid Graph API request URLs, methods, headers, and bodies
- Validates required and optional parameters
- Returns required Microsoft Graph permissions (delegated and application) for every operation
- Returns ready-to-use fetch code examples
- Links to official Microsoft documentation for every operation
- Explains cross-cutting concepts: pagination, OData queries, throttling, delta sync, batching, and auth
- Loads resource categories on demand — only what you need

## Tools

### Always available

These tools are loaded immediately — no setup required.

| Tool | Description |
|---|---|
| `list_categories` | List all resource categories and which are currently loaded |
| `load_category` | Load tools for a category; triggers `tools/list_changed` |
| `search_graph_api` | Search Graph API endpoints by keyword; returns a `suggestedCategory` to load |
| `graph_build_batch` | Build a valid `/$batch` request body from up to 20 operations |
| `graph_explain_pagination` | How `@odata.nextLink` works; iterate all pages of results |
| `graph_explain_odata` | `$filter`, `$select`, `$expand`, `$orderby`, `$count`, `$search` with examples |
| `graph_explain_throttling` | 429 handling, `Retry-After`, exponential backoff pattern |
| `graph_explain_delta` | Delta tokens, change tracking, initial sync vs incremental sync |
| `graph_explain_batch` | JSON batching, `dependsOn`, response handling |
| `graph_explain_permissions` | Delegated vs application permissions, consent flows, token acquisition |

### Resource categories (loaded on demand)

| Category | Tools | What you get |
|---|---|---|
| **users** | 8 | Get, list, create, update, delete users; manager; direct reports; delta sync |
| **files** | 11 | OneDrive/SharePoint document library: list, get, upload (<4MB and resumable), create folder, delete, move, copy, search, download URL, delta sync |
| **mail** | 9 | List and get messages, send, create draft, reply, delete, move, list folders, delta sync |
| **calendar** | 8 | List, get, create, update, delete events; find meeting times; get free/busy schedule; delta sync |
| **groups** | 8 | List, get, create, delete groups; list and manage members and owners |
| **notes** | 8 | OneNote notebooks, sections, pages; create and read content |
| **tasks** | 10 | Planner plans and tasks (CRUD); Microsoft To Do lists and tasks |
| **sites** | 12 | SharePoint sites; lists; list items (CRUD); columns |
| **subscriptions** | 5 | Create, list, get, delete, and renew webhook change notification subscriptions |

### Example tool output

Every tool returns a structured object with the full request details and required permissions:

```json
{
  "endpoint": "https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items?expand=fields",
  "method": "GET",
  "headers": {
    "Authorization": "Bearer {token}"
  },
  "pathParams": { "siteId": "contoso.sharepoint.com,abc123", "listId": "list456" },
  "queryParams": { "expand": "fields" },
  "body": null,
  "description": "List items in list list456.",
  "docsUrl": "https://learn.microsoft.com/en-us/graph/api/listitem-list",
  "codeExample": "const response = await fetch('...', { method: 'GET', headers: { Authorization: 'Bearer {token}' } });\nconst data = await response.json();",
  "requiredPermissions": {
    "delegated": ["Sites.Read.All"],
    "application": ["Sites.Read.All"]
  },
  "notes": "Column values are returned under the 'fields' property. This request already includes ?expand=fields. Without it the items array would contain only metadata, not column data."
}
```

---

## Installation

The recommended way to run this server is via `npx` — no local install needed.

```
npx -y @stubbedev/ms-graph-dev-mcp
```

### Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows):

```json
{
  "mcpServers": {
    "ms-graph-dev": {
      "command": "npx",
      "args": ["-y", "@stubbedev/ms-graph-dev-mcp"]
    }
  }
}
```

### Claude Code (CLI)

```bash
claude mcp add ms-graph-dev -- npx -y @stubbedev/ms-graph-dev-mcp
```

Or add to your project's `.mcp.json`:

```json
{
  "mcpServers": {
    "ms-graph-dev": {
      "command": "npx",
      "args": ["-y", "@stubbedev/ms-graph-dev-mcp"]
    }
  }
}
```

### Cursor

Open **Settings → MCP** and add a new server, or edit `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "ms-graph-dev": {
      "command": "npx",
      "args": ["-y", "@stubbedev/ms-graph-dev-mcp"]
    }
  }
}
```

### Windsurf

Edit `~/.codeium/windsurf/mcp_config.json`:

```json
{
  "mcpServers": {
    "ms-graph-dev": {
      "command": "npx",
      "args": ["-y", "@stubbedev/ms-graph-dev-mcp"]
    }
  }
}
```

### Zed

Edit your `settings.json` (open via **Zed → Settings → Open Settings**):

```json
{
  "context_servers": {
    "ms-graph-dev": {
      "command": {
        "path": "npx",
        "args": ["-y", "@stubbedev/ms-graph-dev-mcp"]
      }
    }
  }
}
```

### OpenCode

Edit `~/.config/opencode/config.json`:

```json
{
  "mcp": {
    "ms-graph-dev": {
      "type": "local",
      "command": ["npx", "-y", "@stubbedev/ms-graph-dev-mcp"]
    }
  }
}
```

### Codex (OpenAI)

Edit `~/.codex/config.json`:

```json
{
  "mcpServers": {
    "ms-graph-dev": {
      "command": "npx",
      "args": ["-y", "@stubbedev/ms-graph-dev-mcp"]
    }
  }
}
```

---

## Development

```bash
git clone https://github.com/stubbedev/ms-graph-dev-mcp.git
cd ms-graph-dev-mcp
npm install
npm run build
npm start
```

For live reload during development:

```bash
npm run dev
```

### Test with MCP Inspector

```bash
npx @modelcontextprotocol/inspector npx -y @stubbedev/ms-graph-dev-mcp
```

---

## License

MIT
