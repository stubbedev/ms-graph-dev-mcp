# Microsoft Graph Dev MCP

An MCP server that helps you construct and validate [Microsoft Graph REST API](https://learn.microsoft.com/en-us/graph/api/overview) calls — no authentication required in the server itself.

Tools are loaded on demand by resource category. Ask about files and the file tools appear. Ask about mail and the mail tools appear. The server starts lean and grows with your needs.

## What it does

- Constructs valid Graph API request URLs, methods, headers, and bodies
- Validates required and optional parameters
- Returns ready-to-use code examples
- Links to official Microsoft documentation for every operation
- Loads tool categories dynamically — only what you need

### Available categories

| Category | What you get |
|---|---|
| **users** | Get, list, create, update, delete users; manager and direct reports |
| **files** | OneDrive/SharePoint drive operations: list, upload, download, move, copy, search |
| **mail** | Read, send, draft, reply, move messages; list folders |
| **calendar** | Events CRUD, find meeting times, free/busy schedule |
| **groups** | Microsoft 365 and security groups; members and owners |
| **notes** | OneNote notebooks, sections, and pages |
| **tasks** | Planner plans and tasks; Microsoft To Do lists and tasks |
| **sites** | SharePoint sites, lists, items, and columns |

### Bootstrap tools (always available)

| Tool | Description |
|---|---|
| `list_categories` | Shows all categories and which ones are currently loaded |
| `load_category` | Loads tools for a category; triggers `tools/list_changed` |
| `search_graph_api` | Keyword search across common Graph API operations |

### Example tool output

Every tool returns a structured object:

```json
{
  "endpoint": "https://graph.microsoft.com/v1.0/users/{userId}/messages",
  "method": "POST",
  "headers": {
    "Authorization": "Bearer {token}",
    "Content-Type": "application/json"
  },
  "pathParams": { "userId": "me" },
  "queryParams": {},
  "body": {
    "message": {
      "subject": "Hello",
      "body": { "contentType": "HTML", "content": "<p>Hello world</p>" },
      "toRecipients": [{ "emailAddress": { "address": "user@example.com" } }]
    },
    "saveToSentItems": true
  },
  "description": "Send an email message on behalf of a user",
  "docsUrl": "https://learn.microsoft.com/en-us/graph/api/user-sendmail",
  "codeExample": "const response = await fetch('https://graph.microsoft.com/v1.0/users/me/messages', { method: 'POST', headers: { Authorization: 'Bearer {token}', 'Content-Type': 'application/json' }, body: JSON.stringify({...}) });"
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
