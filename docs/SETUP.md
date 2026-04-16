# Multi-Platform Setup Guide

This MCP server works with any tool that supports the Model Context Protocol.
It supports two transports:
- **stdio** (default) — for CLI tools (Claude Code, Codex CLI, Gemini CLI)
- **sse** — for IDE integrations (Cursor, VS Code, JetBrains, web clients)

## Prerequisites (all platforms)

```powershell
git clone https://github.com/StealthyLabsHQ/powerbi-mcp-local.git
cd powerbi-mcp-local
pip install -r requirements.txt
```

Power BI Desktop must be running with an open `.pbix` file.

---

## Claude Code (CLI)

### Project-level (recommended)

Create `.claude/settings.json` in your project root:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

### Global

Add to `%USERPROFILE%\.claude\\settings.json`:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"],
      "env": {
        "PYTHONPATH": "C:\\Program Files\\Microsoft Power BI Desktop\\bin"
      }
    }
  }
}
```

### Claude Desktop App

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

---

## Codex CLI (OpenAI)

Codex CLI supports MCP servers via its config file.

Create or edit `~/.codex/config.json`:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

Then use it:
```powershell
codex --model o3 "Connect to Power BI and list all tables"
```

---

## Gemini CLI (Google)

Gemini CLI reads MCP config from its settings file.

Create or edit `~/.gemini/settings.json`:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

Then use it:
```powershell
gemini "Connect to Power BI and create measures from my DAX file"
```

---

## Cursor IDE

Cursor supports MCP servers natively.

### Option A — stdio (project-level)

Create `.cursor/mcp.json` in your project root:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

### Option B — SSE (global, shared across projects)

Start the server in SSE mode:
```powershell
python server.py --transport sse --port 8765
```

Then in Cursor settings (`File > Preferences > Cursor Settings > MCP`):

```json
{
  "mcpServers": {
    "powerbi": {
      "url": "http://localhost:8765/sse"
    }
  }
}
```

---

## VS Code (with MCP extension)

If using an MCP-compatible VS Code extension (e.g. Continue, Cline):

### stdio

Add to `.vscode/mcp.json`:

```json
{
  "servers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

### SSE

Start server: `python server.py --transport sse --port 8765`

```json
{
  "servers": {
    "powerbi": {
      "url": "http://localhost:8765/sse"
    }
  }
}
```

---

## JetBrains IDEs (IntelliJ, PyCharm, WebStorm)

JetBrains IDEs with AI Assistant support MCP via settings.

`Settings > AI Assistant > MCP Servers > Add`:

- **Name**: powerbi
- **Command**: `python C:\path\to\powerbi-mcp-local\server.py`
- **Transport**: stdio

Or for SSE:
- **URL**: `http://localhost:8765/sse`

---

## Windsurf / Cline / Continue

These tools all support the same MCP config format.

Create `.mcp/config.json` in your project root:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

---

## SSE Mode (any HTTP client)

For tools that don't support stdio, run the server as an HTTP endpoint:

```powershell
python server.py --transport sse --port 8765
```

The server exposes:
- `http://localhost:8765/sse` — SSE event stream (for MCP clients)

Keep the terminal open while using the MCP.

---

## Verifying the Connection

After configuring any tool, test with:

```
"Call pbi_connect to verify the Power BI connection"
```

Expected response: port number, database name, table count.

If Power BI Desktop is not running, you'll get a structured JSON error
explaining the issue.

---

## Transport Comparison

| Transport | Best for | Requires |
|---|---|---|
| **stdio** | CLI tools (Claude Code, Codex, Gemini) | Direct process spawn |
| **sse** | IDEs (Cursor, VS Code, JetBrains) | Server running in background |

stdio is simpler and recommended for CLI tools. SSE is useful when the
IDE can't spawn child processes or when you want one server shared across
multiple clients.
