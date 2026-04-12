# MCP Server and VS Code Integration

## Overview

AxonASP includes an MCP (Model Context Protocol) server (`axonasp-mcp.exe`) that exposes the AxonASP documentation and tooling to AI assistants and IDEs that support MCP. The MCP server enables context-aware code generation, documentation lookup, and script analysis directly from your editor.

## Starting the MCP Server

```powershell
.\axonasp-mcp.exe
```

The MCP server operates in `stdio` mode by default, reading commands from standard input and writing responses to standard output. It can also be configured to run in SSE (Server-Sent Events) mode via `config/axonasp.toml`.

Configure the MCP mode in `config/axonasp.toml`:

```toml
[mcp]
mode = "stdio"
```

## Configuring VS Code

To connect VS Code to the AxonASP MCP server, create a `.vscode` directory in the root of your project and add a `mcp.json` file with the following content:

**File: `.vscode/mcp.json`**

```json
{
    "servers": {
        "AxonASP MCP": {
            "type": "stdio",
            "command": ".\\axonasp-mcp.exe",
            "args": []
        }
    },
    "inputs": []
}
```

**Steps:**

1. Open your project root in VS Code.
2. Create a `.vscode` folder if it does not already exist.
3. Inside `.vscode`, create a file named `mcp.json`.
4. Paste the JSON configuration above into the file.
5. Ensure `axonasp-mcp.exe` is present in the project root directory.
6. Reload VS Code or activate the MCP extension to connect.

Once connected, AI tools in VS Code (such as GitHub Copilot) can query the AxonASP MCP server for documentation, object references, and code examples.

## SSE Mode

The MCP server also supports SSE mode, which exposes an HTTP endpoint for browser-based or remote MCP clients.

```toml
[mcp]
mode = "sse"
sse_port = 8900
```

Start the server in SSE mode:

```powershell
.\axonasp-mcp.exe
```

Connect remote clients to `http://localhost:8900`.

## Remarks

- In `stdio` mode, the MCP server does not open a network port. VS Code launches the process directly and communicates through its stdin/stdout pipes.
- The MCP server shares the same configuration file (`config/axonasp.toml`) as the HTTP and FastCGI servers.
- Only one MCP transport mode (`stdio` or `sse`) is active at a time.
- The `.vscode/mcp.json` file is local to your workspace. Add it to `.gitignore` if you do not want to share the MCP configuration with other repository contributors.
- Maintainer note: The MCP server exposes two core tools: `search_axonasp_docs` for fuzzy documentation lookup and `get_asp_coding_style` for returning the full Classic ASP/VBScript formatting guide from `mcp/aspcodingstyle.md`. AI agents should call `get_asp_coding_style` whenever code-style or formatting rules are needed before generating or refactoring ASP code.
