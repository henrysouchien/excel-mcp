# excel-mcp

Connect Excel to AI through MCP (Model Context Protocol). Give Claude, or any MCP-compatible AI, the ability to read, write, and manipulate Excel spreadsheets.

## What it does

excel-mcp provides three components that work together:

1. **MCP Server** — stdio server that exposes Excel tools to AI clients (Claude Code, Claude Desktop, etc.)
2. **Relay Backend** — FastAPI service that bridges MCP requests to the Excel add-in via SSE
3. **Office.js Add-in** — runs inside Excel, executes tool calls against the live workbook

```
AI Client ←→ MCP Server ←→ Relay Backend ←→ Excel Add-in ←→ Workbook
 (stdio)                    (HTTP/SSE)       (Office.js)
```

## Built-in tools

25 tools for working with Excel:

| Tool | Description |
|------|-------------|
| `list_workbooks` | List workbooks currently connected through the taskpane |
| `switch_active_workbook` | Set the default workbook target for later tool calls |
| `read_cells` | Read values and formulas from a range |
| `write_cells` | Write values to a range; returns a restore token for `restore_written_cells` |
| `restore_written_cells` | Restore cell contents overwritten by `write_cells` |
| `read_range_csv` | Read a range as CSV text |
| `read_sheet_csv` | Read an entire sheet as CSV |
| `get_selection` | Get the current selection |
| `get_used_range` | Get the used range dimensions |
| `list_sheets` | List all sheets with metadata |
| `switch_sheet` | Activate a sheet |
| `create_sheet` | Create a new sheet |
| `rename_sheet` | Rename a sheet; returns a restore token for `restore_renamed_sheet` |
| `restore_renamed_sheet` | Undo a sheet rename |
| `delete_sheet` | Delete a sheet; returns a restore token for `restore_deleted_sheet` |
| `insert_row` | Insert rows |
| `delete_row` | Delete rows; returns a restore token for `restore_deleted_row` |
| `restore_deleted_row` | Restore deleted rows |
| `insert_column` | Insert columns |
| `delete_column` | Delete columns; returns a restore token for `restore_deleted_column` |
| `restore_deleted_sheet` | Restore a deleted sheet |
| `restore_deleted_column` | Restore deleted columns |
| `create_table` | Create an Excel table |
| `format_cells` | Format cells (bold, color, number format, etc.) |
| `find_cells` | Search for values or formulas |

## Quick start

### 1. Install the Python package

```bash
pip install -e ./python
```

### 2. Set environment variables

```bash
export EXCEL_MCP_API_KEY="your-channel-mcp-gateway-user-key"
export SESSION_API_KEY="your-channel-excel-gateway-user-key"
export EXCEL_MCP_BACKEND_URL="https://localhost:8000/api/mcp/execute"
```

### 3. Start the authenticated relay backend

```bash
uvicorn api.main:app --host 0.0.0.0 --port 8000
```

The package `create_relay_app()` factory is intended to be embedded with an
application-provided request authenticator. The product gateway wires that to
gateway JWT sessions issued from `GATEWAY_USER_KEYS`.

In the `AI-excel-addin` product repo, prefer services-mcp instead of raw
uvicorn for local product work:

```text
service_start research_gateway
```

That launcher hydrates the local risk_module resolver bridge before the gateway
starts. Raw uvicorn is fine for package harness development, but it skips local
product ops wiring.

### 4. Start the MCP server

```bash
python -m excel_mcp
```

### 5. Start a taskpane

```bash
cd addin
npm install
npm run dev-server
```

This package-level add-in is an internal standalone harness for developing the
generic MCP bridge. It runs on `https://localhost:3102` with a separate
`Excel MCP Bridge (Internal)` manifest so it does not replace the product
taskpane.

In the `AI-excel-addin` product repo, use the root Hank AI add-in instead:

```bash
cd /Users/henrychien/Documents/Jupyter/AI-excel-addin
npm run dev-server
npm start
```

The root Hank AI taskpane on `https://localhost:3002` is the user-facing Excel
surface. It embeds the same generic Office/MCP primitives behind the chat and
artifact UI.

Taskpane startup behavior:

- First use in a workbook still requires a manual open through Insert > My Add-ins.
- After that first manual open, the taskpane is configured to auto-load when the same workbook is reopened.
- New workbooks need that initial manual open once before auto-load applies to them too.

### 6. Connect your AI client

Add to your Claude Code MCP config:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python3",
      "args": ["-m", "excel_mcp"],
      "env": {
        "EXCEL_MCP_API_KEY": "your-channel-mcp-gateway-user-key",
        "EXCEL_MCP_BACKEND_URL": "https://localhost:8000/api/mcp/execute"
      }
    }
  }
}
```

## Adding custom tools

Use `register_tools()` to add domain-specific tools before starting the MCP server:

```python
from excel_mcp.tool_registry import register_tools

MY_TOOL = {
    "name": "my_custom_tool",
    "description": "Does something specific to my workflow",
    "input_schema": {
        "type": "object",
        "properties": {
            "param": {"type": "string", "description": "A parameter"}
        },
        "required": ["param"]
    }
}

register_tools(MY_TOOL)
```

Custom tools are dispatched through the same relay → add-in pipeline. Implement the handler in the Office.js add-in's tool dispatch.

## Architecture

- **`python/excel_mcp/`** — pip-installable Python package
  - `mcp_server.py` — MCP stdio server
  - `relay.py` — FastAPI relay with SSE, `create_relay_app()` factory
  - `tool_registry.py` — tool specs + `register_tools()` extensibility
- **`addin/`** — internal standalone Office.js bridge harness (npm package)
  - `src/taskpane/services/OfficeService.ts` — tool implementations
  - `src/taskpane/taskpane.ts` — MCP event loop

## License

MIT
