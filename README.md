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

18 tools for working with Excel:

| Tool | Description |
|------|-------------|
| `read_cells` | Read values and formulas from a range |
| `write_cells` | Write values to a range |
| `read_range_csv` | Read a range as CSV text |
| `read_sheet_csv` | Read an entire sheet as CSV |
| `get_selection` | Get the current selection |
| `get_used_range` | Get the used range dimensions |
| `list_sheets` | List all sheets with metadata |
| `switch_sheet` | Activate a sheet |
| `create_sheet` | Create a new sheet |
| `rename_sheet` | Rename a sheet |
| `delete_sheet` | Delete a sheet |
| `insert_row` | Insert rows |
| `delete_row` | Delete rows |
| `insert_column` | Insert columns |
| `delete_column` | Delete columns |
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
export EXCEL_MCP_SECRET="your-shared-secret"
export EXCEL_MCP_BACKEND_URL="https://localhost:8000/api/mcp/execute"
```

### 3. Start the relay backend

```bash
uvicorn excel_mcp.relay:app --host 0.0.0.0 --port 8000
```

### 4. Start the MCP server

```bash
python -m excel_mcp
```

### 5. Sideload the Office add-in

```bash
cd addin
npm install
npm run dev-server
```

Then sideload `manifest.xml` in Excel ([instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)).

### 6. Connect your AI client

Add to your Claude Code MCP config:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "excel_mcp"],
      "env": {
        "EXCEL_MCP_SECRET": "your-shared-secret",
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
- **`addin/`** — Office.js add-in (npm package)
  - `src/taskpane/services/OfficeService.ts` — tool implementations
  - `src/taskpane/taskpane.ts` — MCP event loop

## License

MIT
