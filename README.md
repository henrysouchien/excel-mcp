# excel-mcp

Reusable Excel MCP infrastructure extracted from the private app.

## Included components
- `python/excel_mcp/relay.py`: FastAPI relay + SSE endpoints
- `python/excel_mcp/tool_registry.py`: built-in Excel tools + `register_tools()` extension point
- `python/excel_mcp/mcp_server.py`: MCP stdio server that proxies to the relay
- `addin/`: Office.js add-in sources for MCP event handling and tool execution

## Quick start
See [docs/quickstart.md](docs/quickstart.md).
