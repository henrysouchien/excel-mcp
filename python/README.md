# excel-mcp (Python)

Shared Python components for Excel MCP integrations:
- `excel_mcp.relay`: FastAPI relay app (`uvicorn excel_mcp.relay:app`)
- `excel_mcp.tool_registry`: built-in Excel tools + `register_tools()`
- `excel_mcp.mcp_server`: MCP stdio server implementation

## Install

```bash
pip install -e ./packages/excel-mcp/python
```
