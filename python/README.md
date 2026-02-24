# excel-mcp

Python package for connecting Excel to AI via MCP.

See the [main README](../README.md) for full documentation.

## Install

```bash
pip install -e .
```

## Components

- `excel_mcp.mcp_server` — MCP stdio server (`python -m excel_mcp`)
- `excel_mcp.relay` — FastAPI relay backend (`uvicorn excel_mcp.relay:app`)
- `excel_mcp.tool_registry` — 18 built-in Excel tools + `register_tools()` extensibility
