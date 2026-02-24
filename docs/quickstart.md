# excel-mcp quickstart

## 1) Install Python package

```bash
pip install -e ./packages/excel-mcp/python
```

## 2) Start relay backend

```bash
uvicorn excel_mcp.relay:app --host 0.0.0.0 --port 8000
```

## 3) Start MCP server

```bash
python -m excel_mcp
```

Set these env vars first:
- `EXCEL_MCP_SECRET`
- `EXCEL_MCP_BACKEND_URL` (defaults to `https://localhost:8000/api/mcp/execute`)
- `EXCEL_MCP_BACKEND_BASE_URL` (optional override)

## 4) Load Office add-in
Build and sideload from `packages/excel-mcp/addin`.
