# Quickstart

Get Excel connected to Claude in 5 minutes.

## Prerequisites

- Python 3.10+
- Node.js 18+
- Excel (desktop, with add-in sideloading support)

## Setup

### 1. Install the Python package

```bash
cd packages/excel-mcp
pip install -e ./python
```

### 2. Create a shared secret

```bash
export EXCEL_MCP_SECRET=$(python -c "import secrets; print(secrets.token_urlsafe(32))")
export EXCEL_MCP_USER_ID="your-user-id"
echo "EXCEL_MCP_SECRET=$EXCEL_MCP_SECRET"
```

Save this — both the relay and MCP server need it.

### 3. Start the relay backend

```bash
uvicorn excel_mcp.relay:app --host 0.0.0.0 --port 8000
```

The relay bridges MCP requests to the Excel add-in. It exposes endpoints at `/api/mcp/execute`, `/api/mcp/tool-result`, `/api/mcp/events`, and `/health`.

### 4. Build and sideload the add-in

```bash
cd addin
npm install
npm run dev-server
```

This starts the add-in dev server at `https://localhost:3000`. Sideload `manifest.xml` in Excel:
- **Windows**: Insert > My Add-ins > Upload My Add-in
- **Mac**: Insert > Add-ins > My Add-ins > Upload My Add-in

See [Microsoft's sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) for details.

### 5. Start the MCP server

```bash
EXCEL_MCP_SECRET="your-secret" python -m excel_mcp
```

### 6. Configure Claude Code

Add to your MCP settings (`.claude/settings.json` or project config):

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

## Verify

Open Excel with the add-in loaded, then ask Claude Code:

> "Read the cells in A1:C10"

You should see the MCP server relay the request through to Excel and return the cell values.

## Environment variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `EXCEL_MCP_SECRET` | Yes | — | Shared secret for relay authentication |
| `EXCEL_MCP_USER_ID` | Yes for add-in | — | Explicit user id attached to the workbook SSE session |
| `EXCEL_MCP_BACKEND_URL` | No | `https://localhost:8000/api/mcp/execute` | Relay execute endpoint |
| `EXCEL_MCP_BACKEND_BASE_URL` | No | Derived from `BACKEND_URL` | Base URL for relay (used for events/status) |
