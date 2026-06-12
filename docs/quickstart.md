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

### 2. Configure a gateway MCP key

```bash
export EXCEL_MCP_API_KEY="your-channel-mcp-gateway-user-key"
export SESSION_API_KEY="your-channel-excel-gateway-user-key"
export EXCEL_MCP_BACKEND_URL="https://localhost:8000/api/mcp/execute"
```

Use a `channel="mcp"` key for the stdio MCP proxy and a `channel="excel"` key
for the taskpane session. The relay path uses gateway JWT sessions; the old
shared-secret header is not supported.

### 3. Start the gateway

From a session with services-mcp enabled:

```text
service_start research_gateway
```

The gateway bridges MCP requests to the Excel add-in. The catalog service name
is `research_gateway`, but it starts the single local gateway process
(`api.main:app`) through the risk_module launcher. Local raw uvicorn skips the
launcher bridge, so it does not hydrate SSM/resolver/web-operator-key env and
Research will not serve. This does not apply to prod, where systemd correctly
runs uvicorn on EC2:8001. See
[LOCAL_STACK_RUNBOOK.md](../../../docs/setup/LOCAL_STACK_RUNBOOK.md).

Package-level `create_relay_app()` is only a factory; embedded applications
must provide a request authenticator.

### 4. Start a taskpane

For the `AI-excel-addin` product repo, use the root Hank AI taskpane. It is the
user-facing add-in and embeds the Excel MCP bridge behind the chat/artifact UI:

```bash
cd /Users/henrychien/Documents/Jupyter/AI-excel-addin
npm run dev-server
npm start
```

This starts the product add-in dev server at `https://localhost:3002` and
sideloads the root `manifest.xml`.

The package-level `packages/excel-mcp/addin` manifest is an internal standalone
bridge harness for package development. It uses a separate add-in id and
`https://localhost:3102` so it does not replace the Hank AI taskpane.

Manual sideloading:
- **Windows**: Insert > My Add-ins > Upload My Add-in
- **Mac**: Insert > Add-ins > My Add-ins > Upload My Add-in

See [Microsoft's sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) for details.

### 5. Start the MCP server

```bash
EXCEL_MCP_API_KEY="$EXCEL_MCP_API_KEY" python -m excel_mcp
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
        "EXCEL_MCP_API_KEY": "your-channel-mcp-gateway-user-key",
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
| `EXCEL_MCP_API_KEY` | Yes | — | User-scoped `channel="mcp"` gateway key used to obtain JWT sessions |
| `EXCEL_MCP_BACKEND_URL` | No | `https://localhost:8000/api/mcp/execute` | Relay execute endpoint |
| `EXCEL_MCP_BACKEND_BASE_URL` | No | Derived from `BACKEND_URL` | Base URL for relay (used for events/status) |
