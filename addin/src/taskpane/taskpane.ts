/* global document, Office */

import { OfficeService } from "./services/OfficeService";

const API_BASE = process.env.API_BASE || "https://localhost:8000";
const MCP_RECONNECT_DELAY_MS = 3000;

interface ToolErrorPayload {
  code: string;
  message: string;
  details?: any;
}

interface ToolExecutionOutcome {
  result: Record<string, any> | null;
  error: ToolErrorPayload | null;
}

interface McpToolRequestEvent {
  type: "mcp_tool_request";
  request_id: string;
  nonce: string;
  tool_name: string;
  tool_input: Record<string, any>;
  replay?: boolean;
}

const elements: {
  sideloadMsg?: HTMLElement | null;
  app?: HTMLElement | null;
  status?: HTMLElement | null;
  statusText?: HTMLElement | null;
} = {};

let mcpStopped = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initialize();
  }
});

function initialize(): void {
  elements.sideloadMsg = document.getElementById("sideload-msg");
  elements.app = document.getElementById("app");
  elements.status = document.getElementById("status-dot");
  elements.statusText = document.getElementById("status-text");

  if (elements.sideloadMsg) elements.sideloadMsg.style.display = "none";
  if (elements.app) elements.app.style.display = "block";

  void connectLoop();
}

function setStatus(text: string, state: "connected" | "disconnected" | "working"): void {
  if (elements.statusText) elements.statusText.textContent = text;
  if (elements.status) {
    elements.status.className = `status-dot ${state}`;
  }
}

function getMcpSecret(): string {
  const envSecret = process.env.EXCEL_MCP_SECRET || "";
  const sessionSecret = sessionStorage.getItem("excel_mcp_secret") || "";
  return (envSecret || sessionSecret).trim();
}

async function connectLoop(): Promise<void> {
  while (!mcpStopped) {
    const secret = getMcpSecret();
    if (!secret) {
      setStatus("Waiting for EXCEL_MCP_SECRET", "disconnected");
      await sleep(MCP_RECONNECT_DELAY_MS);
      continue;
    }

    try {
      await connectOnce(secret);
    } catch (error) {
      console.error("[MCP] stream error", error);
      setStatus("Disconnected", "disconnected");
    }

    await sleep(MCP_RECONNECT_DELAY_MS);
  }
}

async function connectOnce(secret: string): Promise<void> {
  setStatus("Connecting...", "working");
  const response = await fetch(`${API_BASE}/api/mcp/events?secret=${encodeURIComponent(secret)}`, {
    method: "GET",
    headers: { Accept: "text/event-stream" },
  });

  if (!response.ok || !response.body) {
    const text = await response.text().catch(() => "");
    throw new Error(text || `MCP stream failed (${response.status})`);
  }

  setStatus("Connected", "connected");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const chunk = await reader.read();
    if (chunk.done) {
      break;
    }
    buffer += decoder.decode(chunk.value, { stream: true });
    const parts = buffer.split("\n\n");
    buffer = parts.pop() || "";
    for (const part of parts) {
      const line = part
        .split("\n")
        .find((candidate) => candidate.startsWith("data: "));
      if (!line) continue;

      const payload = line.slice(6);
      let event: any;
      try {
        event = JSON.parse(payload);
      } catch {
        continue;
      }

      if (event.type === "mcp_tool_request") {
        setStatus(`Running ${event.tool_name}`, "working");
        void handleToolRequest(event as McpToolRequestEvent);
      } else if (event.type === "heartbeat") {
        setStatus("Connected", "connected");
      } else if (event.type === "replaced") {
        setStatus("Replaced by newer session", "disconnected");
        return;
      }
    }
  }
}

async function handleToolRequest(event: McpToolRequestEvent): Promise<void> {
  const secret = getMcpSecret();
  if (!secret) return;

  await postToolResult({ request_id: event.request_id, nonce: event.nonce, ack: true }, secret);

  const outcome = await executeOfficeServiceTool(event.tool_name, event.tool_input);

  await postToolResult(
    {
      request_id: event.request_id,
      nonce: event.nonce,
      result: outcome.result,
      error: outcome.error,
    },
    secret
  );

  setStatus("Connected", "connected");
}

async function postToolResult(payload: Record<string, any>, secret: string): Promise<void> {
  const response = await fetch(`${API_BASE}/api/mcp/tool-result`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-MCP-Secret": secret,
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => "");
    throw new Error(text || `Failed to post tool result (${response.status})`);
  }
}

async function executeOfficeServiceTool(toolName: string, toolInput: Record<string, any>): Promise<ToolExecutionOutcome> {
  try {
    let result: any;
    switch (toolName) {
      case "read_cells":
        result = await OfficeService.readCells(toolInput.range, toolInput.include_formulas);
        break;
      case "write_cells":
        result = await OfficeService.writeCells(toolInput.range, toolInput.values, toolInput.number_format);
        break;
      case "get_selection":
        result = await OfficeService.getSelection();
        break;
      case "get_used_range":
        result = await OfficeService.getUsedRange(toolInput.sheet_name);
        break;
      case "list_sheets":
        result = await OfficeService.listSheets();
        break;
      case "switch_sheet":
        result = await OfficeService.switchSheet(toolInput.sheet_name);
        break;
      case "create_sheet":
        result = await OfficeService.createSheet(toolInput.sheet_name, toolInput.activate);
        break;
      case "rename_sheet":
        result = await OfficeService.renameSheet(toolInput.new_name, toolInput.sheet_name);
        break;
      case "delete_sheet":
        result = await OfficeService.deleteSheet(toolInput.sheet_name, toolInput.force);
        break;
      case "delete_row":
        result = await OfficeService.deleteRow(toolInput.row, toolInput.count, toolInput.sheet_name, toolInput.force);
        break;
      case "insert_row":
        result = await OfficeService.insertRow(toolInput.row, toolInput.count, toolInput.sheet_name);
        break;
      case "insert_column":
        result = await OfficeService.insertColumn(toolInput.column, toolInput.count, toolInput.sheet_name);
        break;
      case "delete_column":
        result = await OfficeService.deleteColumn(toolInput.column, toolInput.count, toolInput.sheet_name, toolInput.force);
        break;
      case "create_table":
        result = await OfficeService.createTable(toolInput.range, toolInput.has_headers, toolInput.table_name);
        break;
      case "format_cells":
        result = await OfficeService.formatCells(toolInput as any);
        break;
      case "find_cells":
        result = await OfficeService.findCells(toolInput as any);
        break;
      case "read_range_csv":
        result = await OfficeService.readRangeCsv(toolInput.range, toolInput.sheet_name);
        break;
      case "read_sheet_csv":
        result = await OfficeService.readSheetCsv(toolInput.sheet_name);
        break;
      default:
        throw new Error(`Unsupported tool: ${toolName}`);
    }

    return { result: result as Record<string, any>, error: null };
  } catch (error) {
    return {
      result: null,
      error: {
        code: "tool_error",
        message: error instanceof Error ? error.message : String(error),
      },
    };
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
