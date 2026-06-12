from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List, Set


EXCEL_TOOL_SPECS: List[Dict[str, Any]] = [{'name': 'read_cells',
  'description': 'Read values from a small range on the active worksheet and return a 2D array. Active-sheet tool: '
                 'switch sheets first if needed. Best for small ranges (<50 cells), spot checks, and header validation '
                 'before updates. Set include_formulas=true to return formulas instead of computed values. For larger '
                 'reads, prefer read_range_csv or read_sheet_csv.',
  'input_schema': {'type': 'object',
                   'properties': {'range': {'type': 'string', 'description': 'Excel range (e.g., A1:C10). Keep small.'},
                                  'include_formulas': {'type': 'boolean', 'default': False}},
                   'required': ['range']}},
 {'name': 'write_cells',
  'description': 'Write values to a range on the active worksheet. Active-sheet tool: switch sheets first if needed. '
                 'Accepts a scalar for single-cell writes or a 2D array matching the target range dimensions. '
                 'Overwrites content but does not change formatting unless number_format is provided. A committed '
                 'write returns restore_token and undo_tool=restore_written_cells so callers can restore the prior '
                 'cell contents and formatting.',
  'input_schema': {'type': 'object',
                   'properties': {'range': {'type': 'string'},
                                  'values': {'oneOf': [{'type': 'string'},
                                                       {'type': 'number'},
                                                       {'type': 'null'},
                                                       {'type': 'array',
                                                        'items': {'type': 'array',
                                                                  'items': {'oneOf': [{'type': 'string'},
                                                                                      {'type': 'number'},
                                                                                      {'type': 'null'}]}}}]},
                                  'number_format': {'type': 'string'}},
                   'required': ['range', 'values']}},
 {'name': 'restore_written_cells',
  'description': 'Undo write_cells using the restore_token returned by that write.',
  'input_schema': {'type': 'object',
                   'properties': {'restore_token': {'type': 'string',
                                                    'description': 'restore_token returned by write_cells.'}},
                   'required': ['restore_token']}},
 {'name': 'get_selection',
  'description': 'Get the currently selected range, including address, values, and formulas. Useful when the user '
                 "refers to 'this cell' or 'selected range'.",
  'input_schema': {'type': 'object', 'properties': {}}},
 {'name': 'get_used_range',
  'description': 'Get the used-range bounds for a sheet (or active sheet by default). Returns the bounding rectangle '
                 'with content, useful for sizing reads and determining last used rows/columns.',
  'input_schema': {'type': 'object', 'properties': {'sheet_name': {'type': 'string'}}}},
 {'name': 'list_sheets',
  'description': 'List workbook sheets with visibility, active-sheet status, and used-range metadata for quick '
                 'workbook structure discovery.',
  'input_schema': {'type': 'object', 'properties': {}}},
 {'name': 'switch_sheet',
  'description': 'Activate a worksheet by name so active-sheet tools (read_cells, write_cells, format_cells, '
                 'create_table) run on the intended tab.',
  'input_schema': {'type': 'object',
                   'properties': {'sheet_name': {'type': 'string', 'description': 'Exact worksheet name to activate.'}},
                   'required': ['sheet_name']}},
 {'name': 'create_sheet',
  'description': 'Create a worksheet tab, optionally naming it and activating it.',
  'input_schema': {'type': 'object',
                   'properties': {'sheet_name': {'type': 'string',
                                                 'description': 'Optional name for the new worksheet.'},
                                  'activate': {'type': 'boolean',
                                               'default': True,
                                               'description': 'Activate the created sheet.'}}}},
 {'name': 'rename_sheet',
  'description': 'Rename a worksheet tab. By default renames the active sheet; pass sheet_name to rename a specific '
                 'sheet. A committed rename returns restore_token and undo_tool=restore_renamed_sheet so callers can '
                 'undo the rename.',
  'input_schema': {'type': 'object',
                   'properties': {'new_name': {'type': 'string', 'description': 'New worksheet name.'},
                                  'sheet_name': {'type': 'string',
                                                 'description': 'Worksheet to rename (optional; defaults to active '
                                                                'sheet).'}},
                   'required': ['new_name']}},
 {'name': 'restore_renamed_sheet',
  'description': 'Undo rename_sheet using the restore_token returned by that rename.',
  'input_schema': {'type': 'object',
                   'properties': {'restore_token': {'type': 'string',
                                                    'description': 'restore_token returned by rename_sheet.'}},
                   'required': ['restore_token']}},
 {'name': 'delete_sheet',
  'description': 'Delete a worksheet tab by name. Destructive action: requires force=true. A committed delete returns '
                 'restore_token and undo_tool=restore_deleted_sheet so callers can restore the sheet.',
  'input_schema': {'type': 'object',
                   'properties': {'sheet_name': {'type': 'string', 'description': 'Exact worksheet name to delete.'},
                                  'force': {'type': 'boolean',
                                            'default': False,
                                            'description': 'Must be true to confirm deletion after reviewing the '
                                                           'restore path.'}},
                   'required': ['sheet_name']}},
 {'name': 'delete_row',
  'description': 'Delete one or more whole rows and shift cells up. Destructive action: requires force=true. A '
                 'committed delete returns restore_token and undo_tool=restore_deleted_row so callers can restore the '
                 'rows.',
  'input_schema': {'type': 'object',
                   'properties': {'row': {'type': 'integer',
                                          'minimum': 1,
                                          'description': '1-based starting row index.'},
                                  'count': {'type': 'integer',
                                            'minimum': 1,
                                            'default': 1,
                                            'description': 'Number of rows to delete.'},
                                  'sheet_name': {'type': 'string',
                                                 'description': 'Worksheet to update (optional; defaults to active '
                                                                'sheet).'},
                                  'force': {'type': 'boolean',
                                            'default': False,
                                            'description': 'Must be true to confirm deletion after reviewing the '
                                                           'restore path.'}},
                   'required': ['row']}},
 {'name': 'restore_deleted_row',
  'description': 'Restore rows deleted by delete_row using the restore_token returned by that delete.',
  'input_schema': {'type': 'object',
                   'properties': {'restore_token': {'type': 'string',
                                                    'description': 'restore_token returned by delete_row.'}},
                   'required': ['restore_token']}},
 {'name': 'insert_row',
  'description': 'Insert one or more whole rows and shift cells down.',
  'input_schema': {'type': 'object',
                   'properties': {'row': {'type': 'integer',
                                          'minimum': 1,
                                          'description': '1-based row index where rows are inserted.'},
                                  'count': {'type': 'integer',
                                            'minimum': 1,
                                            'default': 1,
                                            'description': 'Number of rows to insert.'},
                                  'sheet_name': {'type': 'string',
                                                 'description': 'Worksheet to update (optional; defaults to active '
                                                                'sheet).'}},
                   'required': ['row']}},
 {'name': 'insert_column',
  'description': 'Insert one or more whole columns and shift cells right.',
  'input_schema': {'type': 'object',
                   'properties': {'column': {'oneOf': [{'type': 'string'}, {'type': 'integer'}],
                                             'description': "Starting column as letter (e.g., 'C') or 1-based index."},
                                  'count': {'type': 'integer',
                                            'minimum': 1,
                                            'default': 1,
                                            'description': 'Number of columns to insert.'},
                                  'sheet_name': {'type': 'string',
                                                 'description': 'Worksheet to update (optional; defaults to active '
                                                                'sheet).'}},
                   'required': ['column']}},
 {'name': 'delete_column',
  'description': 'Delete one or more whole columns and shift cells left. Destructive action: requires force=true. A '
                 'committed delete returns restore_token and undo_tool=restore_deleted_column so callers can restore '
                 'the columns.',
  'input_schema': {'type': 'object',
                   'properties': {'column': {'oneOf': [{'type': 'string'}, {'type': 'integer'}],
                                             'description': "Starting column as letter (e.g., 'C') or 1-based index."},
                                  'count': {'type': 'integer',
                                            'minimum': 1,
                                            'default': 1,
                                            'description': 'Number of columns to delete.'},
                                  'sheet_name': {'type': 'string',
                                                 'description': 'Worksheet to update (optional; defaults to active '
                                                                'sheet).'},
                                  'force': {'type': 'boolean',
                                            'default': False,
                                            'description': 'Must be true to confirm irreversible deletion.'}},
                   'required': ['column']}},
 {'name': 'restore_deleted_sheet',
  'description': 'Restore a worksheet deleted by delete_sheet using the restore_token returned by that delete.',
  'input_schema': {'type': 'object',
                   'properties': {'restore_token': {'type': 'string',
                                                    'description': 'restore_token returned by delete_sheet.'}},
                   'required': ['restore_token']}},
 {'name': 'restore_deleted_column',
  'description': 'Restore columns deleted by delete_column using the restore_token returned by that delete.',
  'input_schema': {'type': 'object',
                   'properties': {'restore_token': {'type': 'string',
                                                    'description': 'restore_token returned by delete_column.'}},
                   'required': ['restore_token']}},
 {'name': 'create_table',
  'description': 'Create an Excel table from a range on the active worksheet. Active-sheet tool: switch sheets first '
                 'if needed. Tables add filter controls and structured references. Set has_headers=true when first row '
                 'has column names.',
  'input_schema': {'type': 'object',
                   'properties': {'range': {'type': 'string'},
                                  'has_headers': {'type': 'boolean', 'default': True},
                                  'table_name': {'type': 'string'}},
                   'required': ['range']}},
 {'name': 'format_cells',
  'description': 'Apply formatting to cells on the active worksheet. Active-sheet tool: switch sheets first if needed. '
                 'Only provided properties are changed; omitted properties remain unchanged.',
  'input_schema': {'type': 'object',
                   'properties': {'range': {'type': 'string',
                                            'description': "Cell or range to format (e.g. 'A1', 'B2:D10')."},
                                  'bold': {'type': 'boolean', 'description': 'Set true to bold, false to un-bold.'},
                                  'italic': {'type': 'boolean',
                                             'description': 'Set true for italic, false to remove italic.'},
                                  'fill_color': {'type': 'string',
                                                 'description': "Background color as hex (e.g. '#FFFF00') or 'none' to "
                                                                'clear.'},
                                  'font_color': {'type': 'string',
                                                 'description': "Font color as hex (e.g. '#FF0000') or 'auto' to reset "
                                                                'to black.'},
                                  'number_format': {'type': 'string',
                                                    'description': "Excel number format (e.g. '#,##0', '0.00%')."},
                                  'horizontal_alignment': {'type': 'string',
                                                           'enum': ['left',
                                                                    'center',
                                                                    'right',
                                                                    'justify',
                                                                    'distributed'],
                                                           'description': 'Horizontal alignment for the range.'}},
                   'required': ['range']}},
 {'name': 'find_cells',
  'description': 'Find cells by value, fill color, or font formatting. Returns matching addresses and values. One '
                 'search criterion per call. Value searches use substring matching. Results are capped at 100 matches; '
                 'format searches cap scanning at 10,000 cells.',
  'input_schema': {'type': 'object',
                   'properties': {'search_type': {'type': 'string',
                                                  'enum': ['value', 'fill_color', 'font'],
                                                  'description': "What to search by: 'value' (cell contents), "
                                                                 "'fill_color' (background color), 'font' "
                                                                 '(bold/italic/color)'},
                                  'query': {'type': 'string',
                                            'description': 'For value: text/number to match (substring). For '
                                                           "fill_color: color name ('yellow','red','green','blue') or "
                                                           'hex (\'#FFFF00\'). For font: JSON e.g. \'{"bold":true}\' '
                                                           'or \'{"color":"#FF0000"}\'.'},
                                  'range': {'type': 'string',
                                            'description': "Range to search (e.g., 'F5:F190'). Defaults to used range "
                                                           'on active sheet.'},
                                  'sheet_name': {'type': 'string',
                                                 'description': 'Sheet to search. Defaults to active sheet.'}},
                   'required': ['search_type', 'query']}},
 {'name': 'read_range_csv',
  'description': 'Read a specific range as CSV text. Best for large ranges (50+ cells), financial tables, or when you '
                 'already know the range. Response is capped at 500 rows per call; use hasMore=true to continue '
                 'pagination.',
  'input_schema': {'type': 'object',
                   'properties': {'range': {'type': 'string', 'description': 'Excel range (e.g., A1:S100).'},
                                  'sheet_name': {'type': 'string',
                                                 'description': 'Sheet name (optional, defaults to active sheet).'}},
                   'required': ['range']}},
 {'name': 'read_sheet_csv',
  'description': 'Read the used range of an entire sheet as CSV text. Best for understanding full model layout before '
                 'targeted operations. Response is capped at 500 rows per call; use hasMore=true to continue '
                 'pagination.',
  'input_schema': {'type': 'object',
                   'properties': {'sheet_name': {'type': 'string',
                                                 'description': 'Sheet name (optional, defaults to active sheet).'}},
                   'required': []}}]

_TOOL_NAMES: Set[str] = {spec["name"] for spec in EXCEL_TOOL_SPECS}
_REGISTERED_TOOL_SPECS: List[Dict[str, Any]] = []
_RELAY_LOCAL_TOOL_SPECS: List[Dict[str, Any]] = [
  {
    "name": "list_workbooks",
    "description": "List Excel workbooks currently connected via the excel-mcp taskpane.",
    "input_schema": {
      "type": "object",
      "properties": {},
    },
  },
  {
    "name": "switch_active_workbook",
    "description": "Set the active workbook target for subsequent tool calls. Use session token from list_workbooks.",
    "input_schema": {
      "type": "object",
      "properties": {
        "session": {
          "type": "string",
          "description": "Session token from list_workbooks",
        },
      },
      "required": ["session"],
    },
  },
]
_WORKBOOK_BOUND_TOOL_NAMES: Set[str] = {spec["name"] for spec in EXCEL_TOOL_SPECS}
_WORKBOOK_OVERRIDE_SCHEMA: Dict[str, Any] = {
  "type": "string",
  "description": "Advanced: session token from list_workbooks; omit to use the active workbook.",
}
_TOOL_NAMES.update(spec["name"] for spec in _RELAY_LOCAL_TOOL_SPECS)

_ACTIVE_SHEET_TOOLS: Set[str] = {
  "read_cells",
  "write_cells",
  "get_selection",
  "create_table",
  "format_cells",
  "find_cells",
}
_SHEET_SCOPED_TOOLS: Set[str] = {
  "get_used_range",
  "rename_sheet",
  "delete_sheet",
  "delete_row",
  "insert_row",
  "insert_column",
  "delete_column",
  "read_range_csv",
  "read_sheet_csv",
}
_RESTORE_TOOLS: Set[str] = {
  "restore_renamed_sheet",
  "restore_written_cells",
  "restore_deleted_sheet",
  "restore_deleted_row",
  "restore_deleted_column",
}


def _agent_ready_description(spec: Dict[str, Any]) -> str:
  """Return the tool description exposed through MCP list-tools."""
  name = str(spec["name"])
  schema = spec.get("input_schema", {})
  properties = schema.get("properties", {}) if isinstance(schema, dict) else {}
  required = schema.get("required", []) if isinstance(schema, dict) else []
  pieces = [str(spec["description"]).strip()]

  if name in _WORKBOOK_BOUND_TOOL_NAMES:
    pieces.append(
      "Workbook routing: omit `_workbook` to use the active workbook, or pass a session token from "
      "`list_workbooks` when multiple Excel taskpanes are connected."
    )
  if name in _ACTIVE_SHEET_TOOLS:
    pieces.append(
      "Discovery: run `list_sheets` first when the target tab is ambiguous, then `switch_sheet` before "
      "calling this active-sheet tool."
    )
  elif name in _SHEET_SCOPED_TOOLS or "sheet_name" in properties:
    pieces.append(
      "Discovery: run `list_sheets` first to obtain the exact worksheet name and current used-range "
      "metadata before choosing sheet_name."
    )
  if name == "list_workbooks":
    pieces.append(
      "Use this before `_workbook` overrides or `switch_active_workbook`; the returned session token is "
      "the stable workbook selector for the current relay."
    )
  elif name == "switch_active_workbook":
    pieces.append(
      "Discovery: run `list_workbooks` first and pass the exact session token for the workbook that "
      "should become active."
    )

  if required:
    pieces.append(f"Required fields: {', '.join(str(item) for item in required)}.")
  if properties.get("force") is not None:
    pieces.append(
      "Safety: do not set force=true until the caller has reviewed the target workbook, sheet, row, or "
      "column and accepted the restore-token recovery path."
    )
  if name in _RESTORE_TOOLS:
    pieces.append(
      "Use only the restore_token returned by the paired mutating tool; tokens are operation-specific "
      "and should not be guessed or reused across workbooks."
    )
  if "restore_token" in str(spec.get("description", "")):
    pieces.append(
      "Next action: keep the returned restore_token with the undo_tool value so a later agent can reverse "
      "the committed operation without inferring workbook state."
    )

  pieces.append(
    "Result contract: successful calls return a JSON object; proxy failures return a structured "
    "status=error envelope with code, message, recoverable, tool_name, and suggested_tool_calls."
  )
  pieces.append(
    "Common mistake: do not rely on the visible Excel tab or another user's workbook when a workflow "
    "already has an explicit session or sheet target."
  )
  return " ".join(piece for piece in pieces if piece)


def register_tools(*specs: Dict[str, Any]) -> None:
  """Register custom tool specifications before MCP server startup."""
  for spec in specs:
    if not isinstance(spec, dict):
      raise TypeError(f"Tool spec must be a dict, got {type(spec).__name__}")
    for key in ("name", "description", "input_schema"):
      if key not in spec:
        raise ValueError(f"Tool spec missing required key: {key}")

    name = spec["name"]
    if not isinstance(name, str) or not name.strip():
      raise ValueError("Tool spec 'name' must be a non-empty string")

    if name in _TOOL_NAMES:
      raise ValueError(f"Duplicate tool name: {name}")

    _TOOL_NAMES.add(name)
    _REGISTERED_TOOL_SPECS.append(spec)


def get_tool_specs() -> List[Dict[str, Any]]:
  """Return all tool specs (built-in + registered)."""
  specs: List[Dict[str, Any]] = []

  for spec in EXCEL_TOOL_SPECS + _RELAY_LOCAL_TOOL_SPECS + _REGISTERED_TOOL_SPECS:
    copied = deepcopy(spec)
    if spec["name"] in _WORKBOOK_BOUND_TOOL_NAMES:
      input_schema = copied.setdefault("input_schema", {})
      properties = input_schema.setdefault("properties", {})
      properties.setdefault("_workbook", deepcopy(_WORKBOOK_OVERRIDE_SCHEMA))
    copied["description"] = _agent_ready_description(copied)
    specs.append(copied)

  return specs


def get_tool_names() -> Set[str]:
  """Return all tool names (built-in + registered)."""
  return set(_TOOL_NAMES)
