#!/bin/bash
set -e
cd "$(dirname "$0")"
python3 -m pip install -e ./python && python3 -m excel_mcp setup "$@"
