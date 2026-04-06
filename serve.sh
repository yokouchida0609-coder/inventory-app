#!/usr/bin/env bash
PORT="${PORT:-3000}"
cd "$(dirname "$0")"
echo "Serving on http://localhost:$PORT"
exec powershell.exe -ExecutionPolicy Bypass -File "$(pwd)/serve.ps1"
