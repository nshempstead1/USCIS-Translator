#!/usr/bin/env bash
set -e

HOST="${HOST:-0.0.0.0}"
PORT="${PORT:-8000}"

# Load .env if it exists
if [ -f .env ]; then
    set -a
    source .env
    set +a
fi

# Check for API key
if [ -z "$GEMINI_API_KEY" ]; then
    echo "ERROR: GEMINI_API_KEY is not set."
    echo "  1. Copy .env.example to .env:  cp .env.example .env"
    echo "  2. Add your Gemini API key to .env"
    exit 1
fi

# Install dependencies if needed
if ! python3 -c "import fastapi" 2>/dev/null; then
    echo "Installing dependencies..."
    pip install -r requirements.txt
fi

echo "Starting USCIS Certified Translation server on http://${HOST}:${PORT}"
exec uvicorn app:app --host "$HOST" --port "$PORT" --reload
