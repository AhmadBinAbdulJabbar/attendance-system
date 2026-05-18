#!/bin/bash
set -e
cd "$(dirname "$0")"

if [ ! -d .venv ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate

echo "Installing dependencies..."
pip install -r requirements.txt --quiet

if [ ! -f .env ] && [ -f .env.example ]; then
  echo "Tip: copy .env.example to .env and set SMTP_* variables to enable Send Report email."
fi

echo "Starting server at http://localhost:8000"
exec python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
