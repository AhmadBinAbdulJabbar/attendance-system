#!/bin/bash
cd "$(dirname "$0")"
pip install -r requirements.txt --quiet
uvicorn server:app --host 0.0.0.0 --port 8000 --reload
