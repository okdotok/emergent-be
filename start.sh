#!/bin/bash
python seed_data.py

exec uvicorn server:app --host 0.0.0.0 --port ${PORT:-10000}

