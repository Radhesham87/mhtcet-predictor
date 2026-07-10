#!/usr/bin/env bash
# Start script for Render / Railway
exec gunicorn wsgi:app --bind 0.0.0.0:${PORT:-8000} --workers 1 --timeout 120
