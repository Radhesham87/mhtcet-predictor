"""Gunicorn entry point for deployment.  Run: gunicorn wsgi:app"""
import os
from app import app, init_db, DATA_DIR

os.makedirs(DATA_DIR, exist_ok=True)
init_db()
