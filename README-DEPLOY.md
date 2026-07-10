# Deploying the web/ version

This folder is the deployable build of the MHT-CET College Predictor.
Same features as local/, prepared for cloud hosting.

## Environment variables (set these on your host)
- SECRET_KEY   - long random string (required in production)
- ADMIN_EMAIL  - admin login email (default admin@mhtcet.local)
- ADMIN_PASSWORD - admin password (default admin123 - CHANGE THIS)

## Render.com (free tier works)
1. Push this web/ folder to a GitHub repo.
2. Render -> New -> Web Service -> connect the repo.
3. Build command:   pip install -r requirements.txt
4. Start command:   ./start.sh
5. Add the environment variables above. Deploy.

## Railway.app
1. New Project -> Deploy from GitHub repo.
2. Railway auto-detects the Procfile. Add env vars. Deploy.

## PythonAnywhere
1. Upload this folder, create a virtualenv, pip install -r requirements.txt.
2. Web tab -> Manual config -> point WSGI file to wsgi.py's `app`.

## Notes
- SQLite database (predictor.db) is created automatically on first boot.
- Upload new cutoff xlsx files from the Admin panel after deploying.
- On free tiers with ephemeral disks, uploaded data and users reset on
  redeploy; move DB/uploads to a persistent disk or object storage for
  production use.
