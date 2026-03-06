# Docker Containerization Plan

> **Status**: PENDING — waiting for Azure app registration to get required permissions
>
> **Prerequisites before implementing**:
> 1. Add `AZURE_CLIENT_SECRET` to `.env`
> 2. Azure app registration needs `Sites.ReadWrite.All` (Application permission) for SharePoint file upload
> 3. Install Docker Desktop for Windows

## Context

Head of IT wants the tool containerized so it runs independently of any specific machine. Email JSON output goes to the **OneDrive-synced SharePoint folder** on the host, which syncs to SharePoint automatically.

## Azure Credentials

Add `AZURE_CLIENT_SECRET` to your `.env` file:
```
AZURE_CLIENT_SECRET=your-new-secret-value
```

This enables non-interactive auth (required for Docker — no browser popup). The code already supports it (see `src/graph_client.py` lines 42-50).

---

## Live Code Changes — No Rebuild Needed

Your source code (`main.py`, `src/`, `config/`) is **volume-mounted** into the container in the local setup:

- Edit files on your PC as usual (VS Code, etc.)
- Changes are **instantly visible** inside the container
- Next cron run picks them up automatically
- Only rebuild when adding new Python packages to `requirements.txt`

## Architecture

```
docker compose up -d
    └── accounting-mailbox-reader container
            ├── cron (scheduler - replaces Windows Task Scheduler)
            │     ├── Every hour 08-17: python main.py process --upload-sharepoint
            │     └── At 17:00: python main.py cleanup --upload-sharepoint
            ├── Source code mounted from your PC (live edits)
            └── Volumes:
                  ├── .env (secrets — includes AZURE_CLIENT_SECRET)
                  ├── data/ (categories cache)
                  ├── SharePoint sync folder (OneDrive auto-syncs to SharePoint)
                  └── app-data (token caches, stats, logs)
```

---

## Files to Create

### 1. `Dockerfile`

```dockerfile
FROM python:3.13-slim

# Install ODBC Driver 18 for SQL Server + cron
RUN apt-get update && \
    apt-get install -y --no-install-recommends curl gnupg2 cron && \
    curl -fsSL https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg && \
    echo "deb [signed-by=/usr/share/keyrings/microsoft-prod.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" > /etc/apt/sources.list.d/mssql-release.list && \
    apt-get update && \
    ACCEPT_EULA=Y apt-get install -y msodbcsql18 && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .
COPY src/ src/
COPY config/ config/

RUN mkdir -p /root/.accounting_mailbox_reader

COPY docker/crontab /etc/cron.d/mailbox-cron
RUN chmod 0644 /etc/cron.d/mailbox-cron && crontab /etc/cron.d/mailbox-cron

COPY docker/entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]
```

### 2. `docker-compose.yml`

```yaml
services:
  mailbox-reader:
    build: .
    container_name: accounting-mailbox-reader
    env_file: .env
    environment:
      - LOCAL_FOLDER_PATH=/data/emails
      - TZ=Europe/Zurich
    volumes:
      - ./data:/app/data                          # categories cache
      - app-data:/root/.accounting_mailbox_reader  # token caches, stats, logs
    restart: unless-stopped

volumes:
  app-data:
```

### 3. `docker-compose.override.yml` (git-ignored, local dev only)

```yaml
services:
  mailbox-reader:
    volumes:
      # Live code mount — edit files and changes apply instantly
      - ./main.py:/app/main.py
      - ./src:/app/src
      - ./config:/app/config
      # OneDrive-synced SharePoint folder — files auto-sync to SharePoint
      - "C:\\Users\\MatthiasDiBianco\\planted foods AG\\Circle Finance - Documents\\Accounting\\Mailbox\\emails:/data/emails"
```

Mounts live source code (instant edits) + SharePoint sync folder. When deploying to another machine, don't create this file — use Graph API SharePoint upload instead (already implemented in codebase).

### 4. `docker/crontab`

```cron
# Process emails hourly 08:00-16:00 (Mon-Fri)
0 8-16 * * 1-5 cd /app && python main.py process --upload-sharepoint >> /proc/1/fd/1 2>&1

# At 17:00: process + cleanup (Mon-Fri)
0 17 * * 1-5 cd /app && python main.py process --upload-sharepoint >> /proc/1/fd/1 2>&1 && python main.py cleanup --upload-sharepoint >> /proc/1/fd/1 2>&1
```

`>> /proc/1/fd/1` routes output to `docker logs`.

### 5. `docker/entrypoint.sh`

```bash
#!/bin/bash
set -e

# If a command is passed (e.g. "process", "cleanup"), run it directly
if [ $# -gt 0 ]; then
    exec python main.py "$@"
fi

# Otherwise start cron daemon (scheduled mode)
echo "Starting cron scheduler (TZ=$TZ)..."
printenv | grep -v "no_proxy" >> /etc/environment
exec cron -f
```

Dual-mode: `docker compose up` starts cron scheduler; `docker compose run --rm mailbox-reader process` runs a single command.

### 6. `.dockerignore`

```
.venv/
.git/
__pycache__/
*.pyc
.env
docs/
powerquery/
*.bat
```

---

## Files to Modify

### 7. `main.py` (minor)

Update the hardcoded log path in error notification (line 59) to use `Path.home()` dynamically instead of `C:\\Users\\MatthiasDiBianco\\...`.

### 8. `.gitignore`

Add: `docker-compose.override.yml`

---

## How to Use (end-to-end)

### First Time Setup

```bash
# 1. Install Docker Desktop for Windows (one-time)
# 2. Add AZURE_CLIENT_SECRET to your .env file
# 3. Build the image
docker compose build

# 4. Test with a dry run
docker compose run --rm mailbox-reader process --dry-run

# 5. Test a real run (small batch)
docker compose run --rm mailbox-reader process

# 6. Start scheduled mode (replaces Task Scheduler)
docker compose up -d

# 7. Check it's running
docker compose ps
docker logs accounting-mailbox-reader
```

### Daily Operations

```bash
docker logs accounting-mailbox-reader --tail 100   # view logs
docker compose run --rm mailbox-reader process     # manual run
docker compose down                                 # stop
docker compose up -d                                # start
```

### After Code Changes

**Code edits** (main.py, src/, config/): Nothing to do — volume-mounted, picked up next run.

**Dependency changes** (requirements.txt):
```bash
docker compose build && docker compose up -d
```

---

## Verification Plan

1. `docker compose build` — completes without errors
2. `docker compose run --rm mailbox-reader config-show` — shows AZURE_CLIENT_SECRET as `[OK] Configured`
3. `docker compose run --rm mailbox-reader process --dry-run` — runs pipeline successfully
4. `docker compose run --rm mailbox-reader process` — real run, verify email JSONs appear in the SharePoint sync folder
5. `docker compose up -d` + `docker logs` — shows "Starting cron scheduler"
6. Stop & restart container — verify token cache persists (no re-auth)
7. Disable Windows Task Scheduler once confirmed working

## Machine Independence

Docker solves this completely:
- **Today**: Runs on your PC via Docker Desktop
- **Tomorrow**: Copy project folder to any Linux server → `docker compose up -d`
- **Azure**: Push image to ACR → run as Azure Container Instance
- Only requirement on any machine: Docker + `.env` file
