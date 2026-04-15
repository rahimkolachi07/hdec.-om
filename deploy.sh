#!/bin/bash
# HDEC Deployment Script
# Run this on the server after every git pull
# Usage: bash deploy.sh

set -e

APP_DIR="/var/www/hdec"
LOG_DIR="/var/log/hdec"
SERVICE_NAME="hdec"

echo "=== HDEC Deploy Script ==="

# 1. Make sure log directory exists
mkdir -p "$LOG_DIR"

# 2. Kill any old runserver or rogue gunicorn processes
echo "[1] Stopping old processes..."
pkill -f "manage.py runserver" 2>/dev/null && echo "  Killed runserver" || echo "  No runserver running"
pkill -f "gunicorn.*enterprise_hub" 2>/dev/null && echo "  Killed old gunicorn" || echo "  No gunicorn running"
sleep 1

# 3. Pull latest code
echo "[2] Pulling latest code..."
cd "$APP_DIR"
git pull

# 4. Install / update Python dependencies
echo "[3] Installing requirements..."
pip install -r requirements.txt --quiet

# 5. Copy systemd service file and reload if changed
echo "[4] Installing systemd service..."
cp "$APP_DIR/hdec.service" /etc/systemd/system/hdec.service
systemctl daemon-reload
systemctl enable hdec

# 6. Install Caddyfile if Caddy is present
if command -v caddy &>/dev/null; then
    echo "[5] Reloading Caddy..."
    cp "$APP_DIR/Caddyfile" /etc/caddy/Caddyfile
    systemctl reload caddy 2>/dev/null || systemctl restart caddy
else
    echo "[5] Caddy not found — see README for Caddy install instructions"
fi

# 7. Restart gunicorn via systemd
echo "[6] Restarting Django app..."
systemctl restart hdec

# 8. Show status
echo ""
echo "=== Status ==="
systemctl status hdec --no-pager -l
echo ""
echo "=== Done! Site should be live at https://hdec-om.live ==="
