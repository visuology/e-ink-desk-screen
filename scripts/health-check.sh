#!/bin/bash
set -euo pipefail

ENV_FILE="/etc/info-screen-desk.env"
HEARTBEAT_FILE="${HOME}/.cache/info-screen-desk/heartbeat.txt"

echo "Time"
date
echo

echo "Services"
systemctl status info-screen-desk.service --no-pager || true
systemctl status info-screen-desk-buttons.service --no-pager || true
echo

echo "Logs"
journalctl -u info-screen-desk.service -n 30 --no-pager || true
journalctl -u info-screen-desk-buttons.service -n 30 --no-pager || true
echo

echo "Heartbeat"
ls -l "${HEARTBEAT_FILE}" 2>/dev/null || true
tail -n 3 "${HEARTBEAT_FILE}" 2>/dev/null || true
echo

echo "Environment"
sudo ls -l "${ENV_FILE}" 2>/dev/null || true

