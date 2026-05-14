#!/bin/bash
set -euo pipefail

if [ "${EUID}" -eq 0 ]; then
  echo "Run this installer as the target user, not as root."
  exit 1
fi

TARGET_USER="${USER}"
HOME_DIR="/home/${TARGET_USER}"
REPO_DIR="${HOME_DIR}/info_screen_desk"
SYSTEMD_DIR="/etc/systemd/system"
ENV_FILE="/etc/info-screen-desk.env"

echo "Installing base packages"
sudo apt update
sudo apt install -y git python3-pip python3-pillow python3-requests python3-dateutil fonts-dejavu python3-gpiozero python3-spidev python3-lgpio python3-pigpio python3-setuptools

echo "Installing Python dependencies"
python3 -m pip install --break-system-packages -r "${REPO_DIR}/requirements.txt"

echo "Installing Waveshare e-Paper library"
if [ ! -d "${HOME_DIR}/e-Paper" ]; then
  git clone https://github.com/waveshare/e-Paper.git "${HOME_DIR}/e-Paper"
fi
cd "${HOME_DIR}/e-Paper/RaspberryPi_JetsonNano/python"
sudo python3 setup.py install

echo "Preparing cache"
mkdir -p "${HOME_DIR}/.cache/info-screen-desk"

if [ ! -f "${ENV_FILE}" ]; then
  echo "Creating ${ENV_FILE}"
  sudo install -m 640 -o root -g "${TARGET_USER}" /dev/null "${ENV_FILE}"
  sudo tee "${ENV_FILE}" >/dev/null <<EOF
DESK_AUTH_MODE=device_code
DESK_TENANT_ID=
DESK_CLIENT_ID=
DESK_CLIENT_SECRET=
DESK_CALENDAR_EMAIL=
DESK_DELEGATED_SCOPES="https://graph.microsoft.com/User.Read https://graph.microsoft.com/Calendars.Read https://graph.microsoft.com/Mail.Read"
DESK_USER_NAME=Desk
DESK_TIMEZONE=Europe/Amsterdam
DESK_OFFICE_START_HOUR=8
DESK_OFFICE_END_HOUR=18
DESK_PRIVACY_MODE=normal
DESK_MEETING_ALERT_MINUTES=5
DESK_REFRESH_INTERVAL_SECONDS=300
DESK_TIMER_REFRESH_SECONDS=60
DESK_HEARTBEAT_INTERVAL_SECONDS=25
DESK_BUTTON_HOLD_SECONDS=1.8
DESK_EMAIL_ENABLED=1
DESK_EMAIL_REFRESH_MINUTES=10
DESK_WEATHER_ENABLED=1
DESK_WEATHER_LATITUDE=52.3676
DESK_WEATHER_LONGITUDE=4.9041
DESK_WEATHER_LABEL=Amsterdam
DESK_WEATHER_REFRESH_MINUTES=30
DESK_COMMUTE_WEATHER_HOUR=18
DESK_WEATHER_RAIN_THRESHOLD_PERCENT=35
DESK_WEATHER_PRECIP_THRESHOLD_MM=0.2
DESK_WEATHER_WIND_THRESHOLD_KMH=35
DESK_WEATHER_COLD_THRESHOLD_C=3
DESK_WEATHER_HOT_THRESHOLD_C=28
DESK_NIGHTLY_CLEAR_ENABLED=1
DESK_NIGHTLY_CLEAR_HOUR=3
DESK_PROFILE_MODE=auto
DESK_PROFILE_MEETING_HEAVY_COUNT=5
DESK_PROFILE_MEETING_HEAVY_BUSY_MINUTES=240
DESK_PROFILE_FOCUS_DAY_OPEN_HOURS=4
DESK_SAVE_PREVIEWS=1
DESK_PREVIEW_DIR=${REPO_DIR}/previews
DESK_STATE_FILE=${HOME_DIR}/.cache/info-screen-desk/state.json
DESK_TOKEN_CACHE_FILE=${HOME_DIR}/.cache/info-screen-desk/token-cache.json
DESK_HEARTBEAT_FILE=${HOME_DIR}/.cache/info-screen-desk/heartbeat.txt
DESK_LOCK_FILE=/tmp/info-screen-desk.lock
DESK_LOG_LEVEL=INFO
DESK_EPD_DRIVER=epd2in7_V2
DESK_DISPLAY_ROTATION=0
DESK_BUTTON_HOME_PIN=5
DESK_BUTTON_AGENDA_PIN=6
DESK_BUTTON_FOCUS_PIN=13
DESK_BUTTON_REFRESH_PIN=19
EOF
fi

echo "Installing systemd units"
sed "s/__DESK_USER__/${TARGET_USER}/g" "${REPO_DIR}/systemd/info-screen-desk.service" | sudo tee "${SYSTEMD_DIR}/info-screen-desk.service" >/dev/null
sed "s/__DESK_USER__/${TARGET_USER}/g" "${REPO_DIR}/systemd/info-screen-desk-buttons.service" | sudo tee "${SYSTEMD_DIR}/info-screen-desk-buttons.service" >/dev/null
sudo systemctl daemon-reload
sudo systemctl enable info-screen-desk-buttons.service

echo "Installation complete."
echo "Edit ${ENV_FILE}, then run device login if needed:"
echo "set -a; . ${ENV_FILE}; set +a; python3 ${REPO_DIR}/src/info_screen.py --login"
