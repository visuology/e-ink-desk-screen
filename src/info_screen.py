import argparse
import json
import logging
import os
import socket
import subprocess
import sys
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, time as dt_time, timedelta, timezone
from importlib import import_module
from pathlib import Path
from queue import Empty, Queue

import pytz
import requests
from dateutil.parser import isoparse
from msal import ConfidentialClientApplication, PublicClientApplication, SerializableTokenCache
from PIL import Image, ImageChops, ImageDraw, ImageFont


LOG_LEVEL = os.environ.get("DESK_LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s - %(levelname)s - %(message)s",
    stream=sys.stdout,
)
LOGGER = logging.getLogger("info-screen-desk")

BUSY_SHOW_AS = {"busy", "tentative", "oof", "workingElsewhere", "unknown"}
FOCUS_PRESETS = [25, 50, 90]
VIEWS = {"home", "agenda", "focus", "tomorrow"}
RESERVED_MSAL_SCOPES = {"offline_access", "openid", "profile"}
MENU_WIDTH = 32
CONTENT_LEFT = MENU_WIDTH + 8
CONTENT_RIGHT_MARGIN = 8
WEATHER_CODES = {
    0: "clear",
    1: "mainly clear",
    2: "partly cloudy",
    3: "cloudy",
    45: "fog",
    48: "fog",
    51: "drizzle",
    53: "drizzle",
    55: "drizzle",
    61: "rain",
    63: "rain",
    65: "heavy rain",
    71: "snow",
    73: "snow",
    75: "heavy snow",
    80: "showers",
    81: "showers",
    82: "heavy showers",
    95: "thunder",
}


@dataclass(frozen=True)
class Config:
    auth_mode: str
    tenant_id: str
    client_id: str
    client_secret: str
    calendar_email: str
    delegated_scopes: tuple[str, ...]
    user_name: str
    timezone_name: str
    office_start_hour: int
    office_end_hour: int
    privacy_mode: str
    meeting_alert_minutes: int
    refresh_interval_seconds: int
    timer_refresh_seconds: int
    heartbeat_interval_seconds: int
    button_hold_seconds: float
    email_enabled: bool
    email_refresh_minutes: int
    weather_enabled: bool
    weather_latitude: float | None
    weather_longitude: float | None
    weather_label: str
    weather_refresh_minutes: int
    commute_weather_hour: int
    save_previews: bool
    preview_dir: Path
    state_file: Path
    token_cache_file: Path
    heartbeat_file: Path
    lock_file: Path
    epd_driver: str
    display_rotation: int
    button_home_pin: int
    button_agenda_pin: int
    button_focus_pin: int
    button_refresh_pin: int

    @property
    def timezone(self):
        return pytz.timezone(self.timezone_name)


class MockEPD:
    width = 176
    height = 264

    def init(self):
        return None

    def Clear(self):
        return None

    def getbuffer(self, image):
        return image

    def display(self, _buffer):
        return None

    def sleep(self):
        return None


def env_bool(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def env_int(name: str, default: int, minimum: int | None = None) -> int:
    value = int(os.environ.get(name, str(default)))
    return max(minimum, value) if minimum is not None else value


def env_float(name: str) -> float | None:
    value = os.environ.get(name, "").strip()
    if not value:
        return None
    return float(value)


def load_config(require_credentials: bool = True) -> Config:
    auth_mode = os.environ.get("DESK_AUTH_MODE", "device_code").lower()
    if auth_mode not in {"device_code", "client_credentials"}:
        raise RuntimeError("DESK_AUTH_MODE must be device_code or client_credentials")

    tenant_id = os.environ.get("DESK_TENANT_ID", "")
    client_id = os.environ.get("DESK_CLIENT_ID", "")
    client_secret = os.environ.get("DESK_CLIENT_SECRET", "")
    calendar_email = os.environ.get("DESK_CALENDAR_EMAIL", "")

    if require_credentials:
        missing = []
        if not tenant_id:
            missing.append("DESK_TENANT_ID")
        if not client_id:
            missing.append("DESK_CLIENT_ID")
        if auth_mode == "client_credentials":
            if not client_secret:
                missing.append("DESK_CLIENT_SECRET")
            if not calendar_email:
                missing.append("DESK_CALENDAR_EMAIL")
        if missing:
            raise RuntimeError("Missing required environment variables: " + ", ".join(missing))

    cache_dir = Path.home() / ".cache" / "info-screen-desk"
    scopes = tuple(
        scope.strip()
        for scope in os.environ.get(
            "DESK_DELEGATED_SCOPES",
            "https://graph.microsoft.com/User.Read https://graph.microsoft.com/Calendars.Read https://graph.microsoft.com/Mail.Read",
        ).split()
        if scope.strip() and scope.strip() not in RESERVED_MSAL_SCOPES
    )

    return Config(
        auth_mode=auth_mode,
        tenant_id=tenant_id or "sample-tenant",
        client_id=client_id or "sample-client",
        client_secret=client_secret,
        calendar_email=calendar_email,
        delegated_scopes=scopes,
        user_name=os.environ.get("DESK_USER_NAME", "Desk"),
        timezone_name=os.environ.get("DESK_TIMEZONE", "Europe/Amsterdam"),
        office_start_hour=env_int("DESK_OFFICE_START_HOUR", 8),
        office_end_hour=env_int("DESK_OFFICE_END_HOUR", 18),
        privacy_mode=os.environ.get("DESK_PRIVACY_MODE", "normal").lower(),
        meeting_alert_minutes=env_int("DESK_MEETING_ALERT_MINUTES", 5, minimum=1),
        refresh_interval_seconds=env_int("DESK_REFRESH_INTERVAL_SECONDS", 300, minimum=60),
        timer_refresh_seconds=env_int("DESK_TIMER_REFRESH_SECONDS", 60, minimum=30),
        heartbeat_interval_seconds=env_int("DESK_HEARTBEAT_INTERVAL_SECONDS", 25, minimum=10),
        button_hold_seconds=float(os.environ.get("DESK_BUTTON_HOLD_SECONDS", "1.8")),
        email_enabled=env_bool("DESK_EMAIL_ENABLED", True),
        email_refresh_minutes=env_int("DESK_EMAIL_REFRESH_MINUTES", 10, minimum=1),
        weather_enabled=env_bool("DESK_WEATHER_ENABLED", True),
        weather_latitude=env_float("DESK_WEATHER_LATITUDE"),
        weather_longitude=env_float("DESK_WEATHER_LONGITUDE"),
        weather_label=os.environ.get("DESK_WEATHER_LABEL", ""),
        weather_refresh_minutes=env_int("DESK_WEATHER_REFRESH_MINUTES", 30, minimum=5),
        commute_weather_hour=env_int("DESK_COMMUTE_WEATHER_HOUR", env_int("DESK_OFFICE_END_HOUR", 18)),
        save_previews=env_bool("DESK_SAVE_PREVIEWS", False),
        preview_dir=Path(os.environ.get("DESK_PREVIEW_DIR", str(cache_dir / "previews"))),
        state_file=Path(os.environ.get("DESK_STATE_FILE", str(cache_dir / "state.json"))),
        token_cache_file=Path(os.environ.get("DESK_TOKEN_CACHE_FILE", str(cache_dir / "token-cache.json"))),
        heartbeat_file=Path(os.environ.get("DESK_HEARTBEAT_FILE", str(cache_dir / "heartbeat.txt"))),
        lock_file=Path(os.environ.get("DESK_LOCK_FILE", "/tmp/info-screen-desk.lock")),
        epd_driver=os.environ.get("DESK_EPD_DRIVER", "epd2in7_V2"),
        display_rotation=env_int("DESK_DISPLAY_ROTATION", 0),
        button_home_pin=env_int("DESK_BUTTON_HOME_PIN", 5),
        button_agenda_pin=env_int("DESK_BUTTON_AGENDA_PIN", 6),
        button_focus_pin=env_int("DESK_BUTTON_FOCUS_PIN", 13),
        button_refresh_pin=env_int("DESK_BUTTON_REFRESH_PIN", 19),
    )


def default_state() -> dict:
    return {
        "view": "home",
        "focus_until_iso": None,
        "focus_started_iso": None,
        "focus_duration_minutes": None,
        "focus_preset_index": 1,
        "last_sync_iso": None,
        "last_error": None,
        "schedule": None,
        "mail": None,
        "weather": None,
    }


def load_state(config: Config) -> dict:
    state = default_state()
    try:
        if config.state_file.exists():
            state.update(json.loads(config.state_file.read_text()))
    except Exception as exc:
        LOGGER.warning("Could not read state: %s", exc)
    if state.get("view") not in VIEWS:
        state["view"] = "home"
    return state


def save_state(config: Config, state: dict):
    config.state_file.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = config.state_file.with_suffix(".tmp")
    tmp_path.write_text(json.dumps(state, indent=2, sort_keys=True))
    tmp_path.replace(config.state_file)


@contextmanager
def display_lock(config: Config):
    import fcntl

    config.lock_file.parent.mkdir(parents=True, exist_ok=True)
    with open(config.lock_file, "a+", encoding="utf-8") as handle:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def get_epd(config: Config):
    if config.epd_driver.lower() == "mock":
        return MockEPD()
    try:
        module = import_module(f"waveshare_epd.{config.epd_driver}")
    except ModuleNotFoundError as exc:
        raise RuntimeError(f"Unsupported Waveshare driver: {config.epd_driver}") from exc
    return module.EPD()


def get_fonts(config: Config) -> dict:
    regular = os.environ.get("DESK_FONT_REGULAR_PATH", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
    bold = os.environ.get("DESK_FONT_BOLD_PATH", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")
    try:
        return {
            "clock": ImageFont.truetype(bold, 24),
            "status": ImageFont.truetype(bold, 22),
            "headline": ImageFont.truetype(bold, 18),
            "body": ImageFont.truetype(regular, 14),
            "body_bold": ImageFont.truetype(bold, 14),
            "tiny": ImageFont.truetype(regular, 10),
            "tiny_bold": ImageFont.truetype(bold, 10),
        }
    except Exception:
        fallback = ImageFont.load_default()
        return {key: fallback for key in ("clock", "status", "headline", "body", "body_bold", "tiny", "tiny_bold")}


def new_canvas(config: Config):
    epd = get_epd(config)
    canvas = Image.new("1", (epd.height, epd.width), 255)
    draw = ImageDraw.Draw(canvas)
    draw.fontmode = "1"
    return epd, canvas, draw, epd.height, epd.width


def load_token_cache(config: Config) -> SerializableTokenCache:
    cache = SerializableTokenCache()
    try:
        if config.token_cache_file.exists():
            cache.deserialize(config.token_cache_file.read_text())
    except Exception as exc:
        LOGGER.warning("Could not read token cache: %s", exc)
    return cache


def save_token_cache(config: Config, cache: SerializableTokenCache):
    if not cache.has_state_changed:
        return
    config.token_cache_file.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = config.token_cache_file.with_suffix(".tmp")
    tmp_path.write_text(cache.serialize())
    tmp_path.replace(config.token_cache_file)
    try:
        config.token_cache_file.chmod(0o600)
    except OSError:
        pass


def login_with_device_code(config: Config):
    if config.auth_mode != "device_code":
        raise RuntimeError("--login only works with DESK_AUTH_MODE=device_code")
    cache = load_token_cache(config)
    app = PublicClientApplication(
        config.client_id,
        authority=f"https://login.microsoftonline.com/{config.tenant_id}",
        token_cache=cache,
    )
    flow = app.initiate_device_flow(scopes=list(config.delegated_scopes))
    if "user_code" not in flow:
        raise RuntimeError(f"Could not start device login: {json.dumps(flow, indent=2)}")
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    save_token_cache(config, cache)
    if "access_token" not in result:
        raise RuntimeError(f"Device login failed: {result.get('error_description')}")
    LOGGER.info("Device login saved token cache to %s", config.token_cache_file)


def get_token(config: Config) -> str:
    if config.auth_mode == "client_credentials":
        app = ConfidentialClientApplication(
            config.client_id,
            authority=f"https://login.microsoftonline.com/{config.tenant_id}",
            client_credential=config.client_secret,
        )
        result = app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            return result["access_token"]
        raise RuntimeError(f"Token error: {result.get('error_description')}")

    cache = load_token_cache(config)
    app = PublicClientApplication(
        config.client_id,
        authority=f"https://login.microsoftonline.com/{config.tenant_id}",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    result = app.acquire_token_silent(list(config.delegated_scopes), account=accounts[0]) if accounts else None
    save_token_cache(config, cache)
    if result and "access_token" in result:
        return result["access_token"]
    raise RuntimeError("No cached Microsoft token. Run `python3 src/info_screen.py --login` first.")


def parse_local(config: Config, value: str, zone_name: str | None) -> datetime:
    dt = isoparse(value)
    if dt.tzinfo is None:
        try:
            dt = pytz.timezone(zone_name).localize(dt) if zone_name else config.timezone.localize(dt)
        except Exception:
            dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(config.timezone)


def parse_state_dt(value: str | None) -> datetime | None:
    return datetime.fromisoformat(value) if value else None


def graph_time(value: datetime) -> str:
    return value.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def hide_subject(config: Config, event: dict) -> bool:
    if config.privacy_mode in {"private", "hide", "all"}:
        return True
    return config.privacy_mode in {"smart", "sensitive"} and event.get("sensitivity") in {"private", "confidential"}


def event_to_state(config: Config, start_dt: datetime, end_dt: datetime, raw: dict) -> dict:
    show_as = raw.get("showAs", "busy") or "busy"
    return {
        "id": raw.get("id", ""),
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "subject": "Busy" if hide_subject(config, raw) else (raw.get("subject") or "Untitled"),
        "organizer": raw.get("organizer", {}).get("emailAddress", {}).get("name", "") or "",
        "location": raw.get("location", {}).get("displayName", "") or "",
        "show_as": show_as,
        "is_all_day": bool(raw.get("isAllDay")),
        "is_busy": show_as in BUSY_SHOW_AS,
    }


def graph_user_path(config: Config) -> str:
    return f"users/{config.calendar_email}" if config.auth_mode == "client_credentials" else "me"


def fetch_schedule(config: Config) -> dict:
    now = datetime.now(config.timezone)
    start_local = config.timezone.localize(datetime.combine(now.date(), dt_time.min))
    end_local = start_local + timedelta(days=2)
    path = f"{graph_user_path(config)}/calendarView"
    url = (
        f"https://graph.microsoft.com/v1.0/{path}"
        f"?startDateTime={graph_time(start_local)}&endDateTime={graph_time(end_local)}"
        f"&$top=100&$orderby=start/dateTime"
    )
    response = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {get_token(config)}",
            "Prefer": f'outlook.timezone="{config.timezone_name}"',
        },
        timeout=20,
    )
    response.raise_for_status()

    events = []
    for raw in response.json().get("value", []):
        if raw.get("isCancelled"):
            continue
        try:
            start_dt = parse_local(config, raw["start"]["dateTime"], raw["start"].get("timeZone"))
            end_dt = parse_local(config, raw["end"]["dateTime"], raw["end"].get("timeZone"))
        except Exception as exc:
            LOGGER.warning("Skipping unparseable event: %s", exc)
            continue
        events.append(event_to_state(config, start_dt, end_dt, raw))

    today = [event for event in events if parse_state_dt(event["start"]).date() == now.date()]
    tomorrow = [event for event in events if parse_state_dt(event["start"]).date() == now.date() + timedelta(days=1)]
    timed_today = [event for event in today if not event["is_all_day"] and event["is_busy"]]
    current = next((event for event in timed_today if parse_state_dt(event["start"]) <= now < parse_state_dt(event["end"])), None)
    upcoming = [event for event in timed_today if parse_state_dt(event["start"]) > now]
    next_tomorrow = next((event for event in tomorrow if not event["is_all_day"] and event["is_busy"]), None)
    return {
        "current": current,
        "upcoming": upcoming[:8],
        "today": today,
        "tomorrow": tomorrow[:8],
        "next_tomorrow": next_tomorrow,
        "events_today_count": len(timed_today),
        "all_day_count": len([event for event in today if event["is_all_day"]]),
        "fetched_at": now.isoformat(),
    }


def fetch_mail_signal(config: Config) -> dict:
    if not config.email_enabled:
        return {"enabled": False}

    base = f"https://graph.microsoft.com/v1.0/{graph_user_path(config)}/mailFolders/inbox"
    headers = {"Authorization": f"Bearer {get_token(config)}"}
    response = requests.get(base, headers=headers, timeout=15)
    response.raise_for_status()
    folder = response.json()

    important_unread = None
    try:
        important_url = (
            f"{base}/messages?$filter=isRead eq false and importance eq 'high'"
            "&$count=true&$top=1&$select=id"
        )
        important_response = requests.get(
            important_url,
            headers={**headers, "ConsistencyLevel": "eventual"},
            timeout=15,
        )
        important_response.raise_for_status()
        important_unread = int(important_response.json().get("@odata.count", 0))
    except Exception as exc:
        LOGGER.debug("Important mail count unavailable: %s", exc)

    return {
        "enabled": True,
        "unread": int(folder.get("unreadItemCount") or 0),
        "total": int(folder.get("totalItemCount") or 0),
        "important_unread": important_unread,
        "fetched_at": datetime.now(config.timezone).isoformat(),
        "error": None,
    }


def weather_summary(code: int | None) -> str:
    if code is None:
        return "weather"
    return WEATHER_CODES.get(code, "weather")


def fetch_weather_signal(config: Config) -> dict:
    if not config.weather_enabled or config.weather_latitude is None or config.weather_longitude is None:
        return {"enabled": False}

    params = {
        "latitude": config.weather_latitude,
        "longitude": config.weather_longitude,
        "current": "temperature_2m,precipitation,weather_code,wind_speed_10m",
        "hourly": "precipitation_probability,precipitation",
        "forecast_days": 2,
        "timezone": config.timezone_name,
    }
    response = requests.get("https://api.open-meteo.com/v1/forecast", params=params, timeout=15)
    response.raise_for_status()
    payload = response.json()
    current = payload.get("current", {})
    hourly = payload.get("hourly", {})

    commute_precip_probability = None
    commute_precipitation = None
    commute_label = None
    hourly_times = hourly.get("time") or []
    if hourly_times:
        now = datetime.now(config.timezone)
        commute_dt = config.timezone.localize(datetime.combine(now.date(), dt_time(config.commute_weather_hour)))
        if commute_dt < now - timedelta(hours=1):
            commute_dt += timedelta(days=1)
        candidates = []
        for index, value in enumerate(hourly_times):
            try:
                candidate_dt = parse_local(config, value, config.timezone_name)
            except Exception:
                continue
            candidates.append((abs((candidate_dt - commute_dt).total_seconds()), index, candidate_dt))
        if candidates:
            _distance, index, candidate_dt = min(candidates, key=lambda item: item[0])
            probabilities = hourly.get("precipitation_probability") or []
            precipitation = hourly.get("precipitation") or []
            if index < len(probabilities):
                commute_precip_probability = probabilities[index]
            if index < len(precipitation):
                commute_precipitation = precipitation[index]
            commute_label = candidate_dt.strftime("%H:%M")

    code = current.get("weather_code")
    return {
        "enabled": True,
        "label": config.weather_label,
        "temperature": current.get("temperature_2m"),
        "precipitation": current.get("precipitation"),
        "wind": current.get("wind_speed_10m"),
        "code": code,
        "summary": weather_summary(code),
        "commute_label": commute_label,
        "commute_precip_probability": commute_precip_probability,
        "commute_precipitation": commute_precipitation,
        "fetched_at": datetime.now(config.timezone).isoformat(),
        "error": None,
    }


def is_stale_signal(signal: dict | None, timezone_obj, max_age_minutes: int) -> bool:
    if not signal or not signal.get("fetched_at"):
        return True
    fetched_at = parse_state_dt(signal.get("fetched_at"))
    if not fetched_at:
        return True
    age = datetime.now(timezone_obj) - fetched_at.astimezone(timezone_obj)
    return age.total_seconds() > max_age_minutes * 60


def refresh_optional_signals(config: Config, state: dict):
    if config.email_enabled and is_stale_signal(state.get("mail"), config.timezone, config.email_refresh_minutes):
        try:
            state["mail"] = fetch_mail_signal(config)
        except Exception as exc:
            LOGGER.warning("Mail signal unavailable: %s", exc)
            previous = state.get("mail") if isinstance(state.get("mail"), dict) else {}
            state["mail"] = {
                **previous,
                "enabled": True,
                "error": str(exc),
                "fetched_at": datetime.now(config.timezone).isoformat(),
            }

    if config.weather_enabled and is_stale_signal(state.get("weather"), config.timezone, config.weather_refresh_minutes):
        try:
            state["weather"] = fetch_weather_signal(config)
        except Exception as exc:
            LOGGER.warning("Weather signal unavailable: %s", exc)
            previous = state.get("weather") if isinstance(state.get("weather"), dict) else {}
            state["weather"] = {
                **previous,
                "enabled": bool(config.weather_latitude and config.weather_longitude),
                "error": str(exc),
                "fetched_at": datetime.now(config.timezone).isoformat(),
            }


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font, max_width: int, max_lines: int) -> list[str]:
    if not text:
        return []
    words = text.split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    consumed = 1
    for word in words[1:]:
        trial = f"{current} {word}"
        if draw.textlength(trial, font=font) <= max_width:
            current = trial
            consumed += 1
            continue
        lines.append(current)
        current = word
        consumed += 1
        if len(lines) == max_lines - 1:
            break
    remaining = words[consumed:]
    if remaining:
        while draw.textlength(f"{current}...", font=font) > max_width and current:
            current = current[:-1]
        current = f"{current.rstrip()}..."
    lines.append(current)
    return lines[:max_lines]


def truncate(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> str:
    lines = wrap_text(draw, text, font, max_width, 1)
    return lines[0] if lines else ""


def draw_fit_text(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str, fonts: dict, max_width: int, fill: int = 0):
    for font_name in ("status", "headline", "body_bold"):
        font = fonts[font_name]
        if draw.textlength(text, font=font) <= max_width:
            draw.text(xy, text, font=font, fill=fill)
            return
    draw.text(xy, truncate(draw, text, fonts["body_bold"], max_width), font=fonts["body_bold"], fill=fill)


def format_range(event: dict) -> str:
    start_dt = parse_state_dt(event.get("start"))
    end_dt = parse_state_dt(event.get("end"))
    if not start_dt or not end_dt:
        return ""
    return f"{start_dt:%H:%M}-{end_dt:%H:%M}"


def combine_local(config: Config, day: date, hour: int) -> datetime:
    return config.timezone.localize(datetime.combine(day, dt_time(hour=hour)))


def resolve_schedule(schedule: dict, now: datetime) -> tuple[dict | None, list[dict]]:
    current = schedule.get("current")
    if current:
        start_dt = parse_state_dt(current.get("start"))
        end_dt = parse_state_dt(current.get("end"))
        if not start_dt or not end_dt or not start_dt <= now < end_dt:
            current = None
    if not current:
        for event in schedule.get("today", []):
            if event.get("is_all_day") or not event.get("is_busy"):
                continue
            start_dt = parse_state_dt(event.get("start"))
            end_dt = parse_state_dt(event.get("end"))
            if start_dt and end_dt and start_dt <= now < end_dt:
                current = event
                break
    current_end = parse_state_dt(current["end"]) if current else None
    upcoming = []
    for event in schedule.get("today", []):
        if event.get("is_all_day") or not event.get("is_busy"):
            continue
        start_dt = parse_state_dt(event.get("start"))
        if not start_dt or start_dt <= now:
            continue
        if current_end and start_dt < current_end:
            continue
        upcoming.append(event)
    return current, upcoming


def focus_until(state: dict, now: datetime) -> datetime | None:
    value = parse_state_dt(state.get("focus_until_iso"))
    if value and value > now:
        return value
    if state.get("focus_until_iso"):
        state["focus_until_iso"] = None
        state["focus_started_iso"] = None
        state["focus_duration_minutes"] = None
    return None


def sync_label(state: dict, config: Config) -> str:
    last_sync = parse_state_dt(state.get("last_sync_iso"))
    if not last_sync:
        return "Sync never"
    return f"Sync {last_sync.astimezone(config.timezone):%H:%M}"


def content_width(width: int) -> int:
    return width - CONTENT_LEFT - CONTENT_RIGHT_MARGIN


def draw_top(draw: ImageDraw.ImageDraw, width: int, fonts: dict, now: datetime, title: str, subtitle: str):
    draw.rectangle((MENU_WIDTH, 0, width, 34), fill=0)
    draw.text((CONTENT_LEFT, 4), f"{now:%H:%M}", font=fonts["clock"], fill=255)
    meta_x = CONTENT_LEFT + 78
    draw.text((meta_x, 5), truncate(draw, title.upper(), fonts["tiny_bold"], width - meta_x - CONTENT_RIGHT_MARGIN), font=fonts["tiny_bold"], fill=255)
    draw.text((meta_x, 19), truncate(draw, subtitle, fonts["tiny"], width - meta_x - CONTENT_RIGHT_MARGIN), font=fonts["tiny"], fill=255)


def draw_left_menu(draw: ImageDraw.ImageDraw, _width: int, height: int, fonts: dict, active: str):
    labels = [("home", "Hm"), ("agenda", "Ag"), ("focus", "Fc"), ("refresh", "Rf")]
    segment = height // 4
    draw.rectangle((0, 0, MENU_WIDTH, height), outline=0, fill=255)
    for index, (key, label) in enumerate(labels):
        y0 = index * segment
        y1 = height if index == 3 else (index + 1) * segment
        fill = 0
        if key == active:
            draw.rectangle((0, y0, MENU_WIDTH, y1), fill=0)
            fill = 255
        text_width = draw.textlength(label, font=fonts["tiny_bold"])
        text_y = y0 + max(2, (y1 - y0 - 10) // 2)
        draw.text((max(2, (MENU_WIDTH - text_width) // 2), text_y), label, font=fonts["tiny_bold"], fill=fill)
        if index:
            draw.line((4, y0, MENU_WIDTH - 4, y0), fill=0)
    draw.line((MENU_WIDTH, 0, MENU_WIDTH, height), fill=0)


def busy_minutes(events: list[dict], start: datetime, end: datetime) -> int:
    total = 0
    for event in events:
        if event.get("is_all_day") or not event.get("is_busy"):
            continue
        start_dt = parse_state_dt(event.get("start"))
        end_dt = parse_state_dt(event.get("end"))
        if start_dt and end_dt:
            total += max(0, int((min(end, end_dt) - max(start, start_dt)).total_seconds() // 60))
    return total


def meeting_starts_soon(config: Config, now: datetime, upcoming: list[dict]) -> dict | None:
    if not upcoming:
        return None
    start_dt = parse_state_dt(upcoming[0].get("start"))
    if not start_dt:
        return None
    seconds = (start_dt - now).total_seconds()
    if 0 <= seconds <= config.meeting_alert_minutes * 60:
        return upcoming[0]
    return None


def tomorrow_busy_events(schedule: dict) -> list[dict]:
    return [event for event in schedule.get("tomorrow", []) if event.get("is_busy") and not event.get("is_all_day")]


def signal_line(state: dict) -> str:
    parts: list[str] = []
    mail = state.get("mail") or {}
    if mail.get("enabled") and not mail.get("error"):
        unread = int(mail.get("unread") or 0)
        important = mail.get("important_unread")
        if important:
            parts.append(f"Mail {unread}/{important}")
        else:
            parts.append(f"Mail {unread}")
    elif mail.get("enabled") and mail.get("error"):
        parts.append("Mail unavailable")

    weather = state.get("weather") or {}
    if weather.get("enabled") and not weather.get("error"):
        temp = weather.get("temperature")
        summary = weather.get("summary") or "weather"
        if temp is not None:
            parts.append(f"{round(float(temp))}C {summary}")
        else:
            parts.append(summary.title())
        commute_probability = weather.get("commute_precip_probability")
        commute_label = weather.get("commute_label")
        if commute_probability is not None and commute_label:
            parts.append(f"{commute_probability}% rain {commute_label}")
    elif weather.get("enabled") and weather.get("error"):
        parts.append("Weather unavailable")

    return " | ".join(parts)


def render_home(config: Config, state: dict) -> Image.Image:
    _, canvas, draw, width, height = new_canvas(config)
    fonts = get_fonts(config)
    now = datetime.now(config.timezone)
    schedule = state.get("schedule") or {}
    current, upcoming = resolve_schedule(schedule, now)
    active_focus = focus_until(state, now)
    work_end = combine_local(config, now.date(), config.office_end_hour)

    soon = meeting_starts_soon(config, now, upcoming)

    if active_focus:
        status = "FOCUS"
        detail = f"Until {active_focus:%H:%M}"
    elif current:
        end_dt = parse_state_dt(current.get("end"))
        status = "IN MEETING"
        detail = f"Until {end_dt:%H:%M}" if end_dt else "Busy now"
    elif soon:
        start_dt = parse_state_dt(soon.get("start"))
        minutes = max(0, int((start_dt - now).total_seconds() // 60)) if start_dt else 0
        status = "MEETING SOON"
        detail = f"{soon.get('subject', 'Meeting')} in {minutes} min"
    elif config.office_start_hour <= now.hour < config.office_end_hour:
        next_start = parse_state_dt(upcoming[0]["start"]) if upcoming else work_end
        free_until = min(next_start, work_end)
        minutes = max(0, int((free_until - now).total_seconds() // 60))
        status = "AVAILABLE"
        detail = f"{minutes} min free" if minutes < 180 else f"Free until {free_until:%H:%M}"
    else:
        tomorrow = tomorrow_busy_events(schedule)
        first_tomorrow = tomorrow[0] if tomorrow else schedule.get("next_tomorrow")
        first_start = parse_state_dt(first_tomorrow.get("start")) if first_tomorrow else None
        status = "TOMORROW" if first_tomorrow else "OFF HOURS"
        detail = f"First {first_start:%H:%M} / {len(tomorrow)} meetings" if first_start else f"Work starts {config.office_start_hour:02d}:00"

    draw_top(draw, width, fonts, now, config.user_name, f"{now:%a %d %b}")
    draw_fit_text(draw, (CONTENT_LEFT, 44), status, fonts, content_width(width))
    draw.text((CONTENT_LEFT, 69), truncate(draw, detail, fonts["body_bold"], content_width(width)), font=fonts["body_bold"], fill=0)

    y = 91
    if current:
        draw.text((CONTENT_LEFT, y), "Now", font=fonts["tiny_bold"], fill=0)
        subject_x = CONTENT_LEFT + 36
        for i, line in enumerate(wrap_text(draw, current.get("subject", "Busy"), fonts["body"], width - subject_x - CONTENT_RIGHT_MARGIN, 2)):
            draw.text((subject_x, y - 3 + i * 15), line, font=fonts["body"], fill=0)
    elif upcoming:
        next_event = upcoming[0]
        start_dt = parse_state_dt(next_event.get("start"))
        draw.text((CONTENT_LEFT, y), f"Next {start_dt:%H:%M}" if start_dt else "Next", font=fonts["tiny_bold"], fill=0)
        subject_x = CONTENT_LEFT + 78
        draw.text((subject_x, y - 3), truncate(draw, next_event.get("subject", "Scheduled"), fonts["body"], width - subject_x - CONTENT_RIGHT_MARGIN), font=fonts["body"], fill=0)
    elif schedule.get("next_tomorrow"):
        next_event = schedule["next_tomorrow"]
        start_dt = parse_state_dt(next_event.get("start"))
        draw.text((CONTENT_LEFT, y), f"Tom {start_dt:%H:%M}" if start_dt else "Tomorrow", font=fonts["tiny_bold"], fill=0)
        subject_x = CONTENT_LEFT + 70
        draw.text((subject_x, y - 3), truncate(draw, next_event.get("subject", "Scheduled"), fonts["body"], width - subject_x - CONTENT_RIGHT_MARGIN), font=fonts["body"], fill=0)
    else:
        draw.text((CONTENT_LEFT, y), "No more meetings", font=fonts["body_bold"], fill=0)

    office_start = combine_local(config, now.date(), config.office_start_hour)
    office_end = combine_local(config, now.date(), config.office_end_hour)
    open_minutes = max(0, int((office_end - office_start).total_seconds() // 60) - busy_minutes(schedule.get("today", []), office_start, office_end))
    draw.line((CONTENT_LEFT, 122, width - CONTENT_RIGHT_MARGIN, 122), fill=0)
    summary = f"Today {schedule.get('events_today_count', 0)} meetings"
    if schedule.get("all_day_count"):
        summary = f"{summary} + {schedule['all_day_count']} all-day"
    draw.text((CONTENT_LEFT, 130), truncate(draw, summary, fonts["tiny_bold"], content_width(width)), font=fonts["tiny_bold"], fill=0)
    draw.text((CONTENT_LEFT, 144), truncate(draw, f"{open_minutes // 60}h {open_minutes % 60:02d}m open in work day", fonts["tiny"], content_width(width)), font=fonts["tiny"], fill=0)
    signals = signal_line(state)
    if signals:
        draw.text((CONTENT_LEFT, 158), truncate(draw, signals, fonts["tiny"], content_width(width)), font=fonts["tiny"], fill=0)
    draw_left_menu(draw, width, height, fonts, "home")
    return canvas


def render_agenda(config: Config, state: dict) -> Image.Image:
    _, canvas, draw, width, height = new_canvas(config)
    fonts = get_fonts(config)
    now = datetime.now(config.timezone)
    schedule = state.get("schedule") or {}
    current, upcoming = resolve_schedule(schedule, now)
    draw_top(draw, width, fonts, now, "Agenda", sync_label(state, config))
    rows = [("NOW", current)] if current else []
    rows.extend((format_range(event), event) for event in upcoming)
    if not rows:
        draw.text((CONTENT_LEFT, 56), "Clear for today", font=fonts["headline"], fill=0)
    else:
        y = 40
        for label, event in rows[:4]:
            label_x = CONTENT_LEFT
            subject_x = CONTENT_LEFT + 78
            draw.text((label_x, y), truncate(draw, label, fonts["tiny_bold"], 72), font=fonts["tiny_bold"], fill=0)
            draw.text((subject_x, y - 1), truncate(draw, event.get("subject", "Untitled"), fonts["body_bold"], width - subject_x - CONTENT_RIGHT_MARGIN), font=fonts["body_bold"], fill=0)
            meta = " / ".join(part for part in (event.get("location"), event.get("organizer")) if part and config.privacy_mode == "normal")
            if meta:
                draw.text((subject_x, y + 13), truncate(draw, meta, fonts["tiny"], width - subject_x - CONTENT_RIGHT_MARGIN), font=fonts["tiny"], fill=0)
            y += 28
    draw_left_menu(draw, width, height, fonts, "agenda")
    return canvas


def render_tomorrow(config: Config, state: dict) -> Image.Image:
    _, canvas, draw, width, height = new_canvas(config)
    fonts = get_fonts(config)
    now = datetime.now(config.timezone)
    schedule = state.get("schedule") or {}
    tomorrow = tomorrow_busy_events(schedule)
    draw_top(draw, width, fonts, now, "Tomorrow", sync_label(state, config))

    if not tomorrow:
        draw.text((CONTENT_LEFT, 54), "No meetings", font=fonts["headline"], fill=0)
        draw.text((CONTENT_LEFT, 80), "Calendar is open", font=fonts["body_bold"], fill=0)
    else:
        first = tomorrow[0]
        first_start = parse_state_dt(first.get("start"))
        total_minutes = sum(
            max(0, int((parse_state_dt(event["end"]) - parse_state_dt(event["start"])).total_seconds() // 60))
            for event in tomorrow
            if parse_state_dt(event.get("start")) and parse_state_dt(event.get("end"))
        )
        header = f"First {first_start:%H:%M} / {len(tomorrow)} meetings" if first_start else f"{len(tomorrow)} meetings"
        draw.text((CONTENT_LEFT, 45), truncate(draw, header, fonts["body_bold"], content_width(width)), font=fonts["body_bold"], fill=0)
        draw.text((CONTENT_LEFT, 61), f"{total_minutes // 60}h {total_minutes % 60:02d}m booked", font=fonts["tiny"], fill=0)
        y = 82
        for event in tomorrow[:3]:
            draw.text((CONTENT_LEFT, y), truncate(draw, format_range(event), fonts["tiny_bold"], 72), font=fonts["tiny_bold"], fill=0)
            subject_x = CONTENT_LEFT + 78
            draw.text((subject_x, y - 1), truncate(draw, event.get("subject", "Meeting"), fonts["body_bold"], width - subject_x - CONTENT_RIGHT_MARGIN), font=fonts["body_bold"], fill=0)
            y += 24

    signals = signal_line(state)
    if signals:
        draw.line((CONTENT_LEFT, 146, width - CONTENT_RIGHT_MARGIN, 146), fill=0)
        draw.text((CONTENT_LEFT, 154), truncate(draw, signals, fonts["tiny"], content_width(width)), font=fonts["tiny"], fill=0)
    draw_left_menu(draw, width, height, fonts, "agenda")
    return canvas


def render_focus(config: Config, state: dict) -> Image.Image:
    _, canvas, draw, width, height = new_canvas(config)
    fonts = get_fonts(config)
    now = datetime.now(config.timezone)
    schedule = state.get("schedule") or {}
    current, upcoming = resolve_schedule(schedule, now)
    active_focus = focus_until(state, now)
    draw_top(draw, width, fonts, now, "Focus", sync_label(state, config))
    if active_focus:
        remaining = max(0, int((active_focus - now).total_seconds() // 60))
        draw_fit_text(draw, (CONTENT_LEFT, 48), "FOCUS ACTIVE", fonts, content_width(width))
        draw.text((CONTENT_LEFT, 75), truncate(draw, f"Until {active_focus:%H:%M} / {remaining} min", fonts["body_bold"], content_width(width)), font=fonts["body_bold"], fill=0)
        started = parse_state_dt(state.get("focus_started_iso")) or now
        total = max(1, int((active_focus - started).total_seconds()))
        elapsed = max(0, int((now - started).total_seconds()))
        bar_left = CONTENT_LEFT
        bar_right = width - CONTENT_RIGHT_MARGIN
        draw.rectangle((bar_left, 100, bar_right, 110), outline=0, fill=255)
        draw.rectangle((bar_left + 1, 101, bar_left + 1 + int((bar_right - bar_left - 2) * min(1, elapsed / total)), 109), fill=0)
    else:
        preset_index = int(state.get("focus_preset_index", 1)) % len(FOCUS_PRESETS)
        preset = FOCUS_PRESETS[preset_index]
        draw_fit_text(draw, (CONTENT_LEFT, 48), "READY TO FOCUS", fonts, content_width(width))
        draw.text((CONTENT_LEFT, 76), f"{preset} minute block", font=fonts["headline"], fill=0)
    next_event = current or (upcoming[0] if upcoming else None)
    separator_y = 108 if not active_focus else 124
    draw.line((CONTENT_LEFT, separator_y, width - CONTENT_RIGHT_MARGIN, separator_y), fill=0)
    if next_event:
        start_dt = parse_state_dt(next_event.get("start"))
        y = 119 if not active_focus else 134
        draw.text((CONTENT_LEFT, y), "Now" if current else f"Next {start_dt:%H:%M}", font=fonts["tiny_bold"], fill=0)
        subject_x = CONTENT_LEFT + 78
        draw.text((subject_x, y - 3), truncate(draw, next_event.get("subject", "Scheduled"), fonts["body"], width - subject_x - CONTENT_RIGHT_MARGIN), font=fonts["body"], fill=0)
    draw_left_menu(draw, width, height, fonts, "focus")
    return canvas


def get_ip_address() -> str:
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("8.8.8.8", 80))
        return sock.getsockname()[0]
    except OSError:
        return "Unavailable"
    finally:
        sock.close()


def run_command(command: list[str]) -> str:
    try:
        result = subprocess.run(command, capture_output=True, text=True, timeout=3, check=False)
    except Exception:
        return "Unavailable"
    return (result.stdout or "").strip() or "Unavailable"


def render_diagnostics(config: Config, state: dict) -> Image.Image:
    _, canvas, draw, width, height = new_canvas(config)
    fonts = get_fonts(config)
    now = datetime.now(config.timezone)
    draw_top(draw, width, fonts, now, "Diagnostics", sync_label(state, config))
    lines = [
        f"Wi-Fi: {run_command(['iwgetid', '-r'])}",
        f"IP: {get_ip_address()}",
        f"Driver: {config.epd_driver}",
        f"Auth: {config.auth_mode}",
        f"View: {state.get('view', 'home')}",
    ]
    for index, line in enumerate(lines):
        draw.text((CONTENT_LEFT, 50 + index * 15), truncate(draw, line, fonts["tiny"], content_width(width)), font=fonts["tiny"], fill=0)
    if state.get("last_error"):
        draw.line((CONTENT_LEFT, 132, width - CONTENT_RIGHT_MARGIN, 132), fill=0)
        draw.text((CONTENT_LEFT, 140), truncate(draw, state["last_error"], fonts["tiny"], content_width(width)), font=fonts["tiny"], fill=0)
    draw_left_menu(draw, width, height, fonts, "refresh")
    return canvas


def render_error(config: Config, state: dict, message: str) -> Image.Image:
    _, canvas, draw, width, height = new_canvas(config)
    fonts = get_fonts(config)
    now = datetime.now(config.timezone)
    draw_top(draw, width, fonts, now, "Error", sync_label(state, config))
    draw.text((CONTENT_LEFT, 52), truncate(draw, "Calendar unavailable", fonts["headline"], content_width(width)), font=fonts["headline"], fill=0)
    for index, line in enumerate(wrap_text(draw, message, fonts["body"], content_width(width), 4)):
        draw.text((CONTENT_LEFT, 84 + index * 16), line, font=fonts["body"], fill=0)
    draw_left_menu(draw, width, height, fonts, "refresh")
    return canvas


def render_view(config: Config, state: dict) -> Image.Image:
    if state.get("view") == "tomorrow":
        return render_tomorrow(config, state)
    if state.get("view") == "agenda":
        return render_agenda(config, state)
    if state.get("view") == "focus":
        return render_focus(config, state)
    return render_home(config, state)


def push_display(config: Config, canvas: Image.Image):
    if config.save_previews:
        config.preview_dir.mkdir(parents=True, exist_ok=True)
        canvas.save(config.preview_dir / "current-render.png")
    epd = get_epd(config)
    epd.init()
    image = canvas.rotate(config.display_rotation) if config.display_rotation else canvas
    epd.display(epd.getbuffer(image))
    epd.sleep()


def systemd_notify(message: str):
    notify_socket = os.environ.get("NOTIFY_SOCKET")
    if not notify_socket:
        return
    address = "\0" + notify_socket[1:] if notify_socket.startswith("@") else notify_socket
    try:
        with socket.socket(socket.AF_UNIX, socket.SOCK_DGRAM) as sock:
            sock.connect(address)
            sock.sendall(message.encode("utf-8"))
    except Exception:
        LOGGER.debug("systemd notify failed", exc_info=True)


def emit_heartbeat(config: Config, status: str):
    config.heartbeat_file.parent.mkdir(parents=True, exist_ok=True)
    config.heartbeat_file.write_text(f"{datetime.now(config.timezone).isoformat()} {status}\n")
    systemd_notify(f"WATCHDOG=1\nSTATUS={status}")


def clear_display(config: Config):
    epd = get_epd(config)
    epd.init()
    if hasattr(epd, "Clear"):
        epd.Clear()
    epd.sleep()


def refresh_display(config: Config):
    with display_lock(config):
        state = load_state(config)
        try:
            state["schedule"] = fetch_schedule(config)
            state["last_sync_iso"] = datetime.now(config.timezone).isoformat()
            state["last_error"] = None
            refresh_optional_signals(config, state)
            canvas = render_view(config, state)
            status = "refresh ok"
        except Exception as exc:
            LOGGER.exception("Refresh failed")
            state["last_error"] = str(exc)
            canvas = render_error(config, state, str(exc)) if not state.get("schedule") else render_view(config, state)
            status = "refresh degraded"
        save_state(config, state)
        push_display(config, canvas)
    emit_heartbeat(config, status)


def local_redraw(config: Config, status: str = "local redraw"):
    with display_lock(config):
        state = load_state(config)
        save_state(config, state)
        push_display(config, render_view(config, state))
    emit_heartbeat(config, status)


def set_view(config: Config, view: str):
    with display_lock(config):
        state = load_state(config)
        state["view"] = view
        save_state(config, state)
        push_display(config, render_view(config, state))


def toggle_focus(config: Config):
    with display_lock(config):
        state = load_state(config)
        now = datetime.now(config.timezone)
        if focus_until(state, now):
            state["focus_until_iso"] = None
            state["focus_started_iso"] = None
            state["focus_duration_minutes"] = None
        else:
            index = int(state.get("focus_preset_index", 1)) % len(FOCUS_PRESETS)
            minutes = FOCUS_PRESETS[index]
            state["focus_started_iso"] = now.isoformat()
            state["focus_duration_minutes"] = minutes
            state["focus_until_iso"] = (now + timedelta(minutes=minutes)).isoformat()
            state["focus_preset_index"] = (index + 1) % len(FOCUS_PRESETS)
        state["view"] = "focus"
        save_state(config, state)
        push_display(config, render_focus(config, state))


def cycle_focus_preset(config: Config):
    with display_lock(config):
        state = load_state(config)
        index = int(state.get("focus_preset_index", 1)) % len(FOCUS_PRESETS)
        state["focus_preset_index"] = (index + 1) % len(FOCUS_PRESETS)
        state["view"] = "focus"
        save_state(config, state)
        push_display(config, render_focus(config, state))


def show_diagnostics(config: Config):
    with display_lock(config):
        state = load_state(config)
        push_display(config, render_diagnostics(config, state))


def should_redraw_for_clock(config: Config, state: dict, now: datetime) -> bool:
    if state.get("focus_until_iso"):
        return True
    schedule = state.get("schedule") or {}
    current, upcoming = resolve_schedule(schedule, now)
    if current:
        return True
    if meeting_starts_soon(config, now, upcoming):
        return True
    for event in schedule.get("today", []):
        start_dt = parse_state_dt(event.get("start"))
        end_dt = parse_state_dt(event.get("end"))
        if start_dt and abs((start_dt - now).total_seconds()) <= config.timer_refresh_seconds:
            return True
        if end_dt and abs((end_dt - now).total_seconds()) <= config.timer_refresh_seconds:
            return True
    return False


def run_listener(config: Config):
    try:
        from gpiozero import Button
    except Exception as exc:
        raise RuntimeError("gpiozero is required on the Pi for button support") from exc

    queue: Queue[str] = Queue()
    held = {"home": False, "agenda": False, "focus": False, "refresh": False}

    def make_button(pin: int, short_action: str, long_action: str):
        button = Button(pin, pull_up=True, bounce_time=0.1, hold_time=config.button_hold_seconds)

        def on_held():
            held[short_action] = True

        def on_released():
            queue.put(long_action if held[short_action] else short_action)
            held[short_action] = False

        button.when_held = on_held
        button.when_released = on_released
        return button

    home_button = make_button(config.button_home_pin, "home", "diagnostics")
    agenda_button = make_button(config.button_agenda_pin, "agenda", "tomorrow")
    focus_button = make_button(config.button_focus_pin, "focus", "focus_preset")
    refresh_button = make_button(config.button_refresh_pin, "refresh", "clear_refresh")
    systemd_notify("READY=1\nSTATUS=Button listener started")
    refresh_display(config)
    next_refresh = time.monotonic() + config.refresh_interval_seconds
    next_timer = time.monotonic() + config.timer_refresh_seconds
    next_heartbeat = time.monotonic() + config.heartbeat_interval_seconds

    while True:
        timeout = max(0.0, min(next_refresh, next_timer, next_heartbeat) - time.monotonic())
        try:
            action = queue.get(timeout=timeout)
        except Empty:
            now = time.monotonic()
            if now >= next_heartbeat:
                emit_heartbeat(config, "listener idle")
                next_heartbeat = now + config.heartbeat_interval_seconds
            if now >= next_refresh:
                refresh_display(config)
                now = time.monotonic()
                next_refresh = now + config.refresh_interval_seconds
                next_timer = now + config.timer_refresh_seconds
            elif now >= next_timer:
                with display_lock(config):
                    state = load_state(config)
                    if should_redraw_for_clock(config, state, datetime.now(config.timezone)):
                        save_state(config, state)
                        push_display(config, render_view(config, state))
                next_timer = time.monotonic() + config.timer_refresh_seconds
            continue

        if action == "home":
            set_view(config, "home")
        elif action == "agenda":
            set_view(config, "agenda")
        elif action == "tomorrow":
            set_view(config, "tomorrow")
        elif action == "focus":
            toggle_focus(config)
        elif action == "focus_preset":
            cycle_focus_preset(config)
        elif action == "diagnostics":
            show_diagnostics(config)
        elif action == "refresh":
            refresh_display(config)
            now = time.monotonic()
            next_refresh = now + config.refresh_interval_seconds
            next_timer = now + config.timer_refresh_seconds
        elif action == "clear_refresh":
            with display_lock(config):
                clear_display(config)
            refresh_display(config)
            now = time.monotonic()
            next_refresh = now + config.refresh_interval_seconds
            next_timer = now + config.timer_refresh_seconds


def make_sample_schedule(config: Config) -> dict:
    now = datetime.now(config.timezone)
    today = now.date()

    def event(hour: int, minute: int, duration: int, subject: str, location: str = "") -> dict:
        start_dt = config.timezone.localize(datetime.combine(today, dt_time(hour, minute)))
        end_dt = start_dt + timedelta(minutes=duration)
        return {
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
            "subject": subject,
            "organizer": "Visuology",
            "location": location,
            "is_all_day": False,
            "is_busy": True,
        }

    events = [
        event(9, 15, 30, "Team standup", "Teams"),
        event(11, 0, 45, "Client review", "Room B"),
        event(14, 30, 60, "Design decisions"),
        event(16, 0, 30, "Wrap-up"),
    ]
    current = next((item for item in events if parse_state_dt(item["start"]) <= now < parse_state_dt(item["end"])), None)
    upcoming = [item for item in events if parse_state_dt(item["start"]) > now]
    tomorrow = event(9, 0, 30, "Planning")
    tomorrow["start"] = (parse_state_dt(tomorrow["start"]) + timedelta(days=1)).isoformat()
    tomorrow["end"] = (parse_state_dt(tomorrow["end"]) + timedelta(days=1)).isoformat()
    tomorrow_2 = event(13, 30, 45, "Roadmap review")
    tomorrow_2["start"] = (parse_state_dt(tomorrow_2["start"]) + timedelta(days=1)).isoformat()
    tomorrow_2["end"] = (parse_state_dt(tomorrow_2["end"]) + timedelta(days=1)).isoformat()
    return {
        "current": current,
        "upcoming": upcoming,
        "today": events,
        "tomorrow": [tomorrow, tomorrow_2],
        "next_tomorrow": tomorrow,
        "events_today_count": len(events),
        "all_day_count": 0,
        "fetched_at": now.isoformat(),
    }


def render_sample(config: Config, output: Path, view: str):
    state = default_state()
    state["view"] = view
    state["last_sync_iso"] = datetime.now(config.timezone).isoformat()
    state["schedule"] = make_sample_schedule(config)
    state["mail"] = {
        "enabled": True,
        "unread": 4,
        "total": 120,
        "important_unread": 1,
        "fetched_at": datetime.now(config.timezone).isoformat(),
        "error": None,
    }
    state["weather"] = {
        "enabled": True,
        "temperature": 15.7,
        "summary": "showers",
        "commute_label": "18:00",
        "commute_precip_probability": 40,
        "fetched_at": datetime.now(config.timezone).isoformat(),
        "error": None,
    }
    canvas = render_view(config, state)
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output)
    blank = Image.new("1", canvas.size, 255)
    if ImageChops.difference(canvas, blank).getbbox() is None:
        raise RuntimeError(f"Blank render: {output}")
    LOGGER.info("Saved %s", output)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--listen", action="store_true", help="Run the GPIO button listener")
    parser.add_argument("--login", action="store_true", help="Run one-time Microsoft device-code login")
    parser.add_argument("--sample", action="store_true", help="Render with sample data")
    parser.add_argument("--preview", type=Path, help="Render to PNG instead of e-ink")
    parser.add_argument("--view", choices=sorted(VIEWS), default="home")
    return parser.parse_args()


def main():
    args = parse_args()
    if args.preview:
        os.environ["DESK_EPD_DRIVER"] = "mock"
    config = load_config(require_credentials=not args.sample)
    if args.login:
        login_with_device_code(config)
    elif args.sample:
        if not args.preview:
            raise RuntimeError("--sample requires --preview")
        render_sample(config, args.preview, args.view)
    elif args.listen:
        run_listener(config)
    else:
        refresh_display(config)


if __name__ == "__main__":
    main()
