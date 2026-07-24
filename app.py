
import io
import json
import hashlib
import os
import random
import re
import secrets
import sqlite3
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urlencode
from streamlit_ace import st_ace
import docx
import fitz
import requests
import speech_recognition as sr
import streamlit as st

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_TAGS_URL = "http://localhost:11434/api/tags"
DEFAULT_MODEL = "mistral:latest"
USER_DATABASE_PATH = Path(__file__).with_name("users.db")
LEGACY_USERS_JSON_PATH = Path(__file__).with_name("users.json")
LEVEL_ONE_THRESHOLD = 6.0
LEVEL_TWO_SLOT_DAY_COUNT = 5
LEVEL_TWO_SLOT_HOURS = (10, 14, 18)
LEVEL_TWO_SLOT_DURATION_MINUTES = 45
ALLOWED_INTERVIEW_SKILLS = ("DSA", "SQL")
st.set_page_config(page_title="AI Interview Simulator", layout="wide")
st.markdown(
    """
<style>
:root {
    color-scheme: light;
    --bg-a: #f4fbff;
    --bg-b: #eef6ff;
    --card: rgba(255, 255, 255, 0.82);
    --border: #d8e6ff;
    --text: #13213f;
    --accent-a: #0ea5e9;
    --accent-b: #22c55e;
}
.stApp {
    background:
      radial-gradient(900px 420px at 92% -8%, #d8f4ff 0%, transparent 62%),
      radial-gradient(850px 460px at -8% 20%, #e5ffef 0%, transparent 58%),
      linear-gradient(180deg, var(--bg-a), var(--bg-b));
}
html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
    background: transparent !important;
    color: var(--text) !important;
}
[data-testid="stMainBlockContainer"] {
    background: transparent !important;
}
[data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] {
    background: transparent !important;
}
.stApp, .stApp p, .stApp span, .stApp label, .stApp li, .stApp h1, .stApp h2, .stApp h3, .stApp h4 {
    color: var(--text);
}
.main .block-container {
    max-width: 1120px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ffffff, #f4f8ff);
    border-right: 1px solid var(--border);
}
[data-testid="stHeader"] {
    background: transparent !important;
}
[data-testid="stSidebar"] * {
    color: var(--text) !important;
}
[data-testid="stSidebar"] .stButton > button {
    color: #ffffff !important;
}
[data-testid="stSidebar"] [data-baseweb="input"] input {
    background: #ffffff !important;
    color: var(--text) !important;
}
section[data-testid="stFileUploaderDropzone"] {
    background: #ecf4ff !important;
    border: 1px dashed #9ab8ff !important;
}
section[data-testid="stFileUploaderDropzone"] * {
    color: var(--text) !important;
}
section[data-testid="stFileUploaderDropzone"] button {
    background: #ffffff !important;
    color: var(--text) !important;
    border: 1px solid #c8d9ff !important;
    border-radius: 10px !important;
}
.hero {
    background: linear-gradient(120deg, #0ea5e9, #22c55e);
    color: #ffffff;
    border-radius: 20px;
    padding: 1.1rem 1.2rem;
    box-shadow: 0 10px 32px rgba(14, 165, 233, 0.22);
    margin-bottom: 0.9rem;
}
.hero h1 {
    margin: 0;
    font-size: 1.6rem;
    font-weight: 800;
    letter-spacing: 0.2px;
}
.hero p {
    margin: 0.25rem 0 0;
    font-size: 0.93rem;
    opacity: 0.95;
}
.hero h1, .hero p {
    color: #ffffff !important;
}
.card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 18px;
    padding: 0.85rem 1rem;
    margin-bottom: 0.75rem;
    box-shadow: 0 6px 18px rgba(17, 39, 74, 0.06);
    color: var(--text);
}
.stButton > button {
    border-radius: 12px;
    border: 0;
    font-weight: 700;
    letter-spacing: 0.15px;
    background: linear-gradient(120deg, var(--accent-a), var(--accent-b));
    color: white;
}
.stButton > button:hover {
    filter: brightness(1.02);
}
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] span,
[data-testid="stMetricLabel"],
[data-testid="stMetricValue"],
[data-testid="stAlert"] *,
[data-testid="stInfo"] *,
[data-testid="stSuccess"] *,
[data-testid="stWarning"] *,
[data-testid="stError"] * {
    color: var(--text) !important;
}
[data-testid="stInfo"], [data-testid="stSuccess"], [data-testid="stWarning"], [data-testid="stError"] {
    background: #f7fbff !important;
    border: 1px solid #dce9ff !important;
}
[data-baseweb="input"] input,
[data-baseweb="textarea"] textarea,
[data-baseweb="select"] div,
.stTextInput input,
.stTextArea textarea {
    color: var(--text) !important;
    background: #ffffff !important;
}
[data-baseweb="select"] > div,
[data-baseweb="select"] [role="listbox"],
[data-baseweb="select"] [role="option"] {
    color: var(--text) !important;
    background: #ffffff !important;
}
[data-testid="stAudioInput"] button,
[data-testid="stAudioInput"] label,
[data-testid="stAudioInput"] span {
    color: var(--text) !important;
}
[data-testid="stJson"],
[data-testid="stCodeBlock"],
pre,
code {
    background: #f7fbff !important;
    color: var(--text) !important;
    border: 1px solid #dce9ff !important;
    border-radius: 10px !important;
}
[data-testid="stJson"] *,
[data-testid="stCodeBlock"] *,
pre *,
code * {
    color: var(--text) !important;
}
[data-testid="stAudioInput"],
[data-testid="stAudioInput"] > div {
    background: #eef4ff !important;
    border: 1px solid #d6e5ff !important;
    border-radius: 12px !important;
}
[data-testid="stAudioInput"] audio {
    background: #eef4ff !important;
    border-radius: 10px !important;
}
[data-testid="stAudioInput"] svg,
[data-testid="stAudioInput"] path,
[data-testid="stAudioInput"] time {
    fill: var(--text) !important;
    color: var(--text) !important;
}
</style>
""",
    unsafe_allow_html=True,
)
st.markdown(
    """
<div class="hero">
  <h1>AI Interview Studio</h1>
  
</div>
""",
    unsafe_allow_html=True,
)
def call_ollama(
    prompt,
    model=DEFAULT_MODEL,
    status_text="Generating response...",
    timeout=100,
    options=None,
):
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": True,
        "keep_alive": "30m",
        "options": options or {},
    }
    chunks = []
    status_box = st.empty()
    try:
        with st.spinner(status_text):
            with requests.post(
                OLLAMA_URL,
                json=payload,
                stream=True,
                timeout=(10, timeout),
            ) as response:
                response.raise_for_status()
                for line in response.iter_lines(decode_unicode=True):
                    if not line:
                        continue
                    try:
                        packet = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    piece = packet.get("response", "")
                    if piece:
                        chunks.append(piece)
                        if len(chunks) % 14 == 0:
                            status_box.caption(f"{status_text} ({len(''.join(chunks))} chars)")

                    if packet.get("done"):
                        break
        status_box.empty()
    except requests.RequestException as exc:
        status_box.empty()
        st.error("Ollama connection failed. Start Ollama and ensure the model is available.")
        st.write(exc)
        return None
    except Exception as exc:
        status_box.empty()
        st.error("Unexpected error while calling Ollama.")
        st.write(exc)
        return None

    text = "".join(chunks).strip()
    if not text:
        st.error("Ollama returned an empty response.")
        return None
    return text
def extract_json(text):
    if not text:
        return None
    cleaned = re.sub(r"^```(?:json)?\\s*|\\s*```$", "", text.strip(), flags=re.MULTILINE).strip()
    match = re.search(r"\{[\s\S]*\}", cleaned)
    if not match:
        return None
    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError:
        return None
@st.cache_data(ttl=30)
def get_ollama_models():
    try:
        response = requests.get(OLLAMA_TAGS_URL, timeout=10)
        response.raise_for_status()
        data = response.json()
        models = []
        for model in data.get("models", []):
            name = model.get("name")
            if isinstance(name, str) and name.strip():
                models.append(name.strip())
        return models, None
    except requests.RequestException as exc:
        return [], str(exc)
def normalize_skills(raw_skills):
    if isinstance(raw_skills, str):
        candidates = [raw_skills]
    elif isinstance(raw_skills, list):
        candidates = raw_skills
    else:
        return []
    normalized = []
    seen = set()
    for item in candidates:
        if not isinstance(item, str):
            continue
        skill = item.strip()
        if not skill:
            continue
        skill_key = skill.lower()
        if skill_key in seen:
            continue
        seen.add(skill_key)
        normalized.append(skill)
    return normalized
def parse_score(raw_score):
    try:
        score = float(raw_score)
    except (TypeError, ValueError):
        return 0.0
    return max(0.0, min(10.0, score))
def normalize_email(email):
    # Remove all whitespace so pasted emails like "user @gmail.com" still validate.
    return re.sub(r"\s+", "", (email or "")).lower()
def is_valid_email(email):
    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email or "") is not None
def parse_iso_datetime(value):
    if not value or not isinstance(value, str):
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None
def format_datetime_display(value):
    if not isinstance(value, datetime):
        return "Not scheduled"
    return value.strftime("%d %b %Y %I:%M %p")
def build_level_two_slot_options(now=None):
    base_now = now or datetime.now()
    options = []
    for day_offset in range(1, LEVEL_TWO_SLOT_DAY_COUNT + 1):
        slot_date = (base_now + timedelta(days=day_offset)).date()
        for hour in LEVEL_TWO_SLOT_HOURS:
            start_dt = datetime(slot_date.year, slot_date.month, slot_date.day, hour, 0, 0)
            if start_dt <= base_now:
                continue
            end_dt = start_dt + timedelta(minutes=LEVEL_TWO_SLOT_DURATION_MINUTES)
            slot_key = start_dt.strftime("%Y%m%dT%H%M")
            label = f"{format_datetime_display(start_dt)} to {format_datetime_display(end_dt)}"
            options.append(
                {
                    "slot_key": slot_key,
                    "start_dt": start_dt,
                    "end_dt": end_dt,
                    "label": label,
                }
            )
    return options
def get_level_two_slot_by_key(slot_key):
    clean_key = (slot_key or "").strip()
    if not clean_key:
        return None
    for slot in build_level_two_slot_options():
        if slot.get("slot_key") == clean_key:
            return slot
    return None
def build_level_two_booking_link(app_link, slot_key):
    clean_link = (app_link or "").strip()
    if not re.match(r"^https?://", clean_link, flags=re.IGNORECASE):
        return ""
    separator = "&" if "?" in clean_link else "?"
    return f"{clean_link}{separator}{urlencode({'l2_slot': slot_key})}"
def read_query_param_value(name):
    key = (name or "").strip()
    if not key:
        return ""
    try:
        value = st.query_params.get(key)
        if isinstance(value, list):
            return str(value[0]).strip() if value else ""
        return str(value).strip() if value is not None else ""
    except Exception:
        try:
            legacy_params = st.experimental_get_query_params()
            values = legacy_params.get(key, [])
            if isinstance(values, list):
                return str(values[0]).strip() if values else ""
            return str(values).strip() if values is not None else ""
        except Exception:
            return ""
def clear_query_param_value(name):
    key = (name or "").strip()
    if not key:
        return
    try:
        if key in st.query_params:
            del st.query_params[key]
    except Exception:
        pass
    try:
        legacy_params = st.experimental_get_query_params()
        if key in legacy_params:
            legacy_params.pop(key, None)
            st.experimental_set_query_params(**legacy_params)
    except Exception:
        return


def apply_level_two_slot_from_query():
    slot_key = read_query_param_value("l2_slot")
    if not slot_key:
        return

    clear_query_param_value("l2_slot")

    current_level = int(st.session_state.get("current_level", 1))
    if current_level < 2:
        st.session_state.level2_splash = (
            "Level 2 slot link detected, but you are not qualified yet. "
            "Complete Level 1 first."
        )
        return

    slot = get_level_two_slot_by_key(slot_key)
    if not slot:
        st.session_state.level2_splash = (
            "Selected slot from email is unavailable. "
            "Please choose another slot from the app."
        )
        return

    current_email = (st.session_state.get("auth_user") or {}).get("email")
    if not current_email:
        st.session_state.level2_splash = "Please login first to confirm your Level 2 slot."
        return

    saved = save_level_two_slot(current_email, slot["start_dt"], slot["end_dt"])
    if not saved:
        st.session_state.level2_splash = "Could not save the email-selected slot. Please retry in the app."
        return

    st.session_state.level2_slot_start = slot["start_dt"]
    st.session_state.level2_slot_end = slot["end_dt"]
    slot_state = get_level_two_window_state(slot["start_dt"], slot["end_dt"])
    if slot_state == "open":
        st.session_state.level2_splash = "Your Level 2 slot is confirmed and Level 2 is open now. You can start."
    else:
        st.session_state.level2_splash = (
            f"Your Level 2 slot is confirmed for {format_datetime_display(slot['start_dt'])}."
        )


def hash_password(password, salt):
    payload = f"{salt}:{password or ''}".encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def get_user_db_connection():
    conn = sqlite3.connect(USER_DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_user_database():
    try:
        with get_user_db_connection() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    email TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    level INTEGER NOT NULL DEFAULT 1,
                    level2_slot_start TEXT,
                    level2_slot_end TEXT,
                    level2_mail_sent_at TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )
            conn.commit()
    except sqlite3.Error:
        return
    migrate_legacy_user_json()


def migrate_legacy_user_json():
    if not LEGACY_USERS_JSON_PATH.exists():
        return

    try:
        with LEGACY_USERS_JSON_PATH.open("r", encoding="utf-8") as handle:
            legacy_data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return

    if not isinstance(legacy_data, dict):
        return

    now_iso = datetime.now().isoformat()
    try:
        with get_user_db_connection() as conn:
            for email_key, payload in legacy_data.items():
                if not isinstance(payload, dict):
                    continue

                clean_email = normalize_email(payload.get("email", email_key))
                if not clean_email:
                    continue

                name = str(payload.get("name", clean_email.split("@")[0] or "Candidate")).strip() or "Candidate"
                salt = str(payload.get("salt", "")).strip()
                password_hash = str(payload.get("password_hash", "")).strip()
                if not salt or not password_hash:
                    continue

                try:
                    level = int(payload.get("level", 1))
                except (TypeError, ValueError):
                    level = 1
                level = max(1, min(2, level))

                conn.execute(
                    """
                    INSERT OR IGNORE INTO users (
                        email, name, salt, password_hash, level,
                        level2_slot_start, level2_slot_end, level2_mail_sent_at,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        clean_email,
                        name,
                        salt,
                        password_hash,
                        level,
                        payload.get("level2_slot_start"),
                        payload.get("level2_slot_end"),
                        payload.get("level2_mail_sent_at"),
                        now_iso,
                        now_iso,
                    ),
                )
            conn.commit()
    except sqlite3.Error:
        return

    try:
        migrated_path = LEGACY_USERS_JSON_PATH.with_suffix(".json.migrated")
        LEGACY_USERS_JSON_PATH.replace(migrated_path)
    except OSError:
        pass


def get_user_by_email(email):
    clean_email = normalize_email(email)
    if not clean_email:
        return None
    try:
        with get_user_db_connection() as conn:
            row = conn.execute("SELECT * FROM users WHERE email = ?", (clean_email,)).fetchone()
            return row
    except sqlite3.Error:
        return None


def update_user_fields(email, **fields):
    clean_email = normalize_email(email)
    if not clean_email:
        return False

    allowed_fields = {
        "name",
        "salt",
        "password_hash",
        "level",
        "level2_slot_start",
        "level2_slot_end",
        "level2_mail_sent_at",
    }
    filtered = {key: value for key, value in fields.items() if key in allowed_fields}
    if not filtered:
        return False

    assignments = ", ".join(f"{key} = ?" for key in filtered.keys())
    params = list(filtered.values())
    params.append(datetime.now().isoformat())
    params.append(clean_email)

    try:
        with get_user_db_connection() as conn:
            cursor = conn.execute(
                f"UPDATE users SET {assignments}, updated_at = ? WHERE email = ?",
                params,
            )
            conn.commit()
            return cursor.rowcount > 0
    except sqlite3.Error:
        return False


def send_level_two_qualification_email(candidate_name, recipient_email):
    smtp_host = os.getenv("AI_INTERVIEW_SMTP_HOST", "").strip()
    smtp_user = os.getenv("AI_INTERVIEW_SMTP_USER", "").strip()
    smtp_password = os.getenv("AI_INTERVIEW_SMTP_PASSWORD", "")
    smtp_from = os.getenv("AI_INTERVIEW_SMTP_FROM", smtp_user).strip()
    smtp_port = int(os.getenv("AI_INTERVIEW_SMTP_PORT", "587"))
    app_link = os.getenv("AI_INTERVIEW_APP_LINK", "").strip()
    fallback_app_text = "Please open the app and login again."

    if not smtp_host or not smtp_user or not smtp_password or not smtp_from:
        return False, "SMTP is not configured. Set AI_INTERVIEW_SMTP_* environment variables."

    slot_lines = []
    for idx, slot in enumerate(build_level_two_slot_options(), start=1):
        line = f"{idx}) {slot['label']}"
        booking_link = build_level_two_booking_link(app_link, slot["slot_key"])
        if booking_link:
            line += f"\n   Book this slot: {booking_link}"
        slot_lines.append(line)
    slot_section = "\n".join(slot_lines) if slot_lines else "No slots are currently available."

    msg = EmailMessage()
    msg["Subject"] = "Round 1 Cleared - Select Your Level 2 Interview Slot"
    msg["From"] = smtp_from
    msg["To"] = recipient_email
    msg.set_content(
        f"""Hi {candidate_name},

Congratulations. You are selected in Round 1 of the AI Interview process.
You are now qualified for Session 2 (Level 2).

Next steps:
1) Login to the app using your registered email.
2) Select one of the available slots below (or click a direct booking link).
3) Come back at your scheduled time. Level 2 will open automatically.

Available slots:
{slot_section}

App link:
{app_link or fallback_app_text}
"""
    )

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
        return True, "Qualification email sent."
    except Exception as exc:
        return False, f"Email failed: {exc}"


def register_user(name, email, password):
    clean_name = (name or "").strip()
    clean_email = normalize_email(email)
    raw_password = password or ""

    if not clean_name:
        return False, "Name is required."
    if not is_valid_email(clean_email):
        return False, "Enter a valid email address."
    if len(raw_password) < 6:
        return False, "Password must be at least 6 characters."

    salt = secrets.token_hex(16)
    now_iso = datetime.now().isoformat()

    try:
        with get_user_db_connection() as conn:
            existing = conn.execute(
                "SELECT email FROM users WHERE email = ?",
                (clean_email,),
            ).fetchone()
            if existing is not None:
                return False, "An account already exists for this email."

            conn.execute(
                """
                INSERT INTO users (
                    email, name, salt, password_hash, level,
                    level2_slot_start, level2_slot_end, level2_mail_sent_at,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    clean_email,
                    clean_name,
                    salt,
                    hash_password(raw_password, salt),
                    1,
                    None,
                    None,
                    None,
                    now_iso,
                    now_iso,
                ),
            )
            conn.commit()
    except sqlite3.Error:
        return False, "Failed to save user account."

    return True, "Registration successful. Please log in."


def authenticate_user(email, password):
    clean_email = normalize_email(email)
    user = get_user_by_email(clean_email)
    if user is None:
        return None

    salt = str(user["salt"] or "").strip()
    expected_hash = str(user["password_hash"] or "").strip()
    if not salt or not expected_hash:
        return None

    candidate_hash = hash_password(password or "", salt)
    if candidate_hash != expected_hash:
        return None

    try:
        level = int(user["level"] if user["level"] is not None else 1)
    except (TypeError, ValueError):
        level = 1
    level = max(1, min(2, level))
    slot_start = parse_iso_datetime(user["level2_slot_start"])
    slot_end = parse_iso_datetime(user["level2_slot_end"])
    return {
        "name": str(user["name"] or clean_email.split("@")[0]).strip() or "Candidate",
        "email": clean_email,
        "level": level,
        "level2_slot_start": slot_start,
        "level2_slot_end": slot_end,
    }


def update_user_level(email, level):
    try:
        target_level = int(level)
    except Exception:
        target_level = 1
    target_level = max(1, min(2, target_level))

    updates = {"level": target_level}
    if target_level == 1:
        updates["level2_slot_start"] = None
        updates["level2_slot_end"] = None
    update_user_fields(email, **updates)


def clear_interview_state(clear_skills=False):
    keys = [
        "question",
        "question_mode",
        "code_language",
        "asked_language",
        "answer_text",
        "code_answer",
        "reference_answer",
        "reference_type",
        "mistake_feedback",
        "audio_key",
        "code_language_locked",
    ]
    if clear_skills:
        keys.extend(
            [
                "skills",
                "scores",
                "current_skill",
                "resume_signature",
                "resume_skills",
                "resume_text",
                "resume_ready",
                "auto_generate_after_resume",
                "question_topic",
                "question_style",
            ]
        )
    for key in keys:
        st.session_state.pop(key, None)


def init_app_state():
    defaults = {
        "auth_user": None,
        "current_level": 1,
        "level2_slot_start": None,
        "level2_slot_end": None,
        "auth_mode": "Login",
        "pending_auth_mode": "",
        "login_notice": "",
        "completion_notice": "",
        "level2_splash": "",
        "skills": [],
        "scores": [],
        "question_topic": "",
        "question_style": "Theory",
        "resume_signature": "",
        "resume_skills": [],
        "resume_text": "",
        "resume_ready": False,
        "auto_generate_after_resume": False,
    }
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value


def get_question_difficulty(level):
    return "easy to medium" if int(level) <= 1 else "medium to high"


def build_google_calendar_link(title, details, start_dt, end_dt):
    params = {
        "action": "TEMPLATE",
        "text": title,
        "details": details,
        "dates": f"{start_dt.strftime('%Y%m%dT%H%M%S')}/{end_dt.strftime('%Y%m%dT%H%M%S')}",
    }
    return f"https://calendar.google.com/calendar/render?{urlencode(params)}"


def save_level_two_slot(email, start_dt, end_dt):
    if not email or not isinstance(start_dt, datetime) or not isinstance(end_dt, datetime):
        return False
    if end_dt <= start_dt:
        return False
    return update_user_fields(
        email,
        level2_slot_start=start_dt.isoformat(),
        level2_slot_end=end_dt.isoformat(),
    )


def get_level_two_window_state(slot_start, slot_end):
    if not isinstance(slot_start, datetime) or not isinstance(slot_end, datetime):
        return "not_scheduled"
    now = datetime.now()
    if now < slot_start:
        return "upcoming"
    if slot_start <= now <= slot_end:
        return "open"
    return "expired"


def reset_candidate_to_level_one(notice=None):
    user = st.session_state.get("auth_user") or {}
    user_email = user.get("email")
    if user_email:
        update_user_level(user_email, 1)

    st.session_state.current_level = 1
    st.session_state.level2_slot_start = None
    st.session_state.level2_slot_end = None
    st.session_state.completion_notice = (
        notice
        or "Level 2 completed. You are redirected to a fresh Level 1 round."
    )
    clear_interview_state(clear_skills=True)
    st.rerun()


def apply_level_progression(score):
    current_level = int(st.session_state.get("current_level", 1))
    if current_level != 1:
        return

    if score >= LEVEL_ONE_THRESHOLD:
        st.session_state.current_level = 2
        st.session_state.level2_slot_start = None
        st.session_state.level2_slot_end = None
        user = st.session_state.get("auth_user") or {}
        user_name = user.get("name", "Candidate")
        user_email = user.get("email")
        if user_email:
            update_user_fields(
                user_email,
                level=2,
                level2_slot_start=None,
                level2_slot_end=None,
            )
            email_ok, email_msg = send_level_two_qualification_email(user_name, user_email)
            if email_ok:
                update_user_fields(user_email, level2_mail_sent_at=datetime.now().isoformat())
                st.session_state.level2_splash = (
                    f"Success! You scored {score}/10, cleared Round 1, and are selected for Session 2 (Level 2). "
                    "An email has been sent to your registered email. Choose one of the given slots to continue."
                )
            else:
                st.session_state.level2_splash = (
                    f"Success! You scored {score}/10, cleared Round 1, and are selected for Session 2 (Level 2). "
                    f"Email was not sent: {email_msg}. "
                    "You can still select one of the given slots in the app."
                )
        else:
            st.session_state.level2_splash = (
                f"Success! You scored {score}/10 and are selected for Session 2 (Level 2). "
                "Please select one of the given slots in the app."
            )

        clear_interview_state(clear_skills=False)
        st.rerun()
    else:
        st.info(f"Score {score}/10. Get at least {LEVEL_ONE_THRESHOLD}/10 to move to Level 2.")


def render_auth_gate():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Candidate Access")

    pending_mode = str(st.session_state.get("pending_auth_mode", "")).strip()
    if pending_mode in {"Login", "Register"}:
        st.session_state.auth_mode = pending_mode
    st.session_state.pending_auth_mode = ""

    auth_mode = st.radio(
        "Access Mode",
        options=["Login", "Register"],
        horizontal=True,
        key="auth_mode",
    )

    if st.session_state.get("login_notice"):
        st.info(st.session_state.login_notice)
        st.session_state.login_notice = ""

    if auth_mode == "Login":
        login_email = st.text_input("Email", key="login_email")
        login_password = st.text_input("Password", type="password", key="login_password")
        if st.button("Login", key="login_button"):
            user = authenticate_user(login_email, login_password)
            if user is None:
                st.error("Invalid email or password.")
            else:
                clear_interview_state(clear_skills=True)
                st.session_state.auth_user = {
                    "name": user["name"],
                    "email": user["email"],
                }
                st.session_state.current_level = user["level"]
                st.session_state.level2_slot_start = user.get("level2_slot_start")
                st.session_state.level2_slot_end = user.get("level2_slot_end")
                st.success("Login successful.")
                st.rerun()
    else:
        register_name = st.text_input("Name", key="register_name")
        register_email = st.text_input("Email", key="register_email")
        register_password = st.text_input("Password", type="password", key="register_password")
        confirm_password = st.text_input("Confirm Password", type="password", key="register_confirm_password")
        if st.button("Register", key="register_button"):
            if register_password != confirm_password:
                st.error("Password and confirm password do not match.")
            else:
                ok, message = register_user(register_name, register_email, register_password)
                if ok:
                    st.session_state.pending_auth_mode = "Login"
                    st.session_state.login_email = normalize_email(register_email)
                    st.session_state.login_password = ""
                    st.success("Registration successful. Please login.")
                    st.rerun()
                else:
                    st.error(message)

    st.markdown("</div>", unsafe_allow_html=True)


def _normalize_feedback_lines(value):
    if value is None:
        return []

    if isinstance(value, dict):
        return [f"{k}: {v}" for k, v in value.items()]

    if isinstance(value, (list, tuple, set)):
        return [str(item).strip() for item in value if str(item).strip()]

    text = str(value).strip()
    if not text:
        return []

    if text.startswith("[") or text.startswith("{"):
        try:
            parsed = json.loads(text)
            return _normalize_feedback_lines(parsed)
        except Exception:
            pass

    parts = [line.strip(" -•\t") for line in text.splitlines() if line.strip()]
    if len(parts) > 1:
        return parts
    return [text]


def render_feedback_block(title, value):
    st.markdown(f"### {title}")
    lines = _normalize_feedback_lines(value)
    if not lines:
        st.write("Not provided.")
        return
    for line in lines:
        st.markdown(f"- {line}")


def fallback_skill_extract(text):
    if not text:
        return []

    # Try to prioritize explicit skills sections from resume text.
    section_match = re.search(
        r"(?is)(technical skills|skills|tech stack|technologies|tools)\s*[:\-]?\s*(.+?)(?:\n\s*\n|experience|projects|education|certifications|$)",
        text,
    )
    source_text = section_match.group(2) if section_match else text[:3500]

    raw_tokens = re.split(r"[,;\n|/•●\-]+", source_text)
    stop_words = {
        "and",
        "with",
        "using",
        "good",
        "strong",
        "knowledge",
        "familiar",
        "experience",
        "project",
        "projects",
        "developer",
        "development",
        "tools",
        "skills",
        "technical",
    }

    results = []
    seen = set()
    for token in raw_tokens:
        cleaned = token.strip().strip(".:()[]{}")
        cleaned = re.sub(r"\s+", " ", cleaned)
        if not cleaned:
            continue
        if len(cleaned) < 2 or len(cleaned) > 40:
            continue
        if cleaned.lower() in stop_words:
            continue
        if cleaned.isdigit():
            continue
        if len(cleaned.split()) > 4:
            continue
        if not re.search(r"[A-Za-z]", cleaned):
            continue
        key = cleaned.lower()
        if key in seen:
            continue
        seen.add(key)
        results.append(cleaned)

    return results[:30]


def fallback_evaluation(answer):
    words = answer.split()
    word_count = len(words)
    keyword_hits = 0
    for token in ["because", "example", "design", "tradeoff", "performance", "scalable"]:
        if token in answer.lower():
            keyword_hits += 1

    score = min(10.0, round((word_count / 20) + (keyword_hits * 0.8), 1))
    if word_count < 30:
        strengths = "Direct and concise response."
        weaknesses = "Needs deeper technical explanation and clearer structure."
    else:
        strengths = "Good explanation depth with structured detail."
        weaknesses = "Could include more quantified impact and edge cases."

    return {
        "total_score": score,
        "strengths": strengths,
        "weaknesses": weaknesses,
        "improvement_suggestion": "Add one concrete production example and explain one tradeoff.",
    }


def fallback_reference_answer(skill, question):
    skill_name = skill if skill else "this topic"
    return (
        f"A strong answer for {skill_name} should start with a clear definition and your approach. "
        f"Then explain the core technical decisions, mention one real project example, and include measurable impact. "
        f"For this question, state why you chose a specific design, discuss performance and reliability tradeoffs, "
        f"and finish with what you learned or improved in a later iteration."
    )


def is_coding_question(question):
    if not question:
        return False
    q = question.lower().strip()
    coding_patterns = [
        r"\bwrite\s+(a\s+)?(program|code|function|query)\b",
        r"\bimplement\b",
        r"\bsolve\s+(this|the)\s+(problem|question)\b",
        r"\bgiven\s+(an?|the)\s+(array|string|list|matrix|tree|graph|linked list)\b",
        r"\btime complexity\b",
        r"\bspace complexity\b",
        r"\bpseudocode\b",
        r"\bdebug\s+(this|the)\s+code\b",
        r"\bsql\s+query\b",
        r"\bwrite\s+(an?\s+)?sql\b",
        r"\bquery\s+to\b",
    ]
    oral_patterns = [
        r"\bwhat is\b",
        r"\bexplain\b",
        r"\bdescribe\b",
        r"\bdifference between\b",
        r"\badvantages?\b",
        r"\bdisadvantages?\b",
        r"\bwhy\b",
        r"\bhow does\b",
        r"\bwhen would you\b",
    ]

    score = 0
    for pattern in coding_patterns:
        if re.search(pattern, q):
            score += 2
    for pattern in oral_patterns:
        if re.search(pattern, q):
            score -= 1
    return score >= 2


def is_sql_question(question, skill=""):
    q = (question or "").lower()
    s = (skill or "").lower()
    sql_tokens = [
        "sql",
        "query",
        "table",
        "join",
        "group by",
        "having",
        "where",
        "select",
    ]
    return s == "sql" or any(token in q for token in sql_tokens)


def detect_programming_language(question, skill="", use_skill_fallback=True):
    q = (question or "").lower()
    s = (skill or "").lower()

    if is_sql_question(question, skill):
        return "sql"

    question_patterns = [
        ("python", [r"\bpython\b"]),
        ("java", [r"\bjava\b"]),
        ("cpp", [r"\bc\+\+\b", r"\bcpp\b"]),
        ("javascript", [r"\bjavascript\b", r"\bnode\.?js\b", r"\bjs\b"]),
        ("c", [r"\bc language\b", r"\bin c\b", r"\bc program\b", r"\bc code\b"]),
    ]

    for language, checks in question_patterns:
        for pattern in checks:
            if re.search(pattern, q):
                return language

    if not use_skill_fallback:
        return None

    skill_hints = [
        ("sql", ["sql", "database", "postgres", "mysql"]),
        ("python", ["python", "django", "flask", "fastapi"]),
        ("java", ["java", "spring", "spring boot"]),
        ("cpp", ["c++", "cpp"]),
        ("javascript", ["javascript", "node", "react", "angular", "vue"]),
        ("c", ["c language", "embedded c"]),
    ]
    for language, hints in skill_hints:
        if any(hint in s for hint in hints):
            return language

    return None


def is_oral_question(question):
    if not question:
        return False
    q = question.lower().strip()
    oral_patterns = [
        r"\bwhat is\b",
        r"\bexplain\b",
        r"\bdescribe\b",
        r"\bdifference between\b",
        r"\badvantages?\b",
        r"\bdisadvantages?\b",
        r"\bwhy\b",
        r"\bhow does\b",
        r"\bwhen would you\b",
        r"\buse case\b",
        r"\barchitecture\b",
        r"\bconcept\b",
    ]
    coding_hard_patterns = [
        r"\bwrite\s+(a\s+)?(program|code|function|query)\b",
        r"\bimplement\b",
        r"\bsolve\s+(this|the)\s+(problem|question)\b",
        r"\bsql\s+query\b",
    ]
    oral_hits = sum(1 for p in oral_patterns if re.search(p, q))
    coding_hits = sum(1 for p in coding_hard_patterns if re.search(p, q))
    return oral_hits > 0 and coding_hits == 0


def classify_question_mode(question, skill, model, timeout_seconds, fast_mode):
    prompt = f"""
You are a technical interview question classifier.

Decide whether this question requires writing code/query as the answer.

Return ONLY valid JSON:
{{
  "mode": "coding" or "voice",
  "language": "python" | "java" | "cpp" | "javascript" | "c" | "sql" | "none"
}}

Rules:
- Use "coding" when candidate must write code or SQL query.
- Use "voice" for conceptual/theory/explanation-only questions.
- Infer language only if clearly implied; else use "none".
- Example coding: "Write a function...", "Given an array...", "Write SQL query..."
- Example voice: "Explain polymorphism", "Difference between TCP and UDP"

Skill:
{skill}

Question:
{question}
"""
    result = call_ollama(
        prompt,
        model=model,
        status_text="Classifying question type...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.0,
            "num_predict": 90 if fast_mode else 140,
            "num_ctx": 1024,
        },
    )
    parsed = extract_json(result) if result else None
    if not isinstance(parsed, dict):
        return None, None

    mode = str(parsed.get("mode", "")).strip().lower()
    language = str(parsed.get("language", "")).strip().lower()
    if mode not in {"coding", "voice"}:
        mode = None
    if language not in {"python", "java", "cpp", "javascript", "c", "sql", "none"}:
        language = None
    return mode, language


def resolve_question_mode_and_language(question, skill, model, timeout_seconds, fast_mode):
    sql_mode = is_sql_question(question, skill)
    if sql_mode:
        return "coding", "sql", "sql"

    question_lang = detect_programming_language(question, "", use_skill_fallback=False)
    heuristic_lang = detect_programming_language(question, skill, use_skill_fallback=True)
    strong_coding_signal = is_coding_question(question)
    coding_signal = strong_coding_signal or heuristic_lang is not None
    oral_signal = is_oral_question(question)
    heuristic_coding = coding_signal and not oral_signal

    llm_mode, llm_language = classify_question_mode(
        question=question,
        skill=skill,
        model=model,
        timeout_seconds=timeout_seconds,
        fast_mode=fast_mode,
    )

    if llm_mode in {"coding", "voice"}:
        mode = llm_mode
    else:
        mode = "coding" if heuristic_coding else "voice"

    if mode == "coding" and oral_signal and not strong_coding_signal and llm_mode != "coding":
        mode = "voice"

    language = None
    if question_lang:
        language = question_lang
    elif llm_language and llm_language != "none":
        language = llm_language
    elif heuristic_lang:
        language = heuristic_lang

    # Lock language only when it is explicitly mentioned in the question.
    locked_language = question_lang

    if mode == "coding" and not language:
        language = "python"

    return mode, language, locked_language


def extract_code_from_text(text):
    if not text:
        return ""
    match = re.search(r"```(?:[a-zA-Z0-9_+-]+)?\s*([\s\S]*?)```", text)
    if match:
        return match.group(1).strip()
    return text.strip()


def fallback_code_evaluation(code):
    lines = [line for line in code.splitlines() if line.strip()]
    line_count = len(lines)
    has_function = any("def " in line or "function " in line for line in lines)
    has_branching = any(("if " in line) or ("for " in line) or ("while " in line) for line in lines)
    base_score = 3.5
    if line_count >= 6:
        base_score += 1.5
    if has_function:
        base_score += 2.0
    if has_branching:
        base_score += 1.5
    score = min(10.0, round(base_score, 1))
    return {
        "total_score": score,
        "strengths": "Code has a valid structure and attempts a concrete solution.",
        "weaknesses": "Logic correctness is uncertain without full test-case validation.",
        "improvement_suggestion": "Add edge-case handling and explain complexity in comments.",
    }


def evaluate_code_answer(question, skill, code, model, timeout_seconds, fast_mode):
    prompt = f"""
You are a strict technical interviewer.
Evaluate this coding answer for correctness and code quality.

Return ONLY JSON:
{{
  "total_score": 0,
  "strengths": "",
  "weaknesses": "",
  "improvement_suggestion": ""
}}

Skill:
{skill}

Question:
{question}

Candidate Code:
{code}
"""
    result = call_ollama(
        prompt,
        model=model,
        status_text="Evaluating code answer...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.1,
            "num_predict": 200 if fast_mode else 320,
            "num_ctx": 3072,
        },
    )
    parsed_eval = extract_json(result) if result else None
    if parsed_eval:
        return parsed_eval
    return fallback_code_evaluation(code)


def generate_reference_code(question, skill, model, timeout_seconds, fast_mode, language):
    prompt = f"""
You are a senior software engineer.

Solve the coding problem below.

Rules:
- Return ONLY executable code.
- Do NOT explain anything.
- No comments except minimal.
- Use correct syntax.
- Language: {language}
Problem:
{question}

Return only the code.
"""
    result = call_ollama(
        prompt,
        model=model,
        status_text="Generating correct code...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.1,
            "num_predict": 220 if fast_mode else 360,
            "num_ctx": 3072,
        },
    )
    if result:
        return extract_code_from_text(result)
    return "def solve():\n    # Add your implementation here\n    pass"


def generate_code_mistake_feedback(question, user_code, correct_code, model, timeout_seconds, fast_mode):
    prompt = f"""
You are a coding interview coach.
Compare candidate code with the correct code.

Return plain text only with:
1) 3 short bullet points of mistakes or missing parts.
2) 1 short next-step improvement.

Question:
{question}

Candidate Code:
{user_code}

Correct Code:
{correct_code}
"""
    result = call_ollama(
        prompt,
        model=model,
        status_text="Finding code gaps...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.1,
            "num_predict": 120 if fast_mode else 220,
            "num_ctx": 3072,
        },
    )
    if result:
        return result
    return (
        "- Some edge cases are not handled.\n"
        "- Complexity may be higher than needed.\n"
        "- Output/return format is not fully aligned with problem constraints.\n"
        "Next step: validate with 3 edge test cases and optimize loops/data structures."
    )


def generate_mistake_feedback(question, user_answer, correct_answer, model, timeout_seconds, fast_mode):
    prompt = f"""
You are a technical interview coach.
Compare the candidate answer with the correct answer.

Return plain text only with:
1) 3 short bullet points of key points the candidate missed.
2) 1 short improvement step for next attempt.

Question:
{question}

Candidate Answer:
{user_answer}

Correct Answer:
{correct_answer}
"""
    result = call_ollama(
        prompt,
        model=model,
        status_text="Finding learning gaps...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.1,
            "num_predict": 120 if fast_mode else 220,
            "num_ctx": 2048,
        },
    )
    if result:
        return result
    return (
        "- Focus more on approach and architecture choices.\n"
        "- Add one practical example with measurable impact.\n"
        "- Mention tradeoffs and edge cases clearly.\n"
        "Next step: answer using structure: approach -> implementation -> tradeoff -> result."
    )


def generate_reference_answer(question, skill, model, timeout_seconds, fast_mode):
    prompt = f"""
You are a senior technical interviewer.

Write one appropriate sample answer for this interview question.

Rules:
- Return plain text only.
- Keep it realistic and specific.
- Structure: approach, technical detail, example, tradeoff, result.
- Length: {"80 to 120 words" if fast_mode else "120 to 180 words"}.
- Reply in the same natural language as the question.

Skill:
{skill}

Question:
{question}
"""
    result = call_ollama(
        prompt,
        model=model,
        status_text="Generating appropriate sample answer...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.2,
            "num_predict": 120 if fast_mode else 220,
            "num_ctx": 1536,
        },
    )
    if result:
        return result
    return fallback_reference_answer(skill, question)


def transcribe_audio(audio_file):
    recognizer = sr.Recognizer()
    try:
        audio_bytes = io.BytesIO(audio_file.getvalue())
        with sr.AudioFile(audio_bytes) as source:
            audio_data = recognizer.record(source)
        text = recognizer.recognize_google(audio_data)
        return text, None
    except sr.UnknownValueError:
        return None, "Could not understand the audio. Please speak clearly and try again."
    except sr.RequestError as exc:
        return None, f"Speech-to-text service error: {exc}"
    except Exception as exc:
        return None, f"Audio processing failed: {exc}"


def normalize_interview_skill(skill):
    return "SQL" if str(skill or "").strip().upper() == "SQL" else "DSA"


def build_question_prompt(skill, question_style, difficulty_band):
    style = str(question_style or "").strip().lower()
    selected_skill = str(skill or "").strip() or "technical topic"

    if style == "theory":
        return f"""
You are a technical interview generator.

Generate ONE theory interview question based on the candidate skill.

Rules:
- Ask conceptual/theory only.
- Do NOT ask the candidate to write code or SQL query.
- Keep the question concise and practical.
- Question must be answerable by speaking.
- Keep it relevant to the selected skill.

Skill:
{selected_skill}

Difficulty:
{difficulty_band}
"""

    track = normalize_interview_skill(skill)
    if track == "SQL":
        return f"""
You are a technical interview generator.

Generate ONE SQL coding interview problem.

Rules:
- The question MUST require writing an SQL query as the answer.
- Include table schema and enough sample data context.
- Include expected output format.
- Keep it concise and realistic.
- Do NOT ask explanation-only questions.

Skill:
SQL

Difficulty:
{difficulty_band}
"""

    return f"""
You are a technical interview generator.

Generate ONE DSA coding interview problem.

Rules:
- The question MUST require writing code as the answer.
- Focus on arrays/strings/linked lists/trees/graphs/dp/greedy or similar DSA topics.
- Include input description and expected output.
- Keep it concise and clear.
- Do NOT ask explanation-only questions.

Skill:
DSA

Difficulty:
{difficulty_band}
"""


def fallback_question(skill, question_style):
    style = str(question_style or "").strip().lower()
    selected_skill = str(skill or "").strip() or "technical topic"

    if style == "theory":
        return (
            f"Explain a core concept in {selected_skill} with one practical example, "
            "key tradeoffs, and common mistakes."
        )

    track = normalize_interview_skill(skill)
    if track == "SQL":
        return (
            "Given `employees(emp_id, name, dept_id, salary)` and "
            "`departments(dept_id, dept_name)`, write an SQL query to return each department "
            "with the 2 highest paid employees in that department."
        )

    return (
        "Given an array of integers, return the length of the longest subarray with sum equal to K. "
        "Write an efficient DSA solution."
    )


def extract_resume_skills_from_text(resume_text, model_name, timeout_seconds, fast_mode):
    clean_resume = str(resume_text or "").strip()
    if not clean_resume:
        return [], False

    resume_slice = clean_resume[:1500] if fast_mode else clean_resume[:3500]
    prompt = f"""
You are an expert resume analyzer.
Extract ONLY technical skills from this resume.
Return ONLY valid JSON with this format:
{{
  "skills": ["skill1", "skill2", "skill3"]
}}

Resume:
{resume_slice}
"""
    result = call_ollama(
        prompt,
        model=model_name,
        status_text="Analyzing resume...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.1,
            "num_predict": 120 if fast_mode else 220,
            "num_ctx": 1024 if fast_mode else 2048,
        },
    )

    parsed = extract_json(result) if result else None
    skills = normalize_skills(parsed.get("skills", [])) if parsed else []
    used_fallback = False
    if not skills:
        skills = normalize_skills(fallback_skill_extract(clean_resume))
        used_fallback = bool(skills)
    return skills[:30], used_fallback


def generate_and_store_question(skill, question_style, difficulty_band, model_name, timeout_seconds, fast_mode):
    q_prompt = build_question_prompt(skill, question_style, difficulty_band)
    question = call_ollama(
        q_prompt,
        model=model_name,
        status_text="Generating interview question...",
        timeout=timeout_seconds,
        options={
            "temperature": 0.2,
            "num_predict": 70 if fast_mode else 120,
            "num_ctx": 1024,
        },
    )
    st.session_state.question = (question or "").strip() or fallback_question(skill, question_style)
    st.session_state.current_skill = skill

    if str(question_style or "").strip().lower() == "coding":
        st.session_state.question_mode = "coding"
        coding_track = normalize_interview_skill(skill)
        if coding_track == "SQL":
            st.session_state.code_language = "sql"
            st.session_state.asked_language = "sql"
        else:
            default_language = st.session_state.get("code_language", "python")
            allowed_dsa_languages = {"python", "java", "cpp", "javascript", "c"}
            if default_language not in allowed_dsa_languages:
                default_language = "python"
            st.session_state.code_language = default_language
            st.session_state.asked_language = None
    else:
        st.session_state.question_mode = "voice"
        st.session_state.code_language = ""
        st.session_state.asked_language = None

    st.session_state.pop("answer_text", None)
    st.session_state.pop("code_answer", None)
    st.session_state.pop("reference_answer", None)
    st.session_state.pop("reference_type", None)
    st.session_state.pop("mistake_feedback", None)
    st.session_state.audio_key = st.session_state.get("audio_key", 0) + 1

    if not question:
        st.warning("Used fallback question because model response timed out.")


def pick_random_question_plan():
    question_style = random.choice(["Coding", "Theory"])
    if question_style == "Coding":
        topic = random.choice(list(ALLOWED_INTERVIEW_SKILLS))
        skill = normalize_interview_skill(topic)
    else:
        theory_topics = st.session_state.get("resume_skills") or st.session_state.get("skills") or ["General Technical"]
        topic = random.choice(theory_topics)
        skill = topic
    return question_style, topic, skill


init_user_database()
init_app_state()

# Keep user on the last known state if query params were used (e.g., email slot link)
# but still require auth_user if the user is truly logged out.
if not st.session_state.get("auth_user"):
    render_auth_gate()
    st.stop()

apply_level_two_slot_from_query()


with st.sidebar:
    st.header("Account")
    current_user = st.session_state.get("auth_user") or {}
    st.write(f"**Name:** {current_user.get('name', 'Candidate')}")
    st.write(f"**Email:** {current_user.get('email', '')}")
    st.write(f"**Current Level:** {st.session_state.get('current_level', 1)}")

    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()

    st.markdown("---")
    st.header("Interview Controls")
    uploaded_file = st.file_uploader("Upload PDF or DOCX", type=["pdf", "docx"])
    model_name = DEFAULT_MODEL
    fast_mode = True
    timeout_seconds = 240

    if st.button("Reset Session"):
        auth_user = st.session_state.get("auth_user")
        current_level = st.session_state.get("current_level", 1)
        level2_slot_start = st.session_state.get("level2_slot_start")
        level2_slot_end = st.session_state.get("level2_slot_end")
        st.session_state.clear()
        st.session_state.auth_user = auth_user
        st.session_state.current_level = current_level
        st.session_state.level2_slot_start = level2_slot_start
        st.session_state.level2_slot_end = level2_slot_end
        st.session_state.auth_mode = "Login"
        st.session_state.login_notice = ""
        st.session_state.completion_notice = ""
        st.rerun()

current_level = int(st.session_state.get("current_level", 1))
difficulty_band = get_question_difficulty(current_level)
candidate_name = (st.session_state.get("auth_user") or {}).get("name", "Candidate")
level2_slot_start = st.session_state.get("level2_slot_start")
level2_slot_end = st.session_state.get("level2_slot_end")
level2_window_state = get_level_two_window_state(level2_slot_start, level2_slot_end)

st.markdown(f"### Welcome, {candidate_name}")
if st.session_state.get("completion_notice"):
    st.success(st.session_state.completion_notice)
    st.session_state.completion_notice = ""
if st.session_state.get("level2_splash"):
    st.success(st.session_state.level2_splash)
    st.balloons()
    st.session_state.level2_splash = ""

if current_level == 1:
    st.info(
        f"Level 1 active: {difficulty_band} questions. "
        f"Score at least {LEVEL_ONE_THRESHOLD}/10 to unlock Level 2."
    )
else:
    st.success("Level 2 active: medium to high interview questions.")

if current_level >= 2 and level2_window_state == "expired":
    reset_candidate_to_level_one(
        "Sorry, you have no more attempts. Please return back and practice from first round."
    )

if current_level >= 2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Level 2 Slot Selection")
    slot_options = build_level_two_slot_options()
    slot_labels = [slot["label"] for slot in slot_options]
    current_slot_key = ""
    if isinstance(st.session_state.get("level2_slot_start"), datetime):
        current_slot_key = st.session_state["level2_slot_start"].strftime("%Y%m%dT%H%M")
    default_index = 0
    for idx, slot in enumerate(slot_options):
        if slot["slot_key"] == current_slot_key:
            default_index = idx
            break

    selected_slot = None
    if slot_labels:
        selected_label = st.selectbox(
            "Available Slots",
            options=slot_labels,
            index=min(default_index, len(slot_labels) - 1),
            key="level2_slot_choice",
        )
        selected_slot = next((slot for slot in slot_options if slot["label"] == selected_label), None)
        st.caption("Choose one of these given slots and save.")
    else:
        st.warning("No Level 2 slots are currently available. Please check again later.")

    start_dt = selected_slot["start_dt"] if selected_slot else None
    end_dt = selected_slot["end_dt"] if selected_slot else None
    calendar_title = f"Level 2 Interview - {candidate_name}"
    calendar_details = "AI Interview Studio Level 2 interview session (medium to high difficulty)."
    if selected_slot:
        calendar_link = build_google_calendar_link(calendar_title, calendar_details, start_dt, end_dt)
        st.markdown(f"[Add selected slot to Google Calendar]({calendar_link})")
    if st.button("Save Level 2 Slot", key="save_level2_slot_button"):
        if not selected_slot:
            st.error("Please choose one of the available slots.")
        elif start_dt <= datetime.now():
            st.error("Please select a future date and time.")
        else:
            current_email = (st.session_state.get("auth_user") or {}).get("email")
            saved = save_level_two_slot(current_email, start_dt, end_dt)
            if saved:
                st.session_state.level2_slot_start = start_dt
                st.session_state.level2_slot_end = end_dt
                level2_window_state = get_level_two_window_state(start_dt, end_dt)
                st.success("Level 2 slot saved. Come back at your selected time.")
            else:
                st.error("Failed to save Level 2 slot. Please try again.")

    if level2_window_state == "open":
        st.success("Your Level 2 interview is open now. You can start.")
        st.caption(
            f"Slot: {format_datetime_display(st.session_state.get('level2_slot_start'))} "
            f"to {format_datetime_display(st.session_state.get('level2_slot_end'))}"
        )
    elif level2_window_state == "upcoming":
        st.info(
            f"Level 2 will open at {format_datetime_display(level2_slot_start)}. "
            "Please return at that time."
        )
    elif level2_window_state == "expired":
        st.warning("Your previous slot has expired. Please select and save a new slot.")
    else:
        st.warning("Select and save your slot to unlock Level 2 at the scheduled time.")
    st.markdown("</div>", unsafe_allow_html=True)

if current_level == 2 and get_level_two_window_state(
    st.session_state.get("level2_slot_start"),
    st.session_state.get("level2_slot_end"),
) != "open":
    clear_interview_state(clear_skills=False)

resume_text = ""
if uploaded_file is not None:
    try:
        uploaded_bytes = uploaded_file.getvalue()
        file_signature = hashlib.sha256(uploaded_bytes).hexdigest()
        file_name = uploaded_file.name.lower()

        if file_name.endswith(".pdf"):
            pdf = fitz.open(stream=uploaded_bytes, filetype="pdf")
            for page in pdf:
                resume_text += page.get_text()
        elif file_name.endswith(".docx"):
            doc = docx.Document(io.BytesIO(uploaded_bytes))
            resume_text = "\n".join(p.text for p in doc.paragraphs)
        else:
            st.error("Unsupported resume format. Please upload a PDF or DOCX file.")
            resume_text = ""

        if resume_text.strip():
            st.success("Resume uploaded successfully.")
            st.session_state.resume_text = resume_text

            if file_signature != st.session_state.get("resume_signature", ""):
                skills, used_fallback = extract_resume_skills_from_text(
                    resume_text=resume_text,
                    model_name=model_name,
                    timeout_seconds=timeout_seconds,
                    fast_mode=fast_mode,
                )
                if not skills:
                    skills = ["Data Structures and Algorithms", "SQL"]
                    st.warning(
                        "Could not extract enough skills from resume. Using defaults so interview can continue."
                    )
                elif used_fallback:
                    st.warning("Used resume text fallback extraction because model response was invalid.")

                st.session_state.resume_signature = file_signature
                st.session_state.resume_skills = skills
                st.session_state.skills = skills
                st.session_state.resume_ready = True
                st.session_state.scores = []
                st.session_state.question_style = "Theory"
                st.session_state.question_topic = skills[0]
                st.session_state.auto_generate_after_resume = True
                st.session_state.pop("question", None)
                st.session_state.pop("current_skill", None)
                st.session_state.pop("reference_answer", None)
                st.session_state.pop("reference_type", None)
                st.session_state.pop("mistake_feedback", None)
                st.success("Skills extracted successfully from uploaded resume.")

            detected_skills = st.session_state.get("resume_skills") or []
            if detected_skills:
                st.markdown(f"**Detected Resume Skills:** {', '.join(detected_skills)}")
        else:
            st.error("Resume text could not be read. Please try another file.")
    except Exception as exc:
        st.error("Resume reading failed.")
        st.write(exc)
elif not st.session_state.get("resume_ready"):
    st.info("Upload your resume to extract skills and start interview questions.")


st.markdown("---")

if "skills" in st.session_state and st.session_state.skills:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    current_level = int(st.session_state.get("current_level", 1))
    difficulty_band = get_question_difficulty(current_level)
    level2_window_state = get_level_two_window_state(
        st.session_state.get("level2_slot_start"),
        st.session_state.get("level2_slot_end"),
    )
    can_generate_question = current_level == 1 or level2_window_state == "open"
    st.subheader(f"Interview Section - Level {current_level}")
    st.caption(
        f"Question difficulty: {difficulty_band}. "
        f"Level 1 passing threshold: >= {LEVEL_ONE_THRESHOLD}/10."
    )

    if not can_generate_question:
        st.info("Level 2 questions unlock only during your scheduled slot.")

    # st.caption(
    #     "Questions are generated randomly: Theory (voice answer) or Coding (DSA/SQL only). "
    #     "DSA coding supports multiple languages."
    # )

    if can_generate_question and st.session_state.get("auto_generate_after_resume"):
        st.session_state.auto_generate_after_resume = False
        selected_style, selected_topic, skill_for_question = pick_random_question_plan()
        st.session_state.question_style = selected_style
        st.session_state.question_topic = selected_topic
        generate_and_store_question(
            skill=skill_for_question,
            question_style=selected_style,
            difficulty_band=difficulty_band,
            model_name=model_name,
            timeout_seconds=timeout_seconds,
            fast_mode=fast_mode,
        )

    if can_generate_question and st.button("Generate Question"):
        selected_style, selected_topic, skill_for_question = pick_random_question_plan()
        st.session_state.question_style = selected_style
        st.session_state.question_topic = selected_topic
        generate_and_store_question(
            skill=skill_for_question,
            question_style=selected_style,
            difficulty_band=difficulty_band,
            model_name=model_name,
            timeout_seconds=timeout_seconds,
            fast_mode=fast_mode,
        )

    if "question" in st.session_state:
        if st.session_state.get("question_style") and st.session_state.get("question_topic"):
            st.caption(
                f"Generated Type: {st.session_state.get('question_style')} | "
                f"Topic: {st.session_state.get('question_topic')}"
            )
        st.info(st.session_state.question)
        question_mode = st.session_state.get("question_mode", "voice")
        detected_lang = st.session_state.get("code_language", "") if question_mode == "coding" else ""
        if detected_lang:
            st.caption(f"Detected Mode: {question_mode.upper()} | Language: {detected_lang.upper()}")
        else:
            st.caption(f"Detected Mode: {question_mode.upper()}")

        if question_mode == "coding":
            #st.info("Coding/SQL question detected. Write your solution below. Voice input is disabled for this question.")
            active_skill = normalize_interview_skill(st.session_state.get("current_skill", "DSA"))
            if active_skill == "SQL":
                language_options = ["sql"]
            else:
                language_options = ["python", "java", "cpp", "javascript", "c"]
            default_language = st.session_state.get("code_language", "python")
            if default_language not in language_options:
                default_language = language_options[0]
            asked_language = st.session_state.get("asked_language")
            if asked_language in language_options:
                code_language = asked_language
                st.selectbox(
                    "Coding Language (from question)",
                    options=language_options,
                    index=language_options.index(asked_language),
                    key="code_language_locked",
                    disabled=True,
                )
                st.caption(f"Answer is locked to `{asked_language}` because the question asks for it.")
            else:
                code_language = st.selectbox(
                    "Coding Language",
                    options=language_options,
                    index=language_options.index(default_language),
                    key="code_language",
                )
            

            code_answer = st_ace(
                placeholder="Write your code here...",
                language=code_language,
                theme="monokai",
                keybinding="vscode",
                height=400,
                font_size=14,
                tab_size=4,
            )

            if st.button("Evaluate Code Answer"):
                if not code_answer.strip():
                    st.warning("Please write your code before evaluation.")
                    st.stop()

                parsed_eval = evaluate_code_answer(
                    question=st.session_state.question,
                    skill=st.session_state.get("current_skill", ""),
                    code=code_answer,
                    model=model_name,
                    timeout_seconds=timeout_seconds,
                    fast_mode=fast_mode,
                )

                score = parse_score(parsed_eval.get("total_score", 0))
                st.session_state.setdefault("scores", []).append(score)

                st.subheader("Code Evaluation Result")
                if score >= 8:
                    st.success(f"Excellent score: {score}/10")
                elif score >= 5:
                    st.warning(f"Moderate performance: {score}/10")
                else:
                    st.error(f"Needs improvement: {score}/10")

                render_feedback_block("Strengths", parsed_eval.get("strengths", ""))
                render_feedback_block("Weaknesses", parsed_eval.get("weaknesses", ""))
                render_feedback_block("Improvement Suggestion", parsed_eval.get("improvement_suggestion", ""))

                correct_code = generate_reference_code(
                    question=st.session_state.question,
                    skill=st.session_state.get("current_skill", ""),
                    model=model_name,
                    timeout_seconds=timeout_seconds,
                    fast_mode=fast_mode,
                    language=code_language,
                )
                st.session_state.reference_answer = correct_code
                st.session_state.reference_type = "code"
                st.session_state.mistake_feedback = generate_code_mistake_feedback(
                    question=st.session_state.question,
                    user_code=code_answer,
                    correct_code=correct_code,
                    model=model_name,
                    timeout_seconds=timeout_seconds,
                    fast_mode=fast_mode,
                )
                apply_level_progression(score)
                if int(st.session_state.get("current_level", 1)) == 2:
                    if get_level_two_window_state(
                        st.session_state.get("level2_slot_start"),
                        st.session_state.get("level2_slot_end"),
                    ) == "open":
                        reset_candidate_to_level_one()
        else:
            if "audio_key" not in st.session_state:
                st.session_state.audio_key = 0

            st.write("Record your answer:")
            audio_file = st.audio_input(
                "Click to record",
                key=f"answer_audio_{st.session_state.audio_key}",
            )

            if st.button("Transcribe Speech"):
                if audio_file is None:
                    st.warning("Please record your answer first.")
                    st.stop()

                transcript, error = transcribe_audio(audio_file)
                if error:
                    st.error(error)
                    st.stop()

                st.session_state.answer_text = transcript
                st.success("Speech transcribed successfully.")

            answer = st.session_state.get("answer_text", "")
            if answer:
                st.write("### Transcribed Answer")
                st.write(answer)

            if st.button("Evaluate Spoken Answer"):
                if not answer.strip():
                    st.warning("Please record and transcribe your spoken answer before evaluation.")
                    st.stop()

                eval_prompt = f"""
You are a senior technical interviewer.
Evaluate the answer and return ONLY JSON:
{{
  "total_score": 0,
  "strengths": "",
  "weaknesses": "",
  "improvement_suggestion": ""
}}

Evaluation rules:
- Accept different wording and speaking style if technical meaning is correct.
- Score semantic relevance to the expected concept, not exact keyword match.
- If answer is partially correct and related to the question, give partial credit.
- Do not penalize for accent, grammar, or minor transcription noise.

Question:
{st.session_state.question}

Answer:
{answer}
"""
                result = call_ollama(
                    eval_prompt,
                    model=model_name,
                    status_text="Evaluating answer...",
                    timeout=timeout_seconds,
                    options={
                        "temperature": 0.1,
                        "num_predict": 160 if fast_mode else 280,
                        "num_ctx": 2048,
                    },
                )
                parsed_eval = extract_json(result) if result else None
                if not parsed_eval:
                    parsed_eval = fallback_evaluation(answer)
                    st.warning("Used fallback evaluation because Ollama response timed out.")

                score = parse_score(parsed_eval.get("total_score", 0))
                st.session_state.setdefault("scores", []).append(score)

                st.subheader("Evaluation Result")
                if score >= 8:
                    st.success(f"Excellent score: {score}/10")
                elif score >= 5:
                    st.warning(f"Moderate performance: {score}/10")
                else:
                    st.error(f"Needs improvement: {score}/10")

                render_feedback_block("Strengths", parsed_eval.get("strengths", ""))
                render_feedback_block("Weaknesses", parsed_eval.get("weaknesses", ""))
                render_feedback_block("Improvement Suggestion", parsed_eval.get("improvement_suggestion", ""))

                correct_answer = generate_reference_answer(
                    question=st.session_state.question,
                    skill=st.session_state.get("current_skill", ""),
                    model=model_name,
                    timeout_seconds=timeout_seconds,
                    fast_mode=fast_mode,
                )
                st.session_state.reference_answer = correct_answer
                st.session_state.reference_type = "text"
                st.session_state.mistake_feedback = generate_mistake_feedback(
                    question=st.session_state.question,
                    user_answer=answer,
                    correct_answer=correct_answer,
                    model=model_name,
                    timeout_seconds=timeout_seconds,
                    fast_mode=fast_mode,
                )
                apply_level_progression(score)
                if int(st.session_state.get("current_level", 1)) == 2:
                    if get_level_two_window_state(
                        st.session_state.get("level2_slot_start"),
                        st.session_state.get("level2_slot_end"),
                    ) == "open":
                        reset_candidate_to_level_one()

        if st.session_state.get("reference_answer"):
            st.markdown("---")
            if st.session_state.get("reference_type") == "code":
                st.subheader("Correct Code (for Learning)")
                st.code(st.session_state.reference_answer, language=st.session_state.get("code_language", "python"))
            else:
                st.subheader("Correct Answer (for Learning)")
                st.write(st.session_state.reference_answer)
            if st.session_state.get("mistake_feedback"):
                render_feedback_block("What You Missed", st.session_state.mistake_feedback)

    st.markdown("</div>", unsafe_allow_html=True)
