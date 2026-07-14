"""
Shared officer account store — used by both the public intake site (app.py)
and the case tracker (admin.py).

Credentials are stored salted + hashed (PBKDF2-SHA256) in admin_users.json.
Keep that file out of version control.
"""

import hashlib
import hmac
import json
import re
import secrets
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).parent
USERS_FILE = HERE / "admin_users.json"
_PBKDF2_ITERATIONS = 200_000


def hash_password(password: str, salt_hex: str | None = None):
    salt_hex = salt_hex or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode(), bytes.fromhex(salt_hex), _PBKDF2_ITERATIONS
    ).hex()
    return salt_hex, digest


def load_users() -> dict:
    try:
        with open(USERS_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def save_users(users: dict) -> None:
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)


def add_user(users: dict, username: str, password: str) -> str:
    """Add an account. Returns '' on success or an error message."""
    username = username.strip().lower()
    if not re.fullmatch(r"[a-z0-9_.-]{2,30}", username):
        return "Username must be 2–30 characters (letters, numbers, . _ -)."
    if username in users:
        return "That username already exists."
    if len(password) < 8:
        return "Password must be at least 8 characters."
    salt, digest = hash_password(password)
    users[username] = {
        "salt": salt,
        "hash": digest,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    save_users(users)
    return ""


def verify_login(users: dict, username: str, password: str) -> bool:
    u = users.get(username.strip().lower())
    if not u:
        return False
    _, digest = hash_password(password, u.get("salt", ""))
    return hmac.compare_digest(digest, u.get("hash", ""))
