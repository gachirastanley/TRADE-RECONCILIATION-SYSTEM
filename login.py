import json
import hashlib
import os
from datetime import datetime

# =========================
# ✅ FILE PATH
# =========================
USER_FILE = "users.json"


# =========================
# ✅ HELPERS
# =========================
def load_users():
    """Load users from JSON file"""
    if os.path.exists(USER_FILE):
        with open(USER_FILE, "r") as f:
            return json.load(f)
    return {}


def save_users(users):
    """Save users to JSON file"""
    with open(USER_FILE, "w") as f:
        json.dump(users, f, indent=4)


# =========================
# ✅ PASSWORD SECURITY
# =========================
def hash_password(password, salt=None):
    """
    Hash password with optional salt
    """
    if salt is None:
        salt = os.urandom(16).hex()  # generate random salt

    hashed = hashlib.sha256((password + salt).encode()).hexdigest()
    return hashed, salt


# =========================
# ✅ REGISTER USER
# =========================
def register_user(username, password):
    users = load_users()

    if username in users:
        return False, "Username already exists"

    if len(password) < 4:
        return False, "Password must be at least 4 characters"

    hashed, salt = hash_password(password)

    users[username] = {
        "password": hashed,
        "salt": salt,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    save_users(users)

    return True, "Registration successful"


# =========================
# ✅ LOGIN USER
# =========================
def login_user(username, password):
    users = load_users()

    if username not in users:
        return False

    user = users[username]

    # ✅ Handle incorrect/old format safely
    if not isinstance(user, dict):
        return False

    stored_hash = user.get("password")
    salt = user.get("salt")

    if not stored_hash or not salt:
        return False

    hashed, _ = hash_password(password, salt)

    return hashed == stored_hash



# =========================
# ✅ DELETE USER (OPTIONAL)
# =========================
def delete_user(username):
    users = load_users()

    if username in users:
        del users[username]
        save_users(users)
        return True

    return False


# =========================
# ✅ GET ALL USERS (ADMIN)
# =========================
def get_all_users():
    return load_users()
