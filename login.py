import streamlit as st
import json
import hashlib
import os

# =========================
# ✅ FILE TO STORE USERS
# =========================
USER_FILE = "users.json"


# =========================
# ✅ HELPERS
# =========================
def load_users():
    if os.path.exists(USER_FILE):
        with open(USER_FILE, "r") as f:
            return json.load(f)
    return {}


def save_users(users):
    with open(USER_FILE, "w") as f:
        json.dump(users, f)


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


# =========================
# ✅ AUTH FUNCTIONS
# =========================
def register_user(username, password):
    users = load_users()

    if username in users:
        return False, "Username already exists"

    users[username] = hash_password(password)
    save_users(users)

    return True, "Registration successful"


def login_user(username, password):
    users = load_users()

    if username not in users:
        return False

    if users[username] == hash_password(password):
        return True

    return False


# =========================
# ✅ SESSION STATE INIT
# =========================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""


# =========================
# ✅ UI
# =========================
st.title("🔐 Secure Login System")

menu = ["Login", "Register"]
choice = st.sidebar.selectbox("Menu", menu)


# =========================
# ✅ REGISTER
# =========================
if choice == "Register":

    st.subheader("Create New Account")

    new_user = st.text_input("Username")
    new_pass = st.text_input("Password", type="password")

    if st.button("Register"):

        if new_user == "" or new_pass == "":
            st.warning("Please fill all fields")

        else:
            success, message = register_user(new_user, new_pass)

            if success:
                st.success(message)
            else:
                st.error(message)


# =========================
# ✅ LOGIN
# =========================
elif choice == "Login":

    st.subheader("Login to Continue")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):

        if login_user(username, password):
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success(f"Welcome {username} ✅")

        else:
            st.error("Invalid username or password")


# =========================
# ✅ LOGGED-IN AREA
# =========================
if st.session_state.logged_in:

    st.sidebar.success(f"Logged in as {st.session_state.username}")

    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.warning("Logged out successfully")

    # 🔽 PUT YOUR MAIN APP HERE
    st.subheader("✅ Access Granted")

    st.write("This is your protected app area.")

    # 👉 Example:
    st.write("Place your CDC receipting UI here")
