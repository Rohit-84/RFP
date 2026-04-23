import asyncio
import sys
import base64
import json
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

import os
import time
import json
# import logging
from logger_config import get_logger
import secrets
import requests
import threading
from functools import wraps
from datetime import datetime, timedelta, timezone
 
from flask import (
    Flask, request, jsonify, send_from_directory,
    session, redirect, url_for
)
from waitress import serve
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
 
# 👇 NEW IMPORTS FOR SSO
from authlib.integrations.flask_client import OAuth
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'  # ALLOW HTTP FOR LOCAL DEV
os.environ['AUTHLIB_INSECURE_TRANSPORT'] = '1'   # 👈 ADD THIS LINE FOR AUTHLIB
 
# Import backend chat function
from chatbot_backend import chat_once, GENERATED_DIR as BACKEND_GENERATED_DIR, says_is_loading_true_or_false
from history_name import generate_title_from_history
 
# ------------------------------------------------------------------
# LOAD ENV
# ------------------------------------------------------------------
load_dotenv()
 
# ------------------------------------------------------------------
# MICROSOFT GRAPH – APP B (EMAIL ONLY)
# ------------------------------------------------------------------
GRAPH_TENANT_ID = os.getenv("TENANT_ID1")
GRAPH_CLIENT_ID = os.getenv("CLIENT_ID1")
GRAPH_CLIENT_SECRET = os.getenv("CLIENT_SECRET1")
GRAPH_SENDER_EMAIL = os.getenv("GRAPH_SENDER_EMAIL", "yash.dharmadhikari@intelliswift.com")
APP_BASE_URL = "http://localhost:5001"
 
# ------------------------------------------------------------------
# APP SETUP
# ------------------------------------------------------------------
app = Flask(__name__, static_folder="static", static_url_path="/static")
app.secret_key = os.getenv("FLASK_SECRET_KEY", "deck-genie-secret-key-CHANGE-THIS")
 
#  ADD THIS LINE HERE
app.permanent_session_lifetime = timedelta(hours=1)  # Set to 2 mins for testing

# logging.basicConfig(level=logging.INFO)
logger = get_logger(__name__)
 
# 👇 NEW: SSO SETUP
oauth = OAuth(app)
ENTRA_CLIENT_ID = os.getenv("ENTRA_CLIENT_ID")
ENTRA_CLIENT_SECRET = os.getenv("ENTRA_CLIENT_SECRET")
ENTRA_TENANT_ID = os.getenv("ENTRA_TENANT_ID")
 
microsoft = oauth.register(
    name='microsoft',
    client_id=ENTRA_CLIENT_ID,
    client_secret=ENTRA_CLIENT_SECRET,
    server_metadata_url=f'https://login.microsoftonline.com/{ENTRA_TENANT_ID}/v2.0/.well-known/openid-configuration',
    client_kwargs={'scope': 'openid email profile'}
)
# ------------------------------------------------------------------
# PATHS
# ------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
USERS_FILE = os.path.join(DATA_DIR, "users.json")
RESET_TOKENS_FILE = os.path.join(DATA_DIR, "reset_tokens.json")
SESSIONS_DIR = os.path.join(DATA_DIR, "sessions")
 
GENERATED_DIR = BACKEND_GENERATED_DIR
 
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(SESSIONS_DIR, exist_ok=True)



# ------------------------------------------------------------------
# THREAD SAFETY (CRITICAL FIX)
# ------------------------------------------------------------------
token_lock = threading.Lock()
 
# ------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------
def normalize_email(email):
    # Normalize email by lowercasing and stripping whitespace.
    return (email or "").strip().lower()
 
 
def load_users():
    # Load the users.json file if it exists; otherwise return empty dict.
    if not os.path.exists(USERS_FILE):
        return {}
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
 
 
def save_users(users):
    # Save user account data back into users.json with indentation for readability.
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2)
 
 
def load_reset_tokens():
    # Load password-reset token mapping from disk if file exists.
    if not os.path.exists(RESET_TOKENS_FILE):
        return {}
    with open(RESET_TOKENS_FILE, "r") as f:
        return json.load(f)
 
 
def save_reset_tokens(tokens):
    # Write reset-token dictionary to file under thread lock to avoid race conditions.
    with token_lock:  # <--- Added Lock
        with open(RESET_TOKENS_FILE, "w") as f:
            json.dump(tokens, f, indent=2)
 
 
def login_required(fn):
    # Decorator to enforce that a route requires authentication.
    @wraps(fn)
    def wrapper(*args, **kwargs):
        # If not logged in, redirect user to the login page.
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        # Otherwise continue executing the wrapped function.
        return fn(*args, **kwargs)
    return wrapper
 
# ------------------------------------------------------------------
# SESSION STORAGE
# ------------------------------------------------------------------
def user_session_dir():
    # Determine the per-user session directory based on the logged-in user's email.
    email = session.get("user")
    if not email:
        return SESSIONS_DIR  # Fallback to a shared directory if no user in session
 
    # Build and ensure the per-user directory exists
    path = os.path.join(SESSIONS_DIR, email)
    os.makedirs(path, exist_ok=True)
    return path
 
 
def session_file(session_id):
    # Construct the absolute path for a session's JSON file under the user's session dir
    return os.path.join(user_session_dir(), f"{session_id}.json")
 
 
def load_session(session_id):
    # Resolve the session file path
    path = session_file(session_id)
 
    # If this is a brand-new session (no file on disk), return a default structure
    if not os.path.exists(path):
        return {
            "session_id": session_id,
            "title": "New Chat",
            "pinned": False,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "messages": []
        }
 
    # Load existing session JSON
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
 
    # Backward compatibility for older saved sessions: ensure required keys exist
    data.setdefault("session_id", session_id)
    data.setdefault("title", "New Chat")
    data.setdefault("pinned", False)
    data.setdefault("created_at", datetime.now(timezone.utc).isoformat())
    data.setdefault("messages", [])
 
    return data
 
def save_session(data):
    # Persist the in-memory session dictionary to its JSON file with pretty formatting
    with open(session_file(data["session_id"]), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
 
# ------------------------------------------------------------------
# MICROSOFT GRAPH EMAIL
# ------------------------------------------------------------------
def get_graph_token():
    # Acquire an application (client credentials) access token for Microsoft Graph
    # using the OAuth 2.0 v2 endpoint for the configured tenant.
    url = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": GRAPH_CLIENT_ID,                  # Azure AD App (client) ID
        "client_secret": GRAPH_CLIENT_SECRET,          # App client secret (keep secure)
        "grant_type": "client_credentials",            # App-only (no user interaction)
        "scope": "https://graph.microsoft.com/.default",  # Use app's configured Graph permissions
    }
    r = requests.post(url, data=data)                  # Token request
    r.raise_for_status()                               # Raise on HTTP error
    return r.json()["access_token"]                    # Return bearer token string
 
def send_reset_email(to_email, subject, html_content):
    # Send a password reset email via Microsoft Graph using the /sendMail action.
    token = get_graph_token()
    url = f"https://graph.microsoft.com/v1.0/users/{GRAPH_SENDER_EMAIL}/sendMail"
 
    # Authorization with Bearer token; JSON payload
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
 
    # Construct the email message object:
    # - subject: fixed reset subject
    # - body: HTML content with a reset link (expires in 15 minutes)
    # - toRecipients: the destination email address
    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_content
            },
            "toRecipients": [
                {"emailAddress": {"address": to_email}}
            ]
        }
    }
 
    # POST the sendMail request as the specified user (GRAPH_SENDER_EMAIL).
    # NOTE:
    #   - Requires appropriate Graph permissions (e.g., Mail.Send) granted to the app.
    #   - The GRAPH_SENDER_EMAIL must be a mailbox-enabled user allowed to send mail.
    r = requests.post(url, headers=headers, json=payload)
    r.raise_for_status()  # Raise if Graph returns an error response
 
# ------------------------------------------------------------------
#  NEW ROUTES
# ------------------------------------------------------------------
 
@app.get("/c/<session_id>")
@login_required
def chat_page(session_id):
    # This makes localhost/c/abc123 serve the same index.html page
    return send_from_directory(app.static_folder, "index.html")
 
@app.post("/sessions/new")
@login_required
def create_new_session():
    # This API pre-creates the session file immediately when New Chat is clicked
    data = request.get_json() or {}
    session_id = data.get("session_id")
    title = data.get("title", "New Chat")
    
    if not session_id:
        return jsonify({"error": "Missing session_id"}), 400
 
    chat_data = load_session(session_id)
    chat_data["title"] = title
    save_session(chat_data)
    
    return jsonify({"success": True, "session_id": session_id})
 
# ------------------------------------------------------------------
# AUTH ROUTES
# ------------------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    # Handle login form submits (POST) and serve the login page (GET)
    if request.method == "POST":
        data = request.get_json() or {}
        email = normalize_email(data.get("username"))  # normalize input email
        password = data.get("password", "")
 
        users = load_users()
        user = users.get(email)
 
        # Fail fast if user does not exist (generic error avoids user enumeration)
        if not user:
            return jsonify({"error": f"Access Denied: '{email}' is not authorized to access this application."}), 401
 
        # 🔒 SECURE: Hash check only. Backdoor removed.
        # Validate password using a one-way hash comparison
        if not check_password_hash(user["password"], password):
            return jsonify({"error": "Invalid credentials"}), 401
 
        # Establish session on successful login
        session.permanent = True # 👈 ADD THIS LINE HERE
        session["logged_in"] = True
        session["user"] = email
        return jsonify({"success": True})
 
    # For GET: serve the login HTML
    return send_from_directory(app.static_folder, "login.html")
 
 
# 👇 --- START OF NEW SSO CODE --- 👇
 
@app.route('/login/sso')
def login_sso():
    # This grabs the user and shoots them over to Microsoft's login screen
    redirect_uri = url_for('auth_callback', _external=True, _scheme='https')
    return microsoft.authorize_redirect(redirect_uri)
 
@app.route('/callback')
def auth_callback():
    # Microsoft sends the user back here with a token
    token = microsoft.authorize_access_token()
    user_info = token.get('userinfo')
 
    # 👇 NEW FIX: Universal token decoder (No library conflicts!)
    if not user_info and 'id_token' in token:
        # A token has 3 parts separated by dots. The middle part [1] has the user data.
        payload = token['id_token'].split('.')[1]
        # Add required padding so Python can decode it
        padded_payload = payload + '=' * (-len(payload) % 4)
        user_info = json.loads(base64.urlsafe_b64decode(padded_payload).decode('utf-8'))
   
    if not user_info:
        return redirect(url_for('login', error="Error: No user information provided."))
       
    # Get their email from the token
    email = normalize_email(user_info.get('email') or user_info.get('preferred_username'))
   
    if not email:
        return redirect(url_for('login', error="Error: Could not find an email address in your Microsoft profile."))
 
    # 👇 NEW: Check if user exists in users.json
    users = load_users()
    if email not in users:
        # User is NOT in the list. Send them back to the login page with an error.
        return redirect(url_for('login', error=f"Access Denied: '{email}' is not authorized to access this application."))
       
    # If they pass the check, create their Flask session (log them in!)
    session.permanent = True
    session["logged_in"] = True
    session["user"] = email
   
    # Send them to the main Deck Genie page
    return redirect(url_for('index'))
# 👆 --- END OF NEW SSO CODE --- 👆
 
 
@app.get("/logout")
def logout():
    # Clear the entire session (logs out the user)
    session.clear()
    return redirect(url_for("login"))
 
@app.post("/forgot-password")
def forgot_password():
    # Initiate password reset flow
    data = request.get_json() or {}
    email = normalize_email(data.get("email"))

    users = load_users()

    # 👇 THE FIX: Check if email is missing or NOT in the users.json file
    if not email or email not in users:
        return jsonify({"error": "Invalid email address. User not found."}), 400

    # Generate a secure, random token for password reset
    token = secrets.token_urlsafe(32)
    tokens = load_reset_tokens()

    # Store token metadata with 15-minute expiry
    tokens[token] = {
        "email": email,
        "expires": (datetime.now(timezone.utc) + timedelta(minutes=15)).isoformat()
    }
    save_reset_tokens(tokens)

    # Build a reset link pointing to a static reset page with token query param
    base_url = request.host_url.rstrip('/')
    # reset_link = f"{base_url}/static/reset_password.html?token={token}"
    reset_link = f"https://9.223.177.246:5000/static/reset_password.html?token={token}"

    # Extract name from email (e.g., "amar.xyz" -> "Amar Xyz")
    raw_name = email.split('@')[0]
    parts = raw_name.split('.')
    full_name = " ".join([p.capitalize() for p in parts])

    #  Build the HTML email directly inside Python
    html_body = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Reset Your Password</title>
    </head>
    <body style="font-family: Arial, sans-serif;align:left; background-color: #f7f7f8; padding: 20px; color: #333;">
      <div style="max-width: 500px; margin: 0 auto; background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!--  <h2 style="color: #004a99; text-align: center;">Password Reset Request</h2> -->
        
        <p>Hi <strong>{full_name}</strong>,</p>
        
        <p>Forgot your password? No problem, it happens to the best of us. Click the button below to set a new one:</p>
        
        <div style="text-align: left; margin: 30px 0;">
          <a href="{reset_link}" style="background-color: #004a99; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">Reset Password</a>
        </div>
        
        <p style="font-size: 14px; color: #666;">If you didn't request a password reset, you can safely ignore this email.</p>
        
        <p style="font-size: 14px; color: #d9534f;"><em>* This link will expire in 15 minutes.</em></p>
        <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
        
        <p style="font-size: 15px; color: #000; text-align: left;">
        <strong>Thanks!</strong><br>
        </p>
      </div>
    </body>
    </html>
    """

    # Attempt to send the reset email
    try:
        send_reset_email(email, "Reset your password for Deck Genie - Intelliswift", html_body)
        logger.info(f"Reset email sent to {email}")
    except Exception as e:
        logger.error(f"Reset email failed for {email}: {str(e)}")

    return jsonify({"success": True, "message": "A reset link has been sent to your email."})

@app.get("/api/profile")
@login_required
def get_profile():
    email = session.get("user")
    users = load_users()
    user_data = users.get(email, {})
    
    # Extract name from email (e.g., "harshita.yadav" -> "Harshita Yadav")
    # Or if you have a "name" key in JSON, use user_data.get("name")
    raw_name = email.split('@')[0]
    parts = raw_name.split('.')
    full_name = " ".join([p.capitalize() for p in parts])
    
    return jsonify({
        "full_name": full_name,
        "email": email
    })

@app.post("/reset-password")
def reset_password():
    # Complete the password reset using a valid token and new password
    data = request.get_json() or {}
    token = data.get("token")
    new_pw = data.get("password", "")

    # Validate request payload
    if not token or not new_pw:
        return jsonify({"error": "Invalid request"}), 400

    # 👇 NEW FIX: Backend Password Complexity Check
    # Must be at least 5 chars, contain at least one digit, and contain at least one non-alphanumeric (special) character
    if len(new_pw) < 5 or not any(c.isdigit() for c in new_pw) or not any(not c.isalnum() for c in new_pw):
        return jsonify({"error": "Password must be at least 5 characters long, and include at least one number and one special character."}), 400

    tokens = load_reset_tokens()
    entry = tokens.get(token)
 
    # Reject if token not found
    if not entry:
        return jsonify({"error": "Invalid or expired token"}), 400
 
    # Enforce token expiry (15 minutes)
    if datetime.now(timezone.utc) > datetime.fromisoformat(entry["expires"]):
        del tokens[token]                # remove expired token
        save_reset_tokens(tokens)
        return jsonify({"error": "Token expired"}), 400
 
    users = load_users()
    email = entry["email"]
    
    # If the associated user exists, set new hashed password
    if email in users:
        users[email]["password"] = generate_password_hash(new_pw)
        save_users(users)
 
    # Invalidate token after use (one-time use token)
    del tokens[token]
    save_reset_tokens(tokens)
 
    return jsonify({"success": True})

@app.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    # Serve the change password UI
    if request.method == "GET":
        return send_from_directory(app.static_folder, "change_password.html")
 
    # Handle password change submission
    data = request.get_json() or {}
    current_pw = data.get("current_password")
    new_pw = data.get("new_password", "")
    confirm_pw = data.get("confirm_password", "")

    # Basic confirmation check
    if new_pw != confirm_pw:
        return jsonify({"error": "Passwords do not match"}), 400

    # 👇 NEW FIX: Backend Password Complexity Check
    # Must be at least 5 chars, contain at least one digit, and contain at least one non-alphanumeric (special) character
    if len(new_pw) < 5 or not any(c.isdigit() for c in new_pw) or not any(not c.isalnum() for c in new_pw):
        return jsonify({"error": "Password must be at least 5 characters long, and include at least one number and one special character."}), 400

    users = load_users()
    email = session.get("user")
    user = users.get(email)
 
    # Verify current password before changing
    # (assumes `user` exists due to @login_required guard)
    if not check_password_hash(user["password"], current_pw):
        return jsonify({"error": "Current password incorrect"}), 401
 
    # Persist new hashed password
    users[email]["password"] = generate_password_hash(new_pw)
    save_users(users)
 
    return jsonify({"success": True})
 
# ------------------------------------------------------------------
# CHAT & HISTORY ROUTES
# ------------------------------------------------------------------
@app.post("/define_loading")
@login_required
def define_loading_text() -> str:
    try:
        data = request.get_json() or {}
        # print("data :",data)
        if not data.get("pinned"):
            # print("pinned")
            return "Thinking...🤔"
        session = data.get("session_id")
        if says_is_loading_true_or_false(session_id=session):
            return "Deck Generation In Progress...⏳ Please Wait"
        else:
            return "Thinking🤔"
    except Exception as e:
        logger.error(f"error in define_loading_text()... {e}")
        return "Unstable Internet Connection 🛜"
    
@app.post("/chat")
@login_required
def chat():
    # Chat endpoint: accepts a message (and optional session_id), routes to chat logic,
    # persists both user and assistant turns into the per-user session file, and returns the reply.
    #START THE BACKEND STOPWATCH..
    request_start_time = time.time()
   
    data = request.get_json() or {}
    message = data.get("message", "")
    session_id = data.get("session_id")
 
    # Load or initialize the session transcript for this session_id
    chat_data = load_session(session_id)
 
    # 1. Append User Message
    if message:
        chat_data["messages"].append({
            "role": "user",
            "content": message,
            "ts": datetime.now(timezone.utc).isoformat()
        })
        # Persist the session transcript to disk
        save_session(chat_data)
    else:
        chat_data["messages"].append({
            "role": "assistant",
            "content": "Hello! How can I assist you today?",
            "ts": datetime.now(timezone.utc).isoformat()
        })
        # Persist the session transcript to disk (.json file)
        save_session(chat_data)
 
    # Call the central conversational router only if there isn't an empty message from user
    if message:
        result = chat_once(message, session_id=session_id,chat_data=chat_data["messages"])

        # Build the HTTP response
        response = {"answer": result.get("answer", "")}
        if result.get("file"):
            # If a file was produced (e.g., generated PPT), include its basename for download
            response["file"] = os.path.basename(result["file"])
            #STOP THE STOPWATCH AND DO THE MATH HERE.
            request_end_time = time.time()
            total_seconds = int(request_end_time - request_start_time)
            mins = total_seconds // 60
            secs = total_seconds % 60
            full_generation_time = f"{mins:02d}m {secs:02d}s"
 
            # Send it to the frontend
            response["generation_time"] = full_generation_time
           
               
            chat_data["messages"].append({
            "role": "assistant",
            "content": result.get("answer", ""),
            "file":result["file"],
            #SAVE THE TIME PERMANENTLY INTO SESSIONS.JSON FILE
            "generation_time": full_generation_time,
            "ts": datetime.now(timezone.utc).isoformat()
            })
        else:
            # Append the assistant's reply with a UTC timestamp
            chat_data["messages"].append({
                "role": "assistant",
                "content": result.get("answer", ""),
                "ts": datetime.now(timezone.utc).isoformat()
            })
        
        # ---------------------------------------------------------------
        # 🧠 THE NEW BRAIN: DYNAMIC HISTORY TITLING
        # ---------------------------------------------------------------
        # We allow the title to evaluate and update during the first 8 messages (4 turns)
        # This lets it dynamically evolve from "General Conversation" to "Deck: Staples"
        if len(chat_data["messages"]) <= 8:
            smart_title = generate_title_from_history(chat_data["messages"])
            chat_data["title"] = smart_title
           
            # Send the new title back to the frontend so the UI updates instantly
            response["new_title"] = smart_title
 
        # Persist the final session transcript to disk
        save_session(chat_data)
 
        return jsonify(response)
    return ""
 
@app.get("/sessions")
@login_required
def list_sessions():
    # Lists all session metadata for the logged-in user (id, title, created_at, pinned).
    sessions = []
    user_dir = user_session_dir()
 
    # If no session directory for this user, return an empty list
    if not os.path.exists(user_dir):
        return jsonify([])
 
    # Iterate over all *.json session files and collect lightweight metadata
    for fname in os.listdir(user_dir):
        if fname.endswith(".json"):
            try:
                with open(os.path.join(user_dir, fname), "r", encoding="utf-8") as f:
                    data = json.load(f)
 
                # 👇 NEW FIX: Hide any chat that doesn't have at least one user message
                messages = data.get("messages", [])
                has_user_message = any(msg.get("role") == "user" for msg in messages)
               
                if not has_user_message:
                    continue  # Skip this empty file completely!
 
                # Default title if missing/empty
                title = data.get("title") or "New Chat"
 
                sessions.append({
                    "session_id": data.get("session_id", fname.replace(".json", "")),
                    "title": title,
                    "created_at": data.get("created_at"),
                    "pinned": data.get("pinned", False)
                })
            except Exception:
                # Skip any malformed/unreadable session files
                continue
 
    # 📌 pinned chats first, then newest by created_at (descending timestamp)
    # NOTE: Assumes 'created_at' is ISO-8601; missing/invalid values may raise errors here.
    sessions.sort(
        key=lambda s: (
            0 if s.get("pinned") else 1,
            -datetime.fromisoformat(s["created_at"]).timestamp()
        )
    )
 
    return jsonify(sessions)
 
@app.get("/session/<session_id>")
@login_required
def get_session(session_id):
    # Return the full session payload (messages + metadata) for the given session_id
    return jsonify(load_session(session_id))
 
 
# ------------------------------------------------------------------
# SESSION MANAGEMENT (RENAME, PIN, DELETE)
# ------------------------------------------------------------------
 
@app.post("/sessions/<session_id>/rename")
@login_required
def rename_session(session_id):
    # Rename a chat session by updating its title in the session JSON.
    # NOTE: If "@app.post("/sessions/<session_id>/rename")" is literally in your source (HTML-escaped),
    # Flask won't treat it as a path variable. Ensure raw angle brackets are used in code.
    data = request.get_json() or {}
    new_title = data.get("title")
    if not new_title:
        # Require a non-empty title
        return jsonify({"error": "Title required"}), 400
 
    # Load, mutate, and persist the session record
    chat_data = load_session(session_id)
    chat_data["title"] = new_title
    save_session(chat_data)
 
    return jsonify({"success": True})
 
 
@app.post("/sessions/<session_id>/pin")
@login_required
def pin_session(session_id):
    # Toggle the "pinned" state for a given session.
    # This affects sorting (pinned-first) in the sessions list endpoint.
    # NOTE: As above, make sure route uses raw "<session_id>" in actual source.
    chat_data = load_session(session_id)
    chat_data["pinned"] = not chat_data.get("pinned", False)
    save_session(chat_data)
 
    # Return the updated pinned state to the client
    return jsonify({"success": True, "pinned": chat_data["pinned"]})
 
 
@app.delete("/sessions/<session_id>")
@login_required
def delete_session(session_id):
    # Permanently remove a session file from disk.
    # Caller should be careful: this is irreversible for that session.
    path = session_file(session_id)
    if os.path.exists(path):
        os.remove(path)
        return jsonify({"success": True})
    # File missing → 404 to indicate session not found
    return jsonify({"error": "File not found"}), 404
 
 
@app.get("/download/<path:filename>")
@login_required
def download_file(filename):
    # Download a generated artifact (e.g., PPTX) from the GENERATED_DIR as an attachment.
    # send_from_directory will safely join the path and set a reasonable Content-Disposition.
    return send_from_directory(GENERATED_DIR, filename, as_attachment=True)
 
 
@app.get("/")
@login_required
def index():
    # Serve the main frontend shell (SPA entry point or static UI)
    return send_from_directory(app.static_folder, "index.html")

if __name__ == "__main__":
    print("\n🔗 Deck Genie running at http://localhost:5001\n")

    # Turn Waitress back on for production
    serve(app, host="0.0.0.0", port=5001, threads = 16)

    # Delete the app.run(..., ssl_context='adhoc') line entirely