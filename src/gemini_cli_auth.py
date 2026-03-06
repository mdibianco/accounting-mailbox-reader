"""
gemini_cli_auth.py — Reuse Gemini CLI's OAuth credentials to call Gemini from Python.

No API key needed. No GCP project needed. Uses the same free tier as Gemini CLI.
(60 req/min, 1000 req/day with personal Google account)

Prerequisites:
  1. Install Gemini CLI:  npm install -g @google/gemini-cli
  2. Authenticate once:   gemini  (select "Login with Google", follow browser flow)
  3. Verify creds exist:  cat ~/.gemini/oauth_creds.json
  4. pip install requests

How it works:
  - Reads the cached OAuth refresh token from ~/.gemini/oauth_creds.json
  - Uses Gemini CLI's public OAuth client ID to refresh the access token
  - Calls Google's Code Assist API (same endpoint Gemini CLI uses)
  - No API key, no client_secret.json, no GCP project required

⚠️  This uses an undocumented internal API (Code Assist / cloudcode-pa.googleapis.com).
    It works today (Feb 2026) but could break if Google changes the endpoint.
    For production use, consider Vertex AI with gcloud ADC or a free AI Studio API key.
"""

import json
import os
import sys
import requests
from pathlib import Path


# ---------------------------------------------------------------------------
# Constants — extracted from Gemini CLI source & Cline/Roo Code implementations
# ---------------------------------------------------------------------------

# Gemini CLI's public OAuth client — set in .env or use Gemini CLI defaults
OAUTH_CLIENT_ID = os.environ.get("GEMINI_OAUTH_CLIENT_ID", "")
OAUTH_CLIENT_SECRET = os.environ.get("GEMINI_OAUTH_CLIENT_SECRET", "")
OAUTH_TOKEN_URL = "https://oauth2.googleapis.com/token"

# Code Assist API — the endpoint Gemini CLI actually talks to
CODE_ASSIST_ENDPOINT = "https://cloudcode-pa.googleapis.com"
CODE_ASSIST_API_VERSION = "v1internal"

# Where Gemini CLI caches its OAuth credentials
GEMINI_CREDS_PATH = Path.home() / ".gemini" / "oauth_creds.json"


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def load_gemini_cli_creds() -> dict:
    """Load cached OAuth credentials from Gemini CLI."""
    if not GEMINI_CREDS_PATH.exists():
        print(f"❌ No Gemini CLI credentials found at {GEMINI_CREDS_PATH}")
        print()
        print("To fix this:")
        print("  1. Install Gemini CLI:  npm install -g @google/gemini-cli")
        print('  2. Run:                 gemini')
        print('  3. Select:              "Login with Google"')
        print("  4. Complete the browser auth flow")
        print()
        print(f"Then check that {GEMINI_CREDS_PATH} exists.")
        sys.exit(1)

    with open(GEMINI_CREDS_PATH, "r") as f:
        creds = json.load(f)

    if "refresh_token" not in creds:
        print(f"❌ No refresh_token found in {GEMINI_CREDS_PATH}")
        print("Try re-authenticating: run 'gemini' and login again.")
        sys.exit(1)

    return creds


def refresh_access_token(refresh_token: str) -> str:
    """Exchange refresh token for a fresh access token."""
    resp = requests.post(OAUTH_TOKEN_URL, data={
        "grant_type": "refresh_token",
        "client_id": OAUTH_CLIENT_ID,
        "client_secret": OAUTH_CLIENT_SECRET,
        "refresh_token": refresh_token,
    })

    if resp.status_code != 200:
        print(f"❌ Token refresh failed ({resp.status_code}): {resp.text}")
        print("Try re-authenticating: run 'gemini' and login again.")
        sys.exit(1)

    data = resp.json()
    return data["access_token"]


def get_access_token() -> str:
    """Get a valid access token from Gemini CLI's cached credentials."""
    creds = load_gemini_cli_creds()

    # Check if existing access token is still valid
    # (expiry_date is in milliseconds since epoch)
    import time
    expiry = creds.get("expiry_date", 0)
    if expiry > 0 and expiry / 1000 > time.time() + 60:
        # Token still valid (with 60s buffer)
        return creds["access_token"]

    # Token expired or missing — refresh it
    print("🔄 Refreshing access token...")
    access_token = refresh_access_token(creds["refresh_token"])

    # Update cached credentials
    creds["access_token"] = access_token
    creds["expiry_date"] = int((time.time() + 3600) * 1000)
    with open(GEMINI_CREDS_PATH, "w") as f:
        json.dump(creds, f, indent=2)

    return access_token


# ---------------------------------------------------------------------------
# Gemini API call via Code Assist endpoint
# ---------------------------------------------------------------------------

def call_gemini(
    prompt: str,
    system_instruction: str = None,
    model: str = "gemini-2.0-flash",
    temperature: float = 0.7,
    max_tokens: int = 8192,
    json_output: bool = False,
) -> str:
    """
    Call Gemini via the Code Assist API using Gemini CLI's OAuth credentials.

    Args:
        prompt: The text prompt to send
        system_instruction: Optional system instruction (base prompt)
        model: Model name (gemini-2.0-flash, gemini-2.5-pro, etc.)
        temperature: Sampling temperature (0.0 - 2.0)
        max_tokens: Maximum output tokens
        json_output: If True, request JSON response format

    Returns:
        The model's text response
    """
    # Check for API key first (simplest), fall back to OAuth
    api_key = os.environ.get("GEMINI_API_KEY", "")

    if api_key:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}
    else:
        access_token = get_access_token()
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

    payload = {
        "contents": [{
            "role": "user",
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_tokens,
        }
    }

    # Add system instruction if provided
    if system_instruction:
        payload["systemInstruction"] = {
            "parts": [{"text": system_instruction}]
        }

    # Request JSON output format
    if json_output:
        payload["generationConfig"]["responseMimeType"] = "application/json"

    resp = requests.post(url, headers=headers, json=payload)

    if resp.status_code != 200:
        error_msg = resp.text
        try:
            error_data = resp.json()
            error_msg = error_data.get("error", {}).get("message", resp.text)
        except Exception:
            pass
        raise RuntimeError(f"Gemini API error ({resp.status_code}): {error_msg}")

    # Count successful API call
    try:
        from src import api_counter
        total = api_counter.increment(model)
        _cascade_logger.info(f"API call #{total} today (model: {model})")
    except ImportError:
        pass  # Standalone usage outside src package

    data = resp.json()

    # Extract text from response
    try:
        candidates = data["candidates"]
        parts = candidates[0]["content"]["parts"]
        return "".join(part.get("text", "") for part in parts)
    except (KeyError, IndexError) as e:
        raise RuntimeError(f"Unexpected response format: {e}\n{json.dumps(data, indent=2)}")


# ---------------------------------------------------------------------------
# Model cascade — automatic fallback on rate limits (429)
# ---------------------------------------------------------------------------

import logging as _logging

_cascade_logger = _logging.getLogger(__name__)

# Ordered from most capable to least. Each has its own RPD quota.
MODEL_CASCADE = [
    "gemini-2.5-flash",         # 20 RPD  — best quality
    "gemini-3-flash-preview",   # 20 RPD  — new Gemini 3
    "gemini-2.5-flash-lite",    # 20 RPD  — lighter but still Gemini
    "gemma-3-27b-it",           # 14.4K RPD — good open model, last resort
]

# Models that don't support systemInstruction or responseMimeType=application/json
_GEMMA_MODELS = {m for m in MODEL_CASCADE if "gemma" in m}


def call_gemini_cascade(
    prompt: str,
    system_instruction: str = None,
    temperature: float = 0.7,
    max_tokens: int = 8192,
    json_output: bool = False,
    preferred_model: str = None,
) -> tuple:
    """
    Try models in cascade order, falling back on 429 (rate limit) or 404 errors.

    Args:
        prompt: User prompt text
        system_instruction: System prompt (embedded in user prompt for Gemma models)
        temperature: Sampling temperature
        max_tokens: Max output tokens
        json_output: Request JSON mode (skipped for Gemma models)
        preferred_model: Start cascade from this model (default: first in list)

    Returns:
        Tuple of (response_text, model_used)
    """
    cascade = list(MODEL_CASCADE)
    if preferred_model and preferred_model in cascade:
        idx = cascade.index(preferred_model)
        cascade = cascade[idx:] + cascade[:idx]

    last_error = None
    for model in cascade:
        is_gemma = model in _GEMMA_MODELS
        try:
            if is_gemma:
                # Gemma: no system instruction support, no JSON mode
                combined = f"{system_instruction}\n\n---\n\n{prompt}" if system_instruction else prompt
                result = call_gemini(
                    prompt=combined,
                    model=model,
                    temperature=temperature,
                    max_tokens=max_tokens,
                    json_output=False,
                )
            else:
                result = call_gemini(
                    prompt=prompt,
                    system_instruction=system_instruction,
                    model=model,
                    temperature=temperature,
                    max_tokens=max_tokens,
                    json_output=json_output,
                )
            return result, model

        except RuntimeError as e:
            error_str = str(e)
            if "429" in error_str or "404" in error_str or "503" in error_str:
                _cascade_logger.warning(f"{model} unavailable ({error_str[:80]}), trying next...")
                last_error = e
                continue
            raise  # Other errors (400 bad request, 500 server) — don't cascade

    raise RuntimeError(f"All models in cascade exhausted. Last error: {last_error}")


# ---------------------------------------------------------------------------
# Convenience: Outlook email processing
# ---------------------------------------------------------------------------

def fetch_outlook_emails(
    username: str,
    app_password: str,
    folder: str = "INBOX",
    limit: int = 10,
) -> list[dict]:
    """
    Fetch recent emails from Outlook via IMAP.

    Args:
        username: Your Outlook email address
        app_password: App password (not your regular password)
        folder: IMAP folder to read from
        limit: Number of recent emails to fetch

    Returns:
        List of dicts with subject, from, date, body
    """
    import imaplib
    import email
    from email.header import decode_header

    mail = imaplib.IMAP4_SSL("outlook.office365.com")
    mail.login(username, app_password)
    mail.select(folder)

    status, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()

    emails = []
    for eid in email_ids[-limit:]:
        status, msg_data = mail.fetch(eid, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Decode subject
        subject, encoding = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8")

        # Extract body
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                    break
        else:
            body = msg.get_payload(decode=True).decode("utf-8", errors="replace")

        emails.append({
            "subject": subject,
            "from": msg["From"],
            "date": msg["Date"],
            "body": body[:3000],  # Truncate long emails
        })

    mail.logout()
    return emails


# ---------------------------------------------------------------------------
# Main — demo usage
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Call Gemini using Gemini CLI's OAuth credentials (no API key needed)"
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # --- "ask" command: simple prompt ---
    ask_parser = subparsers.add_parser("ask", help="Ask Gemini a question")
    ask_parser.add_argument("prompt", help="The prompt to send")
    ask_parser.add_argument("--model", default="gemini-2.0-flash", help="Model name")

    # --- "check-auth" command: verify setup ---
    subparsers.add_parser("check-auth", help="Verify Gemini CLI auth is working")

    # --- "process-email" command: full pipeline ---
    email_parser = subparsers.add_parser("process-email", help="Fetch & process Outlook emails")
    email_parser.add_argument("--user", required=True, help="Outlook email address")
    email_parser.add_argument("--password", required=True, help="Outlook app password")
    email_parser.add_argument("--limit", type=int, default=5, help="Number of emails to process")
    email_parser.add_argument("--model", default="gemini-2.0-flash", help="Model name")

    args = parser.parse_args()

    if args.command == "check-auth":
        print("🔍 Checking Gemini CLI credentials...")
        creds = load_gemini_cli_creds()
        print(f"✅ Found credentials at {GEMINI_CREDS_PATH}")
        print(f"   Scopes: {creds.get('scope', 'unknown')}")

        print("\n🔄 Testing token refresh...")
        token = get_access_token()
        print(f"✅ Got access token: {token[:20]}...")

        print("\n🤖 Testing Gemini API call...")
        response = call_gemini("Say 'Hello from Gemini CLI OAuth!' in one sentence.")
        print(f"✅ Response: {response.strip()}")

        print("\n🎉 Everything works! You can now use call_gemini() in your scripts.")

    elif args.command == "ask":
        response = call_gemini(args.prompt, model=args.model)
        print(response)

    elif args.command == "process-email":
        print(f"📧 Fetching last {args.limit} emails from {args.user}...")
        emails = fetch_outlook_emails(args.user, args.password, limit=args.limit)

        for i, mail in enumerate(emails, 1):
            print(f"\n{'='*60}")
            print(f"📨 Email {i}/{len(emails)}: {mail['subject']}")
            print(f"{'='*60}")

            prompt = f"""Analyze this email and provide:
1. A one-line summary
2. Action items (if any)
3. Priority (high/medium/low)
4. Suggested response (if needed)

Subject: {mail['subject']}
From: {mail['from']}
Date: {mail['date']}

Body:
{mail['body']}"""

            response = call_gemini(prompt, model=args.model)
            print(response)

    else:
        parser.print_help()
        print("\n💡 Quick start:")
        print("   python gemini_cli_auth.py check-auth")
        print('   python gemini_cli_auth.py ask "What is the capital of Switzerland?"')
