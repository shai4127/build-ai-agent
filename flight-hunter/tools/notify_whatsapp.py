"""Send a WhatsApp message to the owner through WhatsApp Web in a local browser.

Usage:
    python tools/notify_whatsapp.py --login        (one time: opens a browser window, scan the QR code)
    python tools/notify_whatsapp.py --check        (verify phone + logged-in session, sends nothing)
    python tools/notify_whatsapp.py "message text" [--headed]

No API key and no third-party service. The browser profile in .whatsapp-session/
keeps the WhatsApp Web login (a "linked device" on the owner's phone), and the
message goes to the owner's own chat ("Message yourself").

Reads WHATSAPP_PHONE from the .env file in the agent folder
(CALLMEBOT_PHONE is still accepted for older .env files).
"""
import os
import sys
import time
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENV = ROOT / ".env"
PROFILE = ROOT / ".whatsapp-session"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
)
COMPOSE = 'footer div[contenteditable="true"]'
LOGGED_IN = "#side, #pane-side"
QR_CODE = "canvas[aria-label], div[data-ref]"
INVALID_PHONE = "text=/phone number shared via url is invalid/i"
PENDING = 'span[data-icon="msg-time"]'
SEND_BUTTON = 'button[aria-label="Send"], span[data-icon="send"], span[data-icon="wds-ic-send-filled"]'


def load_env() -> dict:
    values = {}
    if ENV.exists():
        for line in ENV.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def owner_phone() -> str:
    env = load_env()
    phone = env.get("WHATSAPP_PHONE") or env.get("CALLMEBOT_PHONE") or ""
    phone = phone.replace("+", "").replace("-", "").replace(" ", "")
    return phone if phone.isdigit() and len(phone) >= 10 else ""


def open_browser(p, headed: bool):
    kwargs = {
        "headless": not headed,
        "user_agent": USER_AGENT,
        "viewport": {"width": 1200, "height": 900},
        "locale": "he-IL",
    }
    # Optional override, e.g. when a preinstalled Chromium must be used.
    if os.environ.get("FLIGHT_HUNTER_CHROMIUM"):
        kwargs["executable_path"] = os.environ["FLIGHT_HUNTER_CHROMIUM"]
    PROFILE.mkdir(exist_ok=True)
    context = p.chromium.launch_persistent_context(str(PROFILE), **kwargs)
    page = context.pages[0] if context.pages else context.new_page()
    return context, page


def wait_for_any(page, selectors: dict, timeout_s: int) -> str:
    """Return the name of the first selector that becomes visible, or 'timeout'."""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        for name, selector in selectors.items():
            try:
                if page.locator(selector).first.is_visible():
                    return name
            except Exception:
                pass
        page.wait_for_timeout(1000)
    return "timeout"


def login() -> int:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        context, page = open_browser(p, headed=True)
        page.goto("https://web.whatsapp.com/", wait_until="domcontentloaded", timeout=60000)
        print("A browser window is open. On your phone: WhatsApp > Settings > Linked devices > "
              "Link a device, then scan the QR code. Waiting up to 3 minutes...")
        state = wait_for_any(page, {"logged_in": LOGGED_IN}, 180)
        if state == "logged_in":
            page.wait_for_timeout(5000)  # let WhatsApp finish saving the session
        context.close()
    print("WHATSAPP: LOGGED_IN" if state == "logged_in" else "WHATSAPP: LOGIN_FAILED - QR code was not scanned in time")
    return 0 if state == "logged_in" else 2


def check(phone: str) -> int:
    if not PROFILE.exists():
        print("WHATSAPP: SETUP_REQUIRED - run: python tools/notify_whatsapp.py --login")
        return 3
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        context, page = open_browser(p, headed=False)
        page.goto("https://web.whatsapp.com/", wait_until="domcontentloaded", timeout=60000)
        state = wait_for_any(page, {"logged_in": LOGGED_IN, "qr": QR_CODE}, 60)
        context.close()
    if state == "logged_in":
        print(f"WHATSAPP: CONFIGURED - logged in, will send to ...{phone[-4:]} (not yet verified by a real message)")
        return 0
    print("WHATSAPP: SETUP_REQUIRED - session expired or not linked, run: python tools/notify_whatsapp.py --login")
    return 3


def send(phone: str, text: str, headed: bool) -> int:
    if not PROFILE.exists():
        print("WHATSAPP: SETUP_REQUIRED - run: python tools/notify_whatsapp.py --login")
        return 3
    from playwright.sync_api import sync_playwright

    url = "https://web.whatsapp.com/send?" + urllib.parse.urlencode({"phone": phone, "text": text})
    with sync_playwright() as p:
        context, page = open_browser(p, headed)
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            state = wait_for_any(
                page, {"chat": COMPOSE, "qr": QR_CODE, "invalid": INVALID_PHONE}, 90
            )
            if state == "qr":
                print("WHATSAPP: SETUP_REQUIRED - session expired, run: python tools/notify_whatsapp.py --login")
                return 3
            if state == "invalid":
                print("WHATSAPP: FAILED - WhatsApp says the phone number in .env is invalid")
                return 2
            if state != "chat":
                print("WHATSAPP: FAILED - WhatsApp Web did not open the chat within 90 seconds")
                return 2

            page.wait_for_timeout(2000)  # the text from the URL is filled in after the chat opens
            button = page.locator(SEND_BUTTON).first
            if button.is_visible():
                button.click()
            else:
                page.locator(COMPOSE).first.press("Enter")

            # Sent = the compose box is empty and no outgoing message shows the pending clock.
            deadline = time.time() + 30
            while time.time() < deadline:
                page.wait_for_timeout(1000)
                empty = not page.locator(COMPOSE).first.inner_text().strip()
                if empty and page.locator(PENDING).count() == 0:
                    page.wait_for_timeout(2000)
                    print("WHATSAPP: SENT")
                    return 0
            print("WHATSAPP: FAILED - message did not leave the outbox within 30 seconds")
            return 2
        except Exception as exc:  # noqa: BLE001 - report any browser failure to the agent
            print(f"WHATSAPP: FAILED - {exc}")
            return 2
        finally:
            context.close()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 1

    if args[0] == "--login":
        return login()

    phone = owner_phone()
    if not phone:
        print("WHATSAPP: SETUP_REQUIRED - fill WHATSAPP_PHONE in .env (e.g. 972501234567)")
        return 3
    if args[0] == "--check":
        return check(phone)

    headed = "--headed" in args
    text = next(a for a in args if a != "--headed")
    return send(phone, text, headed)


if __name__ == "__main__":
    sys.exit(main())
