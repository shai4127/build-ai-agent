"""Send a WhatsApp message to the owner through CallMeBot.

Usage:
    python tools/notify_whatsapp.py "message text"
    python tools/notify_whatsapp.py --check     (verify .env is filled, sends nothing)

Reads CALLMEBOT_PHONE and CALLMEBOT_APIKEY from the .env file in the agent folder.
The message can only reach the phone number registered with CallMeBot.
"""
import sys
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENV = ROOT / ".env"


def load_env() -> dict:
    values = {}
    if ENV.exists():
        for line in ENV.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__)
        return 1

    env = load_env()
    phone = env.get("CALLMEBOT_PHONE", "")
    apikey = env.get("CALLMEBOT_APIKEY", "")
    if not phone or not apikey or "X" in phone or apikey.lower().startswith("your"):
        print("WHATSAPP: SETUP_REQUIRED - fill CALLMEBOT_PHONE and CALLMEBOT_APIKEY in .env")
        return 3
    if sys.argv[1] == "--check":
        print("WHATSAPP: CONFIGURED (not yet verified by a real message)")
        return 0

    query = urllib.parse.urlencode({"phone": phone, "text": sys.argv[1], "apikey": apikey})
    url = f"https://api.callmebot.com/whatsapp.php?{query}"
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            body = resp.read().decode("utf-8", errors="replace")
    except Exception as exc:  # noqa: BLE001
        print(f"WHATSAPP: FAILED - {exc}")
        return 2
    ok = "queued" in body.lower() or "sent" in body.lower()
    print("WHATSAPP: SENT" if ok else f"WHATSAPP: FAILED - {body[:300]}")
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
