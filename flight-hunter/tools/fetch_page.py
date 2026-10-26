"""Render a flight-search page in a headless browser and print its visible text.

Usage:
    python tools/fetch_page.py <url> [--wait 12] [--max-chars 40000]

Output starts with a status line:
    STATUS: OK | BLOCKED | FETCH_FAILED
followed by the final URL, page title and the visible page text.
"""
import argparse
import os
import sys

from playwright.sync_api import sync_playwright

BLOCK_MARKERS = [
    "captcha",
    "are you a robot",
    "are you a person",
    "unusual traffic",
    "verify you are human",
    "access denied",
    "press & hold",
    "אנחנו רוצים לוודא",
]

CONSENT_BUTTONS = [
    "Accept all",
    "I agree",
    "Accept",
    "קבל הכול",
    "אישור",
    "OK",
]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("url")
    parser.add_argument("--wait", type=int, default=12, help="seconds to let results load")
    parser.add_argument("--max-chars", type=int, default=40000)
    args = parser.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    launch_kwargs = {"headless": True}
    # Optional override, e.g. when a preinstalled Chromium must be used.
    if os.environ.get("FLIGHT_HUNTER_CHROMIUM"):
        launch_kwargs["executable_path"] = os.environ["FLIGHT_HUNTER_CHROMIUM"]

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(**launch_kwargs)
            context = browser.new_context(
                locale="en-US",
                timezone_id="Asia/Jerusalem",
                viewport={"width": 1400, "height": 1000},
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
                ),
            )
            page = context.new_page()
            page.goto(args.url, wait_until="domcontentloaded", timeout=60000)

            for label in CONSENT_BUTTONS:
                try:
                    button = page.get_by_role("button", name=label, exact=True)
                    if button.count() > 0:
                        button.first.click(timeout=3000)
                        break
                except Exception:
                    pass

            page.wait_for_timeout(args.wait * 1000)
            text = page.inner_text("body")
            title = page.title()
            final_url = page.url
            browser.close()
    except Exception as exc:  # noqa: BLE001 - report any browser failure to the agent
        print(f"STATUS: FETCH_FAILED\nREASON: {exc}")
        return 2

    lowered = (title + "\n" + text[:3000]).lower()
    status = "BLOCKED" if any(m in lowered for m in BLOCK_MARKERS) else "OK"

    print(f"STATUS: {status}")
    print(f"FINAL_URL: {final_url}")
    print(f"TITLE: {title}")
    print("----- PAGE TEXT -----")
    print(text[: args.max_chars])
    return 0


if __name__ == "__main__":
    sys.exit(main())
