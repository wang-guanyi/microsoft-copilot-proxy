from __future__ import annotations

import argparse
import asyncio
import logging
import msvcrt
import re
import threading
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import unquote

import uvicorn
from playwright.async_api import async_playwright

from .app import create_app

# M365_COPILOT_URL = "https://m365.cloud.microsoft/chat"
M365_COPILOT_URL = "https://copilot.com/chat"
SUBSTRATE_WS_PATTERN = re.compile(
    r"^wss://substrate\.office\.com/m365Copilot/Chathub/([^/?@]+)@([^/?]+)"
)
DEFAULT_CHROMIUM_USER_DATA_DIR = (
    Path.home() / ".m365-copilot-openai-proxy" / "playwright-chromium-profile"
)
_capture_lock = threading.Lock()


class _SuppressCtrlC(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        return "CTRL+C" not in record.getMessage()


logging.getLogger("uvicorn.error").addFilter(_SuppressCtrlC())


def _looks_like_compact_token(token: str) -> bool:
    if not token or any(character.isspace() for character in token):
        return False
    parts = token.split(".")
    return len(parts) in {3, 5} and bool(parts[0]) and all(
        re.fullmatch(r"[A-Za-z0-9_-]*", part) is not None for part in parts
    )


async def _nudge_chat_input(page) -> bool:
    """Trigger Copilot's lazy ChatHub connection without sending a message."""
    await page.wait_for_timeout(5000)

    candidates = [
        page.get_by_role("textbox"),
        page.locator("textarea"),
        page.locator('[contenteditable="true"]'),
    ]

    for candidate in candidates:
        try:
            locator = candidate.last

            await locator.wait_for(
                state="visible",
                timeout=10000,
            )

            await locator.scroll_into_view_if_needed()
            await locator.focus()

            await locator.type("x", delay=50)
            await locator.press("Backspace")

            print("Successfully nudged Copilot composer")
            return True

        except Exception as exc:
            print(f"Nudge candidate failed: {exc}")

    return False


async def _capture_substrate_auth(
    user_data_dir: Path,
    timeout_seconds: int,
    *,
    headless: bool,
) -> tuple[str, str, str] | None:
    user_data_dir = user_data_dir.expanduser().resolve()
    user_data_dir.mkdir(parents=True, exist_ok=True)
    captured: tuple[str, str, str] | None = None
    captured_event = asyncio.Event()
    attached_pages: set[int] = set()

    async with async_playwright() as playwright:
        context = await playwright.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            headless=headless,
        )
        try:
            def on_websocket(websocket) -> None:
                nonlocal captured
                if captured is not None:
                    return
                route_match = SUBSTRATE_WS_PATTERN.match(websocket.url)
                token_match = re.search(r"[?&]access_token=([^&]+)", websocket.url)
                if not route_match or not token_match:
                    return
                token = unquote(token_match.group(1))
                if not _looks_like_compact_token(token):
                    return
                captured = (
                    token,
                    unquote(route_match.group(1)),
                    unquote(route_match.group(2)),
                )
                captured_event.set()

            def attach_to_page(page) -> None:
                if id(page) in attached_pages:
                    return
                attached_pages.add(id(page))
                page.on("websocket", on_websocket)

            for existing_page in context.pages:
                attach_to_page(existing_page)
            context.on("page", attach_to_page)
            page = context.pages[0] if context.pages else await context.new_page()
            attach_to_page(page)
            await page.goto(M365_COPILOT_URL, wait_until="domcontentloaded")
            # Copilot creates the Substrate ChatHub WebSocket lazily after the
            # composer receives keyboard input. Reproduce that interaction so
            # token capture no longer requires a manual keystroke.
            nudged = await _nudge_chat_input(page)
            if not nudged:
                print(
                    "Copilot chat composer was not found; waiting for manual input "
                    "or sign-in."
                )
            try:
                await asyncio.wait_for(captured_event.wait(), timeout_seconds)
            except asyncio.TimeoutError:
                return None
            return captured
        finally:
            await context.close()


def _read_env_value(name: str) -> str | None:
    path = Path(".env")
    if not path.exists():
        return None
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        key, separator, value = stripped.partition("=")
        if separator and key.strip() == name:
            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
                return value[1:-1]
            return value
    return None


def _write_auth_data(
    token: str,
    route_user_id: str,
    route_tenant_id: str,
    refresh_interval_minutes: int = 50,
) -> None:
    path = Path(".env")
    text = path.read_text(encoding="utf-8") if path.exists() else ""
    captured_at = datetime.now(timezone.utc)
    refresh_due_at = captured_at + timedelta(minutes=refresh_interval_minutes)
    values = {
        "M365_ACCESS_TOKEN": token,
        "M365_ROUTE_USER_ID": route_user_id,
        "M365_ROUTE_TENANT_ID": route_tenant_id,
        "M365_TOKEN_CAPTURED_AT": captured_at.isoformat(),
        "M365_TOKEN_REFRESH_DUE_AT": refresh_due_at.isoformat(),
    }
    for key, value in values.items():
        pattern = rf"(?m)^{re.escape(key)}=.*$"
        line = f"{key}={value}"
        if re.search(pattern, text):
            text = re.sub(pattern, lambda _match, replacement=line: replacement, text)
        else:
            if text and not text.endswith("\n"):
                text += "\n"
            text += line + "\n"
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(text, encoding="utf-8")
    temporary.replace(path)


def _capture_and_store(
    user_data_dir: Path,
    timeout_seconds: int,
    *,
    headless: bool,
    refresh_interval_minutes: int = 50,
) -> bool:
    if not _capture_lock.acquire(blocking=False):
        print("Token capture is already in progress.")
        return False
    try:
        try:
            captured = asyncio.run(
                _capture_substrate_auth(
                    user_data_dir,
                    timeout_seconds,
                    headless=headless,
                )
            )
        except Exception as exc:
            print(f"Playwright Chromium capture failed: {exc}")
            return False
        if not captured:
            return False
        _write_auth_data(*captured, refresh_interval_minutes)
        print("Substrate token, routing identifiers, and refresh metadata captured; .env updated.")
        return True
    finally:
        _capture_lock.release()


def _auto_refresh_loop(
    user_data_dir: Path,
    timeout_seconds: int,
    refresh_interval_seconds: int,
    retry_seconds: int,
    stop_event: threading.Event,
    restart_event: threading.Event,
) -> None:
    interval_minutes = max(1, round(refresh_interval_seconds / 60))
    while not stop_event.wait(refresh_interval_seconds):
        print("Scheduled Substrate token capture starting...")
        while not stop_event.is_set():
            if _capture_and_store(
                user_data_dir,
                timeout_seconds,
                headless=True,
                refresh_interval_minutes=interval_minutes,
            ):
                restart_event.set()
                return
            print(f"Scheduled token capture failed; retrying in {retry_seconds} seconds.")
            if stop_event.wait(retry_seconds):
                return


def set_token_command(args: argparse.Namespace) -> None:
    print("Paste the complete Substrate Chathub WebSocket URL, then press Enter:")
    raw = input().strip()
    route_match = SUBSTRATE_WS_PATTERN.match(raw)
    token_match = re.search(r"[?&]access_token=([^&\s]+)", raw)
    if not route_match or not token_match:
        print("Error: complete Chathub WebSocket URL required.")
        return
    token = unquote(token_match.group(1))
    if not _looks_like_compact_token(token):
        print("Error: URL does not contain a compact JWT or JWE token.")
        return
    _write_auth_data(
        token,
        unquote(route_match.group(1)),
        unquote(route_match.group(2)),
        getattr(args, "refresh_interval_minutes", 50),
    )
    print(".env updated with token, routing identifiers, and refresh metadata.")


def capture_token_command(args: argparse.Namespace) -> None:
    print(f"Launching {'visible' if args.headed else 'hidden'} Playwright Chromium...")
    if not _capture_and_store(
        args.chromium_user_data_dir,
        args.timeout_seconds,
        headless=not args.headed,
        refresh_interval_minutes=args.refresh_interval_minutes,
    ):
        print("Error: no Substrate Chathub token captured before timeout.")


def serve_command(args: argparse.Namespace) -> None:
    profile = args.chromium_user_data_dir

    try:
        if not args.no_capture_on_start:
            print("Capturing fresh Substrate token with Playwright Chromium...")
            if not _capture_and_store(
                profile,
                args.playwright_timeout_seconds,
                headless=True,
                refresh_interval_minutes=args.refresh_interval_minutes,
            ):
                complete = all(_read_env_value(name) for name in (
                    "M365_ACCESS_TOKEN", "M365_ROUTE_USER_ID", "M365_ROUTE_TENANT_ID"
                ))
                print(
                    "Startup capture failed; continuing with stored authentication data."
                    if complete else
                    "Startup capture failed and .env is incomplete. Run capture-token --headed."
                )

        while True:
            server = uvicorn.Server(
                uvicorn.Config(create_app(), host=args.host, port=args.port)
            )

            stop_event = threading.Event()
            restart_event = threading.Event()
            refresh_thread = None

            server_thread = threading.Thread(
                target=server.run,
                daemon=True,
            )
            server_thread.start()

            if not args.no_auto_refresh:
                refresh_thread = threading.Thread(
                    target=_auto_refresh_loop,
                    args=(
                        profile,
                        args.playwright_timeout_seconds,
                        args.refresh_interval_minutes * 60,
                        args.refresh_retry_seconds,
                        stop_event,
                        restart_event,
                    ),
                    daemon=True,
                )
                refresh_thread.start()

            while not server.started and server_thread.is_alive():
                time.sleep(0.05)

            print(
                f"\n  [q] quit    [r] capture token    "
                f"auto-refresh: {'off' if args.no_auto_refresh else 'on'}\n"
            )

            action = None

            while server_thread.is_alive():
                if restart_event.is_set():
                    action = "restart"
                    server.should_exit = True
                    break

                if msvcrt.kbhit():
                    key = msvcrt.getwch().lower()
                    if key in {"q", "r"}:
                        action = "quit" if key == "q" else "refresh"
                        server.should_exit = True
                        break

                time.sleep(0.05)

            stop_event.set()
            server_thread.join()

            if refresh_thread:
                refresh_thread.join(timeout=1)

            if action == "refresh":
                if not _capture_and_store(
                    profile,
                    args.playwright_timeout_seconds,
                    headless=True,
                    refresh_interval_minutes=args.refresh_interval_minutes,
                ):
                    set_token_command(args)
                print("Restarting server...")
                continue

            if action == "restart":
                print("Token refreshed. Restarting server...")
                continue

            break

    except KeyboardInterrupt:
        print("\nCTRL+C received, shutting down...")
        stop_event.set()

        try:
            server.should_exit = True
        except Exception:
            pass

        try:
            server_thread.join(timeout=5)
        except Exception:
            pass


def _add_profile_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--chromium-user-data-dir",
        type=Path,
        default=DEFAULT_CHROMIUM_USER_DATA_DIR,
    )


def main() -> None:
    parser = argparse.ArgumentParser(prog="copilot-openai-proxy")
    subparsers = parser.add_subparsers(dest="command", required=True)
    set_token = subparsers.add_parser("set-token")
    set_token.add_argument("--refresh-interval-minutes", type=int, default=50)
    set_token.set_defaults(func=set_token_command)
    capture = subparsers.add_parser("capture-token")
    _add_profile_argument(capture)
    capture.add_argument("--timeout-seconds", type=int, default=600)
    capture.add_argument("--refresh-interval-minutes", type=int, default=50)
    capture.add_argument("--headed", action="store_true")
    capture.set_defaults(func=capture_token_command)
    serve = subparsers.add_parser("serve")
    serve.add_argument("--host", default="127.0.0.1")
    serve.add_argument("--port", type=int, default=9000)
    _add_profile_argument(serve)
    serve.add_argument("--no-auto-refresh", action="store_true")
    serve.add_argument("--no-capture-on-start", action="store_true")
    serve.add_argument("--playwright-timeout-seconds", type=int, default=60)
    serve.add_argument("--refresh-interval-minutes", type=int, default=50)
    serve.add_argument("--refresh-retry-seconds", type=int, default=60)
    serve.set_defaults(func=serve_command)
    args = parser.parse_args()

    try:
        args.func(args)
    except KeyboardInterrupt:
        print("\nExiting...")


if __name__ == "__main__":
    main()
