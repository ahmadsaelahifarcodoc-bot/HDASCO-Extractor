"""Minimal Chrome DevTools Protocol client; intentionally does not use Selenium."""
import json
import os
import subprocess
import time
import urllib.request
import urllib.error
from pathlib import Path

DEBUG_ADDRESS = "http://127.0.0.1:9222"


def launch_chrome(profile_dir):
    profile_dir = Path(profile_dir)
    profile_dir.mkdir(parents=True, exist_ok=True)
    candidates = [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ]
    chrome = next((p for p in candidates if p and os.path.isfile(p)), None)
    if not chrome:
        raise FileNotFoundError("Google Chrome was not found. Install Chrome, then try again.")
    # Chrome's remote debugging is restricted when using the default profile.
    # This is a separate, dedicated profile created only for this application.
    subprocess.Popen([
        chrome,
        "--remote-debugging-port=9222",
        "--remote-allow-origins=*",
        "--no-first-run",
        "--no-default-browser-check",
        f"--user-data-dir={profile_dir}",
        "https://my.hdasco.com/",
    ], close_fds=True)
    for _ in range(30):
        try:
            with urllib.request.urlopen(DEBUG_ADDRESS + "/json/version", timeout=1) as r:
                if r.status == 200:
                    return
        except Exception:
            time.sleep(1)
    raise RuntimeError("Chrome started but its debugging endpoint did not become available.")


class ChromeCDP:
    def __init__(self, ws):
        self.ws = ws
        self.seq = 0
        self.target_id = None

    @classmethod
    def connect(cls):
        try:
            import websocket
        except ImportError:
            raise RuntimeError("Missing package websocket-client. Run: pip install -r requirements.txt")
        with urllib.request.urlopen(DEBUG_ADDRESS + "/json", timeout=4) as r:
            targets = json.loads(r.read().decode("utf-8"))
        pages = [t for t in targets if t.get("type") == "page" and t.get("webSocketDebuggerUrl")]
        if not pages:
            raise RuntimeError("No Chrome page is available. Open Chrome first.")
        # Prefer the current HDASCO page if one exists.
        target = next((t for t in pages if "hdasco.com" in t.get("url", "").lower()), pages[0])
        ws = websocket.create_connection(target["webSocketDebuggerUrl"], timeout=10,
                                         origin="http://localhost")
        ws.settimeout(15)
        obj = cls(ws)
        obj.target_id = target.get("id")
        obj.call("Page.enable")
        obj.call("Runtime.enable")
        return obj

    def call(self, method, params=None, timeout=15):
        self.seq += 1
        ident = self.seq
        self.ws.send(json.dumps({"id": ident, "method": method, "params": params or {}}))
        deadline = time.time() + timeout
        while time.time() < deadline:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == ident:
                if "error" in msg:
                    raise RuntimeError(f"Chrome DevTools error ({method}): {msg['error']}")
                return msg.get("result", {})
        raise TimeoutError("Timed out waiting for Chrome response: " + method)

    def evaluate(self, expression, await_promise=False, timeout=30):
        result = self.call("Runtime.evaluate", {
            "expression": expression,
            "returnByValue": True,
            "awaitPromise": await_promise,
            "userGesture": True
        }, timeout=timeout)
        if result.get("exceptionDetails"):
            detail = result["exceptionDetails"].get("text", "JavaScript evaluation failed")
            raise RuntimeError(detail)
        remote = result.get("result", {})
        if remote.get("subtype") == "error":
            raise RuntimeError(remote.get("description", "JavaScript error"))
        return remote.get("value")

    def navigate(self, url):
        self.call("Page.navigate", {"url": url})
        time.sleep(2)

    def current_url(self):
        return self.evaluate("location.href") or ""

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass
