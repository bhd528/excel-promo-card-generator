from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path

PORTS = set(range(8765, 8785))


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = app_root()
SERVER_INFO = ROOT / ".viewer_server.json"
SERVER_LOG = ROOT / ".viewer_server.log"
SERVE_SCRIPT = ROOT / "serve_viewer.py"
UPDATE_SCRIPT = ROOT / "update_excel_data.py"


def netstat_listener_pids() -> set[int]:
    try:
        output = subprocess.check_output(["netstat", "-ano"], text=True, encoding="mbcs", errors="ignore")
    except Exception:
        return set()

    pids: set[int] = set()
    for line in output.splitlines():
        parts = line.split()
        if len(parts) < 5 or parts[0].upper() != "TCP" or parts[-2].upper() != "LISTENING":
            continue
        local = parts[1]
        try:
            port = int(local.rsplit(":", 1)[-1])
            pid = int(parts[-1])
        except ValueError:
            continue
        if port in PORTS and pid and pid != os.getpid():
            pids.add(pid)
    return pids


def info_pid() -> int | None:
    if not SERVER_INFO.exists():
        return None
    try:
        payload = json.loads(SERVER_INFO.read_text(encoding="utf-8"))
        pid = int(payload.get("pid") or 0)
    except Exception:
        return None
    return pid or None


def stop_pid(pid: int) -> None:
    if pid == os.getpid():
        return
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def stop_viewer() -> None:
    pids = netstat_listener_pids()
    saved_pid = info_pid()
    if saved_pid:
        pids.add(saved_pid)
    for pid in sorted(pids):
        stop_pid(pid)
    if SERVER_INFO.exists():
        SERVER_INFO.unlink()


def run_update() -> None:
    import update_excel_data

    old_argv = sys.argv[:]
    try:
        sys.argv = [str(UPDATE_SCRIPT)]
        result_code = update_excel_data.main()
        if result_code != 0:
            raise SystemExit(result_code)
    finally:
        sys.argv = old_argv


def start_server() -> subprocess.Popen:
    if SERVER_INFO.exists():
        SERVER_INFO.unlink()
    flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
    log = SERVER_LOG.open("a", encoding="utf-8")
    command = [sys.executable, "--serve"] if getattr(sys, "frozen", False) else [sys.executable, str(SERVE_SCRIPT)]
    return subprocess.Popen(
        command,
        cwd=ROOT,
        stdout=log,
        stderr=log,
        stdin=subprocess.DEVNULL,
        creationflags=flags,
    )


def wait_for_url(process: subprocess.Popen, timeout: float = 12.0) -> str:
    deadline = time.time() + timeout
    last_error = ""
    while time.time() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"viewer 服务启动失败，退出码 {process.returncode}")
        if SERVER_INFO.exists():
            try:
                payload = json.loads(SERVER_INFO.read_text(encoding="utf-8"))
                url = str(payload.get("url") or "")
                if url:
                    with urllib.request.urlopen(url, timeout=1) as response:
                        if response.status == 200:
                            return url
            except Exception as exc:
                last_error = str(exc)
        time.sleep(0.3)
    raise RuntimeError(f"viewer 服务启动超时：{last_error}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Start/stop the local viewer and optionally refresh Excel data.")
    parser.add_argument("--update", action="store_true", help="Refresh converted_excel before starting the viewer.")
    parser.add_argument("--stop", action="store_true", help="Only stop the viewer service.")
    parser.add_argument("--no-browser", action="store_true", help="Start the viewer without opening a browser.")
    parser.add_argument("--serve", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.serve:
        import serve_viewer

        old_argv = sys.argv[:]
        try:
            sys.argv = [old_argv[0]]
            return serve_viewer.main()
        finally:
            sys.argv = old_argv

    if args.stop:
        stop_viewer()
        print("Viewer stopped.")
        return 0

    if args.update:
        run_update()

    stop_viewer()
    process = start_server()
    try:
        url = wait_for_url(process)
    except Exception as exc:
        print(f"Viewer start failed: {exc}", file=sys.stderr)
        print(f"See log: {SERVER_LOG}", file=sys.stderr)
        return 1

    print(f"Viewer: {url}")
    if not args.no_browser:
        webbrowser.open(url)
        print("Browser opened.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
