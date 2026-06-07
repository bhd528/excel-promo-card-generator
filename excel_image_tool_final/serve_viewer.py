from __future__ import annotations

import argparse
import json
import os
import posixpath
import sys
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlsplit

import generate_template_images as image_generator


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PROJECT_ROOT = app_root()
SERVER_INFO = PROJECT_ROOT / ".viewer_server.json"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8765
MAX_BODY_BYTES = 256 * 1024


def configure_generator_paths() -> None:
    image_generator.WORKBOOK_JSON = PROJECT_ROOT / "converted_excel" / "workbook_data.json"
    image_generator.WORKBOOK_JS = PROJECT_ROOT / "converted_excel" / "workbook_data.js"
    image_generator.TEMPLATE_DIR = PROJECT_ROOT / "image_templates"
    image_generator.GENERATED_DIR = PROJECT_ROOT / "generated_images"


def find_sheet(data: dict[str, Any], sheet_index: Any) -> dict[str, Any]:
    try:
        wanted_index = int(sheet_index)
    except (TypeError, ValueError) as exc:
        raise ValueError("缺少有效的 sheet_index") from exc

    for sheet in data.get("sheets", []):
        if int(sheet.get("sheet_index", -1)) == wanted_index:
            return sheet
    raise ValueError(f"未找到 sheet_index={wanted_index} 的 sheet")


def find_template(template_name: Any) -> Path:
    name = str(template_name or "").strip()
    if not name:
        return image_generator.find_template(None)

    # The page sends only a file name. Keep it inside image_templates.
    candidate = image_generator.TEMPLATE_DIR / Path(name).name
    if not candidate.exists() or candidate.suffix.lower() not in image_generator.IMAGE_EXTENSIONS:
        raise FileNotFoundError(f"未找到模板图：{name}")
    return candidate


def generate_image(payload: dict[str, Any]) -> dict[str, Any]:
    data = image_generator.load_workbook_data()
    sheet = find_sheet(data, payload.get("sheet_index"))
    template_path = find_template(payload.get("template_name"))

    title = image_generator.get_fields(sheet).get("title") or sheet.get("sheet_name") or "sheet"
    sheet_index = int(sheet["sheet_index"])
    filename = f"{sheet_index:03d}_{image_generator.sanitize_filename(title, 'sheet')}.png"
    output_path = image_generator.GENERATED_DIR / filename

    adjustments = payload.get("adjustments")
    if adjustments is not None and not isinstance(adjustments, dict):
        raise ValueError("adjustments 必须是 JSON 对象")

    for old_file in image_generator.GENERATED_DIR.glob(f"{sheet_index:03d}_*.png"):
        if old_file.resolve() != output_path.resolve():
            old_file.unlink()

    image_generator.render_sheet(template_path, sheet, output_path, adjustments=adjustments)
    image_generator.refresh_workbook_assets(data)

    generated = None
    for asset in data.get("assets", {}).get("generated_images", []):
        if asset.get("name") == output_path.name:
            generated = asset
            break

    return {
        "ok": True,
        "message": f"已生成：{output_path.name}",
        "generated": generated,
        "assets": data.get("assets", {}),
        "summary": data.get("summary", {}),
    }


class ViewerHandler(SimpleHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def translate_path(self, path: str) -> str:
        route = urlsplit(path).path
        route = posixpath.normpath(unquote(route))
        parts = [part for part in route.split("/") if part and part not in {".", ".."}]
        candidate = (PROJECT_ROOT / Path(*parts)).resolve() if parts else PROJECT_ROOT.resolve()
        try:
            candidate.relative_to(PROJECT_ROOT.resolve())
        except ValueError:
            return str(PROJECT_ROOT / "__missing__")
        return str(candidate)

    def send_json(self, status_code: int, payload: dict[str, Any]) -> None:
        raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self) -> None:
        route = urlsplit(self.path).path
        if route == "/api/status":
            self.send_json(
                200,
                {
                    "ok": True,
                    "url": f"http://{self.server.server_address[0]}:{self.server.server_address[1]}/converted_excel/viewer.html",
                },
            )
            return
        super().do_GET()

    def do_POST(self) -> None:
        route = urlsplit(self.path).path
        if route != "/api/generate-image":
            self.send_json(404, {"ok": False, "error": "接口不存在"})
            return

        try:
            length = int(self.headers.get("Content-Length") or "0")
            if length > MAX_BODY_BYTES:
                raise ValueError("请求体过大")
            raw = self.rfile.read(length)
            payload = json.loads(raw.decode("utf-8") or "{}") if raw else {}
            if not isinstance(payload, dict):
                raise ValueError("请求体必须是 JSON 对象")
            result = generate_image(payload)
        except (ValueError, json.JSONDecodeError) as exc:
            self.send_json(400, {"ok": False, "error": str(exc)})
            return
        except FileNotFoundError as exc:
            self.send_json(404, {"ok": False, "error": str(exc)})
            return
        except Exception as exc:
            self.send_json(500, {"ok": False, "error": str(exc)})
            return

        self.send_json(200, result)

    def log_message(self, format: str, *args: Any) -> None:
        print(f"{self.address_string()} - {format % args}")


def create_server(host: str, port: int, tries: int) -> tuple[ThreadingHTTPServer, int]:
    handler = partial(ViewerHandler, directory=str(PROJECT_ROOT))
    last_error: OSError | None = None
    for offset in range(max(1, tries)):
        current_port = port + offset
        try:
            return ThreadingHTTPServer((host, current_port), handler), current_port
        except OSError as exc:
            last_error = exc
    if last_error:
        raise last_error
    raise RuntimeError("无法启动服务")


def main() -> int:
    parser = argparse.ArgumentParser(description="Serve the Excel viewer and image generation API.")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", default=DEFAULT_PORT, type=int)
    parser.add_argument("--port-tries", default=20, type=int)
    args = parser.parse_args()

    configure_generator_paths()
    server, port = create_server(args.host, args.port, args.port_tries)
    display_host = "127.0.0.1" if args.host in {"", "0.0.0.0"} else args.host
    url = f"http://{display_host}:{port}/converted_excel/viewer.html"
    info = {
        "pid": os.getpid(),
        "host": args.host,
        "port": port,
        "url": url,
        "api": f"http://{display_host}:{port}/api/generate-image",
    }
    SERVER_INFO.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"viewer={url}")
    print(f"api={info['api']}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        return 0
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
