from __future__ import annotations

import argparse
import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont


WORKBOOK_JSON = Path("converted_excel") / "workbook_data.json"
WORKBOOK_JS = Path("converted_excel") / "workbook_data.js"
TEMPLATE_DIR = Path("image_templates")
GENERATED_DIR = Path("generated_images")
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}

FILL_RED = (230, 0, 18)
TITLE_COLOR = (90, 8, 20)

FIELD_BOXES = {
    "title": (190, 54, 486, 80),
    "spec": (190, 146, 560, 92),
    "page_price": (170, 255, 150, 44),
    "detail_coupon": (478, 248, 96, 44),
    "buy_count": (670, 248, 74, 44),
    "deposit": (170, 327, 176, 44),
    "tail_payment": (170, 394, 176, 44),
    "category_coupon": (190, 640, 486, 44),
    "vip_coupon": (190, 707, 486, 44),
    "other_discount": (190, 774, 486, 52),
    "final_price": (272, 842, 404, 182),
}

FIELD_STYLES = {
    "title": {"color": TITLE_COLOR, "size": 40, "min_size": 34, "bold": True},
    "spec": {"color": FILL_RED, "size": 32, "min_size": 18, "bold": True},
    "page_price": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "detail_coupon": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "buy_count": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "deposit": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "tail_payment": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "category_coupon": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "vip_coupon": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "other_discount": {"color": FILL_RED, "size": 40, "min_size": 34, "bold": True},
    "final_price": {
        "color": FILL_RED,
        "size": 40,
        "min_size": 34,
        "bold": True,
        "background": (255, 235, 59),
        "background_padding": (10, 6),
        "line_gap_ratio": 0.42,
    },
}


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        Path(r"C:\Windows\Fonts\msyhbd.ttc") if bold else Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
        Path(r"C:\Windows\Fonts\simsun.ttc"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def strip_invisible_text_marks(text: str) -> str:
    if not text:
        return ""
    output: list[str] = []
    for ch in text:
        category = unicodedata.category(ch)
        if category == "Cf":
            continue
        if category == "Cc" and ch not in "\r\n\t":
            continue
        output.append(ch)
    return "".join(output)


def text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> tuple[int, int]:
    if not text:
        return 0, 0
    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
    return right - left, bottom - top


def split_long_token(draw: ImageDraw.ImageDraw, token: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    parts: list[str] = []
    current = ""
    for ch in token:
        candidate = current + ch
        if current and text_size(draw, candidate, font)[0] > max_width:
            parts.append(current)
            current = ch
        else:
            current = candidate
    if current:
        parts.append(current)
    return parts


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    text = strip_invisible_text_marks(str(text or "/")).strip() or "/"
    output: list[str] = []
    for raw_line in text.splitlines() or [text]:
        line = raw_line.strip()
        if not line:
            output.append("")
            continue
        tokens = re.split(r"(\s+)", line)
        current = ""
        for token in tokens:
            if not token:
                continue
            candidate = current + token
            if text_size(draw, candidate, font)[0] <= max_width:
                current = candidate
                continue
            if current.strip():
                output.append(current.strip())
                current = ""
            if text_size(draw, token, font)[0] > max_width:
                output.extend(split_long_token(draw, token.strip(), font, max_width))
            else:
                current = token
        if current.strip():
            output.append(current.strip())
    return output or ["/"]


def line_heights(draw: ImageDraw.ImageDraw, lines: list[str], font: ImageFont.ImageFont, fallback_size: int) -> list[int]:
    heights: list[int] = []
    minimum = max(1, int(fallback_size * 0.95))
    for line in lines:
        _, height = text_size(draw, line or " ", font)
        heights.append(max(minimum, height))
    return heights


def fit_font_and_lines(
    draw: ImageDraw.ImageDraw,
    text: str,
    box: tuple[int, int, int, int],
    start_size: int,
    min_size: int,
    bold: bool = False,
    line_gap_ratio: float = 0.28,
) -> tuple[ImageFont.ImageFont, list[str], list[int], int]:
    _, _, width, height = box
    for size in range(start_size, min_size - 1, -1):
        font = load_font(size, bold=bold)
        lines = wrap_text(draw, text, font, width)
        heights = line_heights(draw, lines, font, size)
        line_gap = max(3, int(size * line_gap_ratio))
        total_height = sum(heights) + max(0, len(lines) - 1) * line_gap
        if total_height <= height:
            return font, lines, heights, line_gap
    font = load_font(min_size, bold=bold)
    lines = wrap_text(draw, text, font, width)
    heights = line_heights(draw, lines, font, min_size)
    line_gap = max(2, int(min_size * max(0.22, line_gap_ratio * 0.8)))
    per_line = max(1, min_size + line_gap)
    max_lines = max(1, height // per_line)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        heights = heights[:max_lines]
        if lines:
            lines[-1] = lines[-1].rstrip("...") + "..."
            heights = line_heights(draw, lines, font, min_size)
    return font, lines, heights, line_gap


def draw_field(
    draw: ImageDraw.ImageDraw,
    key: str,
    value: str,
    *,
    color: tuple[int, int, int],
    size: int,
    min_size: int,
    bold: bool = True,
    box: tuple[int, int, int, int] | None = None,
    background: tuple[int, int, int] | None = None,
    background_padding: tuple[int, int] = (0, 0),
    line_gap_ratio: float = 0.28,
) -> None:
    box = box or FIELD_BOXES[key]
    x, y, _, _ = box
    font, lines, heights, line_gap = fit_font_and_lines(
        draw,
        value,
        box,
        size,
        min_size,
        bold=bold,
        line_gap_ratio=line_gap_ratio,
    )
    cursor_y = y
    for line, line_height in zip(lines, heights):
        if background and line:
            line_width, line_height = text_size(draw, line, font)
            pad_x, pad_y = background_padding
            draw.rectangle(
                (
                    x - pad_x,
                    cursor_y - pad_y,
                    x + line_width + pad_x,
                    cursor_y + line_height + pad_y,
                ),
                fill=background,
            )
        draw.text((x, cursor_y), line, fill=color, font=font)
        cursor_y += line_height + line_gap


def sanitize_filename(value: str, fallback: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value or "").strip(" ._")
    cleaned = re.sub(r"\s+", "_", cleaned)
    return (cleaned or fallback)[:80]


def get_fields(sheet: dict[str, Any]) -> dict[str, str]:
    return dict(sheet.get("template_fields", {}).get("fields", {}))


def clamp_int(value: Any, fallback: int, min_value: int, max_value: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return fallback
    return max(min_value, min(max_value, number))


def field_adjustment(adjustments: dict[str, Any] | None, key: str) -> dict[str, Any]:
    if not isinstance(adjustments, dict):
        return {}
    fields = adjustments.get("fields")
    if not isinstance(fields, dict):
        return {}
    item = fields.get(key)
    return item if isinstance(item, dict) else {}


def adjusted_field_config(
    key: str,
    fields: dict[str, str],
    adjustments: dict[str, Any] | None,
) -> tuple[str, tuple[int, int, int, int], int, int]:
    x, y, width, height = FIELD_BOXES[key]
    style = FIELD_STYLES[key]
    item = field_adjustment(adjustments, key)

    value = item.get("value", fields.get(key, "/"))
    value = str(value or "/").strip() or "/"
    box = (
        clamp_int(item.get("x"), x, 0, 800),
        clamp_int(item.get("y"), y, 0, 1131),
        clamp_int(item.get("w"), width, 20, 800),
        clamp_int(item.get("h"), height, 20, 1131),
    )
    size = clamp_int(item.get("size"), int(style["size"]), 8, 64)
    min_size = clamp_int(item.get("min_size"), int(style["min_size"]), 8, size)
    return value, box, size, min_size


def render_sheet(
    template_path: Path,
    sheet: dict[str, Any],
    output_path: Path,
    adjustments: dict[str, Any] | None = None,
) -> None:
    image = Image.open(template_path).convert("RGB")
    draw = ImageDraw.Draw(image)
    fields = get_fields(sheet)

    for key in FIELD_BOXES:
        value, box, size, min_size = adjusted_field_config(key, fields, adjustments)
        style = FIELD_STYLES[key]
        draw_field(
            draw,
            key,
            value,
            color=style["color"],
            size=size,
            min_size=min_size,
            bold=bool(style["bold"]),
            box=box,
            background=style.get("background"),
            background_padding=style.get("background_padding", (0, 0)),
            line_gap_ratio=float(style.get("line_gap_ratio", 0.28)),
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path, quality=95)


def load_workbook_data() -> dict[str, Any]:
    if not WORKBOOK_JSON.exists():
        raise FileNotFoundError(f"Missing {WORKBOOK_JSON}. Run convert_excel_to_script_data.py first.")
    return json.loads(WORKBOOK_JSON.read_text(encoding="utf-8"))


def find_template(path_arg: str | None = None) -> Path:
    if path_arg:
        path = Path(path_arg)
        if not path.exists():
            raise FileNotFoundError(path)
        return path
    templates = [p for p in sorted(TEMPLATE_DIR.iterdir(), key=lambda item: item.name.lower()) if p.suffix.lower() in IMAGE_EXTENSIONS]
    if not templates:
        raise FileNotFoundError(f"No template images found in {TEMPLATE_DIR}")
    return templates[0]


def select_sheets(data: dict[str, Any], sheet_arg: str | None, all_sheets: bool, limit: int | None) -> list[dict[str, Any]]:
    sheets = data.get("sheets", [])
    if all_sheets:
        selected = sheets
    elif sheet_arg:
        if sheet_arg.isdigit():
            index = int(sheet_arg)
            selected = [sheet for sheet in sheets if sheet.get("sheet_index") == index]
        else:
            needle = sheet_arg.lower()
            selected = [sheet for sheet in sheets if needle in str(sheet.get("sheet_name", "")).lower()]
        if not selected:
            raise ValueError(f"No sheet matched: {sheet_arg}")
    else:
        selected = sheets[:1]

    if limit is not None:
        selected = selected[:limit]
    return selected


def scan_assets(folder: Path, output_dir: Path) -> list[dict[str, Any]]:
    assets: list[dict[str, Any]] = []
    folder.mkdir(exist_ok=True)
    for item in sorted(folder.iterdir(), key=lambda path: path.name.lower()):
        if not item.is_file() or item.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        stat = item.stat()
        rel_src = Path(os.path.relpath(item, output_dir)).as_posix()
        assets.append(
            {
                "name": item.name,
                "stem": item.stem,
                "extension": item.suffix.lower(),
                "path": str(item.resolve()),
                "src": rel_src,
                "size_bytes": stat.st_size,
                "mtime_ms": int(stat.st_mtime * 1000),
            }
        )
    return assets


def refresh_workbook_assets(data: dict[str, Any]) -> None:
    output_dir = Path("converted_excel")
    assets = data.setdefault("assets", {})
    assets["template_dir"] = str(TEMPLATE_DIR.resolve())
    assets["generated_dir"] = str(GENERATED_DIR.resolve())
    assets["templates"] = scan_assets(TEMPLATE_DIR, output_dir)
    assets["generated_images"] = scan_assets(GENERATED_DIR, output_dir)
    summary = data.setdefault("summary", {})
    summary["template_image_count"] = len(assets["templates"])
    summary["generated_image_count"] = len(assets["generated_images"])

    WORKBOOK_JSON.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    WORKBOOK_JS.write_text(
        "window.WORKBOOK_DATA=" + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Fill the price-card image template from converted Excel fields.")
    parser.add_argument("--sheet", help="Sheet index or name substring. Defaults to the first sheet.")
    parser.add_argument("--all", action="store_true", help="Generate images for all sheets.")
    parser.add_argument("--limit", type=int, help="Limit number of generated sheets after selection.")
    parser.add_argument("--template", help="Template image path. Defaults to first image in image_templates.")
    parser.add_argument("--clear", action="store_true", help="Delete existing generated PNG files before generating.")
    args = parser.parse_args()

    data = load_workbook_data()
    template_path = find_template(args.template)
    selected = select_sheets(data, args.sheet, args.all, args.limit)

    GENERATED_DIR.mkdir(exist_ok=True)
    if args.clear:
        for item in GENERATED_DIR.glob("*.png"):
            item.unlink()

    generated = []
    for sheet in selected:
        title = get_fields(sheet).get("title") or sheet.get("sheet_name") or "sheet"
        filename = f"{int(sheet['sheet_index']):03d}_{sanitize_filename(title, 'sheet')}.png"
        output_path = GENERATED_DIR / filename
        render_sheet(template_path, sheet, output_path)
        generated.append(output_path)

    refresh_workbook_assets(data)

    print(f"template={template_path}")
    print(f"generated_count={len(generated)}")
    print(f"output_dir={GENERATED_DIR.resolve()}")
    for path in generated[:10]:
        print(path)
    if len(generated) > 10:
        print(f"... {len(generated) - 10} more")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
