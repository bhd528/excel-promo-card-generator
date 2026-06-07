from __future__ import annotations

import csv
import json
import os
import posixpath
import re
import sys
import unicodedata
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

MAIN = f"{{{MAIN_NS}}}"
REL = f"{{{REL_NS}}}"
PKG_REL = f"{{{PKG_REL_NS}}}"
DRAWING = f"{{{DRAWING_NS}}}"

IMAGE_FORMULA_RE = re.compile(r'(?:_xlfn\.)?DISPIMG\("([^"]+)"', re.IGNORECASE)
INVALID_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
DATE_FORMAT_RE = re.compile(r"(?<!\\)(?:yyyy|yyy|yy|\u5e74|\u6708|\u65e5|dd|d/|/d|m/d|d-m|m-d|hh|h:|ss|s\.0)", re.I)
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
CATEGORY_COUPON_KEYWORDS = ("美妆券", "生活券", "美妆惊喜券", "精致生活券", "服饰天降券", "天猫国际券")
OTHER_DISCOUNT_KEYWORDS = ("购物金充值", "购物金充", "充购物金", "购物金", "红包补贴", "天猫返现卡")
TITLE_SPEC_UNITS = (
    "ml", "mL", "ML", "l", "L", "g", "G", "kg", "KG",
    "毫升", "升", "克", "千克", "斤",
    "片", "瓶", "支", "袋", "盒", "包", "条", "颗", "粒", "枚", "套", "件",
    "罐", "卷", "抽", "只", "双", "杯", "张", "贴", "块",
)
TEMPLATE_FIELD_ORDER = (
    ("title", "产品"),
    ("spec", "规格"),
    ("page_price", "页面价"),
    ("detail_coupon", "详情券"),
    ("buy_count", "拍"),
    ("deposit", "定金"),
    ("tail_payment", "尾款"),
    ("category_coupon", "品类券"),
    ("vip_coupon", "消费券"),
    ("other_discount", "其他优惠"),
    ("final_price", "预估到手价"),
)
BUILTIN_DATE_NUMFMTS = {
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
    45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
}


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


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


def resolve_xl_path(target: str, base_dir: str = "xl") -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return posixpath.normpath(posixpath.join(base_dir, target))


def col_to_number(col_letters: str) -> int:
    value = 0
    for ch in col_letters:
        value = value * 26 + ord(ch.upper()) - 64
    return value


def number_to_col(num: int) -> str:
    out = []
    while num:
        num, rem = divmod(num - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def split_cell_ref(ref: str) -> tuple[int, int, str]:
    match = re.match(r"^([A-Z]+)(\d+)$", ref)
    if not match:
        raise ValueError(f"Unsupported cell reference: {ref}")
    col_letters = match.group(1)
    return int(match.group(2)), col_to_number(col_letters), col_letters


def parse_range(ref: str) -> tuple[int, int, int, int]:
    if ":" not in ref:
        row, col, _ = split_cell_ref(ref)
        return row, col, row, col
    start, end = ref.split(":", 1)
    start_row, start_col, _ = split_cell_ref(start)
    end_row, end_col, _ = split_cell_ref(end)
    return start_row, start_col, end_row, end_col


def sanitize_filename(name: str, fallback: str) -> str:
    cleaned = INVALID_FILENAME_RE.sub("_", name).strip(" ._")
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned:
        cleaned = fallback
    return cleaned[:90]


def next_output_dir(base: Path) -> Path:
    if not base.exists():
        return base
    for i in range(2, 1000):
        candidate = base.with_name(f"{base.name}_{i}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Could not find available output directory for {base}")


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    strings: list[str] = []
    with zf.open("xl/sharedStrings.xml") as stream:
        for _, elem in ET.iterparse(stream, events=("end",)):
            if elem.tag == MAIN + "si":
                parts = [text_elem.text or "" for text_elem in elem.iter(MAIN + "t")]
                strings.append("".join(parts))
                elem.clear()
    return strings


def load_styles(zf: zipfile.ZipFile) -> tuple[dict[int, str], set[int]]:
    if "xl/styles.xml" not in zf.namelist():
        return {}, set()

    root = ET.fromstring(zf.read("xl/styles.xml"))
    numfmts: dict[int, str] = {}
    for numfmt in root.findall(f".//{MAIN}numFmt"):
        numfmts[int(numfmt.attrib["numFmtId"])] = numfmt.attrib.get("formatCode", "")

    style_formats: dict[int, str] = {}
    date_style_ids: set[int] = set()
    cell_xfs = root.find(f"{MAIN}cellXfs")
    if cell_xfs is None:
        return style_formats, date_style_ids

    for style_index, xf in enumerate(cell_xfs.findall(f"{MAIN}xf")):
        numfmt_id = int(xf.attrib.get("numFmtId", "0"))
        fmt = numfmts.get(numfmt_id, "")
        style_formats[style_index] = fmt
        if numfmt_id in BUILTIN_DATE_NUMFMTS or (fmt and DATE_FORMAT_RE.search(fmt)):
            date_style_ids.add(style_index)

    return style_formats, date_style_ids


def excel_serial_to_iso(raw: str, date1904: bool) -> str | None:
    try:
        value = float(raw)
    except ValueError:
        return None
    base = datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 30)
    result = base + timedelta(days=value)
    if result.time() == datetime.min.time():
        return result.date().isoformat()
    return result.isoformat(sep=" ", timespec="seconds")


def load_workbook_sheets(zf: zipfile.ZipFile) -> tuple[list[dict[str, Any]], bool]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(PKG_REL + "Relationship")}
    date1904 = workbook.find(MAIN + "workbookPr") is not None and workbook.find(MAIN + "workbookPr").attrib.get("date1904") in {"1", "true", "True"}

    sheets: list[dict[str, Any]] = []
    sheet_nodes = workbook.find(MAIN + "sheets")
    if sheet_nodes is None:
        return sheets, date1904

    for index, sheet in enumerate(sheet_nodes, start=1):
        rel_id = sheet.attrib[f"{REL}id"]
        target = rel_map[rel_id]
        sheets.append(
            {
                "sheet_index": index,
                "sheet_name": sheet.attrib["name"],
                "sheet_id": sheet.attrib.get("sheetId"),
                "worksheet_path": resolve_xl_path(target, "xl"),
            }
        )
    return sheets, date1904


def load_cell_images(zf: zipfile.ZipFile) -> dict[str, dict[str, Any]]:
    if "xl/cellimages.xml" not in zf.namelist():
        return {}

    rel_map: dict[str, dict[str, str | None]] = {}
    if "xl/_rels/cellimages.xml.rels" in zf.namelist():
        rels = ET.fromstring(zf.read("xl/_rels/cellimages.xml.rels"))
        for rel in rels.findall(PKG_REL + "Relationship"):
            target = rel.attrib.get("Target")
            mode = rel.attrib.get("TargetMode")
            rel_map[rel.attrib["Id"]] = {
                "target": resolve_xl_path(target, "xl") if target and mode != "External" and target != "NULL" else target,
                "target_mode": mode,
            }

    root = ET.fromstring(zf.read("xl/cellimages.xml"))
    images: dict[str, dict[str, Any]] = {}
    for item in root:
        c_nv_pr = None
        blip = None
        for elem in item.iter():
            lname = local_name(elem.tag)
            if lname == "cNvPr":
                c_nv_pr = elem
            elif lname == "blip":
                blip = elem
        if c_nv_pr is None:
            continue
        image_id = c_nv_pr.attrib.get("name")
        if not image_id:
            continue
        rel_id = blip.attrib.get(f"{REL}embed") if blip is not None else None
        rel_info = rel_map.get(rel_id or "", {})
        images[image_id] = {
            "image_id": image_id,
            "relationship_id": rel_id,
            "description": c_nv_pr.attrib.get("descr"),
            "target": rel_info.get("target"),
            "target_mode": rel_info.get("target_mode"),
        }
    return images


def cell_text_and_kind(
    cell: ET.Element,
    shared_strings: list[str],
    style_formats: dict[int, str],
    date_style_ids: set[int],
    date1904: bool,
) -> tuple[str, str, str | None, str | None, bool]:
    data_type = cell.attrib.get("t", "n")
    style_id_text = cell.attrib.get("s")
    style_id = int(style_id_text) if style_id_text is not None else None
    formula_elem = cell.find(MAIN + "f")
    value_elem = cell.find(MAIN + "v")
    formula = formula_elem.text if formula_elem is not None else None
    raw_value = value_elem.text if value_elem is not None else None
    is_date = bool(style_id in date_style_ids and raw_value)

    if data_type == "s":
        text = shared_strings[int(raw_value)] if raw_value not in (None, "") else ""
        return text, "shared_string", raw_value, formula, False
    if data_type == "inlineStr":
        inline = cell.find(MAIN + "is")
        text = "".join(text_elem.text or "" for text_elem in inline.iter(MAIN + "t")) if inline is not None else ""
        return text, "inline_string", raw_value, formula, False
    if data_type == "b":
        return ("true" if raw_value == "1" else "false"), "boolean", raw_value, formula, False
    if data_type == "e":
        return raw_value or "", "error", raw_value, formula, False
    if is_date:
        converted = excel_serial_to_iso(raw_value or "", date1904)
        if converted:
            return converted, "date", raw_value, formula, True

    if formula is not None:
        return raw_value or "", "formula", raw_value, formula, False

    return raw_value or "", "number" if data_type == "n" else data_type, raw_value, formula, False


def parse_sheet(
    zf: zipfile.ZipFile,
    sheet: dict[str, Any],
    shared_strings: list[str],
    style_formats: dict[int, str],
    date_style_ids: set[int],
    date1904: bool,
    image_map: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    worksheet_path = sheet["worksheet_path"]
    cells: list[dict[str, Any]] = []
    merges: list[str] = []
    dimension = None

    with zf.open(worksheet_path) as stream:
        for _, elem in ET.iterparse(stream, events=("end",)):
            if elem.tag == MAIN + "dimension":
                dimension = elem.attrib.get("ref")
            elif elem.tag == MAIN + "c":
                cell_ref = elem.attrib.get("r")
                if not cell_ref:
                    elem.clear()
                    continue
                row, col, col_letter = split_cell_ref(cell_ref)
                text, value_kind, raw_value, formula, is_date = cell_text_and_kind(
                    elem,
                    shared_strings,
                    style_formats,
                    date_style_ids,
                    date1904,
                )
                if text != "" or formula:
                    image_match = IMAGE_FORMULA_RE.search(formula or text or "")
                    image_id = image_match.group(1) if image_match else None
                    image_info = image_map.get(image_id or "", {})
                    style_id_text = elem.attrib.get("s")
                    style_id = int(style_id_text) if style_id_text is not None else None
                    cells.append(
                        {
                            "sheet_index": sheet["sheet_index"],
                            "sheet_name": sheet["sheet_name"],
                            "cell": cell_ref,
                            "row": row,
                            "col": col,
                            "col_letter": col_letter,
                            "text": text,
                            "raw_value": raw_value,
                            "value_kind": value_kind,
                            "data_type": elem.attrib.get("t", "n"),
                            "style_id": style_id,
                            "number_format": style_formats.get(style_id, "") if style_id is not None else "",
                            "is_date": is_date,
                            "formula": formula,
                            "is_formula": formula is not None,
                            "is_image_formula": image_id is not None,
                            "image_id": image_id,
                            "image_target": image_info.get("target"),
                        }
                    )
                elem.clear()
            elif elem.tag == MAIN + "mergeCell":
                ref = elem.attrib.get("ref")
                if ref:
                    merges.append(ref)
                elem.clear()

    merge_ranges = [(ref, *parse_range(ref)) for ref in merges]
    for item in cells:
        merge_range = None
        merge_anchor = None
        for ref, start_row, start_col, end_row, end_col in merge_ranges:
            if start_row <= item["row"] <= end_row and start_col <= item["col"] <= end_col:
                merge_range = ref
                merge_anchor = f"{number_to_col(start_col)}{start_row}"
                break
        item["merge_range"] = merge_range
        item["merge_anchor"] = merge_anchor
        item["is_merge_anchor"] = merge_anchor == item["cell"] if merge_anchor else False

    if cells:
        min_row = min(item["row"] for item in cells)
        max_row = max(item["row"] for item in cells)
        min_col = min(item["col"] for item in cells)
        max_col = max(item["col"] for item in cells)
    else:
        min_row = max_row = min_col = max_col = None

    meta = {
        **sheet,
        "dimension": dimension,
        "non_empty_cell_count": len(cells),
        "formula_count": sum(1 for item in cells if item["is_formula"]),
        "image_formula_count": sum(1 for item in cells if item["is_image_formula"]),
        "merge_count": len(merges),
        "merged_ranges": merges,
        "actual_min_row": min_row,
        "actual_max_row": max_row,
        "actual_min_col": min_col,
        "actual_max_col": max_col,
        "actual_min_col_letter": number_to_col(min_col) if min_col else None,
        "actual_max_col_letter": number_to_col(max_col) if max_col else None,
    }
    return cells, meta


def sheet_text_lines(cells: list[dict[str, Any]]) -> list[str]:
    by_row: dict[int, list[dict[str, Any]]] = {}
    for cell in cells:
        by_row.setdefault(cell["row"], []).append(cell)

    lines: list[str] = []
    for row in sorted(by_row):
        parts = []
        for cell in sorted(by_row[row], key=lambda item: item["col"]):
            text = cell["text"]
            if cell["image_id"]:
                text = f"[IMAGE:{cell['image_id']}]"
            parts.append(f"{cell['cell']}={text}")
        lines.append(" | ".join(parts))
    return lines


def cell_display_text(cell: dict[str, Any] | None) -> str:
    if not cell:
        return ""
    if cell.get("image_id"):
        return f"[IMAGE:{cell['image_id']}]"
    return cell.get("text") or ""


def expanded_row(
    cells: list[dict[str, Any]],
    meta: dict[str, Any],
    row_number: int = 4,
    columns: tuple[int, ...] = (1,),
) -> dict[str, Any]:
    refs = tuple(f"{number_to_col(col)}{row_number}" for col in columns)
    return expanded_cells(cells, meta, refs)


def expanded_cells(cells: list[dict[str, Any]], meta: dict[str, Any], refs: tuple[str, ...]) -> dict[str, Any]:
    cell_by_pos = {(cell["row"], cell["col"]): cell for cell in cells}
    merge_ranges: list[tuple[str, int, int, int, int]] = []

    for ref in meta["merged_ranges"]:
        start_row, start_col, end_row, end_col = parse_range(ref)
        merge_ranges.append((ref, start_row, start_col, end_row, end_col))

    result_cells: list[dict[str, Any]] = []
    seen: set[tuple[int, int]] = set()
    for ref in refs:
        row_number, col, col_letter = split_cell_ref(ref)
        if (row_number, col) in seen:
            continue
        seen.add((row_number, col))
        direct_cell = cell_by_pos.get((row_number, col))
        source_cell = direct_cell
        source_merge_range = direct_cell.get("merge_range") if direct_cell else None
        is_from_merge = False

        if source_cell is None:
            for ref, start_row, start_col, end_row, end_col in merge_ranges:
                if start_row <= row_number <= end_row and start_col <= col <= end_col:
                    source_cell = cell_by_pos.get((start_row, start_col))
                    source_merge_range = ref
                    is_from_merge = source_cell is not None
                    break

        result_cells.append(
            {
                "row": row_number,
                "col": col,
                "col_letter": col_letter,
                "cell": f"{col_letter}{row_number}",
                "text": cell_display_text(source_cell),
                "source_cell": source_cell.get("cell") if source_cell else None,
                "source_text": source_cell.get("text") if source_cell else "",
                "source_value_kind": source_cell.get("value_kind") if source_cell else None,
                "source_formula": source_cell.get("formula") if source_cell else None,
                "source_image_id": source_cell.get("image_id") if source_cell else None,
                "source_image_target": source_cell.get("image_target") if source_cell else None,
                "source_merge_range": source_merge_range,
                "is_from_merge": is_from_merge,
                "is_blank": not bool(cell_display_text(source_cell)),
            }
        )

    return {
        "refs": list(refs),
        "cell_count": len(result_cells),
        "non_blank_cell_count": sum(1 for item in result_cells if not item["is_blank"]),
        "has_merged_sources": any(item["is_from_merge"] for item in result_cells),
        "cells": result_cells,
        "text": " | ".join(f"{item['cell']}={item['text']}" for item in result_cells if not item["is_blank"]),
    }


def clean_inline_text(text: str) -> str:
    text = strip_invisible_text_marks(text or "")
    text = text.replace("\u00a0", " ").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def compact_text(text: str) -> str:
    text = clean_inline_text(text)
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_title(text: str, fallback: str) -> str:
    title = compact_text(text) or compact_text(fallback)
    title = re.sub(r"[【】\[\]「」『』《》〈〉<>]", "", title)
    unit_pattern = "|".join(re.escape(unit) for unit in TITLE_SPEC_UNITS)
    amount_pattern = rf"\d+(?:\.\d+)?\s*(?:{unit_pattern})(?:\s*/\s*(?:{unit_pattern}))?\s*(?:装|量|规格)?"
    repeated_amount_pattern = rf"{amount_pattern}(?:\s*(?:\*|x|X|×)\s*\d+(?:\.\d+)?\s*(?:{unit_pattern})?\s*(?:装|量|规格)?)*"
    title = re.sub(rf"[（(][^）)]*(?:{amount_pattern}|含水量\s*\d+(?:\.\d+)?%)[^）)]*[）)]", " ", title)
    title = re.sub(repeated_amount_pattern, " ", title)
    title = re.sub(r"\d+(?:\.\d+)?\s*%\s*", " ", title)
    title = re.sub(r"\s*([+＋])\s*", r"\1", title)
    title = re.sub(r"\s*[/|｜]\s*(?=[/|｜]|$)", " ", title)
    title = re.sub(r"(?:[/|｜]\s*)+$", "", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title or "/"


def get_expanded_cell_text(important_row: dict[str, Any], cell_ref: str) -> str:
    for cell in important_row.get("cells", []):
        if cell.get("cell") == cell_ref:
            return cell.get("text") or ""
    return ""


def normalize_amount_unit(value: str) -> str:
    value = compact_text(value)
    value = re.sub(r"\s+", "", value)
    return value or "/"


def extract_amount_unit(text: str) -> str:
    match = re.search(r"([0-9]+(?:\.[0-9]+)?\s*元(?:/[^\s，,；;。+]+)?(?:（[^）]+）)?)", text)
    return normalize_amount_unit(match.group(1)) if match else "/"


def first_amount(text: str) -> str:
    match = re.search(r"[0-9]+(?:\.[0-9]+)?\s*元", text)
    return normalize_amount_unit(match.group(0)) if match else "/"


def parse_amount_number(value: str) -> float | None:
    match = re.search(r"[0-9]+(?:\.[0-9]+)?", value or "")
    if not match:
        return None
    return float(match.group(0))


def format_amount_number(value: float) -> str:
    return f"{value:g}元"


def max_amount(matches: list[str]) -> str:
    if not matches:
        return "/"
    def amount_value(value: str) -> float:
        match = re.search(r"[0-9]+(?:\.[0-9]+)?", value)
        return float(match.group(0)) if match else -1
    return normalize_amount_unit(max(matches, key=amount_value))


def max_amount_from_text(text: str, pattern: str) -> str:
    matches = re.findall(pattern, clean_inline_text(text), flags=re.I)
    values = [parse_amount_number(match) for match in matches]
    numbers = [value for value in values if value is not None]
    if not numbers:
        return "/"
    return format_amount_number(max(numbers))


def is_section_break_line(line: str, extra_keywords: tuple[str, ...] = ()) -> bool:
    content = compact_text(line)
    if not content:
        return False
    keywords = (
        "实现方式", "赠品", "88vip", "88VIP", "品类券", "消费券", "其他优惠",
        "预估到手", "最终到手", "香菇省钱凑单攻略", "无需备注", "随单发出",
    ) + tuple(extra_keywords)
    return any(keyword in content for keyword in keywords)


def extract_page_price(text: str) -> str:
    text = clean_inline_text(text)
    lines = [line.rstrip() for line in text.splitlines()]
    for index, raw_line in enumerate(lines):
        line = compact_text(raw_line)
        if "页面价" not in line:
            continue
        candidates: list[str] = []
        suffix = re.split(r"页面价[:：]?", line, maxsplit=1)[-1].strip(" ，,")
        if suffix:
            candidates.append(suffix)
        for next_line in lines[index + 1:]:
            candidate = compact_text(next_line)
            if not candidate:
                if candidates:
                    break
                continue
            if is_section_break_line(candidate):
                break
            candidates.append(candidate)
        amounts: list[float] = []
        for candidate in candidates:
            snippet = compact_text(candidate)
            snippet = re.sub(r"^页面价[:：]?\s*", "", snippet)
            for match in re.findall(r"[0-9]+(?:\.[0-9]+)?\s*元", snippet):
                value = parse_amount_number(match)
                if value is not None:
                    amounts.append(value)
        if amounts:
            return format_amount_number(max(amounts))
    flat_text = compact_text(text)
    match = re.search(r"页面价[:：]?\s*([^。；;\n]+)", flat_text)
    if match:
        amounts = [
            value for value in
            (parse_amount_number(item) for item in re.findall(r"[0-9]+(?:\.[0-9]+)?\s*元", match.group(1)))
            if value is not None
        ]
        if amounts:
            return format_amount_number(max(amounts))
    return "/"


def extract_detail_coupon(text: str) -> str:
    text = clean_inline_text(text)
    patterns = (
        r"详情页[^。\n]*?领\s*([0-9]+(?:\.[0-9]+)?\s*元)\s*券",
        r"详情页[^。\n]*?([0-9]+(?:\.[0-9]+)?\s*元)\s*券",
    )
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            return normalize_amount_unit(match.group(1))
    return "/"


def extract_buy_count(text: str) -> str:
    text = clean_inline_text(text)
    match = re.search(r"拍\s*([0-9]+(?:\s*(?:件|瓶|盒|支|袋|套|单|份|罐|条|个))?)", text, flags=re.I)
    if not match:
        return "/"
    return compact_text(match.group(1)).replace(" ", "") or "/"


def extract_deposit(text: str) -> str:
    return max_amount_from_text(text, r"(?<!免)定金\s*([0-9]+(?:\.[0-9]+)?\s*元?)")


def extract_tail_payment(text: str) -> str:
    return max_amount_from_text(text, r"尾款\s*([0-9]+(?:\.[0-9]+)?\s*元?)")


def spec_variant_score(value: str) -> tuple[float, float, int]:
    multiplier_values = [float(item) for item in re.findall(r"[*xX×]\s*([0-9]+(?:\.[0-9]+)?)", value)]
    multiplier_total = sum(multiplier_values) if multiplier_values else 1.0
    numeric_values = [float(item) for item in re.findall(r"([0-9]+(?:\.[0-9]+)?)\s*(?:ml|mL|ML|L|l|g|G|kg|KG|片|支|瓶|袋|盒|包|粒|颗|套|件|条|枚|罐)", value)]
    amount = max(numeric_values) if numeric_values else 0.0
    return multiplier_total, amount, len(value)


def normalize_spec_source_text(text: str) -> str:
    text = clean_inline_text(text)
    text = re.sub(r"^规格(?:相关)?[:：]?\s*", "", text)
    text = re.sub(r"(?i)^(?:sku|suk)\s*\d+\s*[:：]\s*", "", text)
    return text.strip(" +；;，,。")


def extract_spec_from_source_text(text: str) -> str:
    text = normalize_spec_source_text(text)
    if not text:
        return "/"
    parts = [
        normalize_spec_source_text(part)
        for part in re.split(r"[\n；;]+", text)
    ]
    parts = [part for part in parts if part]
    if not parts:
        return "/"
    if len(parts) == 1:
        return parts[0]
    return max(parts, key=spec_variant_score)


def find_spec_source_cell(cells: list[dict[str, Any]]) -> dict[str, Any] | None:
    a_column = sorted((cell for cell in cells if int(cell.get("col", 0)) == 1), key=lambda item: int(item.get("row", 0)))
    by_row = {int(cell.get("row", 0)): cell for cell in a_column}
    for cell in a_column:
        text = compact_text(cell.get("text") or "")
        if "规格" not in text:
            continue
        row = int(cell.get("row", 0))
        next_cell = by_row.get(row + 1)
        if next_cell and compact_text(next_cell.get("text") or ""):
            return next_cell
        for candidate in a_column:
            candidate_row = int(candidate.get("row", 0))
            if candidate_row > row and compact_text(candidate.get("text") or ""):
                return candidate
    return None


def clean_spec_line(line: str) -> str:
    line = compact_text(line).strip(" +，,；;。")
    if not line:
        return ""
    line = re.sub(r"^赠品[:：]?\s*", "", line)
    if re.match(r"^拍[^：:\s]{1,12}[:：]", line):
        return ""
    if re.fullmatch(r"[（(]?\s*(?:总)?价值\s*[0-9]+(?:\.[0-9]+)?\s*元\s*[）)]?", line):
        return ""
    if re.fullmatch(r"[（(]?\s*总[^）)]*价值\s*[0-9]+(?:\.[0-9]+)?\s*元\s*[）)]?", line):
        return ""
    if re.match(r"^共(?:含赠|到手)", line):
        return ""
    if re.match(r"^[^：:\s]{1,12}[:：]\s*[0-9]+(?:\.[0-9]+)?\s*元", line):
        return ""
    if re.fullmatch(r"[\u4e00-\u9fffA-Za-z0-9]{1,8}\s*[:：]\s*", line):
        return ""
    line = re.sub(r"^[\u4e00-\u9fffA-Za-z0-9]{1,8}\s*[:：]\s*(?=\S)", "", line)
    if "折合" in line:
        line = re.split(r"折合", line, maxsplit=1)[0]
    summary_keywords = ("含赠", "到手", "折合", "正装量", "价值", "无需备注", "随单发出")
    if any(keyword in line for keyword in summary_keywords):
        line = re.split(r"[，,；;。]\s*(?=.*(?:含赠|到手|折合|正装量|价值|无需备注|随单发出))", line, maxsplit=1)[0]
        line = re.split(r"含赠|(?:含赠)?到手|折合|正装量|价值|无需备注|随单发出", line, maxsplit=1)[0]
    line = re.sub(r"^\d+\s*[、.．]\s*", "", line)
    line = re.sub(r"【\s*价值[^】]*】", "", line)
    line = re.sub(r"\[\s*价值[^\]]*\]", "", line)
    line = re.sub(r"[（(](?:价值|折合|到手|正装量|含赠)[^）)]*[）)]", "", line)
    line = re.sub(r"^[（(]\s*总\s*$", "", line)
    line = line.strip(" +，,；;。")
    line = re.sub(r"[【\[]+$", "", line).strip(" +，,；;。")
    line = re.sub(r"[（(]+$", "", line).strip(" +，,；;。")
    if not line or any(keyword in line for keyword in ("含赠", "到手", "折合", "正装量", "价值")):
        return ""
    return line


def unique_keep_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.add(value)
            output.append(value)
    return output


def extract_category_coupon(text: str) -> str:
    text = clean_inline_text(text)
    found: list[str] = []
    for keyword in CATEGORY_COUPON_KEYWORDS:
        pattern = rf"{re.escape(keyword)}[^\n，,；;。+]*(?:[（(]\s*减\s*[0-9]+(?:\.[0-9]+)?\s*元?\s*[）)])?"
        for match in re.finditer(pattern, text):
            snippet = compact_text(match.group(0)).strip(" +，,；;。")
            if snippet:
                found.append(snippet)
    found = unique_keep_order(found)
    return "；".join(found) if found else "/"


def extract_vip_coupon(text: str) -> str:
    text = clean_inline_text(text)
    matches = re.findall(
        r"88\s*vip[^\n，,；;。+]*?消费券[^\n，,；;。+]*?[（(]\s*减\s*([0-9]+(?:\.[0-9]+)?(?:\s*[~-]\s*[0-9]+(?:\.[0-9]+)?)?)\s*元?\s*[）)]",
        text,
        flags=re.I,
    )
    if not matches:
        return "/"
    values: list[float] = []
    for matched in matches:
        amounts = re.findall(r"[0-9]+(?:\.[0-9]+)?", matched)
        values.extend(float(amount) for amount in amounts)
    if not values:
        return "/"
    max_value = max(values)
    return f"预估减{max_value:g}元"


def extract_other_discount(text: str) -> str:
    text = clean_inline_text(text)
    found: list[str] = []
    for line in text.split("\n"):
        line = compact_text(line).strip(" +，,；;。")
        if not line:
            continue
        for pattern in (
            r"(?:购物金充值|购物金充|充购物金|购物金)[^+，,；;。]*(?:[（(][^）)]*[）)])?",
            r"红包补贴\s*[0-9]+(?:\.[0-9]+)?\s*元?",
            r"天猫返现卡[^+，,；;。]*?(?:[（(][^）)]*[）)])?",
        ):
            for match in re.finditer(pattern, line, flags=re.I):
                snippet = compact_text(match.group(0)).strip(" +，,；;。")
                snippet = re.split(r"到手|最终|含赠", snippet, maxsplit=1)[0].strip(" +，,；;。：:")
                if snippet and re.search(r"[0-9]+", snippet):
                    found.append(snippet)
    found = unique_keep_order(found)
    return "；".join(found) if found else "/"


def clean_final_price_snippet(snippet: str) -> str:
    snippet = compact_text(snippet).strip(" +，,；;。")
    snippet = re.sub(r"^到手价[:：]\s*", "", snippet)
    snippet = re.sub(r"^拍[0-9]+[,，]\s*", "", snippet)
    snippet = snippet.strip(" +，,；;。")
    to_hand_index = snippet.rfind("到手")
    if to_hand_index > 0:
        prefix = ""
        for candidate in ("含赠共", "含赠", "共"):
            start = to_hand_index - len(candidate)
            if start >= 0 and snippet[start:to_hand_index] == candidate:
                prefix = candidate
                break
        snippet = prefix + snippet[to_hand_index:]
    if snippet and "到手" not in snippet:
        snippet = "到手" + snippet
    snippet = re.sub(r"(到手)\s+", r"\1", snippet, count=1)
    snippet = re.sub(r"\s*[，,]\s*(折合)", r"\n\1", snippet)
    return snippet or "/"


def extract_final_price(text: str) -> str:
    text = clean_inline_text(text)
    candidates: list[str] = []

    for paragraph in re.split(r"\n\s*\n+", text):
        paragraph = compact_text(paragraph)
        if "到手" not in paragraph:
            continue
        if not re.search(r"[0-9]+(?:\.[0-9]+)?\s*元", paragraph):
            continue
        snippet = clean_final_price_snippet(paragraph)
        if snippet != "/":
            candidates.append(snippet)

    if candidates:
        return candidates[-1]

    flat_text = compact_text(text)
    amount_unit = r"[0-9]+(?:\.[0-9]+)?\s*元(?:/[^\s，,；;。+]+)?"
    pattern = rf"[^。；;]*?到手[^。；;]*?{amount_unit}(?:[，,]\s*折合{amount_unit})?"
    for match in re.finditer(pattern, flat_text):
        snippet = clean_final_price_snippet(match.group(0))
        if snippet != "/":
            candidates.append(snippet)
    return candidates[-1] if candidates else "/"


def extract_template_fields(sheet_name: str, important_row: dict[str, Any], cells: list[dict[str, Any]]) -> dict[str, Any]:
    b1_text = get_expanded_cell_text(important_row, "B1")
    a4_text = get_expanded_cell_text(important_row, "A4")
    spec_source_cell = find_spec_source_cell(cells)
    spec_source_text = spec_source_cell.get("text") if spec_source_cell else ""
    spec_source_ref = spec_source_cell.get("cell") if spec_source_cell else ""
    fields = {
        "title": clean_title(b1_text, sheet_name),
        "spec": extract_spec_from_source_text(spec_source_text),
        "page_price": extract_page_price(a4_text),
        "detail_coupon": extract_detail_coupon(a4_text),
        "buy_count": extract_buy_count(a4_text),
        "deposit": extract_deposit(a4_text),
        "tail_payment": extract_tail_payment(a4_text),
        "category_coupon": extract_category_coupon(a4_text),
        "vip_coupon": extract_vip_coupon(a4_text),
        "other_discount": extract_other_discount(a4_text),
        "final_price": extract_final_price(a4_text),
    }
    items = [
        {
            "key": key,
            "label": label,
            "value": fields.get(key, "/") or "/",
            "source_cell": "B1" if key == "title" else (spec_source_ref or "A4") if key == "spec" else "A4",
        }
        for key, label in TEMPLATE_FIELD_ORDER
    ]
    return {
        "items": items,
        "fields": fields,
        "source_cells": {"title": "B1", "pricing_text": "A4", "spec": spec_source_ref or ""},
        "raw": {"B1": b1_text, "A4": a4_text, "spec": spec_source_text},
    }


def write_sparse_sheet_csv(path: Path, cells: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "cell",
                "row",
                "col",
                "col_letter",
                "text",
                "formula",
                "image_id",
                "image_target",
                "merge_range",
                "merge_anchor",
            ],
        )
        writer.writeheader()
        for cell in sorted(cells, key=lambda item: (item["row"], item["col"])):
            writer.writerow({key: cell.get(key) for key in writer.fieldnames})


def write_grid_sheet_csv(path: Path, cells: list[dict[str, Any]], meta: dict[str, Any]) -> bool:
    if not cells:
        path.write_text("", encoding="utf-8-sig")
        return True

    min_row = meta["actual_min_row"]
    max_row = meta["actual_max_row"]
    min_col = meta["actual_min_col"]
    max_col = meta["actual_max_col"]
    if None in {min_row, max_row, min_col, max_col}:
        return False

    height = max_row - min_row + 1
    width = max_col - min_col + 1
    if width > 200 or height > 1000:
        return False

    values = {(cell["row"], cell["col"]): cell["text"] for cell in cells}
    for cell in cells:
        if cell.get("is_image_formula") and cell.get("image_id"):
            values[(cell["row"], cell["col"])] = f"[IMAGE:{cell['image_id']}]"

    anchor_values = dict(values)
    for ref in meta["merged_ranges"]:
        start_row, start_col, end_row, end_col = parse_range(ref)
        anchor = anchor_values.get((start_row, start_col))
        if not anchor:
            continue
        if end_row - start_row + 1 > 1000 or end_col - start_col + 1 > 200:
            continue
        for row in range(max(start_row, min_row), min(end_row, max_row) + 1):
            for col in range(max(start_col, min_col), min(end_col, max_col) + 1):
                values.setdefault((row, col), anchor)

    headers = [number_to_col(col) for col in range(min_col, max_col + 1)]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["row"] + headers)
        for row in range(min_row, max_row + 1):
            writer.writerow([row] + [values.get((row, col), "") for col in range(min_col, max_col + 1)])
    return True


def write_schema(path: Path) -> None:
    schema = {
        "encoding": "UTF-8",
        "files": {
            "workbook_data.json": "Canonical single JSON structure for scripts. Includes workbook metadata, image mapping, sheet list, cells, merges, and text lines.",
            "workbook_data.js": "Same data assigned to window.WORKBOOK_DATA so viewer.html can run directly from the filesystem.",
            "viewer.html": "Local visual browser for searching, selecting, and inspecting sheets.",
            "../image_templates/": "Default folder for template images used by the viewer and future image-generation script.",
            "../generated_images/": "Default folder for generated output images. The viewer previews files from this folder when present.",
            "schema.json": "This file.",
        },
        "asset_fields": {
            "assets.templates": "Template image files scanned from image_templates.",
            "assets.generated_images": "Generated image files scanned from generated_images.",
            "src": "Relative browser path from converted_excel/viewer.html to the image file.",
        },
        "sheet_fields": {
            "important_row": "Expanded key fields, currently B1 and A4. If a key cell is inside a merged range, the merged range anchor value is used and source_cell/source_merge_range keep traceability.",
            "template_fields": "Parsed fields for filling the price-card template, derived from B1 and A4 using rules/template_fill_rule.md.",
            "cells": "Original non-empty cells from the worksheet. Merged covered cells are not duplicated here.",
            "text_lines": "Readable row text from original non-empty cells.",
        },
        "cell_fields": {
            "sheet_index": "1-based workbook sheet order.",
            "sheet_name": "Original worksheet name.",
            "cell": "Excel cell reference.",
            "row": "1-based row number.",
            "col": "1-based column number.",
            "col_letter": "Excel column letters.",
            "text": "Canonical script-readable cell text.",
            "raw_value": "Raw value from worksheet XML before shared-string/date resolution.",
            "value_kind": "shared_string, inline_string, formula, number, date, boolean, error, or original data type.",
            "formula": "Formula text when present.",
            "image_id": "DISPIMG image ID when the cell references an embedded image.",
            "image_target": "Internal xlsx media path when an image ID could be resolved.",
            "merge_range": "Merged range containing this cell, if any.",
            "merge_anchor": "Top-left cell of the merged range, if any.",
        },
    }
    path.write_text(json.dumps(schema, ensure_ascii=False, indent=2), encoding="utf-8")


def collect_image_assets(workspace_dir: Path, output_dir: Path) -> dict[str, Any]:
    template_dir = workspace_dir / "image_templates"
    generated_dir = workspace_dir / "generated_images"
    template_dir.mkdir(exist_ok=True)
    generated_dir.mkdir(exist_ok=True)

    def scan(folder: Path) -> list[dict[str, Any]]:
        assets: list[dict[str, Any]] = []
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

    return {
        "template_dir": str(template_dir.resolve()),
        "generated_dir": str(generated_dir.resolve()),
        "templates": scan(template_dir),
        "generated_images": scan(generated_dir),
    }


def write_workbook_data(path: Path, workbook_data: dict[str, Any]) -> None:
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(workbook_data, handle, ensure_ascii=False, separators=(",", ":"))


def write_workbook_data_js(path: Path, workbook_data: dict[str, Any]) -> None:
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write("window.WORKBOOK_DATA=")
        json.dump(workbook_data, handle, ensure_ascii=False, separators=(",", ":"))
        handle.write(";\n")


def write_viewer_html(path: Path) -> None:
    html = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Excel Sheet Viewer</title>
  <style>
    :root {
      --bg: #f7f8fa;
      --panel: #ffffff;
      --text: #171a1f;
      --muted: #5f6876;
      --line: #d8dde6;
      --line-strong: #bcc5d2;
      --accent: #1769aa;
      --accent-soft: #e6f1fb;
      --warn: #8a5a00;
      --warn-soft: #fff4d7;
      --shadow: 0 10px 30px rgba(30, 42, 62, 0.08);
    }
    * { box-sizing: border-box; }
    html {
      height: 100%;
      overflow: hidden;
    }
    body {
      margin: 0;
      height: 100%;
      min-height: 100vh;
      overflow: hidden;
      font-family: "Segoe UI", "Microsoft YaHei", Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
    }
    .app {
      display: grid;
      grid-template-columns: minmax(280px, 360px) minmax(0, 1fr);
      height: 100vh;
      overflow: hidden;
    }
    .sidebar {
      display: flex;
      flex-direction: column;
      min-width: 0;
      border-right: 1px solid var(--line);
      background: var(--panel);
    }
    .brand {
      padding: 16px;
      border-bottom: 1px solid var(--line);
    }
    .brand h1 {
      margin: 0 0 8px;
      font-size: 18px;
      font-weight: 650;
      letter-spacing: 0;
    }
    .meta {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.35;
    }
    .meta b {
      display: block;
      color: var(--text);
      font-size: 14px;
      font-weight: 650;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .search {
      padding: 12px 16px;
      border-bottom: 1px solid var(--line);
    }
    .search input {
      width: 100%;
      height: 36px;
      border: 1px solid var(--line-strong);
      border-radius: 6px;
      padding: 0 10px;
      font: inherit;
      color: var(--text);
      background: #fff;
      outline: none;
    }
    .search input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px var(--accent-soft);
    }
    .filters {
      display: flex;
      gap: 6px;
      padding: 0 16px 12px;
      border-bottom: 1px solid var(--line);
      overflow-x: auto;
    }
    button {
      height: 32px;
      border: 1px solid var(--line-strong);
      border-radius: 6px;
      padding: 0 10px;
      color: var(--text);
      background: #fff;
      font: inherit;
      cursor: pointer;
      white-space: nowrap;
    }
    button:hover { border-color: var(--accent); }
    button:disabled {
      cursor: not-allowed;
      opacity: 0.55;
    }
    button.active {
      border-color: var(--accent);
      color: var(--accent);
      background: var(--accent-soft);
    }
    .sheet-list {
      flex: 1 1 0;
      min-height: 0;
      overflow: auto;
      overscroll-behavior: contain;
    }
    .sheet-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr);
      gap: 8px;
      align-items: center;
      width: 100%;
      min-height: 46px;
      padding: 7px 12px 7px 16px;
      border-bottom: 1px solid #edf0f4;
      cursor: pointer;
    }
    .sheet-row:hover,
    .sheet-row.active {
      background: var(--accent-soft);
    }
    .sheet-name {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      font-size: 14px;
    }
    .sheet-sub {
      margin-top: 3px;
      color: var(--muted);
      font-size: 12px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .main {
      display: flex;
      flex-direction: column;
      min-width: 0;
      min-height: 0;
      overflow: hidden;
    }
    .toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
      padding: 12px 16px;
      border-bottom: 1px solid var(--line);
      background: var(--panel);
    }
    .toolbar .grow { flex: 1; min-width: 160px; }
    .pill {
      display: inline-flex;
      align-items: center;
      min-height: 28px;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 4px 10px;
      color: var(--muted);
      background: #fff;
      font-size: 12px;
    }
    .content {
      flex: 1 1 0;
      height: 0;
      min-height: 0;
      overflow: hidden;
      padding: 16px;
    }
    .panel {
      display: flex;
      flex-direction: column;
      height: 100%;
      min-height: 0;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: var(--shadow);
    }
    .sheet-header {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 12px;
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
    }
    .sheet-header h2 {
      margin: 0;
      font-size: 18px;
      line-height: 1.35;
      letter-spacing: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .stats {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      justify-content: flex-end;
    }
    .tabs {
      display: flex;
      gap: 6px;
      padding: 10px 16px 0;
      border-bottom: 1px solid var(--line);
      background: #fbfcfe;
    }
    .tab {
      border-bottom-left-radius: 0;
      border-bottom-right-radius: 0;
      border-bottom-color: transparent;
    }
    .view {
      flex: 1 1 0;
      min-height: 0;
      overflow: auto;
      overscroll-behavior: contain;
      padding: 16px;
    }
    .grid-wrap {
      width: 100%;
      overflow: auto;
      overscroll-behavior: contain;
      border: 1px solid var(--line);
      border-radius: 6px;
    }
    table {
      width: max-content;
      min-width: 100%;
      border-collapse: collapse;
      font-size: 13px;
      table-layout: fixed;
    }
    th, td {
      max-width: 360px;
      min-width: 84px;
      border: 1px solid var(--line);
      padding: 6px 8px;
      vertical-align: top;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
      line-height: 1.35;
      background: #fff;
    }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #eef2f7;
      color: #303946;
      font-weight: 650;
    }
    th.row-head {
      left: 0;
      z-index: 3;
      min-width: 54px;
      width: 54px;
      text-align: right;
    }
    td.row-head {
      position: sticky;
      left: 0;
      z-index: 1;
      min-width: 54px;
      width: 54px;
      color: var(--muted);
      text-align: right;
      background: #f6f8fb;
    }
    .cell-list {
      display: grid;
      gap: 8px;
    }
    .cell-item {
      display: grid;
      grid-template-columns: 76px minmax(0, 1fr);
      gap: 10px;
      padding: 10px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
    }
    .cell-ref {
      color: var(--accent);
      font-weight: 650;
      overflow-wrap: anywhere;
    }
    pre {
      margin: 0;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
      font: 13px/1.5 Consolas, "Courier New", monospace;
    }
    .notice {
      margin-bottom: 12px;
      border: 1px solid #efd289;
      border-radius: 6px;
      padding: 10px 12px;
      color: var(--warn);
      background: var(--warn-soft);
      font-size: 13px;
      line-height: 1.45;
    }
    .action-strip {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 8px;
      margin-bottom: 12px;
    }
    .status-text {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.35;
    }
    .image-workbench {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
      min-height: 100%;
    }
    .image-column {
      display: flex;
      flex-direction: column;
      min-width: 0;
      min-height: 0;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      overflow: hidden;
    }
    .image-column-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      background: #fbfcfe;
    }
    .image-column-head h3 {
      margin: 0;
      font-size: 14px;
      font-weight: 650;
      letter-spacing: 0;
    }
    .image-frame {
      flex: 1 1 0;
      min-height: 320px;
      overflow: auto;
      overscroll-behavior: contain;
      display: grid;
      place-items: center;
      padding: 12px;
      background: #f5f7fa;
    }
    .image-frame img {
      display: block;
      max-width: 74%;
      max-height: 58vh;
      width: auto;
      height: auto;
      object-fit: contain;
      border: 1px solid var(--line);
      background: #fff;
    }
    .key-data-body {
      flex: 1 1 0;
      min-height: 320px;
      overflow: auto;
      overscroll-behavior: contain;
      padding: 12px;
      background: #fff;
    }
    .tune-panel {
      display: grid;
      gap: 10px;
      margin-top: 12px;
      border-top: 1px solid var(--line);
      padding-top: 12px;
    }
    .tune-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
    }
    .tune-head h4 {
      margin: 0;
      font-size: 14px;
      font-weight: 650;
    }
    .tune-field-list {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }
    .tune-field-list button {
      height: 30px;
      max-width: 132px;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .tune-form {
      display: grid;
      gap: 8px;
    }
    .tune-form label {
      display: grid;
      gap: 4px;
      color: var(--muted);
      font-size: 12px;
    }
    .tune-form textarea,
    .tune-form input {
      width: 100%;
      border: 1px solid var(--line-strong);
      border-radius: 6px;
      padding: 6px 8px;
      color: var(--text);
      background: #fff;
      font: inherit;
      outline: none;
    }
    .tune-form textarea {
      min-height: 72px;
      resize: vertical;
      line-height: 1.45;
    }
    .tune-form textarea:focus,
    .tune-form input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px var(--accent-soft);
    }
    .tune-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 8px;
    }
    .asset-list {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      padding: 10px 12px;
      border-top: 1px solid var(--line);
      background: #fff;
    }
    .asset-list button {
      max-width: 180px;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .empty {
      display: grid;
      place-items: center;
      min-height: 280px;
      color: var(--muted);
      text-align: center;
    }
    @media (max-width: 900px) {
      html, body { height: auto; overflow: auto; }
      .app { grid-template-columns: 1fr; height: auto; min-height: 100vh; overflow: visible; }
      .sidebar { height: 42vh; border-right: 0; border-bottom: 1px solid var(--line); }
      .main { min-height: 58vh; overflow: visible; }
      .content { padding: 10px; }
      .sheet-header { grid-template-columns: 1fr; }
      .stats { justify-content: flex-start; }
      .image-workbench { grid-template-columns: 1fr; }
      .tune-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    }
  </style>
</head>
<body>
  <div class="app">
    <aside class="sidebar">
      <div class="brand">
        <h1>Excel Sheet Viewer</h1>
        <div class="meta">
          <span><b id="sheetTotal">0</b>Sheet</span>
          <span><b id="cellTotal">0</b>非空单元格</span>
        </div>
      </div>
      <div class="search">
        <input id="searchInput" type="search" placeholder="搜索 sheet 名称">
      </div>
      <div class="filters">
        <button id="filterAll" class="active" type="button">全部</button>
        <button id="filterImages" type="button">含图片</button>
      </div>
      <div id="sheetList" class="sheet-list"></div>
    </aside>

    <main class="main">
      <div class="toolbar">
        <span id="sourceName" class="pill grow"></span>
      </div>
      <div class="content">
        <section class="panel">
          <div id="sheetHeader" class="sheet-header"></div>
          <div class="tabs">
            <button class="tab active" data-tab="images" type="button">图片预览</button>
            <button class="tab" data-tab="grid" type="button">表格</button>
          </div>
          <div id="view" class="view"></div>
        </section>
      </div>
    </main>
  </div>

  <script src="workbook_data.js"></script>
  <script>
    const data = window.WORKBOOK_DATA;
    const state = {
      query: "",
      filter: "all",
      activeIndex: 0,
      tab: "images",
      templateIndex: 0,
      generatedIndex: 0,
      generating: false,
      generateStatus: "",
      selectedField: "title",
      adjustmentsBySheet: {}
    };

    const tuningFields = [
      { key: "title", label: "产品", x: 190, y: 54, w: 486, h: 80, size: 40, min_size: 34 },
      { key: "spec", label: "规格", x: 190, y: 146, w: 560, h: 92, size: 32, min_size: 18 },
      { key: "page_price", label: "页面价", x: 170, y: 255, w: 150, h: 44, size: 40, min_size: 34 },
      { key: "detail_coupon", label: "详情券", x: 478, y: 248, w: 96, h: 44, size: 40, min_size: 34 },
      { key: "buy_count", label: "拍", x: 670, y: 248, w: 74, h: 44, size: 40, min_size: 34 },
      { key: "deposit", label: "定金", x: 170, y: 327, w: 176, h: 44, size: 40, min_size: 34 },
      { key: "tail_payment", label: "尾款", x: 170, y: 394, w: 176, h: 44, size: 40, min_size: 34 },
      { key: "category_coupon", label: "品类券", x: 190, y: 640, w: 486, h: 44, size: 40, min_size: 34 },
      { key: "vip_coupon", label: "消费券", x: 190, y: 707, w: 486, h: 44, size: 40, min_size: 34 },
      { key: "other_discount", label: "其他优惠", x: 190, y: 774, w: 486, h: 52, size: 40, min_size: 34 },
      { key: "final_price", label: "预估到手价", x: 272, y: 842, w: 404, h: 182, size: 40, min_size: 34 }
    ];
    const generatedHistoryLimit = 3;

    const el = {
      sheetTotal: document.getElementById("sheetTotal"),
      cellTotal: document.getElementById("cellTotal"),
      sourceName: document.getElementById("sourceName"),
      searchInput: document.getElementById("searchInput"),
      sheetList: document.getElementById("sheetList"),
      sheetHeader: document.getElementById("sheetHeader"),
      view: document.getElementById("view")
    };

    function formatNumber(value) {
      return Number(value || 0).toLocaleString("zh-CN");
    }

    function escapeHtml(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
    }

    function sheetSearchText(sheet) {
      if (!sheet._searchText) {
        sheet._searchText = String(sheet.sheet_name || "").toLowerCase();
      }
      return sheet._searchText;
    }

    function filteredSheets() {
      const q = state.query.trim().toLowerCase();
      return data.sheets.filter(sheet => {
        if (state.filter === "images" && !sheet.image_formula_count) return false;
        return !q || sheetSearchText(sheet).includes(q);
      });
    }

    function getActiveSheet() {
      return data.sheets.find(sheet => sheet.sheet_index === state.activeIndex) || data.sheets[0];
    }

    function setActive(sheetIndex) {
      state.activeIndex = sheetIndex;
      renderAll();
    }

    function renderList() {
      const sheets = filteredSheets();
      el.sheetList.innerHTML = sheets.map(sheet => {
        const active = sheet.sheet_index === state.activeIndex ? " active" : "";
        const sub = `${formatNumber(sheet.non_empty_cell_count)} 单元格 · ${sheet.dimension || "无范围"}`;
        return `
          <div class="sheet-row${active}" data-sheet-index="${sheet.sheet_index}">
            <span>
              <span class="sheet-name">${sheet.sheet_index}. ${escapeHtml(sheet.sheet_name)}</span>
              <span class="sheet-sub">${escapeHtml(sub)}</span>
            </span>
          </div>`;
      }).join("") || `<div class="empty">没有匹配的 sheet</div>`;

      el.sheetList.querySelectorAll(".sheet-row").forEach(row => {
        row.addEventListener("click", () => {
          const sheetIndex = Number(row.dataset.sheetIndex);
          setActive(sheetIndex);
        });
      });
    }

    function renderHeader(sheet) {
      if (!sheet) {
        el.sheetHeader.innerHTML = `<div><h2>没有数据</h2></div>`;
        return;
      }
      el.sheetHeader.innerHTML = `
        <div>
          <h2>${sheet.sheet_index}. ${escapeHtml(sheet.sheet_name)}</h2>
          <div class="sheet-sub">${escapeHtml(sheet.dimension || "无范围")} · ${escapeHtml(sheet.worksheet_path || "")}</div>
        </div>
        <div class="stats">
          <span class="pill">${formatNumber(sheet.non_empty_cell_count)} 单元格</span>
          <span class="pill">关键数据 ${formatNumber(sheet.important_row?.non_blank_cell_count || 0)} 项</span>
          <span class="pill">${formatNumber(sheet.merge_count)} 合并</span>
          <span class="pill">${formatNumber(sheet.formula_count)} 公式</span>
          <span class="pill">${formatNumber(sheet.image_formula_count)} 图片</span>
        </div>`;
    }

    function cellDisplay(cell) {
      if (cell.image_id) return `[IMAGE:${cell.image_id}]`;
      return cell.text ?? "";
    }

    function sheetTuningKey(sheet) {
      return String(sheet?.sheet_index ?? "0");
    }

    function fieldValueFromSheet(sheet, key) {
      const fields = sheet?.template_fields?.fields || {};
      return fields[key] || "/";
    }

    function createDefaultAdjustments(sheet) {
      const fields = {};
      tuningFields.forEach(field => {
        fields[field.key] = {
          value: fieldValueFromSheet(sheet, field.key),
          x: field.x,
          y: field.y,
          w: field.w,
          h: field.h,
          size: field.size,
          min_size: field.min_size
        };
      });
      return { fields };
    }

    function getAdjustments(sheet) {
      const key = sheetTuningKey(sheet);
      if (!state.adjustmentsBySheet[key]) {
        state.adjustmentsBySheet[key] = createDefaultAdjustments(sheet);
      }
      return state.adjustmentsBySheet[key];
    }

    function selectedTuningField() {
      return tuningFields.find(field => field.key === state.selectedField) || tuningFields[0];
    }

    function resetAdjustments(sheet) {
      delete state.adjustmentsBySheet[sheetTuningKey(sheet)];
      state.generateStatus = "已重置当前 sheet 的微调参数。";
      renderAll();
    }

    function renderImportantRow(sheet) {
      const important = sheet.important_row || { cells: [], non_blank_cell_count: 0, text: "", refs: ["B1", "A4"] };
      const rows = (important.cells || []).map(cell => {
        const source = [
          cell.source_cell ? `来源 ${cell.source_cell}` : "无来源单元格",
          cell.source_merge_range ? `合并 ${cell.source_merge_range}` : "",
          cell.source_image_target || ""
        ].filter(Boolean).join(" · ");
        return `
          <div class="cell-item">
            <div class="cell-ref">${escapeHtml(cell.cell)}</div>
            <div>
              <pre>${escapeHtml(cell.text || "")}</pre>
              <div class="sheet-sub">${escapeHtml(source)}</div>
            </div>
          </div>`;
      }).join("");

      return `
        <div class="notice">这里显示关键单元格 ${escapeHtml((important.refs || ["B1", "A4"]).join("、"))}。如果关键单元格落在某个合并区域内，会显示该合并区域左上角原始单元格的完整内容。</div>
        <div class="cell-list">${rows || '<div class="empty">没有可展开的关键数据</div>'}</div>`;
    }

    function buildGrid(sheet) {
      const minRow = sheet.actual_min_row || 1;
      const maxRow = sheet.actual_max_row || 1;
      const minCol = sheet.actual_min_col || 1;
      const maxCol = sheet.actual_max_col || 1;
      const height = maxRow - minRow + 1;
      const width = maxCol - minCol + 1;
      const rowLimit = 300;
      const colLimit = 80;
      const visibleMaxRow = Math.min(maxRow, minRow + rowLimit - 1);
      const visibleMaxCol = Math.min(maxCol, minCol + colLimit - 1);
      const values = new Map();

      for (const cell of sheet.cells || []) {
        if (cell.row >= minRow && cell.row <= visibleMaxRow && cell.col >= minCol && cell.col <= visibleMaxCol) {
          values.set(`${cell.row}:${cell.col}`, cellDisplay(cell));
        }
      }

      for (const ref of sheet.merged_ranges || []) {
        const parsed = parseRange(ref);
        if (!parsed) continue;
        const anchor = values.get(`${parsed.startRow}:${parsed.startCol}`);
        if (!anchor) continue;
        for (let row = Math.max(parsed.startRow, minRow); row <= Math.min(parsed.endRow, visibleMaxRow); row++) {
          for (let col = Math.max(parsed.startCol, minCol); col <= Math.min(parsed.endCol, visibleMaxCol); col++) {
            const key = `${row}:${col}`;
            if (!values.has(key)) values.set(key, anchor);
          }
        }
      }

      const limited = height > rowLimit || width > colLimit;
      let html = "";
      if (limited) {
        html += `<div class="notice">当前 sheet 实际范围较大，表格视图只显示前 ${rowLimit} 行、${colLimit} 列；完整内容见“单元格”或“JSON”。</div>`;
      }
      html += `<div class="grid-wrap"><table><thead><tr><th class="row-head">#</th>`;
      for (let col = minCol; col <= visibleMaxCol; col++) {
        html += `<th>${numberToCol(col)}</th>`;
      }
      html += `</tr></thead><tbody>`;
      for (let row = minRow; row <= visibleMaxRow; row++) {
        html += `<tr><td class="row-head">${row}</td>`;
        for (let col = minCol; col <= visibleMaxCol; col++) {
          html += `<td>${escapeHtml(values.get(`${row}:${col}`) || "")}</td>`;
        }
        html += `</tr>`;
      }
      html += `</tbody></table></div>`;
      return html;
    }

    function parseRange(ref) {
      const parts = String(ref || "").split(":");
      const start = parseCellRef(parts[0]);
      const end = parseCellRef(parts[1] || parts[0]);
      if (!start || !end) return null;
      return { startRow: start.row, startCol: start.col, endRow: end.row, endCol: end.col };
    }

    function parseCellRef(ref) {
      const match = /^([A-Z]+)(\d+)$/.exec(ref || "");
      if (!match) return null;
      return { col: colToNumber(match[1]), row: Number(match[2]) };
    }

    function colToNumber(letters) {
      let value = 0;
      for (const ch of letters) value = value * 26 + ch.charCodeAt(0) - 64;
      return value;
    }

    function numberToCol(num) {
      let out = "";
      while (num > 0) {
        const rem = (num - 1) % 26;
        out = String.fromCharCode(65 + rem) + out;
        num = Math.floor((num - 1) / 26);
      }
      return out;
    }

    function renderCells(sheet) {
      const rows = (sheet.cells || []).map(cell => {
        const detail = [
          cell.value_kind,
          cell.merge_range ? `merge ${cell.merge_range}` : "",
          cell.image_target ? cell.image_target : ""
        ].filter(Boolean).join(" · ");
        return `
          <div class="cell-item">
            <div class="cell-ref">${escapeHtml(cell.cell)}</div>
            <div>
              <pre>${escapeHtml(cellDisplay(cell))}</pre>
              <div class="sheet-sub">${escapeHtml(detail)}</div>
            </div>
          </div>`;
      }).join("");
      return `<div class="cell-list">${rows || '<div class="empty">没有非空单元格</div>'}</div>`;
    }

    function clampIndex(index, items) {
      if (!items.length) return 0;
      return Math.min(Math.max(index, 0), items.length - 1);
    }

    function assetEntries(items, limit = null) {
      const entries = items.map((asset, index) => ({ asset, index }));
      if (!limit || items.length <= limit) return entries;
      return entries
        .sort((left, right) => {
          const byTime = Number(right.asset.mtime_ms || 0) - Number(left.asset.mtime_ms || 0);
          return byTime || right.index - left.index;
        })
        .slice(0, limit);
    }

    function renderAssetButtons(items, kind, activeIndex, limit = null) {
      if (!items.length) {
        return `<div class="sheet-sub">暂无图片文件</div>`;
      }
      return assetEntries(items, limit).map(({ asset, index }) => {
        const active = index === activeIndex ? " active" : "";
        return `<button class="asset-choice${active}" data-kind="${kind}" data-index="${index}" type="button" title="${escapeHtml(asset.name)}">${escapeHtml(asset.name)}</button>`;
      }).join("");
    }

    function renderImagePanel(title, folder, items, activeIndex, kind, emptyText, historyLimit = null) {
      const asset = items[activeIndex];
      const version = asset?.mtime_ms ? `?v=${encodeURIComponent(asset.mtime_ms)}` : "";
      const imageHtml = asset
        ? `<img src="${escapeHtml(asset.src + version)}" alt="${escapeHtml(asset.name)}">`
        : `<div class="empty">${escapeHtml(emptyText)}</div>`;
      return `
        <section class="image-column">
          <div class="image-column-head">
            <h3>${escapeHtml(title)}</h3>
            <span class="sheet-sub">${formatNumber(items.length)} 张</span>
          </div>
          <div class="image-frame">${imageHtml}</div>
          <div class="asset-list">${renderAssetButtons(items, kind, activeIndex, historyLimit)}</div>
          <div class="sheet-sub" style="padding: 0 12px 12px;">${escapeHtml(folder || "")}</div>
        </section>`;
    }

    function renderTuningPanel(sheet) {
      const adjustments = getAdjustments(sheet);
      const selected = selectedTuningField();
      const item = adjustments.fields[selected.key] || {};
      const buttons = tuningFields.map(field => {
        const active = field.key === selected.key ? " active" : "";
        return `<button class="tune-field${active}" data-field="${field.key}" type="button" title="${escapeHtml(field.label)}">${escapeHtml(field.label)}</button>`;
      }).join("");

      return `
        <div class="tune-panel">
          <div class="tune-head">
            <h4>手动微调</h4>
            <button id="resetTuning" type="button">重置</button>
          </div>
          <div class="sheet-sub">选择字段后调整内容、位置和字号，再点“生成当前图片”。</div>
          <div class="tune-field-list">${buttons}</div>
          <div class="tune-form" data-field="${escapeHtml(selected.key)}">
            <label>内容
              <textarea data-tune="value">${escapeHtml(item.value ?? fieldValueFromSheet(sheet, selected.key))}</textarea>
            </label>
            <div class="tune-grid">
              <label>X <input data-tune="x" type="number" min="0" max="538" step="1" value="${escapeHtml(item.x ?? selected.x)}"></label>
              <label>Y <input data-tune="y" type="number" min="0" max="780" step="1" value="${escapeHtml(item.y ?? selected.y)}"></label>
              <label>字号 <input data-tune="size" type="number" min="8" max="64" step="1" value="${escapeHtml(item.size ?? selected.size)}"></label>
              <label>宽 <input data-tune="w" type="number" min="20" max="538" step="1" value="${escapeHtml(item.w ?? selected.w)}"></label>
              <label>高 <input data-tune="h" type="number" min="20" max="780" step="1" value="${escapeHtml(item.h ?? selected.h)}"></label>
              <label>最小字号 <input data-tune="min_size" type="number" min="8" max="64" step="1" value="${escapeHtml(item.min_size ?? selected.min_size)}"></label>
            </div>
          </div>
        </div>`;
    }

    function renderKeyDataPanel(sheet) {
      const important = sheet.important_row || { cells: [], non_blank_cell_count: 0 };
      return `
        <section class="image-column">
          <div class="image-column-head">
            <h3>关键数据</h3>
            <span class="sheet-sub">${formatNumber(important.non_blank_cell_count || 0)} 项</span>
          </div>
          <div class="key-data-body">
            ${renderImportantRow(sheet)}
            ${renderTuningPanel(sheet)}
          </div>
        </section>`;
    }

    function bindTuningControls(sheet) {
      el.view.querySelectorAll(".tune-field").forEach(button => {
        button.addEventListener("click", () => {
          state.selectedField = button.dataset.field || "title";
          renderAll();
        });
      });

      const form = el.view.querySelector(".tune-form");
      if (form) {
        const fieldKey = form.dataset.field;
        form.querySelectorAll("[data-tune]").forEach(input => {
          input.addEventListener("input", () => {
            const adjustments = getAdjustments(sheet);
            const item = adjustments.fields[fieldKey] || {};
            const prop = input.dataset.tune;
            item[prop] = prop === "value" ? input.value : Number(input.value);
            adjustments.fields[fieldKey] = item;
          });
        });
      }

      const resetButton = el.view.querySelector("#resetTuning");
      if (resetButton) {
        resetButton.addEventListener("click", () => resetAdjustments(sheet));
      }
    }

    function bindImageAssetButtons(sheet) {
      el.view.querySelectorAll(".asset-choice").forEach(button => {
        button.addEventListener("click", () => {
          const index = Number(button.dataset.index);
          if (button.dataset.kind === "template") state.templateIndex = index;
          if (button.dataset.kind === "generated") state.generatedIndex = index;
          renderAll();
        });
      });
      const generateButton = el.view.querySelector("#generateCurrentImage");
      if (generateButton) {
        generateButton.addEventListener("click", generateCurrentImage);
      }
      bindTuningControls(sheet);
    }

    function mergeGeneratedResponse(payload) {
      if (payload.assets) data.assets = payload.assets;
      if (payload.summary) data.summary = { ...data.summary, ...payload.summary };
      const generated = data.assets?.generated_images || [];
      const generatedName = payload.generated?.name || "";
      const generatedIndex = generated.findIndex(asset => asset.name === generatedName);
      state.generatedIndex = generatedIndex >= 0 ? generatedIndex : clampIndex(state.generatedIndex, generated);
    }

    async function generateCurrentImage() {
      const sheet = getActiveSheet();
      const assets = data.assets || {};
      const templates = assets.templates || [];
      const template = templates[state.templateIndex] || null;

      if (window.location.protocol === "file:") {
        state.generateStatus = "当前是直接打开 HTML，只能浏览。请运行 python serve_viewer.py 后用本地服务地址打开页面再生成。";
        renderView(sheet);
        return;
      }

      if (!sheet) {
        state.generateStatus = "没有可生成的 sheet。";
        renderView(sheet);
        return;
      }

      if (!template) {
        state.generateStatus = "image_templates 中没有可用模板图。";
        renderView(sheet);
        return;
      }

      state.generating = true;
      state.generateStatus = `正在生成：${sheet.sheet_name || sheet.sheet_index}`;
      renderView(sheet);

      try {
        const response = await fetch("/api/generate-image", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            sheet_index: sheet.sheet_index,
            template_name: template.name,
            adjustments: getAdjustments(sheet)
          })
        });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok || !payload.ok) {
          throw new Error(payload.error || `HTTP ${response.status}`);
        }
        mergeGeneratedResponse(payload);
        state.generateStatus = payload.message || "已生成图片并刷新预览。";
      } catch (error) {
        state.generateStatus = `生成失败：${error.message || error}`;
      } finally {
        state.generating = false;
        renderAll();
      }
    }

    function renderImages(sheet) {
      const assets = data.assets || {};
      const templates = assets.templates || [];
      const generated = assets.generated_images || [];
      state.templateIndex = clampIndex(state.templateIndex, templates);
      state.generatedIndex = clampIndex(state.generatedIndex, generated);
      const recentGenerated = assetEntries(generated, generatedHistoryLimit);
      if (recentGenerated.length && !recentGenerated.some(entry => entry.index === state.generatedIndex)) {
        state.generatedIndex = recentGenerated[0].index;
      }
      const status = state.generateStatus || (
        window.location.protocol === "file:"
          ? "直接双击 HTML 时只能浏览；运行 python serve_viewer.py 后可在本页生成图片。"
          : "点击按钮会按当前 sheet 和当前模板生成图片，并写入 generated_images。"
      );
      const disabled = state.generating || !templates.length;
      return `
        <div class="notice">生成图片时会使用 image_templates 文件夹里的默认模板；生成后的图片默认放在 generated_images 文件夹。</div>
        <div class="action-strip">
          <button id="generateCurrentImage" type="button" ${disabled ? "disabled" : ""}>${state.generating ? "生成中..." : "生成当前图片"}</button>
          <span id="generateStatus" class="status-text">${escapeHtml(status)}</span>
        </div>
        <div class="image-workbench">
          ${renderKeyDataPanel(sheet)}
          ${renderImagePanel("生成图片预览", assets.generated_dir, generated, state.generatedIndex, "generated", "generated_images 中还没有生成图", generatedHistoryLimit)}
        </div>`;
    }

    function installWheelScopes() {
      const selector = "textarea, .key-data-body, .image-frame, .grid-wrap, .sheet-list, .view";
      document.addEventListener("wheel", event => {
        const target = event.target instanceof Element ? event.target : null;
        if (!target) return;

        const candidates = [];
        for (let node = target; node && node !== document.body; node = node.parentElement) {
          if (node.matches(selector)) candidates.push(node);
        }
        if (!candidates.length) return;

        const preferHorizontal = Math.abs(event.deltaX) > Math.abs(event.deltaY);
        const canScrollY = node => node.scrollHeight > node.clientHeight + 1;
        const canScrollX = node => node.scrollWidth > node.clientWidth + 1;
        const scrollNode = candidates.find(node => {
          if (preferHorizontal && canScrollX(node)) return true;
          if (!preferHorizontal && canScrollY(node)) return true;
          return canScrollX(node) || canScrollY(node);
        }) || candidates[0];

        if (preferHorizontal) {
          if (canScrollX(scrollNode)) scrollNode.scrollLeft += event.deltaX;
          else if (canScrollY(scrollNode)) scrollNode.scrollTop += event.deltaX;
        } else {
          if (canScrollY(scrollNode)) scrollNode.scrollTop += event.deltaY;
          else if (canScrollX(scrollNode)) scrollNode.scrollLeft += event.deltaY;
        }

        event.preventDefault();
        event.stopPropagation();
      }, { passive: false });
    }

    function renderView(sheet) {
      if (!sheet) {
        el.view.innerHTML = `<div class="empty">没有数据</div>`;
        return;
      }
      if (state.tab === "images") {
        el.view.innerHTML = renderImages(sheet);
        bindImageAssetButtons(sheet);
      }
      if (state.tab === "grid") el.view.innerHTML = buildGrid(sheet);
    }

    function renderAll() {
      const sheet = getActiveSheet();
      el.sheetTotal.textContent = formatNumber(data.summary.sheet_count);
      el.cellTotal.textContent = formatNumber(data.summary.non_empty_cell_count);
      el.sourceName.textContent = data.source_file || "";
      renderList();
      renderHeader(sheet);
      renderView(sheet);
    }

    document.querySelectorAll(".filters button").forEach(button => {
      button.addEventListener("click", () => {
        document.querySelectorAll(".filters button").forEach(item => item.classList.remove("active"));
        button.classList.add("active");
        state.filter = button.id === "filterImages" ? "images" : "all";
        renderAll();
      });
    });

    document.querySelectorAll(".tab").forEach(button => {
      button.addEventListener("click", () => {
        document.querySelectorAll(".tab").forEach(item => item.classList.remove("active"));
        button.classList.add("active");
        state.tab = button.dataset.tab;
        renderAll();
      });
    });

    el.searchInput.addEventListener("input", event => {
      state.query = event.target.value;
      renderAll();
    });

    if (!data || !Array.isArray(data.sheets)) {
      document.body.innerHTML = "<div class='empty'>workbook_data.js 加载失败</div>";
    } else {
      state.activeIndex = data.sheets[0]?.sheet_index || 0;
      installWheelScopes();
      renderAll();
    }
  </script>
</body>
</html>
"""
    path.write_text(html, encoding="utf-8", newline="\n")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    input_path = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    if input_path is None:
        excel_files = sorted(
            list(Path.cwd().glob("*.xlsx"))
            + list(Path.cwd().glob("*.xlsm"))
            + list(Path.cwd().glob("*.xlsb"))
            + list(Path.cwd().glob("*.xls"))
        )
        if len(excel_files) != 1:
            print(f"Expected exactly one Excel file in {Path.cwd()}, found {len(excel_files)}.", file=sys.stderr)
            return 2
        input_path = excel_files[0]

    output_dir = next_output_dir(Path.cwd() / "converted_excel")
    output_dir.mkdir(parents=True, exist_ok=False)

    all_cells_count = 0
    workbook_data: dict[str, Any] = {
        "source_file": str(input_path.resolve()),
        "source_size_bytes": input_path.stat().st_size,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "output_dir": str(output_dir.resolve()),
        "schema_version": 1,
        "summary": {},
        "images": {},
        "assets": collect_image_assets(Path.cwd(), output_dir),
        "sheets": [],
    }

    with zipfile.ZipFile(input_path) as zf:
        shared_strings = load_shared_strings(zf)
        style_formats, date_style_ids = load_styles(zf)
        sheets, date1904 = load_workbook_sheets(zf)
        image_map = load_cell_images(zf)

        media_files = [name for name in zf.namelist() if name.startswith("xl/media/")]
        workbook_data["summary"]["xlsx_entry_count"] = len(zf.namelist())
        workbook_data["summary"]["sheet_count"] = len(sheets)
        workbook_data["summary"]["shared_string_count"] = len(shared_strings)
        workbook_data["summary"]["media_file_count"] = len(media_files)
        workbook_data["summary"]["media_total_uncompressed_bytes"] = sum(zf.getinfo(name).file_size for name in media_files)
        workbook_data["summary"]["cell_image_count"] = len(image_map)
        workbook_data["summary"]["date1904"] = date1904
        workbook_data["summary"]["template_image_count"] = len(workbook_data["assets"]["templates"])
        workbook_data["summary"]["generated_image_count"] = len(workbook_data["assets"]["generated_images"])
        workbook_data["images"] = image_map

        for sheet in sheets:
            cells, meta = parse_sheet(zf, sheet, shared_strings, style_formats, date_style_ids, date1904, image_map)
            all_cells_count += len(cells)
            important_row = expanded_cells(cells, meta, refs=("B1", "A4"))
            workbook_data["sheets"].append(
                {
                    **meta,
                    "cells": cells,
                    "important_row": important_row,
                    "template_fields": extract_template_fields(sheet["sheet_name"], important_row, cells),
                    "text_lines": sheet_text_lines(cells),
                }
            )
            if sheet["sheet_index"] % 100 == 0:
                print(f"processed {sheet['sheet_index']} / {len(sheets)} sheets")

    workbook_data["summary"]["non_empty_cell_count"] = all_cells_count
    write_workbook_data(output_dir / "workbook_data.json", workbook_data)
    write_workbook_data_js(output_dir / "workbook_data.js", workbook_data)
    write_viewer_html(output_dir / "viewer.html")
    write_schema(output_dir / "schema.json")

    print(f"output_dir={output_dir}")
    print(f"sheets={workbook_data['summary']['sheet_count']}")
    print(f"non_empty_cells={all_cells_count}")
    print("files=workbook_data.json, workbook_data.js, viewer.html, schema.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
