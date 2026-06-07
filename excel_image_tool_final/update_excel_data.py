from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
from pathlib import Path


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = app_root()
CONVERT_SCRIPT = ROOT / "convert_excel_to_script_data.py"
STABLE_OUTPUT = ROOT / "converted_excel"
PREVIOUS_OUTPUT = ROOT / "converted_excel_previous"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}


def ensure_inside_root(path: Path) -> Path:
    resolved = path.resolve()
    root = ROOT.resolve()
    try:
        resolved.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"路径不在工具目录内：{resolved}") from exc
    return resolved


def excel_sort_key(path: Path) -> tuple[float, float, str]:
    stat = path.stat()
    return stat.st_ctime, stat.st_mtime, path.name.lower()


def find_excel(path_arg: str | None) -> Path:
    if path_arg:
        path = Path(path_arg)
        if not path.is_absolute():
            path = ROOT / path
        path = path.resolve()
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(path)
        if path.suffix.lower() not in EXCEL_EXTENSIONS:
            raise ValueError(f"不是支持的 Excel 文件：{path.name}")
        return path

    candidates = [
        item
        for item in ROOT.iterdir()
        if item.is_file()
        and item.suffix.lower() in EXCEL_EXTENSIONS
        and not item.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("当前目录没有 Excel 文件")
    candidates.sort(key=excel_sort_key, reverse=True)
    return candidates[0]


def parse_output_dir(stdout: str) -> Path | None:
    match = re.search(r"^output_dir=(.+)$", stdout, flags=re.MULTILINE)
    if not match:
        return None
    return (ROOT / match.group(1).strip()).resolve()


def newest_converted_dir(before: set[Path]) -> Path:
    candidates = [
        path.resolve()
        for path in ROOT.glob("converted_excel*")
        if path.is_dir() and path.resolve() not in before
    ]
    if not candidates:
        raise RuntimeError("转换完成，但没有找到新的 converted_excel 输出目录")
    candidates.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return candidates[0]


def replace_stable_output(new_output: Path) -> None:
    new_output = ensure_inside_root(new_output)
    stable = STABLE_OUTPUT.resolve()
    if new_output == stable:
        return

    if PREVIOUS_OUTPUT.exists():
        shutil.rmtree(ensure_inside_root(PREVIOUS_OUTPUT))
    if STABLE_OUTPUT.exists():
        STABLE_OUTPUT.rename(PREVIOUS_OUTPUT)

    new_output.rename(STABLE_OUTPUT)

    if PREVIOUS_OUTPUT.exists():
        shutil.rmtree(ensure_inside_root(PREVIOUS_OUTPUT))


def run_convert(excel_path: Path) -> Path:
    before = {path.resolve() for path in ROOT.glob("converted_excel*") if path.is_dir()}
    old_argv = sys.argv[:]
    old_cwd = Path.cwd()
    try:
        os.chdir(ROOT)
        sys.argv = [str(CONVERT_SCRIPT), str(excel_path)]
        import convert_excel_to_script_data as converter

        result_code = converter.main()
    finally:
        sys.argv = old_argv
        os.chdir(old_cwd)

    if result_code != 0:
        raise RuntimeError(f"转换失败，退出码 {result_code}")

    output = newest_converted_dir(before)
    replace_stable_output(output)
    return STABLE_OUTPUT.resolve()


def main() -> int:
    parser = argparse.ArgumentParser(description="Pick an Excel file and refresh converted_excel in place.")
    parser.add_argument("excel", nargs="?", help="Excel file path. Defaults to the newest Excel in this folder.")
    args = parser.parse_args()

    try:
        excel_path = find_excel(args.excel)
        print(f"使用 Excel：{excel_path.name}")
        output = run_convert(excel_path)
    except Exception as exc:
        print(f"更新失败：{exc}", file=sys.stderr)
        return 1

    print(f"已更新页面数据：{output}")
    print("下一步运行：python serve_viewer.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
