# Excel Image Tool

Excel Image Tool converts a multi-sheet Excel workbook into structured data, provides a local viewer, and generates promotional template images from extracted fields.

The project is Windows-first. The delivery package can run without Python when `ExcelImageTool.exe` is included.

## What It Does

- Reads all sheets from an Excel workbook.
- Extracts key fields such as `产品`, `规格`, `页面价`, `详情券`, `拍`, `定金`, `尾款`, `品类券`, `消费券`, `其他优惠`, `预估到手价`.
- Builds a local HTML viewer for searching sheets and previewing data.
- Generates images using `image_templates/输出模板_new.jpg`.
- Supports a portable Windows `.exe` build for delivery.

## Project Layout

```text
excel_image_tool_final/      Source and build workspace
excel_image_tool_release/    Windows delivery folder
project_context_handoff.md   Current project memory and maintenance notes
```

Important source files:

```text
excel_image_tool_final/convert_excel_to_script_data.py
excel_image_tool_final/generate_template_images.py
excel_image_tool_final/launch_viewer.py
excel_image_tool_final/serve_viewer.py
excel_image_tool_final/update_excel_data.py
excel_image_tool_final/rules/template_fill_rule.md
```

## Data Policy

Excel files, generated images, converted workbook data, build output, `.exe`, and `.zip` files are intentionally ignored by git.

Reason: these files are large, local, or derived from source/data. Put Excel files into the working folder only when running the tool locally.

## Development Setup

From `excel_image_tool_final`:

```bat
python -m pip install -r requirements.txt
python update_excel_data.py
python launch_viewer.py
```

To build the portable Windows executable:

```bat
build_portable.bat
```

## Delivery Usage

Use `excel_image_tool_release` as the delivery folder.

1. Put an Excel workbook directly inside `excel_image_tool_release`.
2. Double-click `update_and_start.bat`.
3. The tool converts the Excel data and opens the local viewer.
4. Generated images are written to `generated_images/`.

To only open existing converted data:

```bat
start_viewer.bat
```

To stop the viewer:

```bat
stop_viewer.bat
```

## Rules

The canonical rule file is:

```text
excel_image_tool_final/rules/template_fill_rule.md
```

The release copy is:

```text
excel_image_tool_release/rules/template_fill_rule.md
```

When extraction rules, viewer behavior, image generation, runtime steps, or delivery structure change, update `project_context_handoff.md` in the same change.
