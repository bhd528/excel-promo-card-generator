# GitHub Upload Notes

## Recommended Upload Scope

Commit source code, scripts, templates, rules, and documentation.

Do not commit local Excel workbooks, generated images, converted workbook data, build output, portable `.exe`, or `.zip` files. They are excluded in `.gitignore`.

## Before First Commit

```bat
git init
git status --short
git add README.md .gitignore docs excel_image_tool_final excel_image_tool_release project_context_handoff.md
git status --short
git commit -m "Initial project upload"
```

Review `git status --short` before committing. It should not include:

- `*.xlsx`
- `*.xls`
- `*.exe`
- `*.zip`
- `converted_excel/`
- `generated_images/`
- `build/`
- `dist/`
- `__pycache__/`

## Release Package

GitHub source code does not need to include `ExcelImageTool.exe`.

For non-technical users, build the exe locally with:

```bat
cd excel_image_tool_final
build_portable.bat
```

Then copy the built `ExcelImageTool.exe` into `excel_image_tool_release` before packaging or sending the release folder.

## Deployment Model

This project is a local Windows tool, not a hosted web service.

Recommended release flow:

1. Keep source code and docs in git.
2. Build `ExcelImageTool.exe` locally.
3. Copy the exe into `excel_image_tool_release`.
4. Zip `excel_image_tool_release`.
5. Upload the zip as a GitHub Release asset.

The release zip should include:

- `ExcelImageTool.exe`
- `setup_env.bat`
- `start_viewer.bat`
- `stop_viewer.bat`
- `update_and_start.bat`
- `README.md`
- `rules/template_fill_rule.md`
- `image_templates/输出模板_new.jpg`
- `generated_images/README.txt`

Converted Excel data is optional. If users will provide their own Excel, do not include `converted_excel/`.

## Rule Maintenance

Only maintain one rule document:

```text
rules/template_fill_rule.md
```

Keep both source and release copies synchronized when rules change.
