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

## Rule Maintenance

Only maintain one rule document:

```text
rules/template_fill_rule.md
```

Keep both source and release copies synchronized when rules change.
