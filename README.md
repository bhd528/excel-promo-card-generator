# Excel 促销价签生成器

从多 Sheet Excel 中提取促销字段，并基于模板批量生成商品价签图片的本地工具。

项目主要面向 Windows 使用。交付包包含 `ExcelImageTool.exe` 时，使用者不需要安装 Python。

## 功能

- Reads all sheets from an Excel workbook.
- Extracts key fields: `产品`, `规格`, `页面价`, `详情券`, `拍`, `定金`, `尾款`, `品类券`, `消费券`, `其他优惠`, `预估到手价`.
- Builds a local HTML viewer for searching sheets, checking extracted fields, and previewing generated images.
- Generates image files from `image_templates/输出模板_new.jpg`.
- Supports a portable Windows `.exe` for delivery.

## 目录结构

```text
excel_image_tool_final/      Source and build workspace
excel_image_tool_release/    Local delivery folder, ignored by git
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

## 普通使用方式

适用于已经拿到完整交付包的使用者。`excel_image_tool_release` 是本地交付目录，不作为源码提交到 GitHub。

1. 打开 `excel_image_tool_release`。
2. 把 Excel 文件放到 `excel_image_tool_release` 目录下。
3. 双击 `update_and_start.bat`。
4. 等待脚本转换数据并自动打开浏览器。
5. 在页面左侧选择 sheet。
6. 在 `图片预览` 中点击 `生成当前图片`。
7. 生成结果保存在 `excel_image_tool_release/generated_images/`。

如果只是打开已有转换数据，双击：

```bat
excel_image_tool_release/start_viewer.bat
```

停止本地服务：

```bat
excel_image_tool_release/stop_viewer.bat
```

检查交付包环境：

```bat
excel_image_tool_release/setup_env.bat
```

## 开发环境运行

适用于需要改代码或调规则的开发者。

环境要求：

- Windows
- Python 3.10+

安装依赖：

```bat
cd excel_image_tool_final
python -m pip install -r requirements.txt
```

放入 Excel 后转换数据：

```bat
python update_excel_data.py
```

启动本地页面：

```bat
python launch_viewer.py
```

## 构建交付版

安装构建依赖并生成单文件 exe：

```bat
cd excel_image_tool_final
python -m pip install -r requirements-build.txt
build_portable.bat
```

构建完成后会生成：

```text
excel_image_tool_final/ExcelImageTool.exe
```

交付给普通用户时，需要把最新 exe 同步到：

```text
excel_image_tool_release/ExcelImageTool.exe
```

然后把 `excel_image_tool_release` 文件夹作为交付包。

## GitHub 部署/发布方式

本项目不是云端 Web 服务，不需要服务器部署。推荐方式是：

1. GitHub 仓库只保存源码、脚本、模板和文档。
2. 本地构建 `ExcelImageTool.exe`。
3. 将 exe 放入 `excel_image_tool_release`。
4. 把 `excel_image_tool_release` 打包成 zip，作为 GitHub Release 附件发布。

不要把 `excel_image_tool_release/`、业务 Excel、生成图片、转换后的大 JSON、`.exe` 和 `.zip` 直接提交到 git。

## 数据策略

Excel files, generated images, converted workbook data, build output, `.exe`, and `.zip` files are intentionally ignored by git.

Reason: these files are large, local, or derived from source/data. Put Excel files into the working folder only when running the tool locally.

## 规则维护

The canonical rule file is:

```text
excel_image_tool_final/rules/template_fill_rule.md
```

When extraction rules, viewer behavior, image generation, runtime steps, or delivery structure change, update `project_context_handoff.md` in the same change.
