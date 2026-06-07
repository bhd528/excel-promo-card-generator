# Excel 图片生成工具

## 第一步：环境准备

双击 `setup_env.bat`。

- 如果有 `ExcelImageTool.exe`，会直接通过，不需要安装 Python。
- 如果没有 exe，但机器有 Python，会自动安装所需依赖。
- 如果没有 exe 也没有 Python，说明交付包不完整，需补上 `ExcelImageTool.exe`。

## 启动页面

双击 `start_viewer.bat`，会自动打开浏览器。

不用时，双击 `stop_viewer.bat` 停止服务。

## 换新 Excel

1. 把新的 Excel 文件放到本文件夹。
2. 双击 `update_and_start.bat`。

脚本会选择最新放入的 Excel，更新页面数据，然后打开浏览器。

## 生成图片

在页面左侧选择 sheet，在 `图片预览` 中点击 `生成当前图片`。

生成结果保存在 `generated_images/`。

## 目录

- `image_templates/`: 新版图片模板 `输出模板_new.jpg`。
- `generated_images/`: 生成图片输出。
- `converted_excel/`: 页面和转换后的数据。
- `rules/`: 字段提取与填图规则。
