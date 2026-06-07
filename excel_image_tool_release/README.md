# Excel 图片生成工具

## 第一步

双击 `setup_env.bat` 检查环境。

## 启动

双击 `start_viewer.bat`，自动打开浏览器。

不用时，双击 `stop_viewer.bat`。

## 换新 Excel

1. 把新的 Excel 文件放到本文件夹。
2. 双击 `update_and_start.bat`。

## 生成图片

页面左侧选择 sheet，在 `图片预览` 中点击 `生成当前图片`。

生成结果保存在 `generated_images/`。

## 目录

- `image_templates/`: 新版图片模板 `输出模板_new.jpg`。
- `generated_images/`: 生成图片输出。
- `converted_excel/`: 页面和转换数据。
- `rules/`: 字段提取与填图规则。
