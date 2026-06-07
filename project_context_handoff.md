# Excel Image Tool Context Handoff

As of 2026-05-18.

## Maintenance rule

This file is part of the delivery workflow.

Any future bug fix or feature change that affects rules, viewer behavior, image generation behavior, delivery structure, runtime steps, or known project status must update this file in the same turn.

2026-06-07 GitHub preparation note:

- Added repository-level `README.md`.
- Added repository-level `.gitignore`.
- Added repository-level `.gitattributes`.
- Added repository-level `AGENTS.md`.
- Added `docs/GITHUB_UPLOAD.md`.
- Git ignores local Excel files, converted workbook data, generated images, build output, `.exe`, and `.zip` artifacts.
- GitHub source upload should keep source code, templates, rules, scripts, and docs; build/release binaries should be produced locally when needed.
- Updated root `README.md` with normal usage, developer setup, build steps, and GitHub release/deployment guidance.
- Updated `docs/GITHUB_UPLOAD.md` with the local-tool deployment model and recommended GitHub Release packaging contents.
- 2026-06-07 correction: `excel_image_tool_release/` is a local release/package output and must not be tracked in the GitHub source repository. It was removed from git tracking with `git rm --cached` and added to `.gitignore`.

## 1. Project purpose

This project converts a multi-sheet Excel workbook into:

- script-friendly JSON/JS data
- a local HTML viewer
- image generation output based on a template

Main local delivery folder:

- `excel_image_tool_release`

GitHub tracking note:

- `excel_image_tool_release/` is ignored by git.
- It is kept locally for packaging/testing, but GitHub source should not track it.

Local source/build folder:

- `excel_image_tool_final`

## 2. Delivery entry files

In `excel_image_tool_release`:

- `ExcelImageTool.exe`
- `update_and_start.bat`
- `start_viewer.bat`
- `stop_viewer.bat`
- `setup_env.bat`
- `rules/template_fill_rule.md`

## 3. Current important rule set

### Product

- Source cell: `B1`
- Use only the product name
- Remove outer brackets like `【】`
- Remove capacity/spec/count text such as `50g/瓶`, `100ml`, `330g*5袋`, `10条装`, `2颗装`
- Preserve product connector `+` / `＋` when two products are joined
- Preserve model/version-like text such as `2.0`, `5.0`, `8JS4`

### Spec

- Source cell: dynamic A-column cell
- Find the row in column `A` whose text contains `规格`
- Use the next row in column `A` as the spec source cell
- If the source contains multiple spec candidates, choose the largest / most-complete one
- Example:
  - `20ml/瓶；20ml/瓶*2` -> `20ml/瓶*2`
- If no matching A-column `规格` label exists, fill `/`

### Page price

- Source cell: `A4`
- Extract from the `页面价` area
- Keep only the price amount itself
- Ignore the `到手价` section for this field
- If multiple page-price amounts exist, take the largest one
- Examples:
  - `300ml*3：660元/3瓶` -> `660元`
  - `168元/瓶` -> `168元`

### Detail coupon

- Source cell: `A4`
- Extract amount from `详情页领...券`
- Example:
  - `详情页领10元券` -> `10元`

### Buy count

- Source cell: `A4`
- Extract the `拍` quantity only
- Examples:
  - `拍1` -> `1`
  - `拍1瓶` -> `1瓶`

### Deposit

- Source cell: `A4`
- If multiple `定金` values exist, take the maximum value
- Ignore `免定金`
- Support both `50元` and shorthand `50`

### Tail payment

- Source cell: `A4`
- If multiple `尾款` values exist, take the maximum value
- Support both `350元` and shorthand `350`

### Category coupon

- Extract only when keywords exist:
  - `美妆券`
  - `生活券`
  - `美妆惊喜券`
  - `精致生活券`
  - `服饰天降券`
  - `天猫国际券`
- Keep coupon name and amount/threshold together
- Examples:
  - `生活券满200-16元`
  - `生活券200-15（减7.5元）`
  - `美妆券1000减400（减800）`

### Consumption coupon

- Output format is fixed as `预估减XX元`
- Example:
  - `88vip88折消费券（减24元）` -> `预估减24元`
- If multiple 88VIP coupon candidates exist in the same `A4`, take the largest discount value across all candidates
- If a range exists, take the larger boundary
- Example:
  - `88vip88折消费券（减10.78~16.78元）` -> `预估减16.78元`
  - `88vip88折消费券（减15.88-17.88）` -> `预估减17.88元`

### Other discount

- Only extract these kinds of items:
  - `购物金充值`
  - `购物金充`
  - `充购物金`
  - `购物金`
  - `红包补贴`
  - `天猫返现卡`
- Extract only the matching discount snippet itself
- Do not include same-line `88VIP消费券`, `免定金`, `到手价`, `最终到手价`, `含赠到手`
- Examples:
  - `红包补贴20元`
  - `充购物金800得845元(打95折)`
- If only a hint like `购物金到手价` appears with no numeric content, do not fill it

### Estimated final price

- Source cell: `A4`
- Find the final discount-stacked `到手...` short sentence
- If multiple final `到手...` candidates exist, take the last one
- If `到手` and the amount are split across lines, still treat the paragraph as one candidate
- Keep capacity/unit/weight/folded price text
- If `折合...` exists, force a line break before `折合`
- Do not recalculate price

### Text cleaning

- Hidden formatting control characters from Excel are stripped before display and image rendering
- This specifically fixes square-box glyphs caused by invisible Unicode marks such as `U+202C`

## 4. UI / behavior notes

- Search only matches sheet names, not cell content
- Viewer includes image preview and table view
- Generated image overwrite behavior is enabled for repeated clicks on the same sheet
- Generated image history is limited to 3 items in the UI
- Scroll behavior was customized so the hovered panel should consume mouse wheel scrolling
- The template system is now on the new 800x1131 layout
- The old template `输出模板.png` was removed
- The only template kept in both source and release is `image_templates/输出模板_new.jpg`
- The preview tuning panel now targets these fields:
  - `产品`
  - `规格`
  - `页面价`
  - `详情券`
  - `拍`
  - `定金`
  - `尾款`
  - `品类券`
  - `消费券`
  - `其他优惠`
  - `预估到手价`
- `预估到手价` is rendered with a yellow background in generated images
- `预估到手价` now starts on the same line after the template's `~` marker instead of starting below the label
- `预估到手价` field box was moved a bit further right to avoid covering the `~` marker
- `页面价` field box was nudged downward for better vertical alignment with the `页面价` label
- Default font sizes were bumped slightly for `页面价`, `定金`, `尾款`, `消费券`, and `拍`
- Font strategy now is:
  - `产品` / `预估到手价`: auto-adjust up to `40`
  - `规格`: auto-adjust within `18-32`
  - all other fields: auto-adjust with minimum `34`

## 5. Important source files

Core source files in `excel_image_tool_final`:

- `convert_excel_to_script_data.py`
- `generate_template_images.py`
- `launch_viewer.py`
- `serve_viewer.py`
- `update_excel_data.py`
- `rules/template_fill_rule.md`

Rule document policy:

- Keep only one canonical rule file: `rules/template_fill_rule.md`
- The old parallel note file `rules/rule_new.md` was removed to avoid divergence

## 6. Current delivery status

Already synced to delivery:

- `excel_image_tool_release/ExcelImageTool.exe`
- `excel_image_tool_release/rules/template_fill_rule.md`
- `excel_image_tool_release/image_templates/输出模板_new.jpg`
- `excel_image_tool_release/converted_excel/viewer.html`

Potentially stale until Excel is placed back and refreshed:

- `excel_image_tool_release/converted_excel/*`

Reason:

- There is still no Excel workbook in `excel_image_tool_release`, so a full reconvert has not been run from source Excel after the template upgrade.
- However, the existing `workbook_data.json/js` in both source and release were already patched in place from their stored raw `B1/A4` text, so the current viewer data also uses the new field structure.

## 7. Refresh workflow

To refresh data after putting a new Excel file into `excel_image_tool_release`:

1. Put the Excel file directly inside `excel_image_tool_release`
2. Run `update_and_start.bat`

This should:

- detect the Excel
- rebuild `converted_excel`
- start the local viewer

## 8. Recent bug fixes already implemented in code

- Title removes capacity/spec but keeps `+`
- Deposit and tail payment take the maximum value
- Category coupon keeps amount/threshold, not just the coupon name
- 88VIP coupon supports ranges and takes the larger boundary
- Other discount extracts only the actual discount snippet
- Final price supports multi-line `到手`
- Invisible Unicode formatting marks are stripped to prevent square-box glyphs in generated images
- New template field extraction now includes:
  - `规格`
  - `页面价`
  - `详情券`
  - `拍`
- The old template was removed and the new template layout is the only supported one
- Final-price line spacing was increased and its field box was moved to the right of `预估到手价~`
- `规格` now keeps only gift product name + spec and removes `到手` / `折合` / value-note text
- `规格` also drops summary-only lines such as `含赠...`, `正装量...`, and shipping-note lines
- If a spec line contains `折合`, everything from `折合` onward is discarded, including any product/spec text that appears after it
- Pure summary lines such as `（总价值...）`, `共到手...`, `共含赠...`, and empty short labels like `银管:` are discarded from `规格`
- If a short prefix label exists before a colon, like `银管:防晒霜50g*1`, the prefix is removed and only the actual product/spec text is kept
- For numbered gift lists such as `1、商品名 规格【价值...】`, the numbering and `【价值...】` tail are removed, leaving only product name + spec
- Price-stat lines like `银管：97.5元/支` inside a gift block are discarded from `规格`
- `消费券` now takes the maximum discount when multiple 88VIP coupon candidates appear
- `页面价` now outputs amount only from the `页面价` section, without spec suffixes, and takes the maximum amount if multiple prices appear
- `规格` source was moved from `A4` gift parsing to the A-column `规格` label's next-row cell
