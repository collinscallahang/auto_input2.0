# 货拉拉 Excel 批量报价工具

一个 Windows 桌面工具，用于导入任意结构相近的 Excel，自动识别供应商、发货地址、到货地址、距离和可报价车型列，确认后打开货拉拉同城下单页批量查询，并把 `总里程` 与 `运费一口价` 写回新的结果 Excel。

## 运行

```powershell
python -m pip install -r requirements.txt
python -m huolala_quote_tool
```

首次使用网页自动化前，需要安装 Playwright 浏览器驱动：

```powershell
python -m playwright install chromium
```

工具会优先尝试复用本机 Microsoft Edge；如果失败，会退回 Playwright Chromium。登录状态保存在用户目录下的 `.huolala_quote_tool/browser_profile`，通常只需要人工登录一次。

## 使用流程

1. 点击 `选择 Excel`。
2. 在预览区确认自动识别出的必需字段和车型列。
3. 选择复跑策略：
   - `只补空白价格`：默认策略，只处理空白且未标记失败的车型价格。
   - `全部重跑`：覆盖已存在的距离和报价。
   - `只重跑失败项`：只处理带有失败批注的单元格。
4. 点击 `开始`。
5. 浏览器打开后人工登录货拉拉，再回到工具确认继续。

程序不会覆盖原 Excel，会生成：

```text
原文件名_货拉拉报价结果_YYYYMMDD_HHMMSS.xlsx
```

并且每完成一行保存一次。

## 车型规则表

车型规则维护在：

```text
config/vehicle_rules.csv
```

默认规则表可以用 Excel 打开编辑，字段为 `规则名称`、`Excel表头关键词`、`网页车长`、`网页车型要求`、`启用`。多个表头关键词或车型要求用英文分号 `;` 分隔。

例如 `厢式货车&飞翼车 9.6` 会映射到网页 `9米6`，并勾选 `厢式货车`、`飞翼车`。

只有匹配到规则的表头才会被当作车型价格列。疑似车型但没有匹配规则的列会在预览里提示，不会修改。

兼容旧版 `config/vehicle_rules.json`。如果 `vehicle_rules.csv` 存在，程序优先读取 CSV。

## Windows 免安装打包

```powershell
python -m pip install -r requirements.txt
python scripts\build_windows_portable.py
```

生成目录：

```text
dist\HuolalaQuoteToolPortable
```

目录内包含主程序、`config` 规则表、`logs` 日志目录和使用说明。复制整个目录到另一台 Windows 电脑即可运行；工具会优先使用系统自带 Microsoft Edge。

## 验证

```powershell
python -m unittest discover -s tests
```

使用已有结果样本模拟一次网页返回和写回流程：

```powershell
python scripts\simulate_recorded_sample.py "D:\codex\auto_input2.0\input\样本 - F4.xlsx"
```

用真实货拉拉网页做单行 smoke test：

```powershell
python scripts\run_real_sample_smoke.py "D:\codex\auto_input2.0\input\样本 - F4.xlsx" --row 3 --vehicles 1
```

脚本会打开货拉拉同城下单页，人工登录并确认后，自动填写第 3 行地址，读取 `总里程`，并只测试第一个匹配车型列。结果会写入 `real_test_output` 下的新 Excel 副本。

当前自动化测试覆盖：

- 必需字段不依赖固定列位置。
- 车型父表头和短车长表头可以联合识别。
- 无关列保留不处理。
- 车型列只按规则识别。
- `总里程145公里` 解析为 `145`。
- `运费一口价 1072.64元` 解析为 `1072.64`，不会读取优惠后的 `总计`。

## 注意

货拉拉网页结构可能调整。网页交互逻辑集中在 `huolala_quote_tool/browser.py`，如果页面控件文案或选择器变化，优先修改这一层。
