# 开发环境迁移说明

这次操作可以叫“项目迁库”或“开发环境迁移”。代码放到 GitHub，登录态、账号、下载日志、metadata 这些敏感或业务数据不上传，单独打包成本地环境迁移包。

## 在另一台电脑继续开发

1. 克隆代码：

```powershell
git clone https://github.com/2003812zhourui-max/PDF_DDD.git
cd PDF_DDD
```

2. 创建 Python 环境：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
```

3. 恢复本地敏感文件：

把 `PDF_D_sensitive_migration_*.zip` 解压到项目根目录。建议保留目录结构。

常见文件用途：

- `wms_storage_state.json`：Playwright 浏览器登录态，旧浏览器模式可能会用。
- `.wms_token_cache.json`：HTTP 下载 token 缓存，如果存在可以减少重新登录。
- `.env`：本地账号密码环境变量记录，如果你自己创建了才会有。
- `logs/download_log.csv`：下载日志，用于本地 PDF 和 WMS 单号对照。
- `output/download_label_metadata.jsonl`：下载阶段保存的订单 metadata，用于识别结果补充承运商、渠道、客户代码。

如果迁移包里没有账号密码，就在新电脑 PowerShell 里手动设置：

```powershell
$env:WMS_USERNAME="你的WMS账号"
$env:WMS_PASSWORD="你的WMS密码"
```

4. 试运行：

```powershell
python main.py --input-dir pdf_downloads --download-log logs\download_log.csv --output-name local_check
```

如果新电脑没有旧 PDF，就直接下载一批新的：

```powershell
python main.py --start-time "2026-06-10 00:00:00" --end-time "2026-06-10 23:59:59" --wh-codes US02 --statuses 15 --workers 5 --limit 200 --output-name download_200_check
```

## 给 AI 的项目背景

可以这样描述：

```text
这是一个 WMS PDF 面单下载和识别项目。主入口是 main.py。
默认流程：HTTP 并发下载 PDF -> 识别条码/承运商/面单类型 -> 读取 metadata.jsonl 丰富渠道和客户信息 -> 导出 Excel/JSON。

重点逻辑在 barcode_verify_tracking.py：
- CBT 命中就输出 CBT。
- CBS 命中就输出 CBS。
- 0024 命中就输出 0024，并尽量内部区分 0024-01/0024-02。
- 普通面单不需要人工确认，只要识别承运商和追踪号。
- 不要再用 LAX 判断 0024，因为 Yanwen 的 YWLAX 会误判。

下载逻辑在 auto_download.py 和 pdf_download.py：
- 默认 HTTP 并发。
- --browser-mode 才使用旧浏览器下载。
- WMS_USERNAME/WMS_PASSWORD 或 --username/--password 提供账号密码。
```

## 注意事项

- 不要把 `wms_storage_state.json`、`.wms_token_cache.json`、PDF、日志、Excel 输出提交到 GitHub。
- 如果下载报登录失效，重新设置账号密码再跑。
- 如果要提升识别率，优先往 `samples/0024-01`、`samples/0024-02`、`samples/CBT`、`samples/CBS` 放真实错例，再跑 `evaluate_samples.py`。
