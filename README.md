# PDF_DDD

WMS PDF 面单下载、面单类型识别、条码校验和 Excel/JSON 导出工具。

这个项目当前主流程是：

1. 从 WMS 按时间、仓库、状态、渠道下载 PDF 面单，默认走 HTTP 并发下载。
2. 识别 PDF 里的承运商、追踪号、面单类型和 0024 子类型。
3. 结合下载日志和 `metadata.jsonl` 补充渠道、客户、仓库等信息。
4. 导出 Excel 和 JSON，方便人工复核和后续统计。

## 环境准备

建议使用 Python 3.11 或更高版本。

```powershell
git clone https://github.com/2003812zhourui-max/PDF_DDD.git
cd PDF_DDD
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
```

`--ocr` 是可选增强能力。如果新电脑要用 OCR，需要额外安装 Tesseract OCR，并确保 `tesseract.exe` 在系统 PATH 里。

账号密码不要写进代码，也不要提交到 Git。推荐在 PowerShell 当前窗口设置：

```powershell
$env:WMS_USERNAME="你的WMS账号"
$env:WMS_PASSWORD="你的WMS密码"
```

也可以复制 `.env.example` 为 `.env` 做本地记录，`.env` 已经被 `.gitignore` 忽略。

## 常用命令

下载并识别一批面单：

```powershell
python main.py --start-time "2026-06-10 00:00:00" --end-time "2026-06-10 23:59:59" --wh-codes US02 --statuses 15 --workers 5 --limit 200 --output-name download_200_check
```

每次强制重新下载、重新识别，并输出到全新文件夹：

```powershell
$run = "run_0611_0000_0100_" + (Get-Date -Format "yyyyMMdd_HHmmss")
python main.py --start-time "2026-06-11 00:00:00" --end-time "2026-06-11 01:00:00" --wh-codes US02 --statuses "10,15,20,30" --workers 5 --force --output-name $run --output-dir "output\pdf\$run"
```

只识别本地已有 PDF：

```powershell
python main.py --input-dir pdf_downloads --download-log logs\download_log.csv --output-dir output\pdf --output-name local_check
```

按物流渠道下载：

```powershell
python main.py --start-time "2026-06-10 00:00:00" --end-time "2026-06-10 23:59:59" --wh-codes US02 --statuses 15 --channel TikTok-CBT-US --workers 5 --output-name cbt_check
```

使用旧浏览器模式下载：

```powershell
python main.py --browser-mode --start-time "2026-06-10 00:00:00" --end-time "2026-06-10 23:59:59" --wh-codes US02
```

跑样本识别率评估：

```powershell
python evaluate_samples.py --samples-dir samples --output output\evaluation\sample_eval.json
```

## 当前识别规则

面单类型只重点识别三类：

- `0024`
- `CBT`
- `CBS`

业务输出规则：

- 命中 `CBT` 就输出 `CBT`。
- 命中 `CBS` 就输出 `CBS`。
- 命中 `0024` 就输出 `0024`。
- 没命中这三类的，按 `普通面单` 处理，继续识别承运商和追踪号。

`0024` 内部会尽量区分 `0024-01` 和 `0024-02`，但业务表里主要看是否为 `0024`。之前容易把 `YWLAX...` 这类 Yanwen 单误判成 0024，现在已经去掉了 `LAX => 0024` 的旧规则。

## 主要文件

- `main.py`：命令行入口，串起下载、识别、导出。
- `pdf_download.py`：下载模式选择，默认 HTTP 并发，`--browser-mode` 走旧浏览器模式。
- `auto_download.py`：HTTP 并发下载实现，负责登录、查询订单、下载 PDF、写 metadata。
- `barcode_verify_tracking.py`：核心条码识别、承运商识别、面单类型识别。
- `pdf_verify.py`：批量处理 PDF，并读取下载日志、metadata 丰富结果。
- `exporter.py`：导出 Excel/JSON。
- `evaluate_samples.py`：用 `samples/` 下的样本评估识别率。

## 不提交到 Git 的文件

下面这些属于登录态、业务数据、下载结果或构建产物，不进 Git：

- `.env`
- `.wms_token_cache.json`
- `wms_storage_state.json`
- `logs/`
- `output/`
- `pdf_downloads/`
- `downloads/`
- `samples/**/*.pdf`
- `*.csv`
- `*.xlsx`
- `*.zip`
- `build*/`
- `dist*/`
- `__pycache__/`

换电脑开发时，把本地敏感迁移包解压到项目根目录即可。详细步骤见 `MIGRATION.md`。
