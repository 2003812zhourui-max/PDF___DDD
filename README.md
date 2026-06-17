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

使用浏览器兼容模式下载：

```powershell
python main.py --browser-mode --start-time "2026-06-10 00:00:00" --end-time "2026-06-10 23:59:59" --wh-codes US02
```

跑样本识别率评估：

```powershell
python evaluate_samples.py --samples-dir samples --output output\evaluation\sample_eval.json
```

## 按小时批量运行

跑某一天 24 个小时，每个小时一个独立下载目录、独立下载日志、独立 Excel/JSON 输出目录：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_hourly_batches.ps1 -BatchDate 2026-06-11 -WhCodes US02 -Statuses "10,15,20,30" -Workers 5
```

只跑其中几个小时，例如 0 点到 2 点：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_hourly_batches.ps1 -BatchDate 2026-06-11 -StartHour 0 -EndHour 2 -WhCodes US02 -Statuses "10,15,20,30" -Workers 5
```

先检查会跑哪些时间段和命令，但不真正下载：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_hourly_batches.ps1 -BatchDate 2026-06-11 -StartHour 0 -EndHour 24 -DryRun
```

这个脚本默认带 `--force`，所以每一小时都会重新下载、重新识别。输出示例：

```text
pdf_downloads\run_20260611_0000_0059_20260611_120000
logs\run_20260611_0000_0059_20260611_120000_download_log.csv
output\pdf\run_20260611_0000_0059_20260611_120000\run_20260611_0000_0059_20260611_120000.xlsx
```

注册每天自动运行的 Windows 定时任务，默认每天 02:30 跑昨天完整 24 小时：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register_hourly_batches_task.ps1 -Schedule Daily -At "02:30" -WhCodes US02 -Statuses "10,15,20,30" -Workers 5
```

注册每周一自动运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register_hourly_batches_task.ps1 -Schedule Weekly -DaysOfWeek Monday -At "02:30" -WhCodes US02 -Statuses "10,15,20,30" -Workers 5
```

手动触发定时任务：

```powershell
Start-ScheduledTask -TaskName "PDF_DDD hourly batches"
```

## 每天按小时定时跑上一小时

如果你要每天滚动执行：凌晨 1 点跑 `00:00:00 ~ 00:59:59`，凌晨 2 点跑 `01:00:00 ~ 01:59:59`，一直跑到晚上 22 点，可以注册这个任务：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register_previous_hour_task.ps1 -StartAt "01:00" -EndAt "22:30" -WhCodes US02 -Statuses "10,15,20,30" -Workers 5
```

先预演会注册哪些触发点，不真正注册：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register_previous_hour_task.ps1 -StartAt "01:00" -EndAt "22:30" -DryRun
```

它会注册每天这些触发时间：

```text
01:00, 02:00, 03:00, ..., 22:00
```

每次触发时只跑上一小时，并且强制重新下载、重新识别，输出独立目录。例如：

```text
01:00 -> 2026-06-11 00:00:00 ~ 2026-06-11 00:59:59
02:00 -> 2026-06-11 01:00:00 ~ 2026-06-11 01:59:59
22:00 -> 2026-06-11 21:00:00 ~ 2026-06-11 21:59:59
```

手动测试“上一小时”脚本，不真正下载：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_previous_hour_batch.ps1 -RunAt "2026-06-11 01:00:00" -DryRun
```

如果时间窗口是前一天 `22:30` 到当天 `22:30`，每小时滚动一次，可以这样注册：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register_previous_hour_task.ps1 -StartAt "23:30" -EndAt "22:30" -WhCodes US02 -Statuses "10,15,20,30" -Workers 5
```

触发关系是：

```text
23:30 -> 22:30:00 ~ 23:29:59
00:30 -> 23:30:00 ~ 00:29:59
...
22:30 -> 21:30:00 ~ 22:29:59
```

## Excel 简略版

每次导出 Excel 时，会保留原来的 `全部结果`、`异常复核`、`统计汇总` 逻辑，并额外生成或更新 `简略版` sheet。

`简略版` 只包含：

- `追踪号`
- `承运商`
- `最终状态`
- `条码是否一致`

如果同一个 Excel 文件里已经存在 `简略版`，程序会先读取旧数据，再追加本次小时数据，并按追踪号去重，最后覆盖保存原 Excel 文件。

旧 Excel 也可以单独补 `简略版`：

```powershell
python scripts\update_brief_sheet.py "output\pdf\你的文件.xlsx"
```

## 总表模式

如果不想每小时一个 Excel，而是从开始时间一路往下跑，把所有小时结果追加到同一个总表，用这个脚本：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_window_batches_to_master.ps1 `
  -StartTime "2026-06-10 22:30:00" `
  -EndTime "2026-06-11 22:30:00" `
  -MasterOutputName "master_20260610_2230_to_20260611_2230" `
  -WhCodes US02 `
  -Statuses "10,15,20,30" `
  -Workers 5
```

这个模式下：

- PDF 下载目录仍按小时分开，避免文件互相覆盖。
- 下载日志仍按小时分开。
- Excel 固定写入同一个总表：`output\pdf\master_...\master_....xlsx`。
- `全部结果` 会读取旧数据后追加本小时数据，并按追踪号/文件路径去重。
- `简略版` 也会增量追加并按追踪号去重。

先预演，不真正下载：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_window_batches_to_master.ps1 `
  -StartTime "2026-06-10 22:30:00" `
  -EndTime "2026-06-11 22:30:00" `
  -MasterOutputName "master_20260610_2230_to_20260611_2230" `
  -DryRun
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
- `pdf_download.py`：下载模式选择，默认 HTTP 并发，`--browser-mode` 走浏览器兼容模式。
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
