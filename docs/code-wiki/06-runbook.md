# 06. 运行、部署与维护手册

## 环境准备

推荐 Python 3.11 或更高版本。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
```

如需启用 `--ocr`，还需要安装 Tesseract OCR，并确保 `tesseract.exe` 在系统 PATH 中。

## WMS 凭据

推荐用环境变量或本地 `.env`，不要把账号密码写入源码或提交到 Git。

PowerShell 临时设置：

```powershell
$env:WMS_USERNAME="你的WMS账号"
$env:WMS_PASSWORD="你的WMS密码"
```

也可以复制 `.env.example` 为 `.env`，本地填写后保持不提交。

## 常用命令

### 下载并识别一批面单

```powershell
python main.py `
  --start-time "2026-06-10 00:00:00" `
  --end-time "2026-06-10 23:59:59" `
  --wh-codes US02 `
  --statuses "10,15,20,30" `
  --workers 8 `
  --output-name run_20260610
```

### 限量抽样

```powershell
python main.py `
  --start-time "2026-06-10 00:00:00" `
  --end-time "2026-06-10 23:59:59" `
  --wh-codes US02 `
  --statuses 15 `
  --workers 5 `
  --limit 200 `
  --output-name sample_200
```

`--limit 0` 表示不限量。

### 只识别本地已有文件

```powershell
python main.py `
  --input-dir pdf_downloads `
  --download-log logs\download_log.csv `
  --output-dir output\pdf `
  --output-name local_check
```

### 按物流渠道筛选

```powershell
python main.py `
  --start-time "2026-06-10 00:00:00" `
  --end-time "2026-06-10 23:59:59" `
  --wh-codes US02 `
  --statuses 15 `
  --channel TikTok-CBT-US `
  --workers 5 `
  --output-name cbt_check
```

### 强制重新下载

```powershell
python main.py `
  --start-time "2026-06-11 00:00:00" `
  --end-time "2026-06-11 01:00:00" `
  --wh-codes US02 `
  --statuses "10,15,20,30" `
  --workers 5 `
  --force `
  --output-name run_0611_0000_0100
```

### 使用浏览器兼容模式

```powershell
python main.py `
  --browser-mode `
  --start-time "2026-06-10 00:00:00" `
  --end-time "2026-06-10 23:59:59" `
  --wh-codes US02
```

## 批处理脚本

### 按小时跑一天

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_hourly_batches.ps1 `
  -BatchDate 2026-06-11 `
  -WhCodes US02 `
  -Statuses "10,15,20,30" `
  -Workers 5
```

### 预演小时批处理

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_hourly_batches.ps1 `
  -BatchDate 2026-06-11 `
  -StartHour 0 `
  -EndHour 24 `
  -DryRun
```

### 跑上一小时

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_previous_hour_batch.ps1 `
  -RunAt "2026-06-11 01:00:00" `
  -WhCodes US02 `
  -Statuses "10,15,20,30" `
  -Workers 5
```

### 注册上一小时定时任务

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register_previous_hour_task.ps1 `
  -StartAt "01:00" `
  -EndAt "22:30" `
  -WhCodes US02 `
  -Statuses "10,15,20,30" `
  -Workers 5
```

手动触发：

```powershell
Start-ScheduledTask -TaskName "PDF_DDD previous hour batch"
```

### 时间窗口总表模式

```powershell
powershell -ExecutionPolicy Bypass -File scripts\run_window_batches_to_master.ps1 `
  -StartTime "2026-06-10 22:30:00" `
  -EndTime "2026-06-11 22:30:00" `
  -MasterOutputName "master_20260610_2230_to_20260611_2230" `
  -WhCodes US02 `
  -Statuses "10,15,20,30" `
  -Workers 5
```

## GUI 运行

```powershell
python WMS_TOOL\gui_main.py
```

GUI 默认提供：

- 开始时间
- 结束时间
- 仓库
- 并发数
- 进度条
- 日志窗口

GUI 输出默认在 `WMS_TOOL/output/` 和 `WMS_TOOL/logs/` 下。

## 样本评估

目录约定：

```text
samples/
  0024-01/
  0024-02/
  CBT/
  CBS/
  UNKNOWN/
```

运行：

```powershell
python evaluate_samples.py `
  --samples-dir samples `
  --output output\evaluation\sample_eval.json
```

如需识别不匹配时返回失败退出码：

```powershell
python evaluate_samples.py --fail-on-mismatch
```

## 构建 EXE

```powershell
build_exe.bat
```

或直接使用 PyInstaller spec：

```powershell
pyinstaller pdf_label_pipeline.spec
```

冻结模式下，`config.py` 会把 `BASE_DIR` 设为可执行文件所在目录。

## 排障指南

### HTTP 下载提示账号密码缺失

现象：`HTTP 并发模式需要账号密码`。

处理：

- 设置 `WMS_USERNAME` 和 `WMS_PASSWORD`。
- 或传入 `--username`、`--password`。
- 更推荐环境变量或 `.env`，避免命令历史留下密码。

### token 或登录态失效

处理：

- 确认账号密码可用。
- 默认 HTTP 模式会尝试自动登录并刷新登录态。
- 必要时加 `--browser-mode` 走浏览器模式重新生成登录态。

### Excel 中物流渠道名称为空

优先检查：

```powershell
Test-Path output\download_label_metadata.jsonl
```

原因通常是 metadata JSONL 未写入、路径不对，或本次运行未把 metadata 传给识别/导出阶段。

### 非 PDF 来源被当成成功

正确规则：应以 WMS 原始 `source_fields.fileName` 或 `source_file_format` 判断来源格式，而不是只看本地保存路径后缀。

当前异常状态应为：

```text
download_file_error
```

并进入 `异常复核`。

### 条码识别失败

排查方向：

- 是否安装 `zxing-cpp`。
- PDF 是否能被 PyMuPDF 打开。
- 是否需要提高 `--dpi`。
- 是否需要启用 `--rotations`。
- 是否是图片面单，需要人工复核或 OCR。

### 模板识别误判

维护原则：

- 优先往 `samples/` 加真实错例。
- 用 `evaluate_samples.py` 验证。
- 不要用 `LAX` 这类容易误伤 Yanwen 单号的规则判断 `0024`。

### 定时任务存在但没有运行

检查：

```powershell
Get-ScheduledTaskInfo -TaskName "PDF_DDD previous hour batch"
```

也可手动触发：

```powershell
Start-ScheduledTask -TaskName "PDF_DDD previous hour batch"
```

## 开发改动建议

- 改下载字段时，同步检查 `auto_download.py`、`batch_download_wms_pdfs.py`、`exporter.py`。
- 改识别状态时，同步检查 `barcode_verify_tracking.py`、`exporter.py`、`task_pipeline.py`。
- 改 Excel 列时，同步检查 `BUSINESS_COLUMNS`、`BRIEF_COLUMNS` 和 `scripts/update_brief_sheet.py`。
- 改 CLI 参数时，同步检查 PowerShell 脚本和 `WMS_TOOL.worker.make_args()`。
- 涉及真实 WMS 数据时，不要提交运行产物。

