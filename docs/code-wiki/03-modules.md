# 03. 主要模块职责

## 根目录模块

### `main.py`

主命令行入口。它只做入口层编排，不直接实现下载、识别或导出细节。

职责：

- 定义 CLI 参数。
- 解析 WMS 登录态路径。
- 初始化 `ocr_enabled` 等运行参数。
- 顺序执行 `tracked_download_pdf()`、`tracked_run_ocr()`、`tracked_extract_data()`。

### `config.py`

全局路径与默认配置。

职责：

- 识别源码运行模式和 PyInstaller 冻结模式下的 `BASE_DIR`。
- 记录当前工作目录 `CURRENT_WORK_DIR`。
- 从 `BASE_DIR/.env` 和当前工作目录 `.env` 加载环境变量。
- 定义默认目录：`logs/`、`pdf_downloads/`、`output/pdf/`。
- 定义默认识别参数：DPI、超时、最大页数。

### `utils.py`

轻量工具函数。

职责：

- 标准日志输出。
- 目录创建。
- 路径存在性检查。
- 冻结可执行文件模式下的退出暂停。

### `pipeline.py`

三阶段流水线的核心实现。

职责：

- 阶段 1：下载或扫描本地文件，并写 `01_download_manifest.json`。
- 阶段 2：调用识别层，并写 `02_ocr_results.json`。
- 阶段 3：调用导出层，并写 `03_extract_manifest.json`。
- 维护 `_intermediate/<output-name>/` 中间目录。

### `task_pipeline.py`

在 `pipeline.py` 外增加任务状态、重试和异常分类。

职责：

- 每个阶段最多重试 3 次。
- 生成稳定 `task_id`。
- 将下载文件绑定到任务。
- 根据识别结果判断 OCR 成功或失败。
- 对非 PDF、OCR 失败、导出失败写入状态。

### `task_state.py`

任务状态 JSONL 的底层读写。

职责：

- 生成 `task_state.jsonl` 默认路径。
- 用 UUID v5 按文件路径生成稳定任务 ID。
- 追加 JSONL 状态行。
- 从 manifest 建立文件路径到 task ID 的映射。

## 下载模块

### `pdf_download.py`

下载门面，负责选择下载模式。

职责：

- 判断是否需要下载，或直接使用本地 `--input-dir`。
- 生成批次名、下载目录和下载日志路径。
- 默认调用 `auto_download.py` HTTP 并发模式。
- `--browser-mode` 时调用 `batch_download_wms_pdfs.py`。
- HTTP 模式失败时，在有下载需求的情况下回退到浏览器模式。

### `auto_download.py`

默认 HTTP 并发下载实现。

职责：

- 优先复用已有 token。
- token 失效时用账号密码登录并刷新 `wms_storage_state.json`。
- 按仓库、状态、时间范围分页获取订单。
- 使用线程池并发下载面单。
- 写入下载日志 CSV。
- 追加下载 metadata JSONL。
- 发生 token 失效时刷新 session 并重试。

### `batch_download_wms_pdfs.py`

历史较完整的 WMS 下载脚本，兼容 HTTP 轻量模式和 Playwright 浏览器模式。

职责：

- 保存和复用 Playwright 登录态。
- 自动登录和选择仓库。
- 构造 WMS 列表查询 payload。
- 从 storage state 构造 HTTP session。
- 处理 WMS 接口签名和 `track-key`。
- 下载面单文件并校验 PDF 格式。
- 写入下载日志、PDF 校验日志、metadata JSONL 和 metadata 汇总 Excel。
- 支持手动页面筛选和勾选订单。

### `track_key.py`

WMS 接口签名辅助模块。

职责：

- 实现 WMS 侧兼容的 MD5/base64 计算。
- `generate_track_key()` 根据请求体和时间戳生成接口需要的 `track-key`。

## 识别模块

### `pdf_verify.py`

批量识别适配层。

职责：

- 将主入口参数转换为识别参数。
- 读取下载日志中的 WMS 追踪号。
- 加载 zxing-cpp。
- 加载 metadata 索引。
- 遍历输入行并调用 `barcode_verify_tracking.process_row()`。
- 捕获单文件异常并转为 `download_file_error` 结果。

### `barcode_verify_tracking.py`

识别和校验核心模块。

职责：

- 列出输入文件或读取 strict JSON。
- 读取 PDF 文字层。
- 渲染 PDF 页面为图片。
- 读取图片面单。
- 调用 zxing-cpp 反读条码。
- 从文件名、下载日志、PDF 文字层、条码中提取追踪号候选。
- 判断追踪号一致性与最终状态。
- 识别承运商和特殊模板。
- 判断是否需要 OCR 或人工复核。

## 导出模块

### `exporter.py`

业务导出核心。

职责：

- 将 `VerifyResult` 转成业务行。
- 从 metadata 补充订单号、物流渠道名称、客户、仓库等字段。
- 识别下载侧模板并与内容识别模板比对。
- 合并旧 Excel 的 `全部结果` 和 `简略版`。
- 写出 `全部结果`、`异常复核`、`下载不一致`、`统计汇总`、`简略版`。
- 写出业务增强 JSON。

### `scripts/update_brief_sheet.py`

补写或更新旧 Excel 的 `简略版`。

职责：

- 读取 `全部结果`。
- 生成简略行。
- 与已有 `简略版` 合并去重。
- 覆盖保存工作簿。

## 评估与构建模块

### `evaluate_samples.py`

样本识别率评估工具。

职责：

- 扫描 `samples/<expected>/` 下的面单样本。
- 调用 `detect_template()` 识别模板。
- 统计 `0024-01`、`0024-02`、`CBT`、`CBS` 等样本通过率。
- 可通过 `--fail-on-mismatch` 在 CI/本地检查中失败退出。

### `build_exe.bat` 与 `.spec`

PyInstaller 打包入口。

职责：

- 将命令行工具或中文命名工具打包为 Windows 可执行文件。
- 冻结模式下路径由 `config.py` 和下载脚本中的 `sys.frozen` 分支处理。

## `WMS_TOOL/` GUI 模块

### `WMS_TOOL/gui_main.py`

Tkinter 桌面窗口。

职责：

- 提供开始时间、结束时间、仓库、并发数输入。
- 启动后台任务。
- 每秒轮询任务状态。
- 展示日志和进度条。

### `WMS_TOOL/task_runner.py`

内存任务管理器。

职责：

- 为 GUI 任务生成 UUID。
- 后台线程执行 `run_pipeline()`。
- 收集状态、进度、日志、结果和错误。
- 为 GUI 提供 `get_task()` 查询。

### `WMS_TOOL/worker.py`

GUI 到核心流水线的参数适配层。

职责：

- 读取默认配置 `WMS_TOOL/config/default.json`。
- 加载 `.env`。
- 为每次运行生成 run name。
- 设置 GUI 模式下的输出、下载、日志、metadata 路径。
- 可先查询 WMS 总数，再执行三阶段流水线。

### `WMS_TOOL/core/*`

轻量封装层：

- `downloader.py`：调用 `pipeline.download_pdf()`。
- `parser.py`：调用 `pipeline.run_ocr()`。
- `validator.py`：调用 `pipeline.extract_data()`。
- `wms_client.py`：查询 WMS 列表总数，用于 GUI 运行前估算。

## `scripts/` 脚本

| 文件 | 职责 |
| --- | --- |
| `run_hourly_batches.ps1` | 将某一天拆成小时窗口，逐小时运行 `main.py`。 |
| `register_hourly_batches_task.ps1` | 注册每天或每周运行的小时批处理定时任务。 |
| `run_previous_hour_batch.ps1` | 根据触发时间计算上一小时窗口并运行。 |
| `register_previous_hour_task.ps1` | 注册每日多触发点的上一小时滚动任务。 |
| `run_window_batches_to_master.ps1` | 从开始时间到结束时间按小时运行，并追加到总表。 |
| `send_feishu_notification.py` | 发送飞书机器人通知。 |
| `send_feishu_file.py` | 上传并发送飞书文件消息。 |

