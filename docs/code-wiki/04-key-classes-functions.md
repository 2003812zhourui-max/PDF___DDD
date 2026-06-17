# 04. 关键类与函数说明

本文只列维护时最常接触的关键对象和函数。内部 MD5、Playwright DOM 点击、Excel 样式等辅助函数不逐一展开。

## 关键数据类

### `pdf_download.DownloadResult`

```python
@dataclass
class DownloadResult:
    input_dir: Path
    download_log: Path
    skipped: bool
```

下载阶段的标准返回对象。

| 字段 | 说明 |
| --- | --- |
| `input_dir` | 本次识别要读取的面单目录。 |
| `download_log` | 下载日志 CSV。 |
| `skipped` | 是否跳过 WMS 下载，直接使用本地文件。 |

### `barcode_verify_tracking.Candidate`

追踪号候选。

| 字段 | 说明 |
| --- | --- |
| `carrier` | 推断承运商。 |
| `value` | 候选追踪号。 |
| `source` | 来源，例如 `filename`、`pdf_text`、`barcode`。 |
| `score` | 候选评分。 |
| `context` | 候选出现位置附近的文本。 |

### `barcode_verify_tracking.TemplateMatch`

模板识别结果。

| 字段 | 说明 |
| --- | --- |
| `template_code` | 模板类型，例如 `0024`、`CBT`、`CBS`。 |
| `template_sub_code` | 子类型，例如 `01`、`02`。 |
| `template_marker` | 命中的原始标记。 |
| `template_source` | 命中来源，例如文字层、条码文本、OCR 区域。 |
| `template_confidence` | 置信度。 |

### `barcode_verify_tracking.VerifyResult`

识别层的核心输出对象。它是导出层和中间 JSON 的主要数据结构。

字段可按类别理解：

| 类别 | 代表字段 | 说明 |
| --- | --- | --- |
| 文件信息 | `file_name`, `file_path`, `file_type`, `file_format`, `is_image_label` | 文件来源和格式。 |
| 承运商与模板 | `carrier`, `last_mile_carrier`, `template_code`, `template_sub_code`, `template_marker`, `carrier_display` | 承运商和模板识别。 |
| 追踪号来源 | `source_tracking`, `filename_tracking`, `pdf_text_tracking`, `barcode_tracking`, `ocr_tracking` | 不同来源识别出的追踪号。 |
| 下载字段 | `download_order_no`, `download_wave_no`, `download_warehouse` | 下载日志或 metadata 中的订单字段。 |
| 校验状态 | `verify_status`, `verify_status_zh`, `confidence`, `confidence_zh`, `reason`, `reason_zh`, `need_review` | 最终状态、置信度和复核标记。 |
| 条码信息 | `barcode_success`, `barcode_page_numbers`, `barcode_matches_source`, `barcode_matches_pdf_text`, `decoded_formats`, `decoded_raw_values` | 条码反读结果。 |
| OCR 信息 | `ocr_needed`, `ocr_reason`, `ocr_priority` | 是否需要进一步 OCR。 |
| 运行统计 | `processing_seconds`, `page_count`, `recognized_at` | 性能和页数。 |
| metadata 字段 | `meta_carrier`, `meta_channel`, `meta_channel_code`, `meta_customer_code`, `meta_source`, `meta_carrier_conflict`, `meta_carrier_note` | 下载阶段补充字段。 |

### `exporter.TemplateDecision`

导出层的业务模板判断对象。

| 字段 | 说明 |
| --- | --- |
| `carrier` | 业务展示承运商。 |
| `label_template` | 业务面单类型，默认 `普通面单`。 |
| `template_subdivision` | 模板细分，例如 `0024` 的 `01/02`。 |
| `matched_text` | 命中文本。 |
| `matched_rule` | 命中规则说明。 |
| `source` | 判断来源。 |
| `confidence` | 置信度。 |
| `note` | 业务备注。 |
| `need_review` | 是否需要人工复核。 |

`recognized` 属性用于判断是否命中特殊模板：`0024`、`CBT`、`CBS`。

### `evaluate_samples.SampleEvalResult`

样本评估结果对象，用于输出样本识别率报告。

## 主入口函数

### `main.parse_args()`

定义命令行参数。重要参数包括：

| 参数 | 说明 |
| --- | --- |
| `--start-time`, `--end-time` | WMS 查询时间范围。 |
| `--wh-codes` | 仓库代码，逗号分隔。 |
| `--statuses` | WMS 状态码，默认 `15`。 |
| `--input-dir` | 使用本地面单目录时传入。 |
| `--download-log` | 下载日志 CSV。 |
| `--output-dir`, `--output-name` | 输出目录和文件名前缀。 |
| `--workers` | HTTP 并发下载线程数。 |
| `--limit` | 抽样处理数量，`0` 表示不限。 |
| `--channel` | 物流渠道筛选。 |
| `--metadata` | metadata JSONL 路径。 |
| `--task-state` | 任务状态 JSONL 路径。 |
| `--browser-mode` | 使用浏览器兼容模式下载。 |
| `--force` | 强制重新下载。 |

### `main.resolve_storage_state_path(path_text)`

按优先级查找 WMS 登录态：

1. 用户传入路径。
2. `BASE_DIR/wms_storage_state.json`。
3. 当前工作目录 `wms_storage_state.json`。

### `main.main()`

主程序执行函数：

1. 解析参数。
2. 设置 `args.storage_state`。
3. 调用三阶段追踪流水线。
4. 捕获异常并返回退出码。

## 流水线函数

### `pipeline.download_pdf(args)`

阶段 1。它决定是下载面单还是使用本地输入目录，并写出下载 manifest。

核心逻辑：

- 如果 `args.strict_json` 存在，直接报错，因为此阶段只接受文件输入。
- 如果 `--input-dir` 且无下载参数，则跳过下载。
- 否则调用 `pdf_download.run_download()`。
- 优先从下载日志收集成功文件路径，失败时扫描输入目录。

### `pipeline.run_ocr(args, download_manifest_path)`

阶段 2。读取下载 manifest，调用 `verify_pdfs()`，并写出 OCR 结果 JSON。

### `pipeline.extract_data(args, ocr_results_path)`

阶段 3。恢复 `VerifyResult`，调用 `export_results()`，并写出导出 manifest。

### `task_pipeline.retry_step(stage, action)`

对阶段动作做最多 3 次重试。失败日志包含阶段名和尝试次数。

### `task_pipeline.attach_tasks_to_manifest(manifest_path, state_path)`

为下载 manifest 中每个文件生成 `task_id`，并写入 `DOWNLOAD CREATED` 和 `DOWNLOAD DOWNLOADED`。

### `task_pipeline.attach_tasks_to_ocr_results(ocr_path, download_manifest_path, state_path)`

将识别结果和任务 ID 绑定，并根据 `verify_status` 写入 OCR 状态。`download_file_error` 会进入失败状态。

### `task_pipeline.classify_ocr_failure(result)`

将识别失败分为：

- `NOT_PDF`
- `OCR_FAILED`

判断依据包括错误文本和文件后缀。

## 下载函数

### `pdf_download.run_download(args)`

下载模式总入口。

规则：

- 本地输入目录且没有下载参数：跳过下载。
- `--browser-mode`：使用浏览器模式。
- 默认：使用 HTTP 并发模式。
- HTTP 失败后，如果是下载场景，回退到浏览器模式。

### `pdf_download.run_download_http(args)`

调用 `auto_download.main()` 的进程内入口。它会构造底层参数，并屏蔽日志中的密码。

### `pdf_download.run_download_browser(args)`

调用 `batch_download_wms_pdfs.py`。非冻结模式下通过子进程运行；冻结模式下直接导入执行。

### `auto_download.try_existing_token(wh_code)`

从现有 storage state 中读取 token 并测试是否有效，避免不必要登录。

### `auto_download.wms_login(username, password)`

使用账号密码登录 WMS，返回 token 和 tenant code。

### `auto_download.update_storage_state(token, tenant_code, username)`

刷新本地 `wms_storage_state.json`，让后续 HTTP session 或浏览器模式复用登录态。

### `auto_download.fetch_order_page_with_retry(...)`

带重试获取 WMS 订单列表。遇到短时网络或 token 问题时配合 session 刷新逻辑提高稳定性。

### `auto_download.download_task(...)`

并发下载单个订单的核心任务函数。它会获取订单详情、下载面单文件并写 metadata。

### `batch_download_wms_pdfs.run_http_batch(...)`

轻量 HTTP 批量下载路径。按仓库和页码循环，逐订单处理详情、下载和日志。

### `batch_download_wms_pdfs.session_from_storage_state(...)`

从 Playwright storage state 构造可用 HTTP session，并补充 WMS API 必要 header。

### `batch_download_wms_pdfs.append_label_metadata(...)`

将订单来源字段、面单文件名、下载侧识别信息等写入 metadata JSONL。导出层依赖这些数据补充业务字段。

## 识别函数

### `pdf_verify.verify_pdfs(args, input_dir, download_log)`

批量识别入口。它加载 WMS tracking、metadata 索引和 zxing 引擎，然后逐行调用 `process_row()`。

### `barcode_verify_tracking.load_rows(...)`

根据 `--input-dir`、`--strict-json`、下载日志和 metadata 构造待识别行。

### `barcode_verify_tracking.process_row(row, args, zxingcpp, wms_records)`

单文件识别核心函数。主要步骤：

1. 解析文件路径、格式、来源追踪号和 metadata。
2. 对非 PDF 原始来源做异常处理。
3. 读取 PDF 文字层或识别图片输入。
4. 反读条码。
5. 提取并评分追踪号候选。
6. 判断来源、文字层、条码是否一致。
7. 识别承运商、模板和 OCR 需求。
8. 返回 `VerifyResult`。

### `barcode_verify_tracking.determine_status(...)`

根据下载来源、PDF 文字层、条码值和承运商判断最终状态。

主要强通过状态：

- `auto_pass_triple_verified`
- `auto_pass_barcode_verified`
- `auto_pass_text_verified`
- `auto_pass_verified_prefix`

主要复核状态：

- `source_only_low_confidence`
- `review_conflict`
- `review_unknown`
- `barcode_failed`
- `ocr_needed`
- `download_file_error`

### `barcode_verify_tracking.detect_template(...)`

综合 PDF 文字层、条码文本、图片角标、OCR 等信息识别模板类型。

### `barcode_verify_tracking.error_result(row, wms_records, exc)`

将文件打开失败、格式不支持、非 PDF 来源等问题统一转成 `VerifyResult`，状态为 `download_file_error`。

## 导出函数

### `exporter.business_rows(results)`

读取 metadata，构造 metadata 索引，并将每个 `VerifyResult` 转成业务行。

### `exporter.result_to_business_row(result, metadata)`

导出层的核心转换函数。它会：

- 确定承运商。
- 生成内容侧模板判断。
- 生成下载侧模板判断。
- 比较下载侧和内容侧是否一致。
- 补充订单号、追踪号、物流渠道名称。
- 判断是否需要人工复核。

### `exporter.logistics_channel_name(result, metadata)`

物流渠道名称优先级：

1. `metadata.source_fields.logisticsChannelName`
2. `metadata.channel_hint`
3. `result.meta_channel`

### `exporter.write_excel(results, path)`

写出 Excel。若目标文件已存在，会先读取旧数据再合并。

### `exporter.merge_business_rows(existing_rows, new_rows)`

按业务 key 合并旧行和新行。新行会覆盖相同 key 的旧行。

### `exporter.merge_brief_rows(existing_rows, new_rows)`

按追踪号合并 `简略版`，用于小时批次持续追加。

### `exporter.write_json(results, path)`

写出业务增强 JSON，包含原始识别字段、业务模板判断和比对结果。

## GUI 函数

### `WMS_TOOL.worker.run_pipeline(params, callback=None)`

GUI 模式运行入口。它会：

- 校验参数。
- 加载本地环境变量。
- 创建输出、下载、日志目录。
- 设置 `PDF_DDD_METADATA_JSONL` 和 `PDF_DDD_METADATA_SUMMARY` 环境变量。
- 可先查询 WMS 总数。
- 执行三阶段流水线并通过 callback 回传进度。

### `WMS_TOOL.task_runner.start_task(params)`

创建后台线程并返回 GUI 任务 ID。

### `WMS_TOOL.task_runner.get_task(task_id)`

返回任务快照，用于 GUI 轮询展示。

