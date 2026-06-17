from __future__ import annotations

import sys
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk


CURRENT_DIR = Path(__file__).resolve().parent
PARENT_DIR = CURRENT_DIR.parent
if str(PARENT_DIR) not in sys.path:
    sys.path.insert(0, str(PARENT_DIR))

from WMS_TOOL.task_runner import get_task, start_task


class WmsToolApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("WMS_TOOL")
        self.geometry("860x620")
        self.minsize(760, 520)
        self.task_id: str | None = None
        self.rendered_log_count = 0

        self.start_time_var = tk.StringVar()
        self.end_time_var = tk.StringVar()
        self.warehouse_var = tk.StringVar(value="US02")
        self.workers_var = tk.StringVar(value="8")
        self.status_var = tk.StringVar(value="idle")
        self.progress_var = tk.IntVar(value=0)

        self._set_default_time_range()
        self._build_ui()
        self.after(1000, self.poll_task)

    def _set_default_time_range(self) -> None:
        now = datetime.now()
        start = (now - timedelta(days=1)).replace(hour=22, minute=30, second=0, microsecond=0)
        end = now.replace(second=0, microsecond=0)
        self.start_time_var.set(start.strftime("%Y-%m-%d %H:%M:%S"))
        self.end_time_var.set(end.strftime("%Y-%m-%d %H:%M:%S"))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(outer, text="参数")
        form.pack(fill=tk.X)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        ttk.Label(form, text="start_time").grid(row=0, column=0, sticky=tk.W, padx=8, pady=8)
        ttk.Entry(form, textvariable=self.start_time_var).grid(row=0, column=1, sticky=tk.EW, padx=8, pady=8)
        ttk.Label(form, text="end_time").grid(row=0, column=2, sticky=tk.W, padx=8, pady=8)
        ttk.Entry(form, textvariable=self.end_time_var).grid(row=0, column=3, sticky=tk.EW, padx=8, pady=8)

        ttk.Label(form, text="warehouse").grid(row=1, column=0, sticky=tk.W, padx=8, pady=8)
        ttk.Entry(form, textvariable=self.warehouse_var).grid(row=1, column=1, sticky=tk.EW, padx=8, pady=8)
        ttk.Label(form, text="workers").grid(row=1, column=2, sticky=tk.W, padx=8, pady=8)
        ttk.Spinbox(form, from_=1, to=32, textvariable=self.workers_var, width=8).grid(
            row=1, column=3, sticky=tk.W, padx=8, pady=8
        )

        controls = ttk.Frame(outer)
        controls.pack(fill=tk.X, pady=(14, 8))
        self.start_button = ttk.Button(controls, text="Start", command=self.on_start)
        self.start_button.pack(side=tk.LEFT)
        ttk.Label(controls, textvariable=self.status_var).pack(side=tk.LEFT, padx=12)

        self.progress = ttk.Progressbar(outer, variable=self.progress_var, maximum=100)
        self.progress.pack(fill=tk.X, pady=(0, 12))

        log_frame = ttk.LabelFrame(outer, text="日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=18)
        self.log_text.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def on_start(self) -> None:
        try:
            workers = int(self.workers_var.get().strip())
            if workers < 1:
                raise ValueError("workers 必须大于 0")
            params = {
                "start_time": self.start_time_var.get().strip(),
                "end_time": self.end_time_var.get().strip(),
                "warehouse": self.warehouse_var.get().strip() or "US02",
                "workers": workers,
            }
            datetime.strptime(params["start_time"], "%Y-%m-%d %H:%M:%S")
            datetime.strptime(params["end_time"], "%Y-%m-%d %H:%M:%S")
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.log_text.delete("1.0", tk.END)
        self.rendered_log_count = 0
        self.progress_var.set(0)
        self.status_var.set("running")
        self.start_button.configure(state=tk.DISABLED)
        self.task_id = start_task(params)
        self.append_log(f"task_id={self.task_id}")

    def append_log(self, message: str) -> None:
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def poll_task(self) -> None:
        if self.task_id:
            task = get_task(self.task_id)
            if task:
                self.progress_var.set(int(task.get("progress") or 0))
                self.status_var.set(str(task.get("status") or "running"))
                logs = task.get("log") or []
                for line in logs[self.rendered_log_count :]:
                    self.append_log(str(line))
                self.rendered_log_count = len(logs)

                status = str(task.get("status") or "")
                if status in {"done", "failed"}:
                    self.start_button.configure(state=tk.NORMAL)
                    if status == "done":
                        result = task.get("result") or {}
                        excel_path = result.get("excel_path") or ""
                        if excel_path:
                            self.append_log(f"完成: {excel_path}")
                    else:
                        self.append_log(f"失败: {task.get('error') or ''}")
                    self.task_id = None
        self.after(1000, self.poll_task)


def main() -> None:
    app = WmsToolApp()
    app.mainloop()


if __name__ == "__main__":
    main()
