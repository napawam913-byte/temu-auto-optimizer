from __future__ import annotations

import queue
import threading
import webbrowser
from pathlib import Path
from tkinter import BooleanVar, IntVar, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from urllib.parse import quote

from llm_optimizer import LLMConfig
from preview_server import PreviewServer, bulk_tune_excel_images
from processor import DEFAULT_IMAGE_TUNE_PROMPT, ProcessingConfig, TemuProcessor, load_config, save_config, split_excel


APP_NAME = "Temu Auto Optimizer"
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
OUTPUT_ROOT = BASE_DIR / "outputs"


class App(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Temu半托管自动上架助手")
        self.geometry("1100x740")
        self.minsize(1000, 660)

        self.event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.config_data = load_config(CONFIG_PATH)

        self.split_file = StringVar()
        self.split_output_directory = StringVar(value=self.config_data.split_output_directory or str(OUTPUT_ROOT))
        self.rows_per_file = IntVar(value=100)

        self.selected_files: list[Path] = []
        self.selected_folder = StringVar()
        self.last_output_file = StringVar()
        self.preview_file = StringVar()
        self.preview_server: PreviewServer | None = None

        self.api_key = StringVar(value=self.config_data.llm.api_key)
        self.provider = StringVar(value=self.config_data.llm.provider)
        self.base_url = StringVar(value=self.config_data.llm.base_url)
        self.model = StringVar(value=self.config_data.llm.model)
        self.image_api_key = StringVar(value=self.config_data.image_llm.api_key)
        self.image_provider = StringVar(value=self.config_data.image_llm.provider)
        self.image_base_url = StringVar(value=self.config_data.image_llm.base_url)
        self.image_model = StringVar(value=self.config_data.image_llm.model)
        self.image_tune_prompt = StringVar(value=self.config_data.image_tune_prompt or DEFAULT_IMAGE_TUNE_PROMPT)
        self.template_file = StringVar(value=self.config_data.template_file)
        self.output_directory = StringVar(value=self.config_data.output_directory or str(OUTPUT_ROOT))
        self.output_filename = StringVar(value=self.config_data.output_filename)
        self.title_max_length = IntVar(value=self.config_data.title_max_length)
        self.filter_clothing = BooleanVar(value=self.config_data.filter_clothing)
        self.deduplicate = BooleanVar(value=self.config_data.deduplicate)
        self.enable_scraper = BooleanVar(value=self.config_data.enable_scraper)
        self.download_images = BooleanVar(value=self.config_data.download_images)
        self.image_tune_count = IntVar(value=self.config_data.image_tune_count)
        self.dedupe_field = StringVar(value=self.config_data.dedupe_field)
        self.default_price = StringVar(value=str(self.config_data.default_price))
        self.default_stock = IntVar(value=self.config_data.default_stock)
        self.default_ship_days = StringVar(value=self.config_data.default_ship_days)
        self.enhance_white_bg = BooleanVar(value=self.config_data.enhance_white_bg)

        self._build_ui()
        self.after(120, self._drain_queue)

    def _build_ui(self) -> None:
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=12)

        split_tab = ttk.Frame(self.tabs, padding=12)
        process_tab = ttk.Frame(self.tabs, padding=12)
        settings_tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(split_tab, text="拆分工具")
        self.tabs.add(process_tab, text="自动上架")
        self.tabs.add(settings_tab, text="设置")

        self._build_split_tab(split_tab)
        self._build_process_tab(process_tab)
        self._build_settings_tab(settings_tab)

    def _build_split_tab(self, parent: ttk.Frame) -> None:
        file_row = ttk.Frame(parent)
        file_row.pack(fill="x", pady=(0, 10))
        ttk.Label(file_row, text="云启/极鲸云文件：").pack(side="left")
        ttk.Entry(file_row, textvariable=self.split_file).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(file_row, text="选择文件", command=self._choose_split_file).pack(side="left")

        split_output_row = ttk.Frame(parent)
        split_output_row.pack(fill="x", pady=(0, 10))
        ttk.Label(split_output_row, text="拆分输出目录：").pack(side="left")
        ttk.Entry(split_output_row, textvariable=self.split_output_directory).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(split_output_row, text="选择目录", command=self._choose_split_output_directory).pack(side="left")

        options = ttk.Frame(parent)
        options.pack(fill="x", pady=(0, 10))
        ttk.Label(options, text="每个文件行数：").pack(side="left")
        ttk.Spinbox(options, from_=1, to=5000, textvariable=self.rows_per_file, width=10).pack(side="left", padx=8)
        ttk.Button(options, text="开始拆分", command=self._start_split).pack(side="left", padx=8)

        self.split_progress = ttk.Progressbar(parent, mode="determinate")
        self.split_progress.pack(fill="x", pady=8)
        self.split_log = self._log_box(parent)

    def _build_process_tab(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent)
        top.pack(fill="x")
        ttk.Button(top, text="选择 Excel/CSV 文件", command=self._choose_process_files).pack(side="left")
        ttk.Button(top, text="选择文件夹批量处理", command=self._choose_process_folder).pack(side="left", padx=8)
        ttk.Button(top, text="开始处理", command=self._start_process).pack(side="left", padx=8)
        ttk.Button(top, text="查看结果 Excel", command=self._show_result).pack(side="left", padx=8)
        ttk.Button(top, text="一键改图", command=self._start_bulk_image_tune).pack(side="left", padx=8)
        ttk.Button(top, text="网页预览", command=self._open_web_preview_dialog).pack(side="left", padx=8)

        selected_frame = ttk.LabelFrame(parent, text="待处理文件", padding=8)
        selected_frame.pack(fill="x", pady=10)
        self.files_label = ttk.Label(selected_frame, text="尚未选择文件")
        self.files_label.pack(anchor="w")
        ttk.Label(selected_frame, textvariable=self.selected_folder).pack(anchor="w")

        template_row = ttk.Frame(parent)
        template_row.pack(fill="x", pady=(0, 10))
        ttk.Label(template_row, text="最终上架模板：").pack(side="left")
        ttk.Entry(template_row, textvariable=self.template_file).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(template_row, text="选择模板", command=self._choose_template_file).pack(side="left")

        output_row = ttk.Frame(parent)
        output_row.pack(fill="x", pady=(0, 10))
        ttk.Label(output_row, text="输出目录：").pack(side="left")
        ttk.Entry(output_row, textvariable=self.output_directory).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(output_row, text="选择目录", command=self._choose_output_directory).pack(side="left")
        ttk.Label(output_row, text="文件名：").pack(side="left", padx=(12, 0))
        ttk.Entry(output_row, textvariable=self.output_filename, width=28).pack(side="left", padx=8)

        config_frame = ttk.LabelFrame(parent, text="处理配置", padding=8)
        config_frame.pack(fill="x", pady=(0, 10))
        self._grid_entry(config_frame, "标题目标字符数", self.title_max_length, 0, width=10)
        self._grid_entry(config_frame, "默认价格", self.default_price, 1, width=10)
        self._grid_entry(config_frame, "默认库存", self.default_stock, 2, width=10)
        self._grid_entry(config_frame, "发货时效(天)", self.default_ship_days, 3, width=10)
        ttk.Checkbutton(config_frame, text="过滤服装鞋靴类目", variable=self.filter_clothing).grid(row=1, column=0, sticky="w", pady=6)
        ttk.Checkbutton(config_frame, text="开启去重", variable=self.deduplicate).grid(row=1, column=1, sticky="w", pady=6)
        ttk.Checkbutton(config_frame, text="下载并压缩图片", variable=self.download_images).grid(row=1, column=2, sticky="w", pady=6)
        ttk.Checkbutton(config_frame, text="白底增强", variable=self.enhance_white_bg).grid(row=1, column=3, sticky="w", pady=6)
        ttk.Label(config_frame, text="去重字段").grid(row=1, column=4, sticky="e")
        ttk.Combobox(config_frame, textvariable=self.dedupe_field, values=["title", "sku"], width=8, state="readonly").grid(row=1, column=5, sticky="w")
        ttk.Checkbutton(config_frame, text="启用商品详情爬虫补全（推荐）", variable=self.enable_scraper).grid(row=2, column=0, columnspan=3, sticky="w", pady=6)

        self.process_progress = ttk.Progressbar(parent, mode="determinate")
        self.process_progress.pack(fill="x", pady=8)
        self.status_label = ttk.Label(parent, text="就绪")
        self.status_label.pack(anchor="w")
        self.process_log = self._log_box(parent)

    def _build_settings_tab(self, parent: ttk.Frame) -> None:
        text_api = ttk.LabelFrame(parent, text="文字模型（标题优化 / 描述生成）", padding=10)
        text_api.pack(fill="x")
        self._settings_entry(text_api, "Provider", self.provider, row=0, column=0, width=22)
        self._settings_entry(text_api, "Model", self.model, row=0, column=2, width=28)
        self._settings_entry(text_api, "API Key", self.api_key, row=1, column=0, width=58, show="*")
        self._settings_entry(text_api, "Base URL", self.base_url, row=2, column=0, width=58)
        text_api.columnconfigure(1, weight=1)
        text_api.columnconfigure(3, weight=1)

        image_api = ttk.LabelFrame(parent, text="图片模型（图片质检 / 背景处理 / 修图，预留）", padding=10)
        image_api.pack(fill="x", pady=(12, 0))
        self._settings_entry(image_api, "Provider", self.image_provider, row=0, column=0, width=22)
        self._settings_entry(image_api, "Model", self.image_model, row=0, column=2, width=28)
        self._settings_entry(image_api, "API Key", self.image_api_key, row=1, column=0, width=58, show="*")
        self._settings_entry(image_api, "Base URL", self.image_base_url, row=2, column=0, width=58)
        ttk.Label(image_api, text="微调图片数量").grid(row=3, column=0, sticky="e", padx=(0, 6), pady=(10, 4))
        ttk.Spinbox(image_api, from_=0, to=10, textvariable=self.image_tune_count, width=10).grid(row=3, column=1, sticky="w", padx=(0, 12), pady=(10, 4))
        ttk.Label(image_api, text="图片微调提示词模板").grid(row=4, column=0, sticky="ne", padx=(0, 6), pady=(8, 4))
        self.image_prompt_box = ScrolledText(image_api, height=9, wrap="word")
        self.image_prompt_box.grid(row=4, column=1, columnspan=3, sticky="ew", pady=(8, 4))
        self.image_prompt_box.insert("1.0", self.image_tune_prompt.get())
        image_api.columnconfigure(1, weight=1)
        image_api.columnconfigure(3, weight=1)

        hints = ttk.LabelFrame(parent, text="常用 Base URL", padding=10)
        hints.pack(fill="x", pady=12)
        ttk.Label(hints, text="OpenAI: 留空或 https://api.openai.com/v1").pack(anchor="w")
        ttk.Label(hints, text="DeepSeek: https://api.deepseek.com").pack(anchor="w")
        ttk.Label(hints, text="通义千问 DashScope 兼容模式: https://dashscope.aliyuncs.com/compatible-mode/v1").pack(anchor="w")

        ttk.Button(parent, text="保存配置", command=self._save_settings).pack(anchor="w")

    def _grid_entry(self, parent: ttk.Frame, label: str, variable, column: int, width: int = 20, show: str = "") -> None:
        ttk.Label(parent, text=label).grid(row=0, column=column * 2, sticky="e", padx=(0, 6), pady=4)
        ttk.Entry(parent, textvariable=variable, width=width, show=show).grid(row=0, column=column * 2 + 1, sticky="w", padx=(0, 12), pady=4)

    def _settings_entry(
        self,
        parent: ttk.Frame,
        label: str,
        variable,
        row: int,
        column: int,
        width: int = 24,
        show: str = "",
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="e", padx=(0, 6), pady=4)
        ttk.Entry(parent, textvariable=variable, width=width, show=show).grid(
            row=row,
            column=column + 1,
            sticky="ew",
            padx=(0, 16),
            pady=4,
        )

    def _log_box(self, parent: ttk.Frame) -> ScrolledText:
        box = ScrolledText(parent, height=16, wrap="word")
        box.pack(fill="both", expand=True, pady=(8, 0))
        return box

    def _choose_split_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")])
        if path:
            self.split_file.set(path)

    def _choose_split_output_directory(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.split_output_directory.set(folder)

    def _choose_process_files(self) -> None:
        paths = filedialog.askopenfilenames(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")])
        self.selected_files = [Path(path) for path in paths]
        self.selected_folder.set("")
        self.files_label.config(text=f"已选择 {len(self.selected_files)} 个文件")

    def _choose_process_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            root = Path(folder)
            self.selected_files = sorted([*root.glob("*.xlsx"), *root.glob("*.xls"), *root.glob("*.csv")])
            self.selected_folder.set(f"文件夹：{folder}，发现 {len(self.selected_files)} 个文件")
            self.files_label.config(text="文件夹批量处理模式")

    def _choose_template_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.template_file.set(path)

    def _choose_output_directory(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.output_directory.set(folder)

    def _start_split(self) -> None:
        if not self.split_file.get():
            messagebox.showwarning(APP_NAME, "请先选择云启/极鲸云文件")
            return
        self._save_settings(silent=True)
        self._run_worker(self._split_worker)

    def _start_process(self) -> None:
        if not self.selected_files:
            messagebox.showwarning(APP_NAME, "请先选择一个或多个 Excel/CSV 文件")
            return
        self._save_settings(silent=True)
        self._run_worker(self._process_worker)

    def _run_worker(self, target) -> None:
        threading.Thread(target=target, daemon=True).start()

    def _split_worker(self) -> None:
        try:
            zip_path = split_excel(
                Path(self.split_file.get()),
                int(self.rows_per_file.get()),
                Path(self.split_output_directory.get().strip() or OUTPUT_ROOT),
                progress=lambda cur, total, msg: self.event_queue.put(("split_progress", (cur, total, msg))),
            )
            self.event_queue.put(("split_done", zip_path))
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def _process_worker(self) -> None:
        try:
            processor = TemuProcessor(
                self._current_config(),
                progress=lambda cur, total, msg: self.event_queue.put(("process_progress", (cur, total, msg))),
                log=lambda msg: self.event_queue.put(("process_log", msg)),
            )
            output_root = Path(self.output_directory.get().strip() or OUTPUT_ROOT)
            result = processor.process_files(self.selected_files, output_root)
            self.event_queue.put(("process_done", result))
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def _start_bulk_image_tune(self) -> None:
        path = self.last_output_file.get().strip()
        if not path:
            messagebox.showwarning(APP_NAME, "请先完成自动上架处理，生成结果 Excel")
            return
        file_path = Path(path)
        if not file_path.exists():
            messagebox.showerror(APP_NAME, f"结果文件不存在：\n{file_path}")
            return
        if not messagebox.askyesno(
            APP_NAME,
            f"将按设置里的微调图片数量，对结果 Excel 中所有未微调商品批量改图。\n\n文件：{file_path}\n\n是否继续？",
        ):
            return
        self._save_settings(silent=True)
        self._run_worker(lambda: self._bulk_image_tune_worker(file_path))

    def _bulk_image_tune_worker(self, file_path: Path) -> None:
        try:
            result = bulk_tune_excel_images(
                file_path,
                self._current_config(),
                progress=lambda cur, total, msg: self.event_queue.put(("bulk_tune_progress", (cur, total, msg))),
                log=lambda msg: self.event_queue.put(("process_log", msg)),
            )
            self.event_queue.put(("bulk_tune_done", result))
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def _current_config(self) -> ProcessingConfig:
        image_tune_prompt = self.image_tune_prompt.get()
        if hasattr(self, "image_prompt_box"):
            image_tune_prompt = self.image_prompt_box.get("1.0", "end").strip()
        return ProcessingConfig(
            llm=LLMConfig(
                provider=self.provider.get().strip() or "openai",
                api_key=self.api_key.get().strip(),
                base_url=self.base_url.get().strip(),
                model=self.model.get().strip() or "gpt-4o-mini",
            ),
            image_llm=LLMConfig(
                provider=self.image_provider.get().strip() or "openai",
                api_key=self.image_api_key.get().strip(),
                base_url=self.image_base_url.get().strip(),
                model=self.image_model.get().strip() or "gpt-image-1",
            ),
            template_file=self.template_file.get().strip(),
            split_output_directory=self.split_output_directory.get().strip(),
            output_directory=self.output_directory.get().strip(),
            output_filename=self.output_filename.get().strip(),
            title_max_length=int(self.title_max_length.get()),
            filter_clothing=bool(self.filter_clothing.get()),
            deduplicate=bool(self.deduplicate.get()),
            enable_scraper=bool(self.enable_scraper.get()),
            download_images=bool(self.download_images.get()),
            image_tune_count=int(self.image_tune_count.get()),
            image_tune_prompt=image_tune_prompt or DEFAULT_IMAGE_TUNE_PROMPT,
            dedupe_field=self.dedupe_field.get(),
            default_price=float(self.default_price.get() or 0),
            default_stock=int(self.default_stock.get()),
            default_ship_days=self.default_ship_days.get().strip() or "2",
            enhance_white_bg=bool(self.enhance_white_bg.get()),
        )

    def _save_settings(self, silent: bool = False) -> None:
        save_config(self._current_config(), CONFIG_PATH)
        if not silent:
            messagebox.showinfo(APP_NAME, f"配置已保存：{CONFIG_PATH}")

    def _show_result(self) -> None:
        path = self.last_output_file.get()
        if path:
            messagebox.showinfo(APP_NAME, f"结果文件：\n{path}")
        else:
            messagebox.showinfo(APP_NAME, "还没有生成结果文件")

    def _open_web_preview_dialog(self) -> None:
        selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if selected:
            self.preview_file.set(selected)
        elif self.last_output_file.get():
            self.preview_file.set(self.last_output_file.get())
        else:
            return
        self._open_web_preview_from_field()

    def _open_web_preview_from_field(self) -> None:
        path = self.preview_file.get().strip()
        if not path:
            messagebox.showwarning(APP_NAME, "???????? Excel ??")
            return
        file_path = Path(path)
        if not file_path.exists():
            messagebox.showerror(APP_NAME, f"????????\n{file_path}")
            return
        try:
            if not self.preview_server:
                self.preview_server = PreviewServer(CONFIG_PATH)
            base_url = self.preview_server.start()
            preview_url = f"{base_url}/preview?file={quote(str(file_path.resolve()))}"
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"?????????\n{exc}")
            return
        webbrowser.open(preview_url)

    def _drain_queue(self) -> None:
        while True:
            try:
                event, payload = self.event_queue.get_nowait()
            except queue.Empty:
                break
            self._handle_event(event, payload)
        self.after(120, self._drain_queue)

    def _handle_event(self, event: str, payload: object) -> None:
        if event == "split_progress":
            cur, total, msg = payload
            self.split_progress["maximum"] = total
            self.split_progress["value"] = cur
            self._append(self.split_log, msg)
        elif event == "split_done":
            self._append(self.split_log, f"完成：{payload}")
            messagebox.showinfo(APP_NAME, f"拆分完成：\n{payload}")
        elif event == "process_progress":
            cur, total, msg = payload
            self.process_progress["maximum"] = total
            self.process_progress["value"] = cur
            self.status_label.config(text=msg)
        elif event == "process_log":
            self._append(self.process_log, str(payload))
        elif event == "bulk_tune_progress":
            cur, total, msg = payload
            self.process_progress["maximum"] = total
            self.process_progress["value"] = cur
            self.status_label.config(text=msg)
        elif event == "process_done":
            result = payload
            self.last_output_file.set(str(result.output_file))
            self.preview_file.set(str(result.output_file))
            self.status_label.config(
                text=f"完成：成功 {result.success_count} 条，失败 {result.failure_count} 条，跳过 {result.skipped_count} 条"
            )
            self._append(self.process_log, f"输出文件：{result.output_file}")
            if result.failures:
                self._append(self.process_log, "\n".join(result.failures[:30]))
            messagebox.showinfo(APP_NAME, self.status_label.cget("text"))
        elif event == "bulk_tune_done":
            result = payload
            failures = result.get("failures", [])
            text = (
                f"一键改图完成：商品 {result.get('products', 0)} 个，图片 {result.get('images', 0)} 张，"
                f"跳过 {result.get('skipped', 0)} 个，失败 {len(failures)} 个"
            )
            self.status_label.config(text=text)
            self._append(self.process_log, text)
            self._append(self.process_log, f"备份文件：{result.get('backup', '')}")
            if failures:
                self._append(self.process_log, "\n".join(failures[:30]))
            messagebox.showinfo(APP_NAME, text)
        elif event == "error":
            messagebox.showerror(APP_NAME, str(payload))

    @staticmethod
    def _append(box: ScrolledText, message: str) -> None:
        box.insert("end", f"{message}\n")
        box.see("end")


if __name__ == "__main__":
    App().mainloop()
