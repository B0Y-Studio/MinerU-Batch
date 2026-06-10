#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MinerU 文档批量处理工具 - 图形化界面
支持批量上传文档并转换为 Markdown / DOCX / HTML / LaTeX 等格式
"""

import os
import sys
import threading
import queue
import time
import logging
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 确保 mineru 可导入
try:
    import mineru
except ImportError:
    import subprocess
    result = subprocess.run(
        [sys.executable, "-m", "pip", "show", "mineru-open-sdk"],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        messagebox.showerror("缺少依赖", "请先安装 MinerU SDK:\npip install mineru-open-sdk")
        sys.exit(1)

# ── SDK 兼容补丁 ──────────────────────────────────────────
# 问题: 新版 MinerU API 移除了 model_version 字段支持，SDK v0.2.5 仍发送该字段
# 导致服务端返回 [-10002] field "version" is invalid
# 解决方案: 创建自定义提取函数绕过有问题的字段

import httpx
from typing import Iterator

# 保存原 SDK 引用
_mineru_extract = mineru.MinerU.extract
_mineru_extract_batch = mineru.MinerU.extract_batch

_SENTINEL = object()

def _patched_build_options(
    model_version: str,
    formula: object,
    table: object,
    language: object,
    extra_formats: list[str] | None,
) -> dict:
    """改良版 _build_options: 仅在用户明确选择时发送 model_version"""
    opts: dict = {}
    # 仅在用户明确选择了非默认模型时才发送
    if model_version and model_version != "pipeline":
        opts["model_version"] = model_version
    if formula is not _SENTINEL:
        opts["enable_formula"] = formula
    if table is not _SENTINEL:
        opts["enable_table"] = table
    if language is not _SENTINEL:
        opts["language"] = language
    if extra_formats:
        opts["extra_formats"] = extra_formats
    return opts

def _patched_extract(
    self,
    source: str,
    *,
    model: str | None = None,
    ocr: bool | None = None,
    formula: object = _SENTINEL,
    table: object = _SENTINEL,
    language: object = _SENTINEL,
    pages: str | None = None,
    extra_formats: list[str] | None = None,
    file_params: dict[str, "mineru.FileParam"] | None = None,
    timeout: int = 300,
) -> "mineru.ExtractResult":
    """改良版 extract: 不发送 model_version 字段"""
    from mineru.client import _resolve_model

    model_version = _resolve_model(model, source) if model else "pipeline"
    self._require_auth()
    opts = _patched_build_options(model_version, formula, table, language, extra_formats)

    if source.startswith("http://") or source.startswith("https://"):
        batch_id = self._submit_urls_batch([source], opts, ocr, pages, file_params)
    else:
        batch_id = self._upload_and_submit([source], opts, ocr, pages, file_params)
    results = self._wait_batch(batch_id, timeout)
    return results[0]

def _patched_extract_batch(
    self,
    sources: list[str],
    *,
    model: str | None = None,
    ocr: bool | None = None,
    formula: object = _SENTINEL,
    table: object = _SENTINEL,
    language: object = _SENTINEL,
    extra_formats: list[str] | None = None,
    file_params: dict[str, "mineru.FileParam"] | None = None,
    timeout: int = 1800,
) -> Iterator["mineru.ExtractResult"]:
    """改良版 extract_batch: 不发送 model_version 字段"""
    from mineru.client import _resolve_model

    first_source = sources[0] if sources else ""
    model_version = _resolve_model(model, first_source) if model else "pipeline"
    self._require_auth()
    opts = _patched_build_options(model_version, formula, table, language, extra_formats)

    urls = [s for s in sources if s.startswith(("http://", "https://"))]
    files = [s for s in sources if not s.startswith(("http://", "https://"))]

    batch_ids: list[str] = []
    if urls:
        batch_ids.append(self._submit_urls_batch(urls, opts, ocr, None, file_params))
    if files:
        batch_ids.append(self._upload_and_submit(files, opts, ocr, None, file_params))

    yield from self._yield_batch(batch_ids, len(sources), timeout)

# 应用补丁
mineru.MinerU.extract = _patched_extract
mineru.MinerU.extract_batch = _patched_extract_batch

# ========== 常量 ==========
DEFAULT_BASE_URL = "https://mineru.net/api/v4"
SUPPORTED_EXTENSIONS = {
    '.pdf', '.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx',
    '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp',
    '.txt', '.md', '.html', '.htm', '.epub', '.mobi',
    '.csv', '.json', '.xml', '.zip'
}
LANGUAGES = [
    ("中文", "ch"), ("英文", "en"), ("中英混合", "ch-en"),
    ("日文", "ja"), ("韩文", "ko"), ("法文", "fr"),
    ("德文", "de"), ("西班牙文", "es"), ("葡萄牙文", "pt"),
    ("阿拉伯文", "ar"), ("俄文", "ru"), ("自动检测", "auto"),
]
MODELS = [
    ("pipeline（通用解析）", "pipeline"),
    ("vlm（高精度，推荐）", "vlm"),
    ("MinerU-HTML（HTML 专用）", "MinerU-HTML"),
]

# API 限制
MAX_FILE_SIZE = 200 * 1024 * 1024        # 200 MB
MAX_BATCH_COUNT = 200                    # 批量上限 200 个

# ========== 日志 ==========
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("MinerU_GUI")


class LogHandler(logging.Handler):
    """将日志重定向到 tkinter 控件"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                            datefmt="%H:%M:%S"))

    def emit(self, record):
        msg = self.format(record) + "\n"
        self.text_widget.after(0, self._append, msg)

    def _append(self, msg):
        self.text_widget.configure(state=NORMAL)
        self.text_widget.insert(END, msg)
        self.text_widget.see(END)
        self.text_widget.configure(state=DISABLED)


class MinerUApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MinerU 文档批量处理工具 v1.0")
        self.root.geometry("880x780")
        self.root.minsize(800, 700)

        # 设置样式
        style = ttk.Style()
        style.theme_use("vista" if "vista" in style.theme_names() else "clam")
        style.configure("Success.TLabel", foreground="green", font=("Segoe UI", 9, "bold"))
        style.configure("Error.TLabel", foreground="red", font=("Segoe UI", 9, "bold"))
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Section.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Counter.TLabel", font=("Segoe UI", 9))

        # 状态变量
        self.file_list = []           # [(path, status), ...]
        self.processing = False
        self.cancel_flag = False
        self.task_queue = queue.Queue()
        self.progress_queue = queue.Queue()

        # ---------- 变量 ----------
        self.token_var = StringVar(value=os.environ.get("MINERU_TOKEN", ""))
        self.base_url_var = StringVar(value=DEFAULT_BASE_URL)
        self.output_dir_var = StringVar(value=str(Path.home() / "Desktop" / "MinerU_Output"))
        self.save_md_var = BooleanVar(value=True)
        self.save_docx_var = BooleanVar(value=False)
        self.save_html_var = BooleanVar(value=False)
        self.save_latex_var = BooleanVar(value=False)
        self.save_all_var = BooleanVar(value=False)
        self.ocr_var = BooleanVar(value=False)   # 官网默认 false
        self.formula_var = BooleanVar(value=True)
        self.table_var = BooleanVar(value=True)
        self.language_var = StringVar(value="ch")
        self.model_var = StringVar(value="")
        self.select_mode = StringVar(value="file")  # "file" or "folder"

        # ---------- 构建界面 ----------
        self._build_ui()

        # ---------- 日志 ----------
        log_handler = LogHandler(self.log_text)
        logger.addHandler(log_handler)

        # ---------- 定时处理队列 ----------
        self._poll_progress()

        # ---------- 窗口关闭事件 ----------
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ==================================================================
    #  构建界面
    # ==================================================================
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=BOTH, expand=True)

        # ---------- 标题 ----------
        title_frame = ttk.Frame(main)
        title_frame.pack(fill=X, pady=(0, 10))
        ttk.Label(title_frame, text="📄 MinerU 文档批量处理工具",
                  style="Title.TLabel").pack(side=LEFT)
        ttk.Label(title_frame, text="   powered by MinerU API",
                  foreground="gray", font=("Segoe UI", 9)).pack(side=LEFT, padx=6)

        # ---------- Notebook ----------
        notebook = ttk.Notebook(main)
        notebook.pack(fill=BOTH, expand=True)

        # ========== Tab 1: 主操作 ==========
        tab_main = ttk.Frame(notebook, padding=10)
        notebook.add(tab_main, text="⚙ 批量处理")

        # ---- API 配置 ----
        api_frame = ttk.LabelFrame(tab_main, text="API 配置", padding=10)
        api_frame.pack(fill=X, pady=(0, 10))

        row1 = ttk.Frame(api_frame)
        row1.pack(fill=X, pady=2)
        ttk.Label(row1, text="Token / API Key:", width=14).pack(side=LEFT)
        self.token_entry = ttk.Entry(row1, textvariable=self.token_var, width=60, show="*")
        self.token_entry.pack(side=LEFT, fill=X, expand=True, padx=4)
        self.token_btn = ttk.Button(row1, text="👁", width=3, command=self._toggle_token_visible)
        self.token_btn.pack(side=LEFT)

        row2 = ttk.Frame(api_frame)
        row2.pack(fill=X, pady=2)
        ttk.Label(row2, text="Base URL:", width=14).pack(side=LEFT)
        ttk.Entry(row2, textvariable=self.base_url_var, width=60).pack(side=LEFT, fill=X, expand=True, padx=4)
        ttk.Button(row2, text="恢复默认", command=lambda: self.base_url_var.set(DEFAULT_BASE_URL)).pack(side=LEFT)

        # ---- 文件选择 ----
        file_frame = ttk.LabelFrame(tab_main, text="文件选择", padding=10)
        file_frame.pack(fill=BOTH, expand=True, pady=(0, 10))

        # 选择模式
        mode_row = ttk.Frame(file_frame)
        mode_row.pack(fill=X, pady=(0, 6))
        ttk.Radiobutton(mode_row, text="选择单个/多个文件", variable=self.select_mode,
                        value="file", command=self._on_mode_change).pack(side=LEFT, padx=(0, 20))
        ttk.Radiobutton(mode_row, text="选择整个文件夹", variable=self.select_mode,
                        value="folder", command=self._on_mode_change).pack(side=LEFT)
        ttk.Label(mode_row, text="", style="Counter.TLabel").pack(side=LEFT, padx=10)
        self.file_count_label = ttk.Label(mode_row, text="已选择 0 个文件", style="Counter.TLabel")
        self.file_count_label.pack(side=RIGHT)

        # 按钮行
        btn_row = ttk.Frame(file_frame)
        btn_row.pack(fill=X, pady=(0, 6))
        self.add_btn = ttk.Button(btn_row, text="📂 添加文件...", command=self._add_files)
        self.add_btn.pack(side=LEFT, padx=(0, 6))
        self.add_folder_btn = ttk.Button(btn_row, text="📁 添加文件夹...", command=self._add_folder)
        self.add_folder_btn.pack(side=LEFT, padx=(0, 6))
        self.clear_btn = ttk.Button(btn_row, text="🗑 清空列表", command=self._clear_files)
        self.clear_btn.pack(side=LEFT, padx=(0, 6))
        self.remove_sel_btn = ttk.Button(btn_row, text="✂ 移除选中", command=self._remove_selected)
        self.remove_sel_btn.pack(side=LEFT)

        # 文件列表
        list_container = ttk.Frame(file_frame)
        list_container.pack(fill=BOTH, expand=True)
        columns = ("序号", "文件名", "大小", "状态")
        self.file_tree = ttk.Treeview(list_container, columns=columns,
                                       show="headings", height=8, selectmode="extended")
        self.file_tree.heading("序号", text="#", anchor=CENTER)
        self.file_tree.heading("文件名", text="文件名")
        self.file_tree.heading("大小", text="大小", anchor=E)
        self.file_tree.heading("状态", text="状态", anchor=CENTER)
        self.file_tree.column("序号", width=45, anchor=CENTER)
        self.file_tree.column("文件名", width=300)
        self.file_tree.column("大小", width=100, anchor=E)
        self.file_tree.column("状态", width=100, anchor=CENTER)
        self.file_tree["displaycolumns"] = ("序号", "文件名", "大小", "状态")

        vbar = ttk.Scrollbar(list_container, orient=VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=vbar.set)
        self.file_tree.pack(side=LEFT, fill=BOTH, expand=True)
        vbar.pack(side=RIGHT, fill=Y)

        # ---- 输出设置 + 处理选项 (两列) ----
        settings_frame = ttk.Frame(tab_main)
        settings_frame.pack(fill=X, pady=(0, 10))

        # 左列 - 输出格式
        fmt_frame = ttk.LabelFrame(settings_frame, text="输出格式", padding=10)
        fmt_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))

        fmt_grid = ttk.Frame(fmt_frame)
        fmt_grid.pack(fill=X)
        # 提示：Markdown + JSON 默认包含在 Zip 中，无需勾选
        ttk.Label(fmt_grid, text="✅ Markdown + JSON 默认包含在 Zip 中",
                  foreground="#666", font=("Segoe UI", 8)).grid(row=0, column=0, columnspan=2, sticky=W, padx=4, pady=(0, 4))
        ttk.Checkbutton(fmt_grid, text="Markdown (.md) 提取到目录", variable=self.save_md_var).grid(row=1, column=0, columnspan=2, sticky=W, padx=4, pady=1)
        ttk.Checkbutton(fmt_grid, text="额外导出 DOCX (.docx)", variable=self.save_docx_var).grid(row=2, column=0, sticky=W, padx=4, pady=1)
        ttk.Checkbutton(fmt_grid, text="额外导出 HTML (.html)", variable=self.save_html_var).grid(row=2, column=1, sticky=W, padx=4, pady=1)
        ttk.Checkbutton(fmt_grid, text="额外导出 LaTeX (.tex)", variable=self.save_latex_var).grid(row=3, column=0, sticky=W, padx=4, pady=1)
        ttk.Checkbutton(fmt_grid, text="保存全部（含图片到目录）", variable=self.save_all_var).grid(row=3, column=1, sticky=W, padx=4, pady=1)

        # 输出目录
        dir_frame = ttk.Frame(fmt_frame)
        dir_frame.pack(fill=X, pady=(6, 0))
        ttk.Label(dir_frame, text="输出目录:").pack(side=LEFT)
        ttk.Entry(dir_frame, textvariable=self.output_dir_var).pack(side=LEFT, fill=X, expand=True, padx=4)
        ttk.Button(dir_frame, text="浏览...", command=self._browse_output_dir).pack(side=LEFT)

        # 右列 - 处理选项
        opt_frame = ttk.LabelFrame(settings_frame, text="处理选项", padding=10)
        opt_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(5, 0))

        ttk.Checkbutton(opt_frame, text="启用 OCR（文字识别）", variable=self.ocr_var).pack(anchor=W, pady=1)
        ttk.Label(opt_frame, text="  ⚠ 仅对 pipeline / vlm 模型有效",
                  foreground="#888", font=("Segoe UI", 8)).pack(anchor=W, padx=20, pady=(0, 2))
        ttk.Checkbutton(opt_frame, text="启用公式识别", variable=self.formula_var).pack(anchor=W, pady=1)
        ttk.Checkbutton(opt_frame, text="启用表格识别", variable=self.table_var).pack(anchor=W, pady=1)
        # 限制信息
        ttk.Separator(opt_frame, orient=HORIZONTAL).pack(fill=X, pady=(8, 4))
        ttk.Label(opt_frame, text="📋 限制: 单文件 ≤200MB / ≤200页 / 批量 ≤200个",
                  foreground="#888", font=("Segoe UI", 8)).pack(anchor=W)

        # 语言选择
        lang_row = ttk.Frame(opt_frame)
        lang_row.pack(fill=X, pady=(6, 2))
        ttk.Label(lang_row, text="文档语言:").pack(side=LEFT)
        lang_menu = ttk.Combobox(lang_row, textvariable=self.language_var, state="readonly", width=14)
        lang_menu["values"] = [f"{label}" for label, _ in LANGUAGES]
        lang_menu.set("中文")
        lang_menu.pack(side=LEFT, padx=4)
        # 映射显示名 -> 值
        self.lang_map = {label: val for label, val in LANGUAGES}
        self.lang_rev_map = {val: label for label, val in LANGUAGES}
        lang_menu.bind("<<ComboboxSelected>>", lambda e: None)

        # 模型选择
        model_row = ttk.Frame(opt_frame)
        model_row.pack(fill=X, pady=2)
        ttk.Label(model_row, text="处理模型:").pack(side=LEFT)
        model_menu = ttk.Combobox(model_row, textvariable=self.model_var, state="readonly", width=14)
        model_menu["values"] = [f"{label}" for label, _ in MODELS]
        model_menu.set("pipeline（通用解析）")
        model_menu.pack(side=LEFT, padx=4)
        self.model_map = {label: val for label, val in MODELS}
        model_menu.bind("<<ComboboxSelected>>", lambda e: None)

        # ---- 操作按钮 ----
        action_frame = ttk.Frame(tab_main)
        action_frame.pack(fill=X, pady=(0, 6))
        self.start_btn = ttk.Button(action_frame, text="▶ 开始处理", command=self._start_processing,
                                     width=18)
        self.start_btn.pack(side=LEFT, padx=(0, 8))
        self.cancel_btn = ttk.Button(action_frame, text="■ 取消", command=self._cancel_processing,
                                      width=10, state=DISABLED)
        self.cancel_btn.pack(side=LEFT)
        # 进度条
        self.progress_bar = ttk.Progressbar(action_frame, mode="determinate", length=200)
        self.progress_bar.pack(side=LEFT, fill=X, expand=True, padx=(10, 0))
        self.progress_label = ttk.Label(action_frame, text="就绪", width=20, anchor=E)
        self.progress_label.pack(side=LEFT, padx=4)

        # ========== Tab 2: 日志 ==========
        tab_log = ttk.Frame(notebook, padding=10)
        notebook.add(tab_log, text="📋 运行日志")
        self.log_text = scrolledtext.ScrolledText(
            tab_log, state=DISABLED, wrap=WORD,
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white"
        )
        self.log_text.pack(fill=BOTH, expand=True)
        self.log_text.tag_config("ok", foreground="#4ec9b0")
        self.log_text.tag_config("err", foreground="#f44747")
        self.log_text.tag_config("info", foreground="#569cd6")
        self.log_text.tag_config("warn", foreground="#ce9178")

        # ========== Tab 3: 帮助 ==========
        tab_help = ttk.Frame(notebook, padding=15)
        notebook.add(tab_help, text="❓ 帮助")
        help_text = """📄 MinerU 文档批量处理工具 — 官网 API v4 版

使用说明：
1. 在「API 配置」中填入你的 MinerU Token（API Key）
2. 选择要处理的文件或整个文件夹（支持 PDF / Office / 图片等）
3. 选择输出格式和解析模型
4. 点击「开始处理」→ 异步提交 → 自动轮询 → 保存结果

━━━━ 鉴权方式 ━━━━━━━━━━━━━━━━━━━━━━
调用需通过 Token 进行身份验证，填写于上方 API 配置区。
也可设置环境变量 MINERU_TOKEN，启动时自动读取。

━━━━ 接口地址 ━━━━━━━━━━━━━━━━━━━━━━
- 任务提交：  POST /api/v4/extract/task          （URL 直接提交）
- 批量文件：  POST /api/v4/file-urls/batch       （本地上传）
- 结果查询：  GET  /api/v4/extract-results/batch/{batch_id}

━━━━ 模型说明 ━━━━━━━━━━━━━━━━━━━━━━
- pipeline（默认）   通用解析模型，适合大多数文档
- vlm（推荐）        基于视觉语言模型的高精度解析（效果更好）
- MinerU-HTML       HTML 格式输出的专用模型

━━━━ 处理限制 ━━━━━━━━━━━━━━━━━━━━━━
- 单文件大小： ≤ 200 MB
- 单文件页数： ≤ 200 页
- 批量上限：   ≤ 200 个文件

━━━━ 输出格式 ━━━━━━━━━━━━━━━━━━━━━━
- 默认返回： Zip 压缩包（内含 Markdown + JSON）
- 额外导出： DOCX / HTML / LaTeX（需在界面勾选）
- 保存全部： 解压 Zip + 图片到独立子目录

━━━━ 调用方式 ━━━━━━━━━━━━━━━━━━━━━━
异步处理模式：
  提交任务 → 轮询获取结果 → 自动下载保存
支持中途取消当前批次处理。

━━━━ 注意事项 ━━━━━━━━━━━━━━━━━━━━━━
- OCR 功能 默认关闭，仅对 pipeline / vlm 模型有效
- 每次处理会消耗账户配额（每日 1000 页优先额度）
- 上传链接 24 小时有效，超时需重新提交

━━━━ 获取 Token ━━━━━━━━━━━━━━━━━━━━
1. 访问 https://mineru.net 注册账号
2. 进入个人设置 → API Token
3. 复制 Token 粘贴到本工具输入框"""
        help_widget = Text(tab_help, wrap=WORD, font=("Segoe UI", 10),
                           bg=self.root.cget("bg"), relief=FLAT, padx=10, pady=10)
        help_widget.insert("1.0", help_text)
        help_widget.configure(state=DISABLED)
        help_widget.pack(fill=BOTH, expand=True)

        # ---------- 底部状态栏 ----------
        status_frame = ttk.Frame(main)
        status_frame.pack(fill=X, pady=(4, 0))
        self.status_label = ttk.Label(status_frame, text="就绪", relief=SUNKEN, anchor=W,
                                       padding=(6, 2))
        self.status_label.pack(fill=X)

        # 初始化文件列表
        self._refresh_file_list()

    # ==================================================================
    #  事件处理
    # ==================================================================
    def _toggle_token_visible(self):
        if self.token_entry.cget("show") == "*":
            self.token_entry.configure(show="")
            self.token_btn.configure(text="🙈")
        else:
            self.token_entry.configure(show="*")
            self.token_btn.configure(text="👁")

    def _on_mode_change(self):
        pass  # 模式只是改变按钮行为

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="选择要处理的文档",
            filetypes=[
                ("所有支持的格式", "*.pdf *.doc *.docx *.ppt *.pptx *.xls *.xlsx "
                                  "*.jpg *.jpeg *.png *.bmp *.tiff *.tif "
                                  "*.txt *.md *.html *.htm *.epub *.mobi "
                                  "*.csv *.json *.xml *.zip"),
                ("PDF 文档", "*.pdf"),
                ("Office 文档", "*.doc *.docx *.ppt *.pptx *.xls *.xlsx"),
                ("图片", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif"),
                ("文本/标记", "*.txt *.md *.html *.htm"),
                ("电子书", "*.epub *.mobi"),
                ("压缩包", "*.zip"),
                ("所有文件", "*.*"),
            ]
        )
        if not files:
            return
        added = 0
        for f in files:
            if f not in [item[0] for item in self.file_list]:
                self.file_list.append((f, "待处理"))
                added += 1
        if added > 0:
            self._refresh_file_list()
            self.status_label.configure(text=f"已添加 {added} 个文件")
            logger.info(f"添加了 {added} 个文件")

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含文档的文件夹")
        if not folder:
            return
        added = 0
        for root_dir, _, files in os.walk(folder):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in SUPPORTED_EXTENSIONS:
                    path = os.path.join(root_dir, file)
                    if path not in [item[0] for item in self.file_list]:
                        self.file_list.append((path, "待处理"))
                        added += 1
        if added > 0:
            self._refresh_file_list()
            self.status_label.configure(text=f"从文件夹添加了 {added} 个文件")
            logger.info(f"从文件夹 {folder} 添加了 {added} 个文件")

    def _clear_files(self):
        if self.processing:
            messagebox.showwarning("提示", "处理中无法清空列表")
            return
        if not self.file_list:
            return
        if messagebox.askyesno("确认", "确定要清空文件列表吗？"):
            self.file_list.clear()
            self._refresh_file_list()
            self.status_label.configure(text="文件列表已清空")
            logger.info("文件列表已清空")

    def _remove_selected(self):
        if self.processing:
            messagebox.showwarning("提示", "处理中无法移除文件")
            return
        selected = self.file_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选中要移除的文件")
            return
        indices = [int(self.file_tree.item(i, "values")[0]) - 1 for i in selected]
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self.file_list):
                self.file_list.pop(idx)
        self._refresh_file_list()
        logger.info(f"移除了 {len(selected)} 个文件")

    def _browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir_var.set(dir_path)

    def _refresh_file_list(self):
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        for i, (path, status) in enumerate(self.file_list, 1):
            fname = os.path.basename(path)
            size = self._format_size(os.path.getsize(path))
            self.file_tree.insert("", END, values=(i, fname, size, status))
        self.file_count_label.configure(text=f"已选择 {len(self.file_list)} 个文件")

    def _update_file_status(self, index, status):
        """更新列表中某个文件的状态"""
        if 0 <= index < len(self.file_list):
            self.file_list[index] = (self.file_list[index][0], status)
        self.root.after(0, self._refresh_file_list)

    @staticmethod
    def _format_size(size_bytes):
        for unit in ("B", "KB", "MB", "GB"):
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} TB"

    # ==================================================================
    #  处理逻辑
    # ==================================================================
    def _start_processing(self):
        """验证输入并启动后台线程"""
        # 验证 Token
        token = self.token_var.get().strip()
        if not token:
            messagebox.showerror("错误", "请输入 Token / API Key")
            self.token_entry.focus()
            return

        # 验证文件列表
        if not self.file_list:
            messagebox.showerror("错误", "请先添加要处理的文件")
            return

        # 验证批量限制 (≤200)
        if len(self.file_list) > MAX_BATCH_COUNT:
            messagebox.showerror("错误",
                f"批量处理上限为 {MAX_BATCH_COUNT} 个文件\n"
                f"当前选择了 {len(self.file_list)} 个，请分批处理")
            return

        # 验证文件大小限制 (≤200MB)
        oversized = []
        for path, _ in self.file_list:
            size = os.path.getsize(path)
            if size > MAX_FILE_SIZE:
                oversized.append((os.path.basename(path), size))
        if oversized:
            msg = "\n".join([f"  • {name} ({self._format_size(s)})" for name, s in oversized])
            messagebox.showerror("文件过大",
                f"以下文件超过 200MB 限制，无法处理:\n{msg}")
            return

        # 验证输出格式
        if not any([self.save_md_var.get(), self.save_docx_var.get(),
                    self.save_html_var.get(), self.save_latex_var.get(),
                    self.save_all_var.get()]):
            messagebox.showerror("错误", "请至少勾选一种输出格式")
            return

        # 创建输出目录
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            output_dir = str(Path.home() / "Desktop" / "MinerU_Output")
            self.output_dir_var.set(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        # 重置进度
        self.processing = True
        self.cancel_flag = False
        self.start_btn.configure(state=DISABLED)
        self.cancel_btn.configure(state=NORMAL)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.file_list)
        self.status_label.configure(text="处理中...")

        # 重置文件状态
        self.file_list = [(p, "待处理") for p, _ in self.file_list]
        self._refresh_file_list()

        # 解析参数
        lang_display = self.lang_rev_map.get(self.language_var.get(), "ch")
        _lang_val = self.lang_map.get(lang_display, lang_display)
        model_display = self.model_var.get()
        _model_val = self.model_map.get(model_display, "")

        extra_formats = []
        if self.save_docx_var.get():
            extra_formats.append("docx")
        if self.save_html_var.get():
            extra_formats.append("html")
        if self.save_latex_var.get():
            extra_formats.append("latex")

        logger.info("=" * 50)
        logger.info(f"开始处理 {len(self.file_list)} 个文件")
        logger.info(f"输出目录: {output_dir}")
        logger.info(f"OCR: {self.ocr_var.get()}, 公式: {self.formula_var.get()}, 表格: {self.table_var.get()}")
        logger.info(f"语言: {lang_display}, 模型: {model_display}")
        logger.info(f"额外输出格式: {extra_formats if extra_formats else '仅 Markdown'}")
        logger.info(f"保存全部（含图片）: {self.save_all_var.get()}")
        logger.info("=" * 50)

        # 启动后台线程
        thread = threading.Thread(
            target=self._process_files,
            args=(token, self.base_url_var.get().strip(), [p for p, _ in self.file_list],
                  output_dir, extra_formats, _model_val,
                  self.ocr_var.get(), self.formula_var.get(), self.table_var.get(),
                  _lang_val),
            daemon=True
        )
        thread.start()

    def _process_files(self, token, base_url, sources, output_dir,
                       extra_formats, model, ocr, formula, table, language):
        """后台处理线程"""
        start_time = time.time()
        client = None
        successful = 0
        failed = 0

        try:
            # 创建客户端
            logger.info("正在连接 MinerU API...")
            client = mineru.MinerU(token=token, base_url=base_url)

            # 逐文件处理（带进度）
            for idx, source in enumerate(sources):
                if self.cancel_flag:
                    logger.warning("用户取消了处理")
                    self.root.after(0, self.status_label.configure, {"text": "已取消"})
                    break

                fname = os.path.basename(source)
                fsize = self._format_size(os.path.getsize(source))
                logger.info(f"[{idx+1}/{len(sources)}] 正在处理: {fname} ({fsize})")

                # 更新状态
                self._update_file_status(idx, "处理中...")
                self.root.after(0, lambda v=idx+1: self.progress_bar.configure(value=v))
                self.root.after(0, lambda i=idx+1, n=len(sources): self.progress_label.configure(
                    text=f"{i}/{n}"))

                file_start = time.time()
                try:
                    # -- 单文件提取 --
                    # 仅传非空参数，避免服务端收到 null 值时报错
                    extract_kwargs = dict(source=source, timeout=300)
                    if ocr is not None:
                        extract_kwargs["ocr"] = ocr
                    # formula/table/language：默认为 True，传入 False 才表示关闭
                    if formula is not None:
                        extract_kwargs["formula"] = formula
                    if table is not None:
                        extract_kwargs["table"] = table
                    if language is not None:
                        extract_kwargs["language"] = language
                    if extra_formats:
                        extract_kwargs["extra_formats"] = extra_formats
                    result = client.extract(**extract_kwargs)

                    elapsed = time.time() - file_start
                    logger.info(f"  提取完成 (用时 {elapsed:.1f}s)")

                    # -- 保存 --
                    saved_files = []

                    # 创建以文件名命名的子目录
                    base_name = os.path.splitext(fname)[0]
                    file_output_dir = os.path.join(output_dir, base_name)
                    os.makedirs(file_output_dir, exist_ok=True)

                    if self.save_all_var.get():
                        # save_all 保存所有格式 + 图片
                        p = result.save_all(file_output_dir)
                        saved_files.append(str(p))
                    else:
                        if self.save_md_var.get():
                            p = result.save_markdown(os.path.join(file_output_dir, f"{base_name}.md"))
                            saved_files.append(str(p))
                        if self.save_docx_var.get() and result.docx:
                            p = result.save_docx(os.path.join(file_output_dir, f"{base_name}.docx"))
                            saved_files.append(str(p))
                        if self.save_html_var.get() and result.html:
                            p = result.save_html(os.path.join(file_output_dir, f"{base_name}.html"))
                            saved_files.append(str(p))
                        if self.save_latex_var.get() and result.latex:
                            p = result.save_latex(os.path.join(file_output_dir, f"{base_name}.tex"))
                            saved_files.append(str(p))
                        # 如果有图片且没选 save_all, 仍保存 MD + 图片
                        if self.save_md_var.get() and result.images and not self.save_all_var.get():
                            # save_markdown 已自动保存图片
                            pass

                    logger.info(f"  ✓ 已保存 {len(saved_files) or 'Markdown'} 个文件到 {file_output_dir}")
                    self._update_file_status(idx, "✓ 完成")
                    successful += 1

                except mineru.AuthError:
                    elapsed = time.time() - file_start
                    logger.error(f"  ✗ 认证失败，请检查 Token 是否正确")
                    self._update_file_status(idx, "✗ 认证失败")
                    failed += 1
                    self.root.after(0, lambda: messagebox.showerror("认证失败", "请检查 Token 是否正确"))
                    break

                except mineru.FileTooLargeError:
                    logger.error(f"  ✗ 文件过大: {fname}")
                    self._update_file_status(idx, "✗ 文件过大")
                    failed += 1

                except mineru.PageLimitError:
                    logger.error(f"  ✗ 页数超限: {fname}")
                    self._update_file_status(idx, "✗ 页数超限")
                    failed += 1

                except mineru.QuotaExceededError:
                    logger.error(f"  ✗ 配额已用尽")
                    self._update_file_status(idx, "✗ 配额不足")
                    failed += 1
                    break  # 配额用尽后面的也无法处理

                except mineru.ExtractFailedError as e:
                    logger.error(f"  ✗ 提取失败: {e}")
                    self._update_file_status(idx, "✗ 提取失败")
                    failed += 1

                except mineru.TimeoutError:
                    logger.error(f"  ✗ 处理超时: {fname}")
                    self._update_file_status(idx, "✗ 超时")
                    failed += 1

                except Exception as e:
                    logger.error(f"  ✗ 未知错误: {e}")
                    self._update_file_status(idx, "✗ 错误")
                    failed += 1

            # 汇总
            total = successful + failed
            total_elapsed = time.time() - start_time
            logger.info("=" * 50)
            if self.cancel_flag:
                logger.info(f"已取消 - 已处理 {successful}/{len(sources)} 个文件 (用时 {total_elapsed:.1f}s)")
            else:
                logger.info(f"处理完成! 成功: {successful}, 失败: {failed}, 用时 {total_elapsed:.1f}s")
            logger.info(f"输出目录: {output_dir}")
            logger.info("=" * 50)

            self.root.after(0, self._on_processing_done, successful, failed, total_elapsed, self.cancel_flag)

        except Exception as e:
            logger.error(f"严重错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"程序出错:\n{e}"))
            self.root.after(0, self._reset_ui)

        finally:
            if client:
                try:
                    client.close()
                except:
                    pass

    def _on_processing_done(self, successful, failed, elapsed, was_cancelled):
        """处理完成后的 UI 更新"""
        self._reset_ui()
        if was_cancelled:
            self.status_label.configure(text=f"已取消 - 成功 {successful}, 失败 {failed}")
        else:
            self.status_label.configure(
                text=f"处理完成 - 成功 {successful}, 失败 {failed}, 用时 {elapsed:.1f}s")
            if failed == 0:
                messagebox.showinfo("完成", f"全部 {successful} 个文件处理成功!\n输出目录: {self.output_dir_var.get()}")
            else:
                messagebox.showwarning("完成",
                    f"处理完成\n成功: {successful}\n失败: {failed}\n用时: {elapsed:.1f}s\n\n请查看日志了解详情。")

    def _reset_ui(self):
        self.processing = False
        self.start_btn.configure(state=NORMAL)
        self.cancel_btn.configure(state=DISABLED)
        self.progress_bar["value"] = 0
        self.progress_label.configure(text="就绪")

    def _cancel_processing(self):
        self.cancel_flag = True
        logger.warning("正在取消...")
        self.cancel_btn.configure(state=DISABLED)

    def _poll_progress(self):
        """定期检查队列（预留扩展）"""
        self.root.after(200, self._poll_progress)

    def _on_close(self):
        if self.processing:
            if not messagebox.askyesno("确认", "正在处理中，确定要退出吗？"):
                return
            self.cancel_flag = True
        self.root.destroy()


# ==================================================================
#  入口
# ==================================================================
def main():
    # 尝试设置高 DPI 支持
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = Tk()
    app = MinerUApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
