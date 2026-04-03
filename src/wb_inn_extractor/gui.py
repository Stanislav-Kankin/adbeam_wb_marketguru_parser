
from __future__ import annotations

import json
import os
import subprocess
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .excel_io import (
    analyze_workbook,
    extract_research_rows,
    read_research_row,
    read_research_rows_range,
    merge_batch_results_with_compass,
    save_batch_results,
    save_research_sample,
    summarize_research_rows_by_sheet,
)
from .models import AnalyzeSummary
from .wb_research import BatchInspector, inspect_product_row


SETTINGS_PATH = Path(".wb_inn_gui_settings.json")


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("WB INN Extractor — Research MVP")
        self.root.geometry("1440x980")
        self.root.minsize(1240, 820)

        self.input_path_var = tk.StringVar()
        self.sample_output_var = tk.StringVar(value=str(Path("output/research_sample.xlsx")))
        self.artifacts_dir_var = tk.StringVar(value=str(Path("output/artifacts")))
        self.profile_dir_var = tk.StringVar(value=str(Path("output/wb_profile")))
        self.limit_var = tk.StringVar(value="")
        self.row_var = tk.StringVar(value="2")
        self.batch_count_var = tk.StringVar(value="5")
        self.batch_output_var = tk.StringVar(value=str(Path("output/batch_results.xlsx")))
        self.merge_batch_input_var = tk.StringVar(value=str(Path("output/batch_results.xlsx")))
        self.compass_input_var = tk.StringVar()
        self.enriched_output_var = tk.StringVar(value=str(Path("output/final_enriched.xlsx")))
        self.headful_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Готово")
        self.sheet_mode_var = tk.StringVar(value="all")
        self.sheet_summary_var = tk.StringVar(value="Листы ещё не проанализированы")
        self.selection_summary_var = tk.StringVar(value="Выбор листов: все валидные после анализа")
        self.sample_summary_var = tk.StringVar(value="Sample ещё не создавался")

        self.clock_var = tk.StringVar(value="")
        self.progress_text_var = tk.StringVar(value="Ожидание запуска")
        self.progress_detail_var = tk.StringVar(value="Прогресс появится во время пакетного прогона")
        self.progress_percent_var = tk.StringVar(value="0%")
        self.progress_value_var = tk.DoubleVar(value=0.0)

        self._last_analyze_summary: AnalyzeSummary | None = None
        self._task_started_at: float | None = None
        self._progress_total: int = 0
        self._progress_done: int = 0
        self._settings_loaded_selected_sheets: list[str] = []

        self._configure_theme()
        self._load_settings()
        self._build_ui()
        self._bind_settings_save()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._tick_clock()

    def _configure_theme(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        bg = "#f4f7fb"
        card = "#ffffff"
        border = "#dbe4f0"
        muted = "#5f6f86"
        text = "#1f2a37"
        primary = "#2f6fed"
        self.root.configure(background=bg)

        style.configure(".", background=bg, foreground=text, font=("Segoe UI", 10))
        style.configure("Card.TLabelframe", background=card, bordercolor=border, relief="solid")
        style.configure("Card.TLabelframe.Label", background=bg, foreground=text, font=("Segoe UI", 10, "bold"))
        style.configure("SectionTitle.TLabel", background=bg, foreground=text, font=("Segoe UI", 15, "bold"))
        style.configure("Muted.TLabel", background=bg, foreground=muted, font=("Segoe UI", 9))
        style.configure("Body.TLabel", background=card, foreground=text)
        style.configure("HeaderValue.TLabel", background=bg, foreground=muted)
        style.configure("Primary.TButton", padding=(12, 8), foreground="#ffffff", background=primary)
        style.map("Primary.TButton", background=[("active", "#2358c9")])
        style.configure("Secondary.TButton", padding=(10, 8), background=card)
        style.configure("Flat.TButton", padding=(8, 6), background=bg)
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor=border)
        style.configure("TCheckbutton", background=card)
        style.configure("TRadiobutton", background=card)
        style.configure("TFrame", background=bg)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)

        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="WB INN Extractor", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Пайплайн: Excel → листы → sample → WB → batch_results → Compass → final_enriched",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(header, textvariable=self.clock_var, style="Muted.TLabel").grid(row=0, column=1, sticky="e")
        ttk.Label(header, textvariable=self.status_var, style="HeaderValue.TLabel").grid(row=1, column=1, sticky="e")

        self._build_source_block(container).grid(row=1, column=0, sticky="ew", pady=(0, 10))

        top_grid = ttk.Frame(container)
        top_grid.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        top_grid.columnconfigure(0, weight=1, uniform="top_halves")
        top_grid.columnconfigure(1, weight=1, uniform="top_halves")
        top_grid.rowconfigure(0, weight=1)

        self._build_sheet_block(top_grid).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_sample_block(top_grid).grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        mid_grid = ttk.Frame(container)
        mid_grid.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        mid_grid.columnconfigure(0, weight=1, uniform="mid_halves")
        mid_grid.columnconfigure(1, weight=1, uniform="mid_halves")
        mid_grid.rowconfigure(0, weight=1)

        self._build_wb_block(mid_grid).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_compass_block(mid_grid).grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self._build_status_log_block(container).grid(row=4, column=0, sticky="nsew")
        self._restore_settings_to_widgets()

    def _build_source_block(self, parent):
        frame = ttk.LabelFrame(parent, text="1. Источник данных и анализ книги", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Входной Excel", style="Body.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(frame, textvariable=self.input_path_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(frame, text="Выбрать файл", command=self._choose_input_file, style="Secondary.TButton").grid(row=0, column=2, padx=(8, 0))
        ttk.Label(
            frame,
            text="Сначала выбери книгу, затем нажми анализ. После анализа можно выбирать листы.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(
            frame,
            text="Проанализировать Excel",
            command=lambda: self._run_in_thread(self._action_analyze, task_name="Анализ Excel"),
            style="Primary.TButton",
        ).grid(row=1, column=2, sticky="e", pady=(8, 0))
        return frame

    def _build_sheet_block(self, parent):
        frame = ttk.LabelFrame(parent, text="2. Выбор листов для sample", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        ttk.Label(frame, textvariable=self.sheet_summary_var, style="Body.TLabel", justify="left").grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(frame, textvariable=self.selection_summary_var, style="Muted.TLabel", justify="left").grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 8))

        mode_frame = ttk.Frame(frame)
        mode_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Radiobutton(mode_frame, text="Использовать все валидные листы", value="all", variable=self.sheet_mode_var, command=self._update_selection_summary).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(mode_frame, text="Использовать только выбранные листы", value="selected", variable=self.sheet_mode_var, command=self._update_selection_summary).grid(row=0, column=1, sticky="w", padx=(16, 0))

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=3, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.sheet_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            height=10,
            exportselection=False,
            bg="#ffffff",
            fg="#1f2a37",
            selectbackground="#2f6fed",
            selectforeground="#ffffff",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground="#dbe4f0",
            relief="solid",
        )
        self.sheet_listbox.grid(row=0, column=0, sticky="nsew")
        self.sheet_listbox.bind("<<ListboxSelect>>", self._on_sheet_selection_changed)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.sheet_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.sheet_listbox.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=1, sticky="ns", padx=(12, 0))
        ttk.Button(buttons, text="Выбрать все", command=self._select_all_sheets, style="Secondary.TButton").grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(buttons, text="Только валидные", command=self._select_valid_sheets, style="Secondary.TButton").grid(row=1, column=0, sticky="ew", pady=8)
        ttk.Button(buttons, text="Снять выбор", command=self._clear_sheet_selection, style="Secondary.TButton").grid(row=2, column=0, sticky="ew", pady=8)
        return frame

    def _build_sample_block(self, parent):
        frame = ttk.LabelFrame(parent, text="3. Подготовка research sample", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Файл research sample", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.sample_output_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(frame, text="Выбрать", command=self._choose_sample_output, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(frame, text="Лимит строк для sample", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.limit_var, width=12).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=4)

        ttk.Button(
            frame,
            text="Создать research sample",
            command=lambda: self._run_in_thread(self._action_sample, task_name="Создание sample"),
            style="Primary.TButton",
        ).grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        ttk.Label(
            frame,
            text="Создаёт общий sample по выбранным листам и убирает дубли бренд+продавец.",
            style="Muted.TLabel",
            wraplength=380,
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w")
        ttk.Label(frame, textvariable=self.sample_summary_var, style="Body.TLabel", justify="left", wraplength=420).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        return frame

    def _build_wb_block(self, parent):
        frame = ttk.LabelFrame(parent, text="4. Парсинг Wildberries", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(1, weight=1)

        fields = [
            ("Папка артефактов", self.artifacts_dir_var, self._choose_artifacts_dir),
            ("Папка профиля WB", self.profile_dir_var, self._choose_profile_dir),
            ("Стартовая строка в sample", self.row_var, None),
            ("Размер batch", self.batch_count_var, None),
            ("Итоговый Excel batch", self.batch_output_var, self._choose_batch_output),
        ]
        for row_index, (label, var, command) in enumerate(fields):
            ttk.Label(frame, text=label, style="Body.TLabel").grid(row=row_index, column=0, sticky="w", pady=4)
            width = 12 if label in {"Стартовая строка в sample", "Размер batch"} else None
            ttk.Entry(frame, textvariable=var, width=width).grid(row=row_index, column=1, sticky="ew" if width is None else "w", padx=(8, 8), pady=4)
            if command:
                ttk.Button(frame, text="Выбрать", command=command, style="Secondary.TButton").grid(row=row_index, column=2, pady=4)

        ttk.Checkbutton(frame, text="Показывать браузер при обычном inspect", variable=self.headful_var).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 10))

        actions = ttk.Frame(frame)
        actions.grid(row=6, column=0, columnspan=3, sticky="ew")
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        ttk.Button(
            actions,
            text="Проверить строку",
            command=lambda: self._run_in_thread(self._action_inspect, task_name="Inspect строки"),
            style="Secondary.TButton",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(
            actions,
            text="Пакетный прогон",
            command=lambda: self._run_in_thread(self._action_batch, task_name="Пакетный прогон"),
            style="Primary.TButton",
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        return frame

    def _build_compass_block(self, parent):
        frame = ttk.LabelFrame(parent, text="5. Склейка с Compass", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Файл batch_results для склейки", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.merge_batch_input_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(frame, text="Выбрать", command=self._choose_merge_batch_input, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(frame, text="Выгрузка Compass", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.compass_input_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(frame, text="Выбрать", command=self._choose_compass_input, style="Secondary.TButton").grid(row=1, column=2, pady=4)

        ttk.Label(frame, text="Итоговый enriched Excel", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.enriched_output_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(frame, text="Выбрать", command=self._choose_enriched_output, style="Secondary.TButton").grid(row=2, column=2, pady=4)

        ttk.Button(
            frame,
            text="Склеить с Compass",
            command=lambda: self._run_in_thread(self._action_merge_compass, task_name="Склейка с Compass"),
            style="Primary.TButton",
        ).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        ttk.Label(
            frame,
            text="Можно выбрать любой ранее сохранённый batch_results.xlsx, даже если он лежит в другой папке.",
            style="Muted.TLabel",
            wraplength=450,
            justify="left",
        ).grid(row=4, column=0, columnspan=3, sticky="w")
        return frame

    def _build_status_log_block(self, parent):
        frame = ttk.LabelFrame(parent, text="6. Статус, прогресс и лог", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        top = ttk.Frame(frame)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)
        top.columnconfigure(2, weight=1)
        ttk.Label(top, textvariable=self.progress_text_var, style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.progress_percent_var, style="Body.TLabel").grid(row=0, column=1)
        ttk.Label(top, textvariable=self.clock_var, style="Muted.TLabel").grid(row=0, column=2, sticky="e")

        self.progressbar = ttk.Progressbar(frame, variable=self.progress_value_var, maximum=100, mode="determinate")
        self.progressbar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(frame, textvariable=self.progress_detail_var, style="Muted.TLabel").grid(row=2, column=0, sticky="ew", pady=(0, 8))

        tool_row = ttk.Frame(frame)
        tool_row.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        for idx in range(5):
            tool_row.columnconfigure(idx, weight=1)
        ttk.Button(tool_row, text="Открыть артефакты", command=self._open_artifacts_dir, style="Secondary.TButton").grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(tool_row, text="Открыть профиль WB", command=self._open_profile_dir, style="Secondary.TButton").grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(tool_row, text="Открыть batch_results.xlsx", command=self._open_batch_results, style="Secondary.TButton").grid(row=0, column=2, sticky="ew", padx=6)
        ttk.Button(tool_row, text="Открыть final_enriched.xlsx", command=self._open_enriched_output, style="Secondary.TButton").grid(row=0, column=3, sticky="ew", padx=6)
        ttk.Button(tool_row, text="Очистить лог", command=self._clear_log, style="Secondary.TButton").grid(row=0, column=4, sticky="ew", padx=(6, 0))

        log_frame = ttk.Frame(frame)
        log_frame.grid(row=4, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        frame.rowconfigure(4, weight=1)

        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            bg="#ffffff",
            fg="#1f2a37",
            insertbackground="#1f2a37",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground="#dbe4f0",
            relief="solid",
            font=("Consolas", 10),
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        return frame

    def _bind_settings_save(self) -> None:
        self._tracked_vars = [
            self.input_path_var,
            self.sample_output_var,
            self.artifacts_dir_var,
            self.profile_dir_var,
            self.limit_var,
            self.row_var,
            self.batch_count_var,
            self.batch_output_var,
            self.merge_batch_input_var,
            self.compass_input_var,
            self.enriched_output_var,
            self.sheet_mode_var,
        ]
        for var in self._tracked_vars:
            var.trace_add("write", lambda *_: self._save_settings())
        self.headful_var.trace_add("write", lambda *_: self._save_settings())


    @staticmethod
    def _format_duration_hms(seconds: float) -> str:
        total_seconds = max(int(round(seconds)), 0)
        hours, remainder = divmod(total_seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"

    def _tick_clock(self) -> None:
        self.clock_var.set(datetime.now().strftime("Сейчас: %d.%m.%Y %H:%M:%S"))
        self.root.after(1000, self._tick_clock)

    def _load_settings(self) -> None:
        self._loaded_settings = {}
        if not SETTINGS_PATH.exists():
            return
        try:
            self._loaded_settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            self._loaded_settings = {}

    def _restore_settings_to_widgets(self) -> None:
        settings = getattr(self, "_loaded_settings", {})
        if not settings:
            return
        for key, var in {
            "input_path": self.input_path_var,
            "sample_output": self.sample_output_var,
            "artifacts_dir": self.artifacts_dir_var,
            "profile_dir": self.profile_dir_var,
            "limit": self.limit_var,
            "row": self.row_var,
            "batch_count": self.batch_count_var,
            "batch_output": self.batch_output_var,
            "merge_batch_input": self.merge_batch_input_var,
            "compass_input": self.compass_input_var,
            "enriched_output": self.enriched_output_var,
            "sheet_mode": self.sheet_mode_var,
        }.items():
            value = settings.get(key)
            if isinstance(value, str):
                var.set(value)
        if "headful" in settings:
            self.headful_var.set(bool(settings["headful"]))
        loaded = settings.get("selected_sheets")
        if isinstance(loaded, list):
            self._settings_loaded_selected_sheets = [str(x) for x in loaded]
        summary = settings.get("selection_summary")
        if isinstance(summary, str) and summary:
            self.selection_summary_var.set(summary)
        sample_summary = settings.get("sample_summary")
        if isinstance(sample_summary, str) and sample_summary:
            self.sample_summary_var.set(sample_summary)

    def _save_settings(self) -> None:
        try:
            data = {
                "input_path": self.input_path_var.get().strip(),
                "sample_output": self.sample_output_var.get().strip(),
                "artifacts_dir": self.artifacts_dir_var.get().strip(),
                "profile_dir": self.profile_dir_var.get().strip(),
                "limit": self.limit_var.get().strip(),
                "row": self.row_var.get().strip(),
                "batch_count": self.batch_count_var.get().strip(),
                "batch_output": self.batch_output_var.get().strip(),
                "merge_batch_input": self.merge_batch_input_var.get().strip(),
                "compass_input": self.compass_input_var.get().strip(),
                "enriched_output": self.enriched_output_var.get().strip(),
                "sheet_mode": self.sheet_mode_var.get(),
                "selected_sheets": self._get_selected_sheet_names_from_ui() if hasattr(self, "sheet_listbox") else [],
                "selection_summary": self.selection_summary_var.get(),
                "sample_summary": self.sample_summary_var.get(),
                "headful": bool(self.headful_var.get()),
            }
            SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _on_close(self) -> None:
        self._save_settings()
        self.root.destroy()

    def _choose_input_file(self) -> None:
        path = filedialog.askopenfilename(title="Выбери Excel-файл", filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if path:
            self.input_path_var.set(path)
            self._last_analyze_summary = None
            self.sheet_listbox.delete(0, "end")
            self.sheet_summary_var.set("Файл выбран. Нажми 'Проанализировать Excel', чтобы увидеть листы.")
            self.selection_summary_var.set("Выбор листов: все валидные после анализа")
            self._save_settings()

    def _choose_sample_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить research sample",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.sample_output_var.get()).name,
        )
        if path:
            self.sample_output_var.set(path)

    def _choose_artifacts_dir(self) -> None:
        path = filedialog.askdirectory(title="Папка для артефактов")
        if path:
            self.artifacts_dir_var.set(path)

    def _choose_profile_dir(self) -> None:
        path = filedialog.askdirectory(title="Папка профиля WB")
        if path:
            self.profile_dir_var.set(path)

    def _choose_compass_input(self) -> None:
        path = filedialog.askopenfilename(title="Выбери выгрузку Compass", filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if path:
            self.compass_input_var.set(path)

    def _choose_enriched_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить enriched Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.enriched_output_var.get()).name,
        )
        if path:
            self.enriched_output_var.set(path)

    def _choose_batch_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить batch_results",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.batch_output_var.get()).name,
        )
        if path:
            self.batch_output_var.set(path)

    def _choose_merge_batch_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери batch_results.xlsx для склейки",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            self.merge_batch_input_var.set(path)

    def _open_artifacts_dir(self) -> None:
        self._open_dir(self.artifacts_dir_var.get())

    def _open_profile_dir(self) -> None:
        self._open_dir(self.profile_dir_var.get())

    def _open_batch_results(self) -> None:
        path = Path(self.batch_output_var.get().strip() or "output/batch_results.xlsx")
        if not path.exists():
            raise FileNotFoundError(f"Не найден файл: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_enriched_output(self) -> None:
        path = Path(self.enriched_output_var.get().strip() or "output/final_enriched.xlsx")
        if not path.exists():
            raise FileNotFoundError(f"Не найден файл: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    @staticmethod
    def _open_dir(path_str: str) -> None:
        path = Path(path_str)
        path.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif os.name == "posix":
            subprocess.Popen(["xdg-open", str(path)])
        else:
            raise RuntimeError("Открытие папки не поддерживается на этой ОС")

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", "end")

    def _run_in_thread(self, target, task_name: str = "Операция") -> None:
        worker = threading.Thread(target=lambda: self._safe_run(target, task_name=task_name), daemon=True)
        worker.start()

    def _safe_run(self, target, task_name: str = "Операция") -> None:
        self._begin_task(task_name)
        try:
            target()
            self._finish_task(success=True, task_name=task_name)
        except Exception as exc:
            self._append_log("\n[ERROR]\n")
            self._append_log(f"{exc}\n")
            self._append_log(traceback.format_exc())
            self._finish_task(success=False, task_name=task_name)
            self.root.after(0, lambda exc=exc: messagebox.showerror("Ошибка", str(exc)))

    def _begin_task(self, task_name: str) -> None:
        self._task_started_at = time.monotonic()
        self._set_status(f"Выполняется: {task_name}")
        self.root.after(0, lambda: self.progress_text_var.set(f"Текущая операция: {task_name}"))
        self.root.after(0, lambda: self.progress_detail_var.set("Операция запущена"))
        self.root.after(0, lambda: self.progress_percent_var.set("0%"))
        self.root.after(0, lambda: self.progress_value_var.set(0.0))

    def _finish_task(self, success: bool, task_name: str) -> None:
        elapsed = 0.0
        if self._task_started_at is not None:
            elapsed = time.monotonic() - self._task_started_at
        status = "Готово" if success else "Ошибка"
        self._set_status(status)
        self.root.after(0, lambda success=success, task_name=task_name, elapsed=elapsed: self._finish_task_ui(success, task_name, elapsed))
        self._task_started_at = None

    def _finish_task_ui(self, success: bool, task_name: str, elapsed: float) -> None:
        human_elapsed = self._format_duration_hms(elapsed)
        label = f"{task_name}: завершено за {human_elapsed}" if success else f"{task_name}: завершилось с ошибкой через {human_elapsed}"
        self.progress_text_var.set(label)
        self.progress_detail_var.set("Лог ниже содержит детали последней операции")
        self.progress_percent_var.set("100%" if success else self.progress_percent_var.get())
        if success and self.progress_value_var.get() < 100:
            self.progress_value_var.set(100)

    def _set_batch_progress(self, done: int, total: int, row_number: int | None = None, seller: str | None = None) -> None:
        self._progress_done = done
        self._progress_total = total
        elapsed = time.monotonic() - self._task_started_at if self._task_started_at else 0.0
        percent = (done / total * 100.0) if total else 0.0
        avg = (elapsed / done) if done else 0.0
        eta = avg * max(total - done, 0)
        if row_number is not None:
            title = f"Пакетный прогон: {done}/{total} строк | текущая строка {row_number}"
        else:
            title = f"Пакетный прогон: {done}/{total} строк"
        if seller:
            title += f" | seller={seller}"
        detail = (
            f"Прошло: {self._format_duration_hms(elapsed)} | "
            f"ETA: {self._format_duration_hms(eta)} | "
            f"Среднее: {avg:.1f} сек/строка"
        ) if done else "Идёт подготовка batch..."
        self.root.after(0, lambda title=title, detail=detail, percent=percent: self._set_batch_progress_ui(title, detail, percent))

    def _set_batch_progress_ui(self, title: str, detail: str, percent: float) -> None:
        self.progress_text_var.set(title)
        self.progress_detail_var.set(detail)
        self.progress_percent_var.set(f"{percent:.1f}%")
        self.progress_value_var.set(percent)

    def _action_analyze(self) -> None:
        input_path = self._require_input_path()
        summary = analyze_workbook(input_path)
        self._last_analyze_summary = summary
        self.root.after(0, lambda summary=summary: self._apply_analyze_summary(summary))
        self._append_log(self._format_analyze_summary(summary))

    def _action_sample(self) -> None:
        input_path = self._require_input_path()
        output_path = Path(self.sample_output_var.get())
        raw_limit = self.limit_var.get().strip()
        limit = self._parse_positive_int(raw_limit, field_name="Лимит строк") if raw_limit else None
        selected_sheets = self._resolve_selected_sheets_for_sample()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        rows = extract_research_rows(input_path, limit=limit, selected_sheets=selected_sheets)
        if not rows:
            raise ValueError("Не удалось собрать ни одной строки для research sample")

        save_research_sample(output_path, rows)
        rows_by_sheet = summarize_research_rows_by_sheet(rows)

        sheet_scope_label = self._build_sheet_scope_label(selected_sheets)
        self._append_log(f"Создан файл: {output_path}\n")
        self._append_log(f"Листы для sample: {sheet_scope_label}\n")
        if limit is None:
            self._append_log("Лимит строк не задан: в sample сохранены все уникальные продавцы по паре бренд+продавец\n")
        self._append_log(f"Строк (уникальных по продавцу+бренду): {len(rows)}\n")
        self._append_log("Распределение строк по листам в sample:\n")
        for sheet_name, count in rows_by_sheet.items():
            self._append_log(f"  - {sheet_name}: {count}\n")

        self.root.after(0, lambda: self.row_var.set("2"))
        sample_summary = (
            f"Последний sample: {len(rows)} строк\n"
            f"Листов в sample: {len(rows_by_sheet)}\n"
            f"Рекомендуемая стартовая строка для batch: 2"
        )
        self.root.after(0, lambda sample_summary=sample_summary: self.sample_summary_var.set(sample_summary))
        self.root.after(0, lambda: self.selection_summary_var.set(
            f"Последний sample: {len(rows)} строк | листов в sample: {len(rows_by_sheet)} | стартовая строка = 2"
        ))

    def _action_inspect(self) -> None:
        sample_path = self._require_sample_path()
        row_number = self._parse_positive_int(self.row_var.get(), field_name="Номер строки")
        artifacts_dir = Path(self.artifacts_dir_var.get())
        profile_dir = self._optional_profile_dir()

        if profile_dir is not None and artifacts_dir.resolve() == profile_dir.resolve():
            raise ValueError("Папка артефактов и папка профиля WB должны быть разными")

        artifacts_dir.mkdir(parents=True, exist_ok=True)
        if profile_dir is not None:
            profile_dir.mkdir(parents=True, exist_ok=True)
            self._append_log(f"Обычный inspect запущен в режиме persistent profile: {profile_dir}\n")
        else:
            self._append_log("Обычный inspect запущен во временном браузере без профиля WB\n")

        research_row = read_research_row(sample_path, row_number=row_number)
        result = inspect_product_row(
            row_number=row_number,
            research_row=research_row,
            artifacts_dir=artifacts_dir,
            headful=self.headful_var.get() or profile_dir is not None,
            profile_dir=profile_dir,
        )
        self._append_log(f"Inspect: sheet={research_row.source_sheet}, source_row_index={research_row.source_row_index}\n")
        self._append_log(f"Inspect: seller={research_row.seller_name_raw}, brand={research_row.brand}\n")
        self._append_log(f"Inspect: status={result.parse_status}, inn={result.inn}, ogrn={result.ogrn}, ogrnip={result.ogrnip}\n")
        self.root.after(0, lambda: messagebox.showinfo("Проверка завершена", f"Артефакты сохранены в:\n{artifacts_dir}"))

    def _action_batch(self) -> None:
        sample_path = self._require_sample_path()
        start_row = self._parse_positive_int(self.row_var.get(), field_name="Стартовая строка")
        batch_count = self._parse_positive_int(self.batch_count_var.get(), field_name="Количество строк в batch")
        artifacts_dir = Path(self.artifacts_dir_var.get())
        profile_dir = self._required_profile_dir()
        batch_output = Path(self.batch_output_var.get())

        if artifacts_dir.resolve() == profile_dir.resolve():
            raise ValueError("Папка артефактов и папка профиля WB должны быть разными")

        artifacts_dir.mkdir(parents=True, exist_ok=True)
        profile_dir.mkdir(parents=True, exist_ok=True)
        batch_output.parent.mkdir(parents=True, exist_ok=True)

        research_rows = read_research_rows_range(sample_path, start_row=start_row, limit=batch_count)
        total = len(research_rows)
        self._append_log(f"Batch: старт диапазона строк {start_row}-{start_row + max(total - 1, 0)} | sample={sample_path}\n")
        output_rows = []
        self._set_batch_progress(done=0, total=max(total, 1))

        with BatchInspector(artifacts_dir=artifacts_dir, headful=True, profile_dir=profile_dir) as inspector:
            for offset, research_row in enumerate(research_rows, start=0):
                row_number = start_row + offset
                self._set_batch_progress(done=offset, total=total, row_number=row_number, seller=research_row.seller_name_raw)
                self._append_log(
                    f"Batch: обрабатываю строку {row_number}/{start_row + total - 1} | sheet={research_row.source_sheet} | seller={research_row.seller_name_raw}\n"
                )
                result = inspector.inspect_row(
                    row_number=row_number,
                    research_row=research_row,
                )
                output_rows.append({
                    "row_number": row_number,
                    "source_sheet": research_row.source_sheet,
                    "source_row_index": research_row.source_row_index,
                    "product_name": research_row.product_name,
                    "wb_nm_id": research_row.wb_nm_id,
                    "brand": research_row.brand,
                    "seller_name_raw": research_row.seller_name_raw,
                    "wb_candidate_url": research_row.wb_candidate_url,
                    "final_url": result.final_url,
                    "seller_url": result.seller_url,
                    "navigated_to_seller_page": result.navigated_to_seller_page,
                    "seller_display_name": result.seller_display_name,
                    "entity_type": result.entity_type,
                    "inn": result.inn,
                    "ogrn": result.ogrn,
                    "ogrnip": result.ogrnip,
                    "parse_status": result.parse_status,
                    "parse_note": result.note,
                    "http_status": result.http_status,
                    "used_persistent_profile": result.used_persistent_profile,
                    "screenshot_path": result.screenshot_path,
                    "html_path": result.html_path,
                    "text_path": result.text_path,
                    "marketguru_source_sheet": research_row.source_sheet,
                    "marketguru_source_row_index": research_row.source_row_index,
                    "marketguru_product_name": research_row.product_name,
                    "marketguru_brand": research_row.brand,
                    "marketguru_seller_name": research_row.seller_name_raw,
                    "marketguru_wb_nm_id": research_row.wb_nm_id,
                    "marketguru_candidate_url": research_row.wb_candidate_url,
                    "wb_seller_name": result.seller_display_name,
                    "wb_seller_url": result.seller_url,
                })
                self._append_log(f"Batch: строка {row_number} завершена, status={result.parse_status}, inn={result.inn}\n")

        save_batch_results(batch_output, output_rows)
        self.root.after(0, lambda batch_output=batch_output: self.merge_batch_input_var.set(str(batch_output)))
        self._set_batch_progress(done=total, total=max(total, 1), row_number=start_row + total - 1 if total else start_row)
        next_row = start_row + total
        self.root.after(0, lambda next_row=next_row: self.row_var.set(str(next_row)))
        self._append_log(f"Batch: итоговый файл сохранён: {batch_output}\n")
        self._append_log(f"Batch: рекомендованная следующая стартовая строка = {next_row}\n")
        self.root.after(0, lambda: messagebox.showinfo("Batch завершён", f"Итоговый Excel сохранён:\n{batch_output}"))

    def _action_merge_compass(self) -> None:
        batch_results_value = self.merge_batch_input_var.get().strip() or self.batch_output_var.get().strip() or "output/batch_results.xlsx"
        batch_results_path = Path(batch_results_value)
        if not batch_results_path.exists():
            raise FileNotFoundError(f"Не найден batch_results.xlsx: {batch_results_path}")

        compass_value = self.compass_input_var.get().strip()
        if not compass_value:
            raise ValueError("Сначала выбери выгрузку Compass")
        compass_path = Path(compass_value)
        if not compass_path.exists():
            raise FileNotFoundError(f"Не найден файл Compass: {compass_path}")

        output_path = Path(self.enriched_output_var.get().strip() or "output/final_enriched.xlsx")
        output_path.parent.mkdir(parents=True, exist_ok=True)

        self._append_log(f"Merge Compass: batch={batch_results_path}\n")
        self._append_log(f"Merge Compass: compass={compass_path}\n")
        summary = merge_batch_results_with_compass(
            batch_results_path=batch_results_path,
            compass_path=compass_path,
            output_path=output_path,
        )
        match_rate = (summary["matched_rows"] / summary["batch_rows"] * 100) if summary["batch_rows"] else 0.0
        self._append_log(
            "Merge Compass: "
            f"batch_rows={summary['batch_rows']}, matched={summary['matched_rows']}, unmatched={summary['unmatched_rows']}, match_rate={match_rate:.1f}%\n"
        )
        self._append_log(
            "Merge Compass: "
            f"sheet={summary['compass_sheet_name']}, inn_header={summary['compass_inn_header']}, indexed={summary['compass_rows_indexed']}\n"
        )
        self._append_log(f"Merge Compass: создан файл {output_path}\n")
        self.root.after(0, lambda: messagebox.showinfo("Merge Compass завершён", f"Итоговый enriched Excel сохранён:\n{output_path}"))

    def _require_input_path(self) -> Path:
        input_value = self.input_path_var.get().strip()
        if not input_value:
            raise ValueError("Сначала выбери входной Excel")
        input_path = Path(input_value)
        if not input_path.exists():
            raise FileNotFoundError(f"Файл не найден: {input_path}")
        return input_path

    def _require_sample_path(self) -> Path:
        sample_path = Path(self.sample_output_var.get())
        if not sample_path.exists():
            raise FileNotFoundError(f"Не найден sample-файл: {sample_path}")
        return sample_path

    def _required_profile_dir(self) -> Path:
        profile_dir = self._optional_profile_dir()
        if profile_dir is None:
            raise ValueError("Укажи папку профиля WB")
        return profile_dir

    def _optional_profile_dir(self) -> Path | None:
        raw = self.profile_dir_var.get().strip()
        if not raw:
            return None
        return Path(raw)

    def _build_sheet_scope_label(self, selected_sheets: list[str] | None) -> str:
        if not selected_sheets:
            return "все валидные листы"
        if len(selected_sheets) <= 5:
            return ", ".join(selected_sheets)
        preview = ", ".join(selected_sheets[:5])
        return f"{preview} ... (+ещё {len(selected_sheets) - 5})"

    def _update_selection_summary(self) -> None:
        summary = self._last_analyze_summary
        if summary is None:
            self.selection_summary_var.set("Выбор листов: сначала проанализируй книгу")
            return

        if self.sheet_mode_var.get() == "all":
            self.selection_summary_var.set(
                f"Выбор листов: все валидные ({len(summary.valid_sheet_names)})"
            )
            self._save_settings()
            return

        selected_names = self._get_selected_sheet_names_from_ui()
        if not selected_names:
            self.selection_summary_var.set("Выбор листов: ничего не выбрано")
            self._save_settings()
            return

        self.selection_summary_var.set(
            f"Выбор листов: {len(selected_names)} | {self._build_sheet_scope_label(selected_names)}"
        )
        self._save_settings()

    def _on_sheet_selection_changed(self, _event=None) -> None:
        self._update_selection_summary()

    def _resolve_selected_sheets_for_sample(self) -> list[str] | None:
        summary = self._last_analyze_summary
        if summary is None:
            return None

        if self.sheet_mode_var.get() == "all":
            return list(summary.valid_sheet_names)

        selected_names = self._get_selected_sheet_names_from_ui()
        if not selected_names:
            raise ValueError("В режиме 'только выбранные листы' нужно выбрать хотя бы один лист")

        valid_set = set(summary.valid_sheet_names)
        invalid_selected = [sheet_name for sheet_name in selected_names if sheet_name not in valid_set]
        if invalid_selected:
            raise ValueError("Нельзя строить sample по невалидным листам: " + ", ".join(invalid_selected))
        return selected_names

    def _apply_analyze_summary(self, summary: AnalyzeSummary) -> None:
        self.sheet_listbox.delete(0, "end")
        valid_set = set(summary.valid_sheet_names)
        for sheet_name in summary.selected_sheet_names:
            status = "VALID" if sheet_name in valid_set else "SKIPPED"
            self.sheet_listbox.insert("end", f"[{status}] {sheet_name}")

        if self._settings_loaded_selected_sheets:
            self.sheet_listbox.selection_clear(0, "end")
            wanted = set(self._settings_loaded_selected_sheets)
            for index in range(self.sheet_listbox.size()):
                raw_text = self.sheet_listbox.get(index)
                if self._sheet_name_from_list_item(raw_text) in wanted:
                    self.sheet_listbox.selection_set(index)
            if not self.sheet_listbox.curselection():
                self._select_valid_sheets()
            self._settings_loaded_selected_sheets = []
        else:
            self._select_valid_sheets()

        self._update_selection_summary()
        self.sheet_summary_var.set(
            "Всего листов: "
            f"{summary.workbook_sheet_count} | валидных: {len(summary.valid_sheet_names)} | пропущенных: {len(summary.skipped_sheet_names)}"
        )

    def _format_analyze_summary(self, summary: AnalyzeSummary) -> str:
        lines = [
            f"Книга: {summary.input_path}",
            f"Всего листов: {summary.workbook_sheet_count}",
            f"Валидных листов: {len(summary.valid_sheet_names)}",
            f"Пропущенных листов: {len(summary.skipped_sheet_names)}",
        ]
        for sheet in summary.sheets:
            line = f"- {sheet.sheet_name}: {sheet.status}"
            if sheet.status == "VALID":
                line += f", data_rows={sheet.data_rows_count}, header_row={sheet.detected_header_row}"
            elif sheet.reason:
                line += f", reason={sheet.reason}"
            lines.append(line)
        lines.append("")
        return "\n".join(lines)

    def _get_selected_sheet_names_from_ui(self) -> list[str]:
        names: list[str] = []
        for index in self.sheet_listbox.curselection():
            raw_text = self.sheet_listbox.get(index)
            names.append(self._sheet_name_from_list_item(raw_text))
        return names

    def _select_all_sheets(self) -> None:
        if self.sheet_listbox.size() == 0:
            return
        self.sheet_listbox.selection_set(0, "end")
        self._update_selection_summary()

    def _select_valid_sheets(self) -> None:
        self.sheet_listbox.selection_clear(0, "end")
        for index in range(self.sheet_listbox.size()):
            raw_text = self.sheet_listbox.get(index)
            if raw_text.startswith("[VALID]"):
                self.sheet_listbox.selection_set(index)
        self._update_selection_summary()

    def _clear_sheet_selection(self) -> None:
        self.sheet_listbox.selection_clear(0, "end")
        self._update_selection_summary()

    @staticmethod
    def _sheet_name_from_list_item(raw_text: str) -> str:
        parts = raw_text.split("] ", maxsplit=1)
        if len(parts) == 2:
            return parts[1]
        return raw_text

    @staticmethod
    def _parse_positive_int(raw_value: str, field_name: str) -> int:
        try:
            value = int(raw_value)
        except ValueError as exc:
            raise ValueError(f"{field_name}: требуется целое число") from exc
        if value <= 0:
            raise ValueError(f"{field_name}: число должно быть > 0")
        return value

    def _append_log(self, text: str) -> None:
        self.root.after(0, lambda text=text: self._append_log_on_ui(text))

    def _append_log_on_ui(self, text: str) -> None:
        self.log_text.insert("end", text)
        self.log_text.see("end")

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda text=text: self.status_var.set(text))


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
