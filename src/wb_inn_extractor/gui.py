
from __future__ import annotations

from contextlib import ExitStack
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

from adbeam_excel_parser.audit_runner import attach_output_file_path, run_excel_audit
from adbeam_excel_parser.conference_contacts import (
    build_icp_contacts_output_path,
    export_icp_contacts_queue,
    read_icp_contacts_summary,
)
from adbeam_excel_parser.excel_exporter import build_output_path, export_audit_to_excel
from adbeam_excel_parser.excel_reader import read_excel_summary
from adbeam_excel_parser.icp_autosearch import (
    build_icp_autosearch_output_path,
    run_icp_autosearch,
)
from .excel_merger import merge_excel_files_to_tabs
from .excel_io import (
    analyze_workbook,
    build_known_seller_batch_row,
    build_viewed_seller_batch_row,
    discover_batch_results,
    extract_research_rows,
    import_seller_history_from_registry_files,
    load_known_seller_sources,
    merge_inn_registry_files,
    read_research_row,
    read_research_rows_range,
    merge_batch_results_with_compass,
    save_batch_results,
    save_research_sample,
    summarize_research_rows_by_sheet,
    update_inn_registry_from_batch_files,
    update_seller_history_from_batch_files,
)
from .models import AnalyzeSummary
from .wb_research import BatchInspector, build_row_error_result, inspect_product_row


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
        self.registry_path_var = tk.StringVar(value=str(Path("output/inn_registry.xlsx")))
        self.seller_history_path_var = tk.StringVar(value=str(Path("output/seller_history.xlsx")))
        self.registry_new_inn_output_var = tk.StringVar(value=str(Path("output/new_inn_for_kontur.xlsx")))
        self.registry_merge_primary_var = tk.StringVar(value=str(Path("output/inn_registry.xlsx")))
        self.registry_merge_secondary_var = tk.StringVar()
        self.registry_merge_output_var = tk.StringVar(value=str(Path("output/inn_registry_merged.xlsx")))
        self.registry_summary_var = tk.StringVar(value="Реестр ещё не обновлялся")
        self.excel_merge_source_var = tk.StringVar()
        self.excel_merge_output_xlsx_var = tk.StringVar(value=str(Path("output/merged_tabs.xlsx")))
        self.excel_merge_output_zip_var = tk.StringVar(value=str(Path("output/merged_tabs.zip")))
        self.excel_merge_create_zip_var = tk.BooleanVar(value=True)
        self.excel_merge_summary_var = tk.StringVar(value="Объединение Excel ещё не запускалось")
        self.sample_use_known_sellers_var = tk.BooleanVar(value=True)
        self.sample_filter_summary_var = tk.StringVar(
            value="Фильтры повторов: включены, подробности появятся после создания sample"
        )
        self.headful_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Готово")
        self.sheet_mode_var = tk.StringVar(value="all")
        self.sheet_summary_var = tk.StringVar(value="Листы ещё не проанализированы")
        self.selection_summary_var = tk.StringVar(value="Выбор листов: все валидные после анализа")
        self.sample_summary_var = tk.StringVar(value="Sample ещё не создавался")
        self.adbeam_input_path_var = tk.StringVar()
        self.adbeam_output_path_var = tk.StringVar()
        self.adbeam_summary_var = tk.StringVar(value="AdBeam аудит ещё не запускался")
        self.adbeam_progress_var = tk.StringVar(value="Ожидание Excel-файла для аудита сайтов")
        self.conference_input_path_var = tk.StringVar()
        self.conference_output_path_var = tk.StringVar(value=str(Path("output/icp1_contacts_queue.xlsx")))
        self.conference_autosearch_output_path_var = tk.StringVar(value=str(Path("output/icp1_autosearch.xlsx")))
        self.conference_autosearch_limit_var = tk.StringVar(value="1200")
        self.conference_autosearch_delay_var = tk.StringVar(value="1.5")
        self.conference_autosearch_only_p1_var = tk.BooleanVar(value=False)
        self.conference_summary_var = tk.StringVar(value="Конференция ICP-1 ещё не запускалась")
        self.conference_progress_var = tk.StringVar(value="Ожидание Excel-файла со списком компаний")
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
        self._registry_batch_files: list[str] = []

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
        container.rowconfigure(1, weight=1)

        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="WB INN Extractor", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.progress_text_var, style="Body.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )
        ttk.Label(header, textvariable=self.clock_var, style="Muted.TLabel").grid(row=0, column=1, sticky="e")
        header_status = ttk.Frame(header)
        header_status.grid(row=1, column=1, sticky="e", pady=(4, 0))
        ttk.Label(header_status, textvariable=self.status_var, style="HeaderValue.TLabel").grid(row=0, column=0)
        ttk.Label(header_status, textvariable=self.progress_percent_var, style="Body.TLabel").grid(
            row=0, column=1, padx=(12, 0)
        )
        ttk.Progressbar(header, variable=self.progress_value_var, maximum=100, mode="determinate").grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0)
        )
        ttk.Label(header, textvariable=self.progress_detail_var, style="Muted.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )

        notebook = ttk.Notebook(container)
        notebook.grid(row=1, column=0, sticky="nsew")

        extract_tab = ttk.Frame(notebook, padding=8)
        conference_tab = ttk.Frame(notebook, padding=8)
        adbeam_tab = ttk.Frame(notebook, padding=8)
        registry_tab = ttk.Frame(notebook, padding=8)
        kontur_tab = ttk.Frame(notebook, padding=8)
        excel_merge_tab = ttk.Frame(notebook, padding=8)
        notebook.add(extract_tab, text="Извлечение ИНН")
        notebook.add(registry_tab, text="Реестр ИНН")
        notebook.add(kontur_tab, text="Склейка с Контур")
        notebook.add(excel_merge_tab, text="Объединение Excel")
        notebook.add(adbeam_tab, text="Аудит сайтов")
        notebook.add(conference_tab, text="Конференция ICP-1")

        extract_tab.columnconfigure(0, weight=1)
        extract_tab.rowconfigure(0, weight=1)
        self.extract_canvas = tk.Canvas(
            extract_tab,
            background="#f4f7fb",
            borderwidth=0,
            highlightthickness=0,
        )
        self.extract_canvas.grid(row=0, column=0, sticky="nsew")
        extract_scrollbar = ttk.Scrollbar(extract_tab, orient="vertical", command=self.extract_canvas.yview)
        extract_scrollbar.grid(row=0, column=1, sticky="ns")
        self.extract_canvas.configure(yscrollcommand=extract_scrollbar.set)

        extract_content = ttk.Frame(self.extract_canvas, padding=(0, 0, 8, 8))
        extract_content.columnconfigure(0, weight=1)
        extract_window = self.extract_canvas.create_window((0, 0), window=extract_content, anchor="nw")
        extract_content.bind(
            "<Configure>",
            lambda _event: self.extract_canvas.configure(scrollregion=self.extract_canvas.bbox("all")),
        )
        self.extract_canvas.bind(
            "<Configure>",
            lambda event: self.extract_canvas.itemconfigure(extract_window, width=event.width),
        )

        def scroll_extract(event) -> None:
            if event.delta:
                self.extract_canvas.yview_scroll(-int(event.delta / 120), "units")

        self.extract_canvas.bind(
            "<Enter>", lambda _event: self.extract_canvas.bind_all("<MouseWheel>", scroll_extract)
        )
        self.extract_canvas.bind(
            "<Leave>", lambda _event: self.extract_canvas.unbind_all("<MouseWheel>")
        )

        self._build_source_block(extract_content).grid(row=0, column=0, sticky="ew", pady=(0, 10))

        top_grid = ttk.Frame(extract_content)
        top_grid.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        top_grid.columnconfigure(0, weight=1, uniform="top_halves")
        top_grid.columnconfigure(1, weight=1, uniform="top_halves")
        top_grid.rowconfigure(0, weight=1)

        self._build_sheet_block(top_grid).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_sample_block(top_grid).grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        mid_grid = ttk.Frame(extract_content)
        mid_grid.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        mid_grid.columnconfigure(0, weight=1)
        mid_grid.rowconfigure(0, weight=1)

        self._build_wb_block(mid_grid).grid(row=0, column=0, sticky="nsew")

        self._build_status_log_block(extract_content).grid(row=4, column=0, sticky="nsew")

        conference_tab.columnconfigure(0, weight=1)
        conference_tab.rowconfigure(0, weight=1)
        self._build_conference_tab(conference_tab).grid(row=0, column=0, sticky="nsew")

        adbeam_tab.columnconfigure(0, weight=1)
        adbeam_tab.rowconfigure(0, weight=1)
        self._build_adbeam_tab(adbeam_tab).grid(row=0, column=0, sticky="nsew")

        registry_tab.columnconfigure(0, weight=1)
        registry_tab.rowconfigure(0, weight=1)
        self._build_registry_tab(registry_tab).grid(row=0, column=0, sticky="nsew")

        kontur_tab.columnconfigure(0, weight=1)
        kontur_tab.rowconfigure(0, weight=1)
        self._build_compass_block(kontur_tab).grid(row=0, column=0, sticky="new")

        excel_merge_tab.columnconfigure(0, weight=1)
        excel_merge_tab.rowconfigure(0, weight=1)
        self._build_excel_merge_tab(excel_merge_tab).grid(row=0, column=0, sticky="nsew")
        self._restore_settings_to_widgets()

    def _build_conference_tab(self, parent):
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        source = ttk.LabelFrame(frame, text="Конференция ICP-1: очередь обогащения 1200 компаний", padding=12, style="Card.TLabelframe")
        source.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        source.columnconfigure(1, weight=1)

        ttk.Label(source, text="Входной Excel", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(source, textvariable=self.conference_input_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(source, text="Выбрать файл", command=self._choose_conference_input_file, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(source, text="Очередь Excel", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(source, textvariable=self.conference_output_path_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(source, text="Выбрать", command=self._choose_conference_output_file, style="Secondary.TButton").grid(row=1, column=2, pady=4)

        ttk.Label(source, text="Автопоиск Excel", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(source, textvariable=self.conference_autosearch_output_path_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(source, text="Выбрать", command=self._choose_conference_autosearch_output_file, style="Secondary.TButton").grid(row=2, column=2, pady=4)

        options = ttk.Frame(source)
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Label(options, text="Лимит строк", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.conference_autosearch_limit_var, width=8).grid(row=0, column=1, sticky="w", padx=(8, 18))
        ttk.Label(options, text="Пауза, сек", style="Body.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(options, textvariable=self.conference_autosearch_delay_var, width=8).grid(row=0, column=3, sticky="w", padx=(8, 18))
        ttk.Checkbutton(options, text="Только P1", variable=self.conference_autosearch_only_p1_var).grid(row=0, column=4, sticky="w")

        actions = ttk.Frame(source)
        actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        for index in range(5):
            actions.columnconfigure(index, weight=1)
        ttk.Button(
            actions,
            text="Проверить базу",
            command=lambda: self._run_in_thread(self._action_conference_analyze, task_name="ICP-1 анализ базы"),
            style="Secondary.TButton",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(
            actions,
            text="Автопоиск сайтов",
            command=lambda: self._run_in_thread(self._action_conference_autosearch, task_name="ICP-1 автопоиск"),
            style="Primary.TButton",
        ).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(
            actions,
            text="Создать очередь 1200",
            command=lambda: self._run_in_thread(self._action_conference_export, task_name="ICP-1 очередь контактов"),
            style="Secondary.TButton",
        ).grid(row=0, column=2, sticky="ew", padx=6)
        ttk.Button(actions, text="Открыть автопоиск", command=self._open_conference_autosearch_output_file, style="Secondary.TButton").grid(row=0, column=3, sticky="ew", padx=6)
        ttk.Button(actions, text="Очистить вывод", command=self._clear_conference_output, style="Secondary.TButton").grid(row=0, column=4, sticky="ew", padx=(6, 0))

        summary = ttk.LabelFrame(frame, text="Статус", padding=12, style="Card.TLabelframe")
        summary.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        summary.columnconfigure(0, weight=1)
        ttk.Label(summary, textvariable=self.conference_summary_var, style="Body.TLabel", justify="left").grid(row=0, column=0, sticky="ew")
        ttk.Label(summary, textvariable=self.conference_progress_var, style="Muted.TLabel", justify="left").grid(row=1, column=0, sticky="ew", pady=(6, 0))

        output = ttk.LabelFrame(frame, text="JSON-результат", padding=12, style="Card.TLabelframe")
        output.grid(row=2, column=0, sticky="nsew")
        output.columnconfigure(0, weight=1)
        output.rowconfigure(0, weight=1)
        self.conference_output_text = tk.Text(
            output,
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
        self.conference_output_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(output, orient="vertical", command=self.conference_output_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.conference_output_text.configure(yscrollcommand=scrollbar.set)
        return frame

    def _build_adbeam_tab(self, parent):
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        source = ttk.LabelFrame(frame, text="AdBeam: аудит сайтов из Excel", padding=12, style="Card.TLabelframe")
        source.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        source.columnconfigure(1, weight=1)

        ttk.Label(source, text="Входной Excel", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(source, textvariable=self.adbeam_input_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(source, text="Выбрать файл", command=self._choose_adbeam_input_file, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(source, text="Итоговый Excel", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(source, textvariable=self.adbeam_output_path_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(source, text="Выбрать", command=self._choose_adbeam_output_file, style="Secondary.TButton").grid(row=1, column=2, pady=4)

        actions = ttk.Frame(source)
        actions.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        for index in range(4):
            actions.columnconfigure(index, weight=1)
        ttk.Button(
            actions,
            text="Проверить Excel",
            command=lambda: self._run_in_thread(self._action_adbeam_analyze, task_name="AdBeam анализ Excel"),
            style="Secondary.TButton",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(
            actions,
            text="Начать аудит и сохранить",
            command=lambda: self._run_in_thread(self._action_adbeam_audit, task_name="AdBeam аудит сайтов"),
            style="Primary.TButton",
        ).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(actions, text="Открыть итоговый Excel", command=self._open_adbeam_output_file, style="Secondary.TButton").grid(row=0, column=2, sticky="ew", padx=6)
        ttk.Button(actions, text="Очистить вывод", command=self._clear_adbeam_output, style="Secondary.TButton").grid(row=0, column=3, sticky="ew", padx=(6, 0))

        summary = ttk.LabelFrame(frame, text="Статус", padding=12, style="Card.TLabelframe")
        summary.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        summary.columnconfigure(0, weight=1)
        ttk.Label(summary, textvariable=self.adbeam_summary_var, style="Body.TLabel", justify="left").grid(row=0, column=0, sticky="ew")
        ttk.Label(summary, textvariable=self.adbeam_progress_var, style="Muted.TLabel", justify="left").grid(row=1, column=0, sticky="ew", pady=(6, 0))

        output = ttk.LabelFrame(frame, text="JSON-результат", padding=12, style="Card.TLabelframe")
        output.grid(row=2, column=0, sticky="nsew")
        output.columnconfigure(0, weight=1)
        output.rowconfigure(0, weight=1)
        self.adbeam_output_text = tk.Text(
            output,
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
        self.adbeam_output_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(output, orient="vertical", command=self.adbeam_output_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.adbeam_output_text.configure(yscrollcommand=scrollbar.set)
        return frame
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

        ttk.Checkbutton(
            frame,
            text="Исключать sellers из реестра/истории",
            variable=self.sample_use_known_sellers_var,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        ttk.Button(
            frame,
            text="Создать research sample",
            command=lambda: self._run_in_thread(self._action_sample, task_name="Создание sample"),
            style="Primary.TButton",
        ).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        ttk.Label(frame, textvariable=self.sample_summary_var, style="Body.TLabel", justify="left", wraplength=520).grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(frame, textvariable=self.sample_filter_summary_var, style="Muted.TLabel", justify="left", wraplength=520).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
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
        frame = ttk.LabelFrame(parent, text="Склейка с Контур.Поиск клиентов", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Файл batch_results для склейки", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.merge_batch_input_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(frame, text="Выбрать", command=self._choose_merge_batch_input, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(frame, text="Выгрузка Контур", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.compass_input_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(frame, text="Выбрать", command=self._choose_compass_input, style="Secondary.TButton").grid(row=1, column=2, pady=4)

        ttk.Label(frame, text="Итоговый enriched Excel", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.enriched_output_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(frame, text="Выбрать", command=self._choose_enriched_output, style="Secondary.TButton").grid(row=2, column=2, pady=4)

        ttk.Button(
            frame,
            text="Склеить с Контур",
            command=lambda: self._run_in_thread(self._action_merge_compass, task_name="Склейка с Контур"),
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

    def _build_excel_merge_tab(self, parent):
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)

        source = ttk.LabelFrame(frame, text="Источник", padding=12, style="Card.TLabelframe")
        source.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        source.columnconfigure(1, weight=1)
        ttk.Label(source, text="ZIP-архив или папка", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(source, textvariable=self.excel_merge_source_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(source, text="Выбрать ZIP", command=self._choose_excel_merge_zip, style="Secondary.TButton").grid(row=0, column=2, pady=4, padx=(0, 6))
        ttk.Button(source, text="Выбрать папку", command=self._choose_excel_merge_dir, style="Secondary.TButton").grid(row=0, column=3, pady=4)
        ttk.Label(
            source,
            text="Выбери ZIP-архив или папку с Excel-файлами. Каждый найденный файл станет отдельной вкладкой 1, 2, 3 и так далее.",
            style="Muted.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))

        outputs = ttk.LabelFrame(frame, text="Итоговые файлы", padding=12, style="Card.TLabelframe")
        outputs.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        outputs.columnconfigure(1, weight=1)
        ttk.Label(outputs, text="Итоговый Excel", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(outputs, textvariable=self.excel_merge_output_xlsx_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(outputs, text="Выбрать", command=self._choose_excel_merge_output_xlsx, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(outputs, text="Итоговый архив", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(outputs, textvariable=self.excel_merge_output_zip_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(outputs, text="Выбрать", command=self._choose_excel_merge_output_zip, style="Secondary.TButton").grid(row=1, column=2, pady=4)

        ttk.Checkbutton(
            outputs,
            text="Упаковать итоговый Excel в ZIP",
            variable=self.excel_merge_create_zip_var,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        for index in range(4):
            actions.columnconfigure(index, weight=1)
        ttk.Button(
            actions,
            text="Объединить Excel",
            command=lambda: self._run_in_thread(self._action_excel_merge, task_name="Объединение Excel"),
            style="Primary.TButton",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(actions, text="Открыть итоговый Excel", command=self._open_excel_merge_output_xlsx, style="Secondary.TButton").grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(actions, text="Открыть архив", command=self._open_excel_merge_output_zip, style="Secondary.TButton").grid(row=0, column=2, sticky="ew", padx=6)
        ttk.Button(actions, text="Очистить статус", command=self._clear_excel_merge_status, style="Secondary.TButton").grid(row=0, column=3, sticky="ew", padx=(6, 0))

        summary = ttk.LabelFrame(frame, text="Статус", padding=12, style="Card.TLabelframe")
        summary.grid(row=3, column=0, sticky="ew")
        summary.columnconfigure(0, weight=1)
        ttk.Label(
            summary,
            textvariable=self.excel_merge_summary_var,
            style="Body.TLabel",
            justify="left",
            wraplength=900,
        ).grid(row=0, column=0, sticky="ew")
        return frame

    def _build_registry_tab(self, parent):
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1, minsize=220)

        settings = ttk.LabelFrame(frame, text="Реестр ИНН", padding=12, style="Card.TLabelframe")
        settings.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        settings.columnconfigure(1, weight=1)

        ttk.Label(settings, text="Основной реестр", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(settings, textvariable=self.registry_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(settings, text="Выбрать", command=self._choose_registry_path, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(settings, text="История просмотренных sellers", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(settings, textvariable=self.seller_history_path_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(settings, text="Выбрать", command=self._choose_seller_history_path, style="Secondary.TButton").grid(row=1, column=2, pady=4)

        ttk.Label(settings, text="Новые ИНН для Контур", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(settings, textvariable=self.registry_new_inn_output_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(settings, text="Выбрать", command=self._choose_registry_new_inn_output, style="Secondary.TButton").grid(row=2, column=2, pady=4)

        ttk.Label(
            settings,
            text="Реестр ИНН хранит найденные реквизиты, а история sellers хранит уже просмотренных продавцов. Оба файла используются для пропуска повторов.",
            style="Muted.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        merge_registry = ttk.LabelFrame(frame, text="Объединение 2 реестров", padding=12, style="Card.TLabelframe")
        merge_registry.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        merge_registry.columnconfigure(1, weight=1)

        ttk.Label(merge_registry, text="Главный реестр", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(merge_registry, textvariable=self.registry_merge_primary_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(merge_registry, text="Выбрать", command=self._choose_registry_merge_primary, style="Secondary.TButton").grid(row=0, column=2, pady=4)

        ttk.Label(merge_registry, text="Второй реестр", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(merge_registry, textvariable=self.registry_merge_secondary_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(merge_registry, text="Выбрать", command=self._choose_registry_merge_secondary, style="Secondary.TButton").grid(row=1, column=2, pady=4)

        ttk.Label(merge_registry, text="Итоговый реестр", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(merge_registry, textvariable=self.registry_merge_output_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(merge_registry, text="Выбрать", command=self._choose_registry_merge_output, style="Secondary.TButton").grid(row=2, column=2, pady=4)

        ttk.Label(
            merge_registry,
            text="Если ИНН есть в обоих файлах, остаётся строка из первого реестра. Второй файл добавляет только новые ИНН.",
            style="Muted.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        merge_actions = ttk.Frame(merge_registry)
        merge_actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        merge_actions.columnconfigure(0, weight=1)
        merge_actions.columnconfigure(1, weight=1)
        ttk.Button(
            merge_actions,
            text="Объединить реестры",
            command=lambda: self._run_in_thread(self._action_merge_registries, task_name="Объединение реестров ИНН"),
            style="Primary.TButton",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(
            merge_actions,
            text="Открыть итоговый реестр",
            command=self._open_registry_merge_output,
            style="Secondary.TButton",
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        files = ttk.LabelFrame(frame, text="Batch-файлы для импорта", padding=12, style="Card.TLabelframe")
        files.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        files.columnconfigure(0, weight=1)
        files.rowconfigure(0, weight=1, minsize=180)

        list_frame = ttk.Frame(files)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.registry_batch_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            height=6,
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
        self.registry_batch_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.registry_batch_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        xscrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=self.registry_batch_listbox.xview)
        xscrollbar.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.registry_batch_listbox.configure(yscrollcommand=scrollbar.set, xscrollcommand=xscrollbar.set)

        file_buttons = ttk.Frame(files)
        file_buttons.grid(row=0, column=1, sticky="n", padx=(12, 0))
        ttk.Button(file_buttons, text="Добавить файлы", command=self._add_registry_batch_files, style="Secondary.TButton").grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(file_buttons, text="Добавить папку", command=self._add_registry_batch_dir, style="Secondary.TButton").grid(row=1, column=0, sticky="ew", pady=8)
        ttk.Button(file_buttons, text="Очистить список", command=self._clear_registry_batch_files, style="Secondary.TButton").grid(row=2, column=0, sticky="ew", pady=8)

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        for index in range(6):
            actions.columnconfigure(index, weight=1)
        ttk.Button(
            actions,
            text="Обновить реестр и историю",
            command=lambda: self._run_in_thread(self._action_update_registry, task_name="Обновление реестра ИНН и истории sellers"),
            style="Primary.TButton",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(
            actions,
            text="Импорт sellers из реестра",
            command=self._start_import_history_from_registry,
            style="Secondary.TButton",
        ).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(actions, text="Открыть реестр", command=self._open_registry_path, style="Secondary.TButton").grid(row=0, column=2, sticky="ew", padx=6)
        ttk.Button(actions, text="Открыть историю sellers", command=self._open_seller_history_path, style="Secondary.TButton").grid(row=0, column=3, sticky="ew", padx=6)
        ttk.Button(actions, text="Открыть новые ИНН", command=self._open_registry_new_inn, style="Secondary.TButton").grid(row=0, column=4, sticky="ew", padx=6)
        ttk.Button(actions, text="Взять текущий batch", command=self._use_current_batch_for_registry, style="Secondary.TButton").grid(row=0, column=5, sticky="ew", padx=(6, 0))

        summary = ttk.LabelFrame(frame, text="Сводка", padding=12, style="Card.TLabelframe")
        summary.grid(row=4, column=0, sticky="nsew")
        summary.columnconfigure(0, weight=1)
        summary.rowconfigure(0, weight=1)
        self.registry_summary_text = tk.Text(
            summary,
            wrap="word",
            height=8,
            bg="#ffffff",
            fg="#1f2a37",
            insertbackground="#1f2a37",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground="#dbe4f0",
            relief="solid",
            font=("Consolas", 10),
        )
        self.registry_summary_text.grid(row=0, column=0, sticky="nsew")
        summary_scrollbar = ttk.Scrollbar(summary, orient="vertical", command=self.registry_summary_text.yview)
        summary_scrollbar.grid(row=0, column=1, sticky="ns")
        self.registry_summary_text.configure(yscrollcommand=summary_scrollbar.set)
        self.registry_summary_text.configure(state="disabled")
        self._set_registry_summary_ui(self.registry_summary_var.get())
        return frame

    def _build_status_log_block(self, parent):
        frame = ttk.LabelFrame(parent, text="6. Статус, прогресс и лог", padding=12, style="Card.TLabelframe")
        frame.columnconfigure(0, weight=1)

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
            self.conference_input_path_var,
            self.conference_output_path_var,
            self.conference_autosearch_output_path_var,
            self.conference_autosearch_limit_var,
            self.conference_autosearch_delay_var,
            self.adbeam_input_path_var,
            self.adbeam_output_path_var,
            self.artifacts_dir_var,
            self.profile_dir_var,
            self.limit_var,
            self.row_var,
            self.batch_count_var,
            self.batch_output_var,
            self.merge_batch_input_var,
            self.compass_input_var,
            self.enriched_output_var,
            self.registry_path_var,
            self.seller_history_path_var,
            self.registry_new_inn_output_var,
            self.registry_merge_primary_var,
            self.registry_merge_secondary_var,
            self.registry_merge_output_var,
            self.excel_merge_source_var,
            self.excel_merge_output_xlsx_var,
            self.excel_merge_output_zip_var,
            self.sheet_mode_var,
        ]
        for var in self._tracked_vars:
            var.trace_add("write", lambda *_: self._save_settings())
        self.headful_var.trace_add("write", lambda *_: self._save_settings())
        self.excel_merge_create_zip_var.trace_add("write", lambda *_: self._save_settings())
        self.conference_autosearch_only_p1_var.trace_add("write", lambda *_: self._save_settings())
        self.sample_use_known_sellers_var.trace_add("write", lambda *_: self._save_settings())


    def _open_registry_path(self) -> None:
        path = Path(self.registry_path_var.get().strip() or "output/inn_registry.xlsx")
        if not path.exists():
            raise FileNotFoundError(f"Не найден реестр ИНН: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_seller_history_path(self) -> None:
        path = Path(self.seller_history_path_var.get().strip() or "output/seller_history.xlsx")
        if not path.exists():
            raise FileNotFoundError(f"Не найдена история sellers: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_registry_merge_output(self) -> None:
        path = Path(self.registry_merge_output_var.get().strip() or "output/inn_registry_merged.xlsx")
        if not path.exists():
            raise FileNotFoundError(f"Не найден итоговый реестр: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_registry_new_inn(self) -> None:
        path = Path(self.registry_new_inn_output_var.get().strip() or "output/new_inn_for_kontur.xlsx")
        if not path.exists():
            raise FileNotFoundError(f"Не найден файл новых ИНН: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_excel_merge_output_xlsx(self) -> None:
        path = Path(self.excel_merge_output_xlsx_var.get().strip() or "output/merged_tabs.xlsx")
        if not path.exists():
            raise FileNotFoundError(f"Не найден итоговый Excel: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_excel_merge_output_zip(self) -> None:
        path = Path(self.excel_merge_output_zip_var.get().strip() or "output/merged_tabs.zip")
        if not path.exists():
            raise FileNotFoundError(f"Не найден итоговый архив: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _clear_excel_merge_status(self) -> None:
        self.excel_merge_summary_var.set("Объединение Excel ещё не запускалось")
        self._save_settings()

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
            "conference_input_path": self.conference_input_path_var,
            "conference_output_path": self.conference_output_path_var,
            "conference_autosearch_output_path": self.conference_autosearch_output_path_var,
            "conference_autosearch_limit": self.conference_autosearch_limit_var,
            "conference_autosearch_delay": self.conference_autosearch_delay_var,
            "adbeam_input_path": self.adbeam_input_path_var,
            "adbeam_output_path": self.adbeam_output_path_var,
            "artifacts_dir": self.artifacts_dir_var,
            "profile_dir": self.profile_dir_var,
            "limit": self.limit_var,
            "row": self.row_var,
            "batch_count": self.batch_count_var,
            "batch_output": self.batch_output_var,
            "merge_batch_input": self.merge_batch_input_var,
            "compass_input": self.compass_input_var,
            "enriched_output": self.enriched_output_var,
            "registry_path": self.registry_path_var,
            "seller_history_path": self.seller_history_path_var,
            "registry_new_inn_output": self.registry_new_inn_output_var,
            "registry_merge_primary": self.registry_merge_primary_var,
            "registry_merge_secondary": self.registry_merge_secondary_var,
            "registry_merge_output": self.registry_merge_output_var,
            "excel_merge_source": self.excel_merge_source_var,
            "excel_merge_output_xlsx": self.excel_merge_output_xlsx_var,
            "excel_merge_output_zip": self.excel_merge_output_zip_var,
            "sheet_mode": self.sheet_mode_var,
        }.items():
            value = settings.get(key)
            if isinstance(value, str):
                var.set(value)
        if "headful" in settings:
            self.headful_var.set(bool(settings["headful"]))
        if "excel_merge_create_zip" in settings:
            self.excel_merge_create_zip_var.set(bool(settings["excel_merge_create_zip"]))
        if "conference_autosearch_only_p1" in settings:
            self.conference_autosearch_only_p1_var.set(bool(settings["conference_autosearch_only_p1"]))
        if "sample_use_known_sellers" in settings:
            self.sample_use_known_sellers_var.set(bool(settings["sample_use_known_sellers"]))
        loaded = settings.get("selected_sheets")
        if isinstance(loaded, list):
            self._settings_loaded_selected_sheets = [str(x) for x in loaded]
        summary = settings.get("selection_summary")
        if isinstance(summary, str) and summary:
            self.selection_summary_var.set(summary)
        sample_summary = settings.get("sample_summary")
        if isinstance(sample_summary, str) and sample_summary:
            self.sample_summary_var.set(f"{sample_summary.splitlines()[0]} | старт batch: 2")
        sample_filter_summary = settings.get("sample_filter_summary")
        if isinstance(sample_filter_summary, str) and sample_filter_summary:
            filter_state = "включены" if self.sample_use_known_sellers_var.get() else "отключены"
            self.sample_filter_summary_var.set(f"Фильтры повторов: {filter_state}")
        registry_summary = settings.get("registry_summary")
        if isinstance(registry_summary, str) and registry_summary:
            self._set_registry_summary(registry_summary)
        excel_merge_summary = settings.get("excel_merge_summary")
        if isinstance(excel_merge_summary, str) and excel_merge_summary:
            self.excel_merge_summary_var.set(excel_merge_summary)
        adbeam_summary = settings.get("adbeam_summary")
        if isinstance(adbeam_summary, str) and adbeam_summary:
            self.adbeam_summary_var.set(adbeam_summary)
        conference_summary = settings.get("conference_summary")
        if isinstance(conference_summary, str) and conference_summary:
            self.conference_summary_var.set(conference_summary)
        registry_batch_files = settings.get("registry_batch_files")
        if isinstance(registry_batch_files, list):
            self._registry_batch_files = [str(path) for path in registry_batch_files]
            self._refresh_registry_batch_listbox()

    def _save_settings(self) -> None:
        try:
            data = {
                "input_path": self.input_path_var.get().strip(),
                "sample_output": self.sample_output_var.get().strip(),
                "conference_input_path": self.conference_input_path_var.get().strip(),
                "conference_output_path": self.conference_output_path_var.get().strip(),
                "conference_autosearch_output_path": self.conference_autosearch_output_path_var.get().strip(),
                "conference_autosearch_limit": self.conference_autosearch_limit_var.get().strip(),
                "conference_autosearch_delay": self.conference_autosearch_delay_var.get().strip(),
                "conference_autosearch_only_p1": bool(self.conference_autosearch_only_p1_var.get()),
                "conference_summary": self.conference_summary_var.get(),
                "sample_use_known_sellers": bool(self.sample_use_known_sellers_var.get()),
                "adbeam_input_path": self.adbeam_input_path_var.get().strip(),
                "adbeam_output_path": self.adbeam_output_path_var.get().strip(),
                "adbeam_summary": self.adbeam_summary_var.get(),
                "artifacts_dir": self.artifacts_dir_var.get().strip(),
                "profile_dir": self.profile_dir_var.get().strip(),
                "limit": self.limit_var.get().strip(),
                "row": self.row_var.get().strip(),
                "batch_count": self.batch_count_var.get().strip(),
                "batch_output": self.batch_output_var.get().strip(),
                "merge_batch_input": self.merge_batch_input_var.get().strip(),
                "compass_input": self.compass_input_var.get().strip(),
                "enriched_output": self.enriched_output_var.get().strip(),
                "registry_path": self.registry_path_var.get().strip(),
                "seller_history_path": self.seller_history_path_var.get().strip(),
                "registry_new_inn_output": self.registry_new_inn_output_var.get().strip(),
                "registry_merge_primary": self.registry_merge_primary_var.get().strip(),
                "registry_merge_secondary": self.registry_merge_secondary_var.get().strip(),
                "registry_merge_output": self.registry_merge_output_var.get().strip(),
                "excel_merge_source": self.excel_merge_source_var.get().strip(),
                "excel_merge_output_xlsx": self.excel_merge_output_xlsx_var.get().strip(),
                "excel_merge_output_zip": self.excel_merge_output_zip_var.get().strip(),
                "excel_merge_create_zip": bool(self.excel_merge_create_zip_var.get()),
                "excel_merge_summary": self.excel_merge_summary_var.get(),
                "registry_batch_files": list(self._registry_batch_files),
                "registry_summary": self.registry_summary_var.get(),
                "sheet_mode": self.sheet_mode_var.get(),
                "selected_sheets": self._get_selected_sheet_names_from_ui() if hasattr(self, "sheet_listbox") else [],
                "selection_summary": self.selection_summary_var.get(),
                "sample_summary": self.sample_summary_var.get(),
                "sample_filter_summary": self.sample_filter_summary_var.get(),
                "headful": bool(self.headful_var.get()),
            }
            SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _set_registry_summary(self, text: str) -> None:
        self.registry_summary_var.set(text)
        self.root.after(0, lambda text=text: self._set_registry_summary_ui(text))

    def _set_registry_summary_ui(self, text: str) -> None:
        if not hasattr(self, "registry_summary_text"):
            return
        self.registry_summary_text.configure(state="normal")
        self.registry_summary_text.delete("1.0", "end")
        if text:
            self.registry_summary_text.insert("1.0", text)
        self.registry_summary_text.configure(state="disabled")

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

    def _choose_adbeam_input_file(self) -> None:
        path = filedialog.askopenfilename(title="Выбери Excel-файл для AdBeam", filetypes=[("Excel", "*.xlsx")])
        if path:
            input_path = Path(path)
            self.adbeam_input_path_var.set(str(input_path))
            if not self.adbeam_output_path_var.get().strip():
                self.adbeam_output_path_var.set(str(build_output_path(input_path)))
            self.adbeam_summary_var.set("Файл выбран. Можно проверить Excel или запустить аудит сайтов.")
            self.adbeam_progress_var.set("Ожидание запуска")
            self._save_settings()

    def _choose_conference_input_file(self) -> None:
        path = filedialog.askopenfilename(title="Выбери ICP-1 Excel-файл", filetypes=[("Excel", "*.xlsx")])
        if path:
            input_path = Path(path)
            self.conference_input_path_var.set(str(input_path))
            if not self.conference_output_path_var.get().strip():
                self.conference_output_path_var.set(str(build_icp_contacts_output_path(input_path)))
            if not self.conference_autosearch_output_path_var.get().strip():
                self.conference_autosearch_output_path_var.set(str(build_icp_autosearch_output_path(input_path)))
            self.conference_summary_var.set("Файл выбран. Можно проверить базу или создать очередь.")
            self.conference_progress_var.set("Ожидание запуска")
            self._save_settings()

    def _choose_conference_output_file(self) -> None:
        current = self.conference_output_path_var.get().strip()
        initialfile = Path(current).name if current else "icp1_contacts_queue.xlsx"
        path = filedialog.asksaveasfilename(
            title="Куда сохранить ICP-1 очередь",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=initialfile,
        )
        if path:
            self.conference_output_path_var.set(path)

    def _choose_conference_autosearch_output_file(self) -> None:
        current = self.conference_autosearch_output_path_var.get().strip()
        initialfile = Path(current).name if current else "icp1_autosearch.xlsx"
        path = filedialog.asksaveasfilename(
            title="Куда сохранить ICP-1 автопоиск",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=initialfile,
        )
        if path:
            self.conference_autosearch_output_path_var.set(path)

    def _choose_adbeam_output_file(self) -> None:
        current = self.adbeam_output_path_var.get().strip()
        initialfile = Path(current).name if current else "adbeam_audited.xlsx"
        path = filedialog.asksaveasfilename(
            title="Куда сохранить AdBeam audited Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=initialfile,
        )
        if path:
            self.adbeam_output_path_var.set(path)
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
        path = filedialog.askopenfilename(title="Выбери выгрузку Контур", filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
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

    def _choose_registry_path(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Выбери или создай реестр ИНН",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.registry_path_var.get()).name,
        )
        if path:
            self.registry_path_var.set(path)

    def _choose_seller_history_path(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Выбери или создай историю просмотренных sellers",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.seller_history_path_var.get()).name,
        )
        if path:
            self.seller_history_path_var.set(path)

    def _choose_registry_new_inn_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить новые ИНН для Контур",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.registry_new_inn_output_var.get()).name,
        )
        if path:
            self.registry_new_inn_output_var.set(path)

    def _choose_registry_merge_primary(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери главный реестр ИНН",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            self.registry_merge_primary_var.set(path)

    def _choose_registry_merge_secondary(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери второй реестр ИНН",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            self.registry_merge_secondary_var.set(path)

    def _choose_registry_merge_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить объединённый реестр",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.registry_merge_output_var.get()).name,
        )
        if path:
            self.registry_merge_output_var.set(path)

    def _choose_excel_merge_zip(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери ZIP-архив с Excel-файлами",
            filetypes=[("ZIP archive", "*.zip"), ("All files", "*.*")],
        )
        if path:
            self.excel_merge_source_var.set(path)

    def _choose_excel_merge_dir(self) -> None:
        path = filedialog.askdirectory(title="Выбери папку с Excel-файлами")
        if path:
            self.excel_merge_source_var.set(path)

    def _choose_excel_merge_output_xlsx(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить итоговый Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.excel_merge_output_xlsx_var.get()).name,
        )
        if path:
            self.excel_merge_output_xlsx_var.set(path)

    def _choose_excel_merge_output_zip(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Куда сохранить ZIP-архив",
            defaultextension=".zip",
            filetypes=[("ZIP archive", "*.zip")],
            initialfile=Path(self.excel_merge_output_zip_var.get()).name,
        )
        if path:
            self.excel_merge_output_zip_var.set(path)

    def _start_import_history_from_registry(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Выбери реестры для импорта sellers в историю",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if not paths:
            return

        registry_paths = [Path(path) for path in paths]
        self._run_in_thread(
            lambda registry_paths=registry_paths: self._action_import_history_from_registry(registry_paths),
            task_name="Импорт sellers из реестра",
        )

    def _add_registry_batch_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Выбери batch_results.xlsx",
            filetypes=[("batch_results.xlsx", "batch_results.xlsx"), ("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if paths:
            self._add_registry_batch_paths([Path(path) for path in paths])

    def _add_registry_batch_dir(self) -> None:
        path = filedialog.askdirectory(title="Папка с batch_results.xlsx")
        if not path:
            return
        batch_paths = discover_batch_results(Path(path))
        if not batch_paths:
            messagebox.showinfo("Реестр ИНН", "В папке не найдено batch_results.xlsx")
            return
        self._add_registry_batch_paths(batch_paths)

    def _use_current_batch_for_registry(self) -> None:
        path = Path(self.batch_output_var.get().strip())
        if not path.exists():
            raise FileNotFoundError(f"Не найден текущий batch_results.xlsx: {path}")
        self._add_registry_batch_paths([path])

    def _add_registry_batch_paths(self, paths: list[Path]) -> None:
        existing = set(self._registry_batch_files)
        added = 0
        skipped = 0
        for path in paths:
            if path.name.casefold() != "batch_results.xlsx":
                skipped += 1
                continue
            normalized = str(path)
            if normalized in existing:
                continue
            self._registry_batch_files.append(normalized)
            existing.add(normalized)
            added += 1
        self._refresh_registry_batch_listbox()
        suffix = f" Пропущено не-batch файлов: {skipped}." if skipped else ""
        self._set_registry_summary(f"Добавлено batch-файлов: {added}. Всего в списке: {len(self._registry_batch_files)}.{suffix}")
        self._save_settings()

    def _clear_registry_batch_files(self) -> None:
        self._registry_batch_files = []
        self._refresh_registry_batch_listbox()
        self._set_registry_summary("Список batch-файлов очищен")
        self._save_settings()

    def _refresh_registry_batch_listbox(self) -> None:
        if not hasattr(self, "registry_batch_listbox"):
            return
        self.registry_batch_listbox.delete(0, "end")
        for path in self._registry_batch_files:
            self.registry_batch_listbox.insert("end", path)

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

    def _open_adbeam_output_file(self) -> None:
        path = Path(self.adbeam_output_path_var.get().strip())
        if not path.exists():
            raise FileNotFoundError(f"Не найден итоговый AdBeam Excel: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_conference_output_file(self) -> None:
        path = Path(self.conference_output_path_var.get().strip())
        if not path.exists():
            raise FileNotFoundError(f"Не найден итоговый ICP-1 Excel: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _open_conference_autosearch_output_file(self) -> None:
        path = Path(self.conference_autosearch_output_path_var.get().strip())
        if not path.exists():
            raise FileNotFoundError(f"Не найден итоговый ICP-1 автопоиск Excel: {path}")
        os.startfile(path)  # type: ignore[attr-defined]

    def _clear_conference_output(self) -> None:
        self.conference_output_path_var.set("")
        self.conference_autosearch_output_path_var.set("")
        self.conference_summary_var.set("Конференция ICP-1 ещё не запускалась")
        self.conference_progress_var.set("Ожидание Excel-файла со списком компаний")
        if hasattr(self, "conference_output_text"):
            self.conference_output_text.delete("1.0", "end")

    def _clear_adbeam_output(self) -> None:
        self.adbeam_output_path_var.set("")
        self.adbeam_summary_var.set("AdBeam аудит ещё не запускался")
        self.adbeam_progress_var.set("Ожидание Excel-файла для аудита сайтов")
        if hasattr(self, "adbeam_output_text"):
            self.adbeam_output_text.delete("1.0", "end")
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
        self.root.title("WB INN Extractor — Research MVP")
        self.progress_text_var.set(label)
        self.progress_detail_var.set("Лог ниже содержит детали последней операции")
        self.progress_percent_var.set("100%" if success else self.progress_percent_var.get())
        if success and self.progress_value_var.get() < 100:
            self.progress_value_var.set(100)

    def _set_adbeam_progress_ui(self, title: str, detail: str, percent: float) -> None:
        self.adbeam_progress_var.set(f"{title}\n{detail}")
        self.progress_text_var.set(title)
        self.progress_detail_var.set(detail)
        self.progress_percent_var.set(f"{percent:.1f}%")
        self.progress_value_var.set(percent)

    def _set_conference_progress_ui(self, title: str, detail: str, percent: float) -> None:
        self.conference_progress_var.set(f"{title}\n{detail}")
        self.progress_text_var.set(title)
        self.progress_detail_var.set(detail)
        self.progress_percent_var.set(f"{percent:.1f}%")
        self.progress_value_var.set(percent)

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
        self.root.title(
            f"WB INN Extractor — {self._progress_done}/{self._progress_total} ({percent:.1f}%)"
        )

    def _action_adbeam_analyze(self) -> None:
        input_path = self._require_adbeam_input_path()
        summary = read_excel_summary(input_path)
        output_text = summary.model_dump_json(indent=2, exclude_none=True)
        self.root.after(0, lambda output_text=output_text: self._set_adbeam_output(output_text))
        self.root.after(0, lambda summary=summary: self.adbeam_summary_var.set(
            f"Лист: {summary.sheet_name} | строк: {summary.total_rows} | сайтов найдено: {summary.rows_with_websites}"
        ))
        self.root.after(0, lambda: self.adbeam_progress_var.set("Проверка Excel завершена"))
        self._append_log(
            f"AdBeam analyze: sheet={summary.sheet_name}, rows={summary.total_rows}, websites={summary.rows_with_websites}\n"
        )

    def _action_adbeam_audit(self) -> None:
        input_path = self._require_adbeam_input_path()
        raw_output_path = self.adbeam_output_path_var.get().strip()
        output_path = Path(raw_output_path) if raw_output_path else build_output_path(input_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        summary = run_excel_audit(input_path, progress_callback=self._on_adbeam_progress)
        saved_path = export_audit_to_excel(input_path, summary, output_file_path=output_path)
        attach_output_file_path(summary, saved_path)

        output_text = summary.model_dump_json(indent=2, exclude_none=True)
        counts = ", ".join(f"{key}: {value}" for key, value in sorted(summary.status_counts.items())) or "нет результатов"
        self.root.after(0, lambda output_text=output_text: self._set_adbeam_output(output_text))
        self.root.after(0, lambda saved_path=saved_path: self.adbeam_output_path_var.set(str(saved_path)))
        self.root.after(0, lambda summary=summary, counts=counts: self.adbeam_summary_var.set(
            f"Проверено сайтов: {summary.audited_rows}\nСтатусы: {counts}\nИтоговый файл: {summary.output_file_path}"
        ))
        self.root.after(0, lambda: self.adbeam_progress_var.set("Аудит завершён"))
        self._append_log(f"AdBeam audit: checked={summary.audited_rows}, output={saved_path}\n")
        self.root.after(0, lambda saved_path=saved_path: messagebox.showinfo("AdBeam аудит завершён", f"Итоговый Excel сохранён:\n{saved_path}"))

    def _on_adbeam_progress(self, current: int, total: int, website: str) -> None:
        safe_total = max(total, 1)
        percent = min(current / safe_total * 100.0, 100.0)
        title = f"AdBeam аудит: {current}/{total} сайтов"
        detail = f"Текущий сайт: {website}"
        self.root.after(0, lambda title=title, detail=detail, percent=percent: self._set_adbeam_progress_ui(title, detail, percent))

    def _set_adbeam_output(self, text: str) -> None:
        self.adbeam_output_text.delete("1.0", "end")
        self.adbeam_output_text.insert("1.0", text)

    def _action_conference_analyze(self) -> None:
        input_path = self._require_conference_input_path()
        summary = read_icp_contacts_summary(input_path)
        output_text = summary.model_dump_json(indent=2, exclude_none=True)
        self.root.after(0, lambda output_text=output_text: self._set_conference_output(output_text))
        self.root.after(0, lambda summary=summary: self.conference_summary_var.set(
            f"Лист: {summary.sheet_name} | компаний: {summary.total_rows} | "
            f"с ИНН: {summary.rows_with_inn} | с сайтами: {summary.rows_with_website} | "
            f"валидация: {summary.validation_rows} | исключения: {summary.excluded_rows}"
        ))
        self.root.after(0, lambda: self.conference_progress_var.set("Проверка базы завершена"))
        self._append_log(
            f"ICP-1 analyze: sheet={summary.sheet_name}, rows={summary.total_rows}, "
            f"validation={summary.validation_rows}, excluded={summary.excluded_rows}\n"
        )

    def _action_conference_export(self) -> None:
        input_path = self._require_conference_input_path()
        raw_output_path = self.conference_output_path_var.get().strip()
        output_path = Path(raw_output_path) if raw_output_path else build_icp_contacts_output_path(input_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        summary = export_icp_contacts_queue(
            input_path,
            output_file_path=output_path,
            progress_callback=self._on_conference_progress,
        )
        output_text = summary.model_dump_json(indent=2, exclude_none=True)
        self.root.after(0, lambda output_text=output_text: self._set_conference_output(output_text))
        self.root.after(0, lambda output_path=output_path: self.conference_output_path_var.set(str(output_path)))
        self.root.after(0, lambda summary=summary: self.conference_summary_var.set(
            f"Создана очередь: {summary.queue_rows} компаний\n"
            f"Валидация: {summary.validation_rows} | исключения: {summary.excluded_rows}\n"
            f"Итоговый файл: {summary.output_file_path}"
        ))
        self.root.after(0, lambda: self.conference_progress_var.set("Очередь ICP-1 создана"))
        self._append_log(f"ICP-1 queue: rows={summary.queue_rows}, output={output_path}\n")
        self.root.after(0, lambda output_path=output_path: messagebox.showinfo("ICP-1 очередь создана", f"Итоговый Excel сохранён:\n{output_path}"))

    def _action_conference_autosearch(self) -> None:
        input_path = self._require_conference_input_path()
        raw_output_path = self.conference_autosearch_output_path_var.get().strip()
        output_path = Path(raw_output_path) if raw_output_path else build_icp_autosearch_output_path(input_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        limit = self._parse_positive_int(self.conference_autosearch_limit_var.get().strip() or "20", "Лимит строк")
        delay_seconds = self._parse_non_negative_float(
            self.conference_autosearch_delay_var.get().strip() or "1.5",
            "Пауза между запросами",
        )

        summary = run_icp_autosearch(
            source_file_path=input_path,
            output_file_path=output_path,
            limit=limit,
            delay_seconds=delay_seconds,
            only_p1=False,
            progress_callback=self._on_conference_progress,
        )
        output_text = summary.model_dump_json(indent=2, exclude_none=True)
        self.root.after(0, lambda output_text=output_text: self._set_conference_output(output_text))
        self.root.after(0, lambda output_path=output_path: self.conference_autosearch_output_path_var.set(str(output_path)))
        self.root.after(0, lambda summary=summary: self.conference_summary_var.set(
            f"Автопоиск: обработано {summary.processed_rows}/{summary.requested_rows}\n"
            f"Сайты: {summary.found_websites} | email: {summary.found_emails} | телефоны: {summary.found_phones} | ИНН: {summary.found_inns}\n"
            f"Высокая уверенность: {summary.high_confidence}\n"
            f"Итоговый файл: {summary.output_file_path}"
        ))
        self.root.after(0, lambda: self.conference_progress_var.set("Автопоиск ICP-1 завершён"))
        self._append_log(
            f"ICP-1 autosearch: processed={summary.processed_rows}, websites={summary.found_websites}, "
            f"emails={summary.found_emails}, phones={summary.found_phones}, output={output_path}\n"
        )
        self.root.after(0, lambda output_path=output_path: messagebox.showinfo("ICP-1 автопоиск завершён", f"Итоговый Excel сохранён:\n{output_path}"))

    def _on_conference_progress(self, current: int, total: int, brand: str) -> None:
        safe_total = max(total, 1)
        percent = min(current / safe_total * 100.0, 100.0)
        title = f"ICP-1 очередь: {current}/{total} компаний"
        detail = f"Текущая компания: {brand}"
        self.root.after(0, lambda title=title, detail=detail, percent=percent: self._set_conference_progress_ui(title, detail, percent))

    def _set_conference_output(self, text: str) -> None:
        self.conference_output_text.delete("1.0", "end")
        self.conference_output_text.insert("1.0", text)

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
        registry_path = Path(self.registry_path_var.get().strip() or "output/inn_registry.xlsx")
        seller_history_path = Path(self.seller_history_path_var.get().strip() or "output/seller_history.xlsx")
        use_known_sellers = bool(self.sample_use_known_sellers_var.get())
        registry_exists = registry_path.exists()
        seller_history_exists = seller_history_path.exists()

        known_seller_sources = {}
        if use_known_sellers:
            known_seller_sources = load_known_seller_sources(
                registry_path=registry_path if registry_exists else None,
                seller_history_path=seller_history_path if seller_history_exists else None,
            )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        unfiltered_rows = extract_research_rows(
            input_path,
            selected_sheets=selected_sheets,
        )
        filtered_rows = [
            row
            for row in unfiltered_rows
            if not known_seller_sources or (row.seller_name_raw or "").strip().casefold() not in known_seller_sources
        ]
        rows = filtered_rows[:limit] if limit is not None else filtered_rows
        if not rows:
            raise ValueError("Не удалось собрать ни одной строки для research sample")

        save_research_sample(output_path, rows)
        rows_by_sheet = summarize_research_rows_by_sheet(rows)
        excluded_by_filters = len(unfiltered_rows) - len(filtered_rows)
        registry_resolved = self._format_path_status(registry_path)
        seller_history_resolved = self._format_path_status(seller_history_path)

        sheet_scope_label = self._build_sheet_scope_label(selected_sheets)
        self._append_log(f"Создан файл: {output_path}\n")
        self._append_log(f"Листы для sample: {sheet_scope_label}\n")
        self._append_log(f"Sample: рабочая папка приложения: {Path.cwd()}\n")
        self._append_log(f"Sample: реестр ИНН: {registry_resolved}\n")
        self._append_log(f"Sample: история sellers: {seller_history_resolved}\n")
        if not use_known_sellers:
            self._append_log("Sample: фильтр по реестру/истории отключён\n")
        elif known_seller_sources:
            self._append_log(
                f"Sample: ключей продавцов для пропуска: {len(known_seller_sources)} | исключено из текущей выборки: {excluded_by_filters}\n"
            )
        else:
            self._append_log("Sample: фильтр включён, но ключей для пропуска не найдено\n")
        if limit is None:
            self._append_log("Лимит строк не задан: в sample сохранены все уникальные продавцы по имени seller\n")
        else:
            self._append_log(f"Лимит строк задан: {limit}\n")
        self._append_log(f"Строк без фильтра реестра/истории: {len(unfiltered_rows)}\n")
        self._append_log(f"Строк после фильтра реестра/истории: {len(filtered_rows)}\n")
        self._append_log(f"Строк (уникальных по seller): {len(rows)}\n")
        self._append_log("Распределение строк по листам в sample:\n")
        for sheet_name, count in rows_by_sheet.items():
            self._append_log(f"  - {sheet_name}: {count}\n")

        self.root.after(0, lambda: self.row_var.set("2"))
        sample_summary = (
            f"Sample: {len(rows)} | без фильтра: {len(unfiltered_rows)} | "
            f"исключено: {excluded_by_filters} | листов: {len(rows_by_sheet)}"
        )
        if limit is not None:
            sample_summary += f" | лимит: {limit}"
        sample_summary += " | старт batch: 2"
        filter_summary = (
            f"Фильтры повторов: {'включены' if use_known_sellers else 'отключены'} | "
            f"ключей: {len(known_seller_sources)} | исключено: {excluded_by_filters}"
        )
        if use_known_sellers and len(unfiltered_rows) > 0 and len(filtered_rows) < len(unfiltered_rows) * 0.25:
            filter_summary += "\nВнимание: фильтры отсекли больше 75% выборки."
        self.root.after(0, lambda sample_summary=sample_summary: self.sample_summary_var.set(sample_summary))
        self.root.after(0, lambda filter_summary=filter_summary: self.sample_filter_summary_var.set(filter_summary))
        self.root.after(0, lambda: self.selection_summary_var.set(
            f"Последний sample: {len(rows)} строк | без фильтра: {len(unfiltered_rows)} | исключено: {excluded_by_filters} | стартовая строка = 2"
        ))

    @staticmethod
    def _format_path_status(path: Path) -> str:
        resolved = path.resolve() if path.exists() else path.absolute()
        status = "есть" if path.exists() else "нет"
        return f"{resolved} ({status})"

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
        registry_path = Path(self.registry_path_var.get().strip() or "output/inn_registry.xlsx")
        seller_history_path = Path(self.seller_history_path_var.get().strip() or "output/seller_history.xlsx")
        use_known_sellers = bool(self.sample_use_known_sellers_var.get())

        if artifacts_dir.resolve() == profile_dir.resolve():
            raise ValueError("Папка артефактов и папка профиля WB должны быть разными")

        artifacts_dir.mkdir(parents=True, exist_ok=True)
        profile_dir.mkdir(parents=True, exist_ok=True)
        batch_output.parent.mkdir(parents=True, exist_ok=True)

        known_seller_sources = {}
        if use_known_sellers:
            known_seller_sources = load_known_seller_sources(
                registry_path=registry_path if registry_path.exists() else None,
                seller_history_path=seller_history_path if seller_history_path.exists() else None,
            )

        research_rows = read_research_rows_range(sample_path, start_row=start_row, limit=batch_count)
        total = len(research_rows)
        self._append_log(f"Batch: старт диапазона строк {start_row}-{start_row + max(total - 1, 0)} | sample={sample_path}\n")
        self._append_log(f"Batch: рабочая папка приложения: {Path.cwd()}\n")
        self._append_log(f"Batch: реестр ИНН: {self._format_path_status(registry_path)}\n")
        self._append_log(f"Batch: история sellers: {self._format_path_status(seller_history_path)}\n")
        if use_known_sellers:
            self._append_log(f"Batch: фильтр по реестру/истории включён | ключей для пропуска: {len(known_seller_sources)}\n")
        else:
            self._append_log("Batch: фильтр по реестру/истории отключён\n")
        output_rows = []
        skipped_known_sellers = 0
        skipped_viewed_sellers = 0
        self._set_batch_progress(done=0, total=max(total, 1))

        with ExitStack() as exit_stack:
            inspector: BatchInspector | None = None
            for offset, research_row in enumerate(research_rows, start=0):
                row_number = start_row + offset
                self._set_batch_progress(done=offset, total=total, row_number=row_number, seller=research_row.seller_name_raw)
                seller_key = (research_row.seller_name_raw or "").strip().casefold()
                known_payload = known_seller_sources.get(seller_key) if seller_key else None
                if known_payload is not None:
                    if known_payload["source"] == "inn_registry":
                        skipped_row = build_known_seller_batch_row(
                            row_number=row_number,
                            research_row=research_row,
                            registry_row=known_payload["row"],
                            registry_path=known_payload["path"],
                        )
                        skipped_known_sellers += 1
                    else:
                        skipped_row = build_viewed_seller_batch_row(
                            row_number=row_number,
                            research_row=research_row,
                            history_row=known_payload["row"],
                            history_path=known_payload["path"],
                        )
                        skipped_viewed_sellers += 1
                    output_rows.append(skipped_row)
                    self._append_log(
                        f"Batch: строка {row_number} пропущена по известному seller | seller={research_row.seller_name_raw} | status={skipped_row['parse_status']} | inn={skipped_row['inn']}\n"
                    )
                    continue

                self._append_log(
                    f"Batch: обрабатываю строку {row_number}/{start_row + total - 1} | sheet={research_row.source_sheet} | seller={research_row.seller_name_raw}\n"
                )
                if inspector is None:
                    inspector = exit_stack.enter_context(
                        BatchInspector(artifacts_dir=artifacts_dir, headful=True, profile_dir=profile_dir)
                    )
                try:
                    result = inspector.inspect_row(
                        row_number=row_number,
                        research_row=research_row,
                    )
                except Exception as exc:
                    result = build_row_error_result(
                        row_number=row_number,
                        research_row=research_row,
                        error=exc,
                        profile_dir=profile_dir,
                    )
                    self._append_log(
                        f"Batch: СЃС‚СЂРѕРєР° {row_number} СѓРїР°Р»Р° РїРѕСЃР»Рµ РїРѕРІС‚РѕСЂРЅС‹С… РїРѕРїС‹С‚РѕРє, РїРёС€Сѓ ROW_ERROR | note={result.note}\n"
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
        history_summary = update_seller_history_from_batch_files(
            batch_results_paths=[batch_output],
            history_path=seller_history_path,
        )
        self.root.after(0, lambda batch_output=batch_output: self.merge_batch_input_var.set(str(batch_output)))
        self._set_batch_progress(done=total, total=max(total, 1), row_number=start_row + total - 1 if total else start_row)
        next_row = start_row + total
        self.root.after(0, lambda next_row=next_row: self.row_var.set(str(next_row)))
        self._append_log(f"Batch: итоговый файл сохранён: {batch_output}\n")
        self._append_log(
            f"Batch: история sellers обновлена автоматически | файл={seller_history_path} | новых sellers={history_summary['new_seller_added']} | всего sellers={history_summary['history_rows']}\n"
        )
        if known_seller_sources:
            self._append_log(f"Batch: пропущено по реестру ИНН: {skipped_known_sellers}\n")
            self._append_log(f"Batch: пропущено по истории sellers: {skipped_viewed_sellers}\n")
        self._append_log(f"Batch: рекомендованная следующая стартовая строка = {next_row}\n")
        self.root.after(0, lambda: messagebox.showinfo("Batch завершён", f"Итоговый Excel сохранён:\n{batch_output}"))
        self.root.after(0, lambda batch_output=batch_output: self._prompt_add_batch_to_registry(batch_output))

    def _prompt_add_batch_to_registry(self, batch_output: Path) -> None:
        if messagebox.askyesno(
            "Реестр ИНН",
            "Batch завершён. История sellers уже обновлена автоматически. Добавить этот batch_results.xlsx ещё и во вкладку реестра ИНН?",
        ):
            self._add_registry_batch_paths([batch_output])

    def _action_excel_merge(self) -> None:
        source_value = self.excel_merge_source_var.get().strip()
        if not source_value:
            raise ValueError("Сначала выбери ZIP-архив или папку с Excel-файлами")

        source_path = Path(source_value)
        if not source_path.exists():
            raise FileNotFoundError(f"Источник не найден: {source_path}")

        output_xlsx_path = Path(self.excel_merge_output_xlsx_var.get().strip() or "output/merged_tabs.xlsx")
        output_zip_path = Path(self.excel_merge_output_zip_var.get().strip() or "output/merged_tabs.zip")
        create_zip = bool(self.excel_merge_create_zip_var.get())

        self._append_log(f"Excel Merge: source={source_path}\n")
        self._append_log(f"Excel Merge: output_xlsx={output_xlsx_path}\n")
        self._append_log(f"Excel Merge: create_zip={create_zip}\n")
        if create_zip:
            self._append_log(f"Excel Merge: output_zip={output_zip_path}\n")

        summary = merge_excel_files_to_tabs(
            source_path=source_path,
            output_xlsx_path=output_xlsx_path,
            output_zip_path=output_zip_path if create_zip else None,
            create_zip=create_zip,
        )

        zip_line = summary["output_zip_path"] or "не создавался"
        text = (
            f"Источник: {summary['source_path']}\n"
            f"Файлов найдено: {summary['files_found']}\n"
            f"Листов создано: {summary['sheets_created']}\n"
            f"Итоговый Excel: {summary['output_xlsx_path']}\n"
            f"Итоговый архив: {zip_line}"
        )
        self.excel_merge_summary_var.set(text)
        self._append_log("Excel Merge: объединение завершено\n")
        self._append_log(text + "\n")
        self._save_settings()
        self.root.after(0, lambda: messagebox.showinfo("Объединение Excel завершено", text))

    def _action_merge_registries(self) -> None:
        primary_registry_path = Path(self.registry_merge_primary_var.get().strip() or self.registry_path_var.get().strip() or "output/inn_registry.xlsx")
        secondary_registry_value = self.registry_merge_secondary_var.get().strip()
        if not secondary_registry_value:
            raise ValueError("Сначала выбери второй реестр ИНН")
        secondary_registry_path = Path(secondary_registry_value)
        output_path = Path(self.registry_merge_output_var.get().strip() or "output/inn_registry_merged.xlsx")

        self._append_log(f"Registry Merge: primary={primary_registry_path}\n")
        self._append_log(f"Registry Merge: secondary={secondary_registry_path}\n")
        self._append_log(f"Registry Merge: output={output_path}\n")

        summary = merge_inn_registry_files(
            primary_registry_path=primary_registry_path,
            secondary_registry_path=secondary_registry_path,
            output_path=output_path,
        )

        text = (
            f"Главный реестр: {summary['primary_registry_path']}\n"
            f"Второй реестр: {summary['secondary_registry_path']}\n"
            f"Строк в первом: {summary['primary_rows_total']}\n"
            f"Уникальных ИНН в первом: {summary['primary_unique_inn']}\n"
            f"Строк во втором: {summary['secondary_rows_total']}\n"
            f"Уникальных ИНН во втором: {summary['secondary_unique_inn']}\n"
            f"Совпадений по ИНН: {summary['overlaps_with_primary']}\n"
            f"Добавлено из второго: {summary['added_from_secondary']}\n"
            f"Итоговых ИНН: {summary['merged_registry_rows']}\n"
            f"Итоговый реестр: {summary['output_path']}"
        )
        self._set_registry_summary(text)
        self._append_log("Registry Merge: объединение завершено\n")
        self._append_log(text + "\n")
        self._save_settings()
        self.root.after(0, lambda: messagebox.showinfo("Реестры объединены", text))

    def _action_update_registry(self) -> None:
        if not self._registry_batch_files:
            raise ValueError("Сначала добавь хотя бы один batch_results.xlsx")

        batch_paths = [Path(path) for path in self._registry_batch_files]
        missing = [str(path) for path in batch_paths if not path.exists()]
        if missing:
            raise FileNotFoundError("Не найдены batch-файлы:\n" + "\n".join(missing[:10]))

        registry_path = Path(self.registry_path_var.get().strip() or "output/inn_registry.xlsx")
        seller_history_path = Path(self.seller_history_path_var.get().strip() or "output/seller_history.xlsx")
        new_inn_output_path = Path(self.registry_new_inn_output_var.get().strip() or "output/new_inn_for_kontur.xlsx")

        self._append_log(f"Registry: registry={registry_path}\n")
        self._append_log(f"Registry: seller_history={seller_history_path}\n")
        self._append_log(f"Registry: new_inn={new_inn_output_path}\n")
        self._append_log(f"Registry: batch_files={len(batch_paths)}\n")

        registry_summary = update_inn_registry_from_batch_files(
            batch_results_paths=batch_paths,
            registry_path=registry_path,
            new_inn_output_path=new_inn_output_path,
        )
        history_summary = update_seller_history_from_batch_files(
            batch_results_paths=batch_paths,
            history_path=seller_history_path,
        )

        text = (
            f"Файлов batch: {registry_summary['batch_files']}\n"
            f"Строк всего: {registry_summary['rows_total']}\n"
            f"Строк с ИНН: {registry_summary['rows_with_inn']}\n"
            f"Новых ИНН добавлено: {registry_summary['new_inn_added']}\n"
            f"Уже были в реестре: {registry_summary['already_known']}\n"
            f"Дубликаты ИНН в импорте: {registry_summary['duplicate_in_import']}\n"
            f"Пропущено без ИНН: {registry_summary['skipped_without_inn']}\n"
            f"Всего ИНН в реестре: {registry_summary['registry_rows']}\n"
            f"Реестр ИНН: {registry_summary['registry_path']}\n"
            f"Новые ИНН для Контур: {registry_summary['new_inn_output_path']}\n\n"
            f"История sellers: {history_summary['history_path']}\n"
            f"Строк с seller: {history_summary['rows_with_seller']}\n"
            f"Новых sellers добавлено: {history_summary['new_seller_added']}\n"
            f"Уже были в истории: {history_summary['already_known']}\n"
            f"Дубликаты sellers в импорте: {history_summary['duplicate_in_import']}\n"
            f"Пропущено без seller: {history_summary['skipped_without_seller']}\n"
            f"Seller-ов с автопропуском: {history_summary['skip_recommended_total']}\n"
            f"Всего sellers в истории: {history_summary['history_rows']}"
        )
        self._set_registry_summary(text)
        self._append_log("Registry: обновление реестра ИНН и истории sellers завершено\n")
        self._append_log(text + "\n")
        self._save_settings()
        self.root.after(0, lambda: messagebox.showinfo("Реестр И история обновлены", text))

    def _action_import_history_from_registry(self, registry_paths: list[Path]) -> None:
        missing = [str(path) for path in registry_paths if not path.exists()]
        if missing:
            raise FileNotFoundError("Не найдены файлы реестра:\n" + "\n".join(missing[:10]))

        seller_history_path = Path(self.seller_history_path_var.get().strip() or "output/seller_history.xlsx")

        self._append_log(f"History Import: seller_history={seller_history_path}\n")
        self._append_log(f"History Import: registry_files={len(registry_paths)}\n")
        summary = import_seller_history_from_registry_files(
            registry_paths=registry_paths,
            history_path=seller_history_path,
        )

        text = (
            f"Файлов реестра: {summary['source_files']}\n"
            f"Строк всего: {summary['rows_total']}\n"
            f"Строк с seller: {summary['rows_with_seller']}\n"
            f"Новых sellers добавлено: {summary['new_seller_added']}\n"
            f"Уже были в истории: {summary['already_known']}\n"
            f"Дубликаты sellers в импорте: {summary['duplicate_in_import']}\n"
            f"Пропущено без seller: {summary['skipped_without_seller']}\n"
            f"Seller-ов с автопропуском: {summary['skip_recommended_total']}\n"
            f"Всего sellers в истории: {summary['history_rows']}\n"
            f"История sellers: {summary['history_path']}"
        )
        self._set_registry_summary(text)
        self._append_log("History Import: импорт sellers из реестра завершён\n")
        self._append_log(text + "\n")
        self._save_settings()
        self.root.after(0, lambda: messagebox.showinfo("История sellers обновлена", text))

    def _action_merge_compass(self) -> None:
        batch_results_value = self.merge_batch_input_var.get().strip() or self.batch_output_var.get().strip() or "output/batch_results.xlsx"
        batch_results_path = Path(batch_results_value)
        if not batch_results_path.exists():
            raise FileNotFoundError(f"Не найден batch_results.xlsx: {batch_results_path}")

        compass_value = self.compass_input_var.get().strip()
        if not compass_value:
            raise ValueError("Сначала выбери выгрузку Контур")
        compass_path = Path(compass_value)
        if not compass_path.exists():
            raise FileNotFoundError(f"Не найден файл Контур: {compass_path}")

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

    def _require_adbeam_input_path(self) -> Path:
        input_value = self.adbeam_input_path_var.get().strip()
        if not input_value:
            raise ValueError("Сначала выбери Excel-файл во вкладке AdBeam")
        input_path = Path(input_value)
        if not input_path.exists():
            raise FileNotFoundError(f"Файл не найден: {input_path}")
        return input_path

    def _require_conference_input_path(self) -> Path:
        input_value = self.conference_input_path_var.get().strip()
        if not input_value:
            raise ValueError("Сначала выбери Excel-файл во вкладке Конференция ICP-1")
        input_path = Path(input_value)
        if not input_path.exists():
            raise FileNotFoundError(f"Файл не найден: {input_path}")
        return input_path

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

    @staticmethod
    def _parse_non_negative_float(raw_value: str, field_name: str) -> float:
        try:
            value = float(raw_value.replace(",", "."))
        except ValueError as exc:
            raise ValueError(f"{field_name}: требуется число") from exc
        if value < 0:
            raise ValueError(f"{field_name}: число должно быть >= 0")
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
