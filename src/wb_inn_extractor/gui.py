from __future__ import annotations

import json
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .excel_io import analyze_workbook, extract_research_rows, read_research_row, save_research_sample
from .wb_research import inspect_product_row


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("WB INN Extractor — Research MVP")
        self.root.geometry("980x760")
        self.root.minsize(900, 680)

        self.input_path_var = tk.StringVar()
        self.sample_output_var = tk.StringVar(value=str(Path("output/research_sample.xlsx")))
        self.artifacts_dir_var = tk.StringVar(value=str(Path("output/artifacts")))
        self.limit_var = tk.StringVar(value="30")
        self.row_var = tk.StringVar(value="2")
        self.headful_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Готово")

        self._build_layout()

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(4, weight=1)

        ttk.Label(container, text="Входной Excel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(container, textvariable=self.input_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
        ttk.Button(container, text="Выбрать файл", command=self._choose_input_file).grid(row=0, column=2, sticky="ew", pady=(0, 8))

        options = ttk.LabelFrame(container, text="Параметры", padding=12)
        options.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        options.columnconfigure(1, weight=1)

        ttk.Label(options, text="Файл research sample").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.sample_output_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(options, text="Выбрать", command=self._choose_sample_output).grid(row=0, column=2, sticky="ew", pady=4)

        ttk.Label(options, text="Папка для артефактов").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.artifacts_dir_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(options, text="Выбрать", command=self._choose_artifacts_dir).grid(row=1, column=2, sticky="ew", pady=4)

        ttk.Label(options, text="Лимит строк для sample").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.limit_var, width=12).grid(row=2, column=1, sticky="w", padx=(8, 8), pady=4)

        ttk.Label(options, text="Номер строки в sample для inspect").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.row_var, width=12).grid(row=3, column=1, sticky="w", padx=(8, 8), pady=4)

        ttk.Checkbutton(options, text="Показывать браузер при inspect", variable=self.headful_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        actions = ttk.Frame(container)
        actions.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        for idx in range(3):
            actions.columnconfigure(idx, weight=1)

        ttk.Button(actions, text="1. Проанализировать Excel", command=lambda: self._run_in_thread(self._action_analyze)).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(actions, text="2. Создать research sample", command=lambda: self._run_in_thread(self._action_sample)).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="3. Проверить строку", command=lambda: self._run_in_thread(self._action_inspect)).grid(row=0, column=2, sticky="ew", padx=(8, 0))

        log_frame = ttk.LabelFrame(container, text="Лог / результат", padding=8)
        log_frame.grid(row=4, column=0, columnspan=3, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        status_bar = ttk.Label(container, textvariable=self.status_var, anchor="w")
        status_bar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(8, 0))

    def _choose_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери Excel-файл",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            self.input_path_var.set(path)

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
        path = filedialog.askdirectory(title="Папка для artifacts")
        if path:
            self.artifacts_dir_var.set(path)

    def _run_in_thread(self, target) -> None:
        thread = threading.Thread(target=self._safe_run, args=(target,), daemon=True)
        thread.start()

    def _safe_run(self, target) -> None:
        self._set_status("Выполняется...")
        try:
            target()
            self._set_status("Готово")
        except Exception as exc:  # pragma: no cover - UI-level fallback
            self._append_log("\n[ERROR]\n")
            self._append_log(f"{exc}\n")
            self._append_log(traceback.format_exc())
            self._set_status("Ошибка")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(exc)))

    def _action_analyze(self) -> None:
        input_path = self._require_input_path()
        summary = analyze_workbook(input_path)
        self._append_log(json.dumps(summary.model_dump(mode="json"), ensure_ascii=False, indent=2) + "\n")

    def _action_sample(self) -> None:
        input_path = self._require_input_path()
        output_path = Path(self.sample_output_var.get())
        limit = self._parse_positive_int(self.limit_var.get(), field_name="Лимит строк")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        rows = extract_research_rows(input_path, limit=limit)
        save_research_sample(output_path, rows)

        self._append_log(f"Создан файл: {output_path}\n")
        self._append_log(f"Строк: {len(rows)}\n")

    def _action_inspect(self) -> None:
        sample_path = Path(self.sample_output_var.get())
        if not sample_path.exists():
            raise FileNotFoundError(f"Не найден sample-файл: {sample_path}")

        row_number = self._parse_positive_int(self.row_var.get(), field_name="Номер строки")
        artifacts_dir = Path(self.artifacts_dir_var.get())
        artifacts_dir.mkdir(parents=True, exist_ok=True)

        research_row = read_research_row(sample_path, row_number=row_number)
        result = inspect_product_row(
            row_number=row_number,
            research_row=research_row,
            artifacts_dir=artifacts_dir,
            headful=self.headful_var.get(),
        )
        self._append_log(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2) + "\n")
        self.root.after(0, lambda: messagebox.showinfo("Inspect завершён", f"Артефакты сохранены в:\n{artifacts_dir}"))

    def _require_input_path(self) -> Path:
        input_value = self.input_path_var.get().strip()
        if not input_value:
            raise ValueError("Сначала выбери входной Excel")
        input_path = Path(input_value)
        if not input_path.exists():
            raise FileNotFoundError(f"Файл не найден: {input_path}")
        return input_path

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
        self.root.after(0, lambda: self._append_log_on_ui(text))

    def _append_log_on_ui(self, text: str) -> None:
        self.log_text.insert("end", text)
        self.log_text.see("end")

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_var.set(text))


def main() -> None:
    root = tk.Tk()
    ttk.Style().theme_use("clam")
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
