from __future__ import annotations

import os
import subprocess
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .excel_io import (
    analyze_workbook,
    extract_research_rows,
    read_research_row,
    read_research_rows_range,
    save_batch_results,
    save_research_sample,
)
from .models import AnalyzeSummary
from .wb_research import BatchInspector, inspect_product_row


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("WB INN Extractor — Research MVP")
        self.root.geometry("1280x900")
        self.root.minsize(1120, 760)

        self.input_path_var = tk.StringVar()
        self.sample_output_var = tk.StringVar(value=str(Path("output/research_sample.xlsx")))
        self.artifacts_dir_var = tk.StringVar(value=str(Path("output/artifacts")))
        self.profile_dir_var = tk.StringVar(value=str(Path("output/wb_profile")))
        self.limit_var = tk.StringVar(value="")
        self.row_var = tk.StringVar(value="2")
        self.batch_count_var = tk.StringVar(value="5")
        self.batch_output_var = tk.StringVar(value=str(Path("output/batch_results.xlsx")))
        self.headful_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Готово")
        self.sheet_mode_var = tk.StringVar(value="all")
        self.sheet_summary_var = tk.StringVar(value="Листы ещё не проанализированы")
        self._last_analyze_summary: AnalyzeSummary | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        container.rowconfigure(6, weight=1)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Входной Excel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        ttk.Entry(container, textvariable=self.input_path_var).grid(row=0, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(container, text="Выбрать файл", command=self._choose_input_file).grid(row=0, column=2, sticky="ew", padx=(8, 0), pady=(0, 8))

        options = ttk.LabelFrame(container, text="Параметры", padding=12)
        options.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        options.columnconfigure(1, weight=1)

        ttk.Label(options, text="Файл research sample").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.sample_output_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(options, text="Выбрать", command=self._choose_sample_output).grid(row=0, column=2, sticky="ew", pady=4)

        ttk.Label(options, text="Папка для артефактов").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.artifacts_dir_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(options, text="Выбрать", command=self._choose_artifacts_dir).grid(row=1, column=2, sticky="ew", pady=4)

        ttk.Label(options, text="Папка профиля WB").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.profile_dir_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(options, text="Выбрать", command=self._choose_profile_dir).grid(row=2, column=2, sticky="ew", pady=4)

        ttk.Label(options, text="Лимит строк для sample").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.limit_var, width=12).grid(row=3, column=1, sticky="w", padx=(8, 8), pady=4)

        ttk.Label(options, text="Номер строки в sample для inspect").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.row_var, width=12).grid(row=4, column=1, sticky="w", padx=(8, 8), pady=4)

        ttk.Label(options, text="Сколько строк обработать в batch").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.batch_count_var, width=12).grid(row=5, column=1, sticky="w", padx=(8, 8), pady=4)

        ttk.Label(options, text="Итоговый Excel batch").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.batch_output_var).grid(row=6, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(options, text="Выбрать", command=self._choose_batch_output).grid(row=6, column=2, sticky="ew", pady=4)

        ttk.Checkbutton(options, text="Показывать браузер при обычном inspect", variable=self.headful_var).grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))

        sheet_frame = ttk.LabelFrame(container, text="Листы Excel", padding=12)
        sheet_frame.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(0, 12))
        sheet_frame.columnconfigure(0, weight=1)
        sheet_frame.rowconfigure(2, weight=1)

        ttk.Label(sheet_frame, textvariable=self.sheet_summary_var, justify="left").grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))

        mode_frame = ttk.Frame(sheet_frame)
        mode_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        ttk.Radiobutton(mode_frame, text="Использовать все валидные листы", value="all", variable=self.sheet_mode_var).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(mode_frame, text="Использовать только выбранные листы", value="selected", variable=self.sheet_mode_var).grid(row=0, column=1, sticky="w", padx=(16, 0))

        list_frame = ttk.Frame(sheet_frame)
        list_frame.grid(row=2, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.sheet_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=8, exportselection=False)
        self.sheet_listbox.grid(row=0, column=0, sticky="nsew")
        sheet_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.sheet_listbox.yview)
        sheet_scrollbar.grid(row=0, column=1, sticky="ns")
        self.sheet_listbox.configure(yscrollcommand=sheet_scrollbar.set)

        sheet_buttons = ttk.Frame(sheet_frame)
        sheet_buttons.grid(row=2, column=1, sticky="ns", padx=(12, 0))
        ttk.Button(sheet_buttons, text="Выбрать все", command=self._select_all_sheets).grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(sheet_buttons, text="Только валидные", command=self._select_valid_sheets).grid(row=1, column=0, sticky="ew", pady=6)
        ttk.Button(sheet_buttons, text="Снять выбор", command=self._clear_sheet_selection).grid(row=2, column=0, sticky="ew", pady=6)

        hint = (
            "1) Нажми 'Проанализировать Excel'. 2) Посмотри список листов и выбери режим: все валидные или только нужные. "
            "3) Создай sample только по выбранным листам. Обычный inspect использует папку профиля WB, если она заполнена. "
            "Папка профиля WB и папка артефактов должны быть разными."
        )
        ttk.Label(container, text=hint, wraplength=1220, justify="left").grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        actions = ttk.Frame(container)
        actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        for idx in range(4):
            actions.columnconfigure(idx, weight=1)

        ttk.Button(actions, text="1. Проанализировать Excel", command=lambda: self._run_in_thread(self._action_analyze)).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(actions, text="2. Создать research sample", command=lambda: self._run_in_thread(self._action_sample)).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="3. Проверить строку", command=lambda: self._run_in_thread(self._action_inspect)).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(actions, text="4. Пакетный прогон", command=lambda: self._run_in_thread(self._action_batch)).grid(row=0, column=3, sticky="ew", padx=(4, 0))

        tools = ttk.Frame(container)
        tools.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        for idx in range(3):
            tools.columnconfigure(idx, weight=1)
        ttk.Button(tools, text="Открыть папку артефактов", command=self._open_artifacts_dir).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(tools, text="Открыть папку профиля WB", command=self._open_profile_dir).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(tools, text="Открыть результаты", command=self._open_batch_results).grid(row=0, column=2, sticky="ew", padx=(8, 0))

        log_frame = ttk.LabelFrame(container, text="Лог / результат", padding=8)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        status_bar = ttk.Label(container, textvariable=self.status_var, anchor="w")
        status_bar.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 0))

    def _choose_input_file(self) -> None:
        path = filedialog.askopenfilename(title="Выбери Excel-файл", filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if path:
            self.input_path_var.set(path)
            self._last_analyze_summary = None
            self.sheet_listbox.delete(0, "end")
            self.sheet_summary_var.set("Файл выбран. Нажми 'Проанализировать Excel', чтобы увидеть листы.")

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

    def _choose_batch_output(self) -> None:
        current = Path(self.batch_output_var.get()) if self.batch_output_var.get().strip() else Path("output/batch_results.xlsx")
        path = filedialog.asksaveasfilename(
            title="Куда сохранить итоговый Excel batch",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=current.name,
            initialdir=str(current.parent),
        )
        if path:
            self.batch_output_var.set(path)

    def _open_artifacts_dir(self) -> None:
        self._open_dir(Path(self.artifacts_dir_var.get()))

    def _open_profile_dir(self) -> None:
        self._open_dir(Path(self.profile_dir_var.get()))

    def _open_batch_results(self) -> None:
        path = Path(self.batch_output_var.get().strip() or "output/batch_results.xlsx")
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            try:
                os.startfile(str(path))  # type: ignore[attr-defined]
                return
            except AttributeError:
                subprocess.Popen(["xdg-open", str(path)])
                return
        self._open_dir(path.parent)

    def _open_dir(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except AttributeError:
            subprocess.Popen(["xdg-open", str(path)])

    def _run_in_thread(self, target) -> None:
        thread = threading.Thread(target=self._safe_run, args=(target,), daemon=True)
        thread.start()

    def _safe_run(self, target) -> None:
        self._set_status("Выполняется...")
        try:
            target()
            self._set_status("Готово")
        except Exception as exc:
            self._append_log("\n[ERROR]\n")
            self._append_log(f"{exc}\n")
            self._append_log(traceback.format_exc())
            self._set_status("Ошибка")
            self.root.after(0, lambda exc=exc: messagebox.showerror("Ошибка", str(exc)))

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
        save_research_sample(output_path, rows)

        sheet_scope_label = "все валидные листы"
        if selected_sheets:
            sheet_scope_label = ", ".join(selected_sheets)

        self._append_log(f"Создан файл: {output_path}\n")
        self._append_log(f"Листы для sample: {sheet_scope_label}\n")
        if limit is None:
            self._append_log("Лимит строк не задан: в sample сохранены все уникальные продавцы по паре бренд+продавец\n")
        self._append_log(f"Строк (уникальных по продавцу+бренду): {len(rows)}\n")

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
        output_rows = []
        with BatchInspector(artifacts_dir=artifacts_dir, headful=True, profile_dir=profile_dir) as inspector:
            for offset, research_row in enumerate(research_rows, start=0):
                row_number = start_row + offset
                self._append_log(
                    f"Batch: обрабатываю строку {row_number} | sheet={research_row.source_sheet} | seller={research_row.seller_name_raw}\n"
                )
                result = inspector.inspect_row(
                    row_number=row_number,
                    research_row=research_row,
                )
                output_rows.append({
                    "row_number": row_number,
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
                })
                self._append_log(f"Batch: строка {row_number} завершена, status={result.parse_status}, inn={result.inn}\n")

        save_batch_results(batch_output, output_rows)
        self.root.after(0, lambda: messagebox.showinfo("Batch завершён", f"Итоговый Excel сохранён:\n{batch_output}"))

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

        self._select_valid_sheets()
        self.sheet_summary_var.set(
            "Всего листов в книге: "
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

    def _select_valid_sheets(self) -> None:
        self.sheet_listbox.selection_clear(0, "end")
        for index in range(self.sheet_listbox.size()):
            raw_text = self.sheet_listbox.get(index)
            if raw_text.startswith("[VALID]"):
                self.sheet_listbox.selection_set(index)

    def _clear_sheet_selection(self) -> None:
        self.sheet_listbox.selection_clear(0, "end")

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
        self.root.after(0, lambda: self._append_log_on_ui(text))

    def _append_log_on_ui(self, text: str) -> None:
        self.log_text.insert("end", text)
        self.log_text.see("end")

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_var.set(text))


def main() -> None:
    root = tk.Tk()
    ttk.Style().theme_use("clam")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
