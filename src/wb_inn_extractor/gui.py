from __future__ import annotations

import json
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
from .wb_research import inspect_product_row


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("WB INN Extractor — Research MVP")
        self.root.geometry("1180x820")
        self.root.minsize(1040, 720)

        self.input_path_var = tk.StringVar()
        self.sample_output_var = tk.StringVar(value=str(Path("output/research_sample.xlsx")))
        self.artifacts_dir_var = tk.StringVar(value=str(Path("output/artifacts")))
        self.profile_dir_var = tk.StringVar(value=str(Path("output/wb_profile")))
        self.limit_var = tk.StringVar(value="30")
        self.row_var = tk.StringVar(value="2")
        self.batch_count_var = tk.StringVar(value="5")
        self.batch_output_var = tk.StringVar(value=str(Path("output/batch_results.xlsx")))
        self.wait_seconds_var = tk.StringVar(value="90")
        self.headful_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Готово")

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        container.rowconfigure(4, weight=1)
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
        ttk.Button(options, text="Выбрать", command=self._choose_sample_output).grid(row=6, column=2, sticky="ew", pady=4)

        ttk.Label(options, text="Секунд на ручную проверку").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.wait_seconds_var, width=12).grid(row=7, column=1, sticky="w", padx=(8, 8), pady=4)

        ttk.Checkbutton(options, text="Показывать браузер при обычном inspect", variable=self.headful_var).grid(row=8, column=0, columnspan=2, sticky="w", pady=(8, 0))

        hint = (
            "Обычный inspect теперь по умолчанию использует папку профиля WB, если она заполнена. "
            "Это безопаснее для WB, чем каждый раз открывать новый чистый браузер. "
            "Папка профиля WB и папка артефактов должны быть разными."
        )
        ttk.Label(container, text=hint, wraplength=1120, justify="left").grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        actions = ttk.Frame(container)
        actions.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        for idx in range(5):
            actions.columnconfigure(idx, weight=1)

        ttk.Button(actions, text="1. Проанализировать Excel", command=lambda: self._run_in_thread(self._action_analyze)).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(actions, text="2. Создать research sample", command=lambda: self._run_in_thread(self._action_sample)).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="3. Проверить строку", command=lambda: self._run_in_thread(self._action_inspect)).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(actions, text="4. Пакетный прогон", command=lambda: self._run_in_thread(self._action_batch)).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(actions, text="5. Ручная сессия WB", command=lambda: self._run_in_thread(self._action_manual_session)).grid(row=0, column=4, sticky="ew", padx=(8, 0))

        tools = ttk.Frame(container)
        tools.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        for idx in range(2):
            tools.columnconfigure(idx, weight=1)
        ttk.Button(tools, text="Открыть папку артефактов", command=self._open_artifacts_dir).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(tools, text="Открыть папку профиля WB", command=self._open_profile_dir).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        log_frame = ttk.LabelFrame(container, text="Лог / результат", padding=8)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        container.rowconfigure(5, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        status_bar = ttk.Label(container, textvariable=self.status_var, anchor="w")
        status_bar.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 0))

    def _choose_input_file(self) -> None:
        path = filedialog.askopenfilename(title="Выбери Excel-файл", filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
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
        path = filedialog.askdirectory(title="Папка для артефактов")
        if path:
            self.artifacts_dir_var.set(path)

    def _choose_profile_dir(self) -> None:
        path = filedialog.askdirectory(title="Папка профиля WB")
        if path:
            self.profile_dir_var.set(path)

    def _open_artifacts_dir(self) -> None:
        self._open_dir(Path(self.artifacts_dir_var.get()))

    def _open_profile_dir(self) -> None:
        self._open_dir(Path(self.profile_dir_var.get()))

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
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(exec)))

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
        self._append_log(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2) + "\n")
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
        for offset, research_row in enumerate(research_rows, start=0):
            row_number = start_row + offset
            self._append_log(f"Batch: обрабатываю строку {row_number}...")
            result = inspect_product_row(
                row_number=row_number,
                research_row=research_row,
                artifacts_dir=artifacts_dir,
                headful=True,
                profile_dir=profile_dir,
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
            self._append_log(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2) + "")

        save_batch_results(batch_output, output_rows)
        self.root.after(0, lambda: messagebox.showinfo("Batch завершён", f"Итоговый Excel сохранён:{batch_output}"))

    def _action_manual_session(self) -> None:
        sample_path = self._require_sample_path()
        row_number = self._parse_positive_int(self.row_var.get(), field_name="Номер строки")
        wait_seconds = self._parse_positive_int(self.wait_seconds_var.get(), field_name="Секунды на ручную проверку")
        artifacts_dir = Path(self.artifacts_dir_var.get())
        profile_dir = self._required_profile_dir()

        if artifacts_dir.resolve() == profile_dir.resolve():
            raise ValueError("Папка артефактов и папка профиля WB должны быть разными")

        artifacts_dir.mkdir(parents=True, exist_ok=True)
        profile_dir.mkdir(parents=True, exist_ok=True)

        self._append_log(
            "Ручная сессия WB началась. Что делать:\n"
            "1) если открылась защита WB — пройди её руками;\n"
            "2) если страница пустая — попробуй обновить;\n"
            "3) дождись конца таймера, потом смотри JSON и screenshot.\n"
        )
        research_row = read_research_row(sample_path, row_number=row_number)
        result = inspect_product_row(
            row_number=row_number,
            research_row=research_row,
            artifacts_dir=artifacts_dir,
            headful=True,
            profile_dir=profile_dir,
            manual_wait_seconds=wait_seconds,
        )
        self._append_log(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2) + "\n")
        self.root.after(
            0,
            lambda: messagebox.showinfo(
                "Ручная сессия завершена",
                f"Профиль: {profile_dir}\nАртефакты: {artifacts_dir}",
            ),
        )

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
