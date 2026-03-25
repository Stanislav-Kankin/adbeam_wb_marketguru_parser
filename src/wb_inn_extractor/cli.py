from __future__ import annotations

import argparse
import json
from pathlib import Path

from .excel_io import analyze_workbook, extract_research_rows, read_research_row, save_research_sample
from .wb_research import inspect_product_row


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="wb-inn")
    subparsers = parser.add_subparsers(dest="command", required=True)

    analyze_parser = subparsers.add_parser("analyze", help="Проанализировать структуру входного Excel")
    analyze_parser.add_argument("--input", required=True, type=Path)

    sample_parser = subparsers.add_parser("sample", help="Подготовить research-выборку")
    sample_parser.add_argument("--input", required=True, type=Path)
    sample_parser.add_argument("--output", type=Path, default=Path("output/research_sample.xlsx"))
    sample_parser.add_argument("--limit", type=int, default=30)

    inspect_parser = subparsers.add_parser("inspect-row", help="Открыть одну строку через Playwright и сохранить артефакты")
    inspect_parser.add_argument("--input", required=True, type=Path)
    inspect_parser.add_argument("--row", required=True, type=int)
    inspect_parser.add_argument("--artifacts-dir", type=Path, default=Path("output/artifacts"))
    inspect_parser.add_argument("--headful", action="store_true")

    manual_parser = subparsers.add_parser("manual-session", help="Открыть строку в persistent profile режиме и дать время на ручную проверку")
    manual_parser.add_argument("--input", required=True, type=Path)
    manual_parser.add_argument("--row", required=True, type=int)
    manual_parser.add_argument("--artifacts-dir", type=Path, default=Path("output/artifacts"))
    manual_parser.add_argument("--profile-dir", type=Path, default=Path("output/wb_profile"))
    manual_parser.add_argument("--wait-seconds", type=int, default=90)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "analyze":
        summary = analyze_workbook(args.input)
        print(json.dumps(summary.model_dump(mode="json"), ensure_ascii=False, indent=2))
        return

    if args.command == "sample":
        args.output.parent.mkdir(parents=True, exist_ok=True)
        rows = extract_research_rows(args.input, limit=args.limit)
        save_research_sample(args.output, rows)
        print(f"Создан файл: {args.output}")
        print(f"Строк: {len(rows)}")
        return

    if args.command == "inspect-row":
        args.artifacts_dir.mkdir(parents=True, exist_ok=True)
        research_row = read_research_row(args.input, row_number=args.row)
        result = inspect_product_row(
            row_number=args.row,
            research_row=research_row,
            artifacts_dir=args.artifacts_dir,
            headful=args.headful,
        )
        print(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2))
        return

    if args.command == "manual-session":
        if args.artifacts_dir.resolve() == args.profile_dir.resolve():
            raise ValueError("Папка артефактов и папка профиля WB должны быть разными")

        args.artifacts_dir.mkdir(parents=True, exist_ok=True)
        args.profile_dir.mkdir(parents=True, exist_ok=True)
        research_row = read_research_row(args.input, row_number=args.row)
        result = inspect_product_row(
            row_number=args.row,
            research_row=research_row,
            artifacts_dir=args.artifacts_dir,
            headful=True,
            profile_dir=args.profile_dir,
            manual_wait_seconds=args.wait_seconds,
        )
        print(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2))
        return

    raise ValueError(f"Неизвестная команда: {args.command}")


if __name__ == "__main__":
    main()
