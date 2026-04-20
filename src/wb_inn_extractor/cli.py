from __future__ import annotations

import argparse
import json
from pathlib import Path

from .excel_io import (
    analyze_workbook,
    extract_research_rows,
    merge_inn_registry_files,
    merge_batch_results_with_compass,
    export_batch_no_inn,
    export_compass_unmatched,
    export_unique_inn_list,
    read_research_row,
    read_research_rows_range,
    save_batch_results,
    save_research_sample,
    summarize_research_rows_by_sheet,
)
from .wb_research import BatchInspector, build_row_error_result, inspect_product_row


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="wb-inn")
    subparsers = parser.add_subparsers(dest="command", required=True)

    analyze_parser = subparsers.add_parser("analyze", help="Проанализировать структуру входного Excel")
    analyze_parser.add_argument("--input", required=True, type=Path)
    analyze_parser.add_argument("--sheet", dest="selected_sheets", action="append", default=None)

    sample_parser = subparsers.add_parser("sample", help="Подготовить research-выборку")
    sample_parser.add_argument("--input", required=True, type=Path)
    sample_parser.add_argument("--output", type=Path, default=Path("output/research_sample.xlsx"))
    sample_parser.add_argument("--limit", type=int, default=None)
    sample_parser.add_argument("--sheet", dest="selected_sheets", action="append", default=None)

    inspect_parser = subparsers.add_parser("inspect-row", help="Открыть одну строку через Playwright и сохранить артефакты")
    inspect_parser.add_argument("--input", required=True, type=Path)
    inspect_parser.add_argument("--row", required=True, type=int)
    inspect_parser.add_argument("--artifacts-dir", type=Path, default=Path("output/artifacts"))
    inspect_parser.add_argument("--profile-dir", type=Path, default=None)
    inspect_parser.add_argument("--headful", action="store_true")

    manual_parser = subparsers.add_parser("manual-session", help="Открыть строку в persistent profile режиме и дать время на ручную проверку")
    manual_parser.add_argument("--input", required=True, type=Path)
    manual_parser.add_argument("--row", required=True, type=int)
    manual_parser.add_argument("--artifacts-dir", type=Path, default=Path("output/artifacts"))
    manual_parser.add_argument("--profile-dir", type=Path, default=Path("output/wb_profile"))
    manual_parser.add_argument("--wait-seconds", type=int, default=90)

    batch_parser = subparsers.add_parser("batch-run", help="Пакетно обработать несколько строк и сохранить итоговый Excel")
    batch_parser.add_argument("--input", required=True, type=Path)
    batch_parser.add_argument("--start-row", type=int, default=2)
    batch_parser.add_argument("--limit", type=int, default=10)
    batch_parser.add_argument("--output", type=Path, default=Path("output/batch_results.xlsx"))
    batch_parser.add_argument("--artifacts-dir", type=Path, default=Path("output/batch_artifacts"))
    batch_parser.add_argument("--profile-dir", type=Path, default=Path("output/wb_profile"))

    merge_parser = subparsers.add_parser("merge-compass", help="Склеить batch_results.xlsx с выгрузкой Compass по ИНН")

    export_inn_parser = subparsers.add_parser("export-inn", help="Экспортировать уникальные ИНН из batch_results.xlsx для Compass")
    merge_registry_parser = subparsers.add_parser("merge-registries", help="Объединить 2 реестра ИНН с приоритетом первого")
    export_inn_parser.add_argument("--batch-results", required=True, type=Path)
    export_inn_parser.add_argument("--output", type=Path, default=Path("output/inn_for_compass.xlsx"))

    no_inn_parser = subparsers.add_parser("export-no-inn", help="Сохранить строки batch_results.xlsx без найденного ИНН")
    no_inn_parser.add_argument("--batch-results", required=True, type=Path)
    no_inn_parser.add_argument("--output", type=Path, default=Path("output/batch_no_inn.xlsx"))

    unmatched_parser = subparsers.add_parser("export-unmatched", help="Сохранить строки final_enriched.xlsx без совпадения в Compass")
    unmatched_parser.add_argument("--final-enriched", required=True, type=Path)
    unmatched_parser.add_argument("--output", type=Path, default=Path("output/compass_unmatched.xlsx"))
    merge_parser.add_argument("--batch-results", required=True, type=Path)
    merge_parser.add_argument("--compass", required=True, type=Path)
    merge_parser.add_argument("--output", type=Path, default=Path("output/final_enriched.xlsx"))
    merge_registry_parser.add_argument("--primary-registry", required=True, type=Path)
    merge_registry_parser.add_argument("--secondary-registry", required=True, type=Path)
    merge_registry_parser.add_argument("--output", type=Path, default=Path("output/inn_registry_merged.xlsx"))

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "analyze":
        summary = analyze_workbook(args.input, selected_sheets=args.selected_sheets)
        print(json.dumps(summary.model_dump(mode="json"), ensure_ascii=False, indent=2))
        return

    if args.command == "sample":
        args.output.parent.mkdir(parents=True, exist_ok=True)
        rows = extract_research_rows(args.input, limit=args.limit, selected_sheets=args.selected_sheets)
        save_research_sample(args.output, rows)
        print(f"Создан файл: {args.output}")
        print(f"Строк (уникальных по seller): {len(rows)}")
        rows_by_sheet = summarize_research_rows_by_sheet(rows)
        if rows_by_sheet:
            print("Распределение строк по листам:")
            for sheet_name, count in rows_by_sheet.items():
                print(f"  - {sheet_name}: {count}")
        return

    if args.command == "inspect-row":
        args.artifacts_dir.mkdir(parents=True, exist_ok=True)
        if args.profile_dir is not None:
            args.profile_dir.mkdir(parents=True, exist_ok=True)
            if args.artifacts_dir.resolve() == args.profile_dir.resolve():
                raise ValueError("Папка артефактов и папка профиля WB должны быть разными")

        research_row = read_research_row(args.input, row_number=args.row)
        result = inspect_product_row(
            row_number=args.row,
            research_row=research_row,
            artifacts_dir=args.artifacts_dir,
            headful=args.headful or args.profile_dir is not None,
            profile_dir=args.profile_dir,
        )
        print(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2))
        return

    if args.command == "batch-run":
        if args.artifacts_dir.resolve() == args.profile_dir.resolve():
            raise ValueError("Папка артефактов и папка профиля WB должны быть разными")

        args.artifacts_dir.mkdir(parents=True, exist_ok=True)
        args.profile_dir.mkdir(parents=True, exist_ok=True)
        args.output.parent.mkdir(parents=True, exist_ok=True)

        research_rows = read_research_rows_range(args.input, start_row=args.start_row, limit=args.limit)
        output_rows = []
        with BatchInspector(artifacts_dir=args.artifacts_dir, headful=True, profile_dir=args.profile_dir) as inspector:
            for offset, research_row in enumerate(research_rows, start=0):
                row_number = args.start_row + offset
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
                        profile_dir=args.profile_dir,
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
                print(json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2))

        save_batch_results(args.output, output_rows)
        print(f"Итоговый Excel сохранён: {args.output}")
        print(f"Обработано строк: {len(output_rows)}")
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

    if args.command == "merge-compass":
        summary = merge_batch_results_with_compass(
            batch_results_path=args.batch_results,
            compass_path=args.compass,
            output_path=args.output,
        )
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return

    if args.command == "merge-registries":
        summary = merge_inn_registry_files(
            primary_registry_path=args.primary_registry,
            secondary_registry_path=args.secondary_registry,
            output_path=args.output,
        )
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return


    if args.command == "export-inn":
        summary = export_unique_inn_list(args.batch_results, args.output)
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return

    if args.command == "export-no-inn":
        summary = export_batch_no_inn(args.batch_results, args.output)
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return

    if args.command == "export-unmatched":
        summary = export_compass_unmatched(args.final_enriched, args.output)
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return


    raise ValueError(f"Неизвестная команда: {args.command}")


if __name__ == "__main__":
    main()
