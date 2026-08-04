from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from wb_inn_extractor.gui import App


class _Var:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class _Root:
    def __init__(self):
        self.window_title = ""

    def after(self, _delay, _callback):
        return None

    def title(self, value):
        self.window_title = value


class GuiBatchTests(unittest.TestCase):
    def test_batch_progress_updates_header_and_window_title(self) -> None:
        app = App.__new__(App)
        app.root = _Root()
        app.progress_text_var = _Var("")
        app.progress_detail_var = _Var("")
        app.progress_percent_var = _Var("")
        app.progress_value_var = _Var(0.0)
        app._progress_done = 40
        app._progress_total = 200

        app._set_batch_progress_ui(
            "Пакетный прогон: 40/200 строк | текущая строка 42",
            "Прошло: 00:10:00 | ETA: 00:40:00 | Среднее: 15.0 сек/строка",
            20.0,
        )

        self.assertEqual(app.progress_percent_var.get(), "20.0%")
        self.assertEqual(app.progress_value_var.get(), 20.0)
        self.assertIn("ETA: 00:40:00", app.progress_detail_var.get())
        self.assertEqual(app.root.window_title, "WB INN Extractor — 40/200 (20.0%)")

    def test_batch_reads_known_seller_filter_toggle(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            app = App.__new__(App)
            app._require_sample_path = lambda: root / "research_sample.xlsx"
            app._required_profile_dir = lambda: root / "profile"
            app._append_log = lambda _message: None
            app._set_batch_progress = lambda **_kwargs: None
            app.row_var = _Var("2")
            app.batch_count_var = _Var("1")
            app.artifacts_dir_var = _Var(str(root / "artifacts"))
            app.batch_output_var = _Var(str(root / "batch_results.xlsx"))
            app.registry_path_var = _Var(str(root / "inn_registry.xlsx"))
            app.seller_history_path_var = _Var(str(root / "seller_history.xlsx"))
            app.sample_use_known_sellers_var = _Var(False)
            app.merge_batch_input_var = _Var("")
            app.root = _Root()

            with (
                patch("wb_inn_extractor.gui.load_known_seller_sources") as load_known,
                patch("wb_inn_extractor.gui.read_research_rows_range", return_value=[]),
                patch("wb_inn_extractor.gui.save_batch_results"),
                patch(
                    "wb_inn_extractor.gui.update_seller_history_from_batch_files",
                    return_value={"new_seller_added": 0, "history_rows": 0},
                ),
            ):
                app._action_batch()

            load_known.assert_not_called()


if __name__ == "__main__":
    unittest.main()
