import csv
import importlib.util
import os
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa


MODULE_PATH = Path(__file__).resolve().parents[1] / "run_retail_dashboard.py"
SPEC = importlib.util.spec_from_file_location("run_retail_dashboard", MODULE_PATH)
module = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(module)


class FinanceMonthTests(unittest.TestCase):
    columns = ["PERIOD", "CATEGORY", "CHANNEL", "SPEND"]

    def test_english_month_labels_are_compared_chronologically(self):
        self.assertEqual(module.parse_period_month("MAY 2026"), date(2026, 5, 1))
        self.assertEqual(module.parse_period_month("JUN 2026"), date(2026, 6, 1))
        self.assertEqual(module.parse_period_month("June 2026"), date(2026, 6, 1))
        self.assertGreater(
            module.parse_period_month("JUN 2026"),
            module.parse_period_month("MAY 2026"),
        )

    def test_same_month_leaves_file_untouched(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "Marketing_Spend_2026YTD.csv"
            original = b"\xef\xbb\xbfPERIOD,CATEGORY,CHANNEL,SPEND\r\nMay 2026,Promos,Walmart,10\r\n"
            path.write_bytes(original)
            changed = module.refresh_finance_if_new_month(
                self.columns,
                [("May 2026", "Promos", "Walmart", Decimal("99.00"))],
                path,
            )
            self.assertFalse(changed)
            self.assertEqual(path.read_bytes(), original)

    def test_new_month_replaces_full_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "Marketing_Spend_2026YTD.csv"
            path.write_text(
                "PERIOD,CATEGORY,CHANNEL,SPEND\nMay 2026,Promos,Walmart,10\n",
                encoding="utf-8-sig",
            )
            rows = [
                ("MAY 2026", "Promos", "Walmart", Decimal("10.00")),
                ("JUN 2026", "Promos", "Walmart", Decimal("12.34")),
            ]
            changed = module.refresh_finance_if_new_month(self.columns, rows, path)
            self.assertTrue(changed)
            self.assertEqual(module.newest_local_finance_month(path), date(2026, 6, 1))
            with path.open(encoding="utf-8-sig", newline="") as handle:
                saved = list(csv.DictReader(handle))
            self.assertEqual(len(saved), 2)
            self.assertEqual(saved[-1]["SPEND"], "12.34")

    def test_sql_loader_rejects_non_select(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "bad.sql"
            path.write_text("DELETE FROM important_table;", encoding="utf-8")
            with self.assertRaises(ValueError):
                module.load_select_sql(path)


class PrivateKeyTests(unittest.TestCase):
    def test_private_key_is_loaded_directly_from_env_with_escaped_newlines(self):
        private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        pem = private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        ).decode("utf-8")
        inline_pem = pem.strip().replace("\n", "\\n")

        with patch.dict(os.environ, {"SNOWFLAKE_PRIVATE_KEY": inline_pem}, clear=False):
            der = module._private_key_bytes()

        loaded = serialization.load_der_private_key(der, password=None)
        self.assertEqual(
            loaded.private_numbers().public_numbers,
            private_key.private_numbers().public_numbers,
        )


class ProjectLayoutTests(unittest.TestCase):
    def test_runtime_paths_match_handoff_layout(self):
        self.assertEqual(module.SQL_DIR, module.ROOT / "sql")
        self.assertEqual(module.BUILDER, module.ROOT / "scripts" / "build_outputs.py")
        self.assertEqual(
            module.FINANCE_CSV,
            module.ROOT / "inputs" / "finance" / "Marketing_Spend_2026YTD.csv",
        )
        for path in (
            module.WEEKLY_SQL,
            module.FINANCE_SQL,
            module.BUILDER,
            module.FINANCE_CSV,
        ):
            self.assertTrue(path.exists(), path)

    def test_sql_wrappers_use_only_dashboard_views(self):
        weekly_sql = module.load_select_sql(module.WEEKLY_SQL).upper()
        finance_sql = module.load_select_sql(module.FINANCE_SQL).upper()

        self.assertIn(
            "DATA_MART.RETAIL_DASHBOARD.WEEKLY_RETAIL_ACTUALS",
            weekly_sql,
        )
        self.assertIn(
            "DATA_MART.RETAIL_DASHBOARD.MARKETING_SPEND_YTD",
            finance_sql,
        )
        self.assertNotIn("DATA_MART.FINANCE", weekly_sql)
        self.assertNotIn("DATA_MART.FINANCE", finance_sql)


if __name__ == "__main__":
    unittest.main()
