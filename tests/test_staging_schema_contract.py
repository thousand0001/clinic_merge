import unittest
from pathlib import Path


SCHEMA_PATH = (
    Path(__file__).resolve().parents[1]
    / "db_pipeline"
    / "storage"
    / "schema"
    / "001_staging_v1.sql"
)


class StagingSchemaContractTests(unittest.TestCase):
    def test_schema_contains_all_standard_datasets(self):
        sql = SCHEMA_PATH.read_text(encoding="utf-8")
        for table_name in (
            "members",
            "monthly_claims",
            "p4p_cases",
            "p4p_tracks",
            "lab_results",
            "screenings",
            "member_selections",
            "validation_issues",
        ):
            self.assertIn(
                f"CREATE TABLE IF NOT EXISTS staging_v1.{table_name}",
                sql,
            )

    def test_schema_is_transaction_wrapped_and_traceable(self):
        sql = SCHEMA_PATH.read_text(encoding="utf-8")
        self.assertTrue(sql.lstrip().startswith("BEGIN;"))
        self.assertTrue(sql.rstrip().endswith("COMMIT;"))
        for column_name in (
            "batch_id",
            "clinic_id",
            "source_file",
            "source_sheet",
            "source_row",
            "raw_row_hash",
        ):
            self.assertIn(column_name, sql)


if __name__ == "__main__":
    unittest.main()
