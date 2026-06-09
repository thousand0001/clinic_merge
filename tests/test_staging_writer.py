# -*- coding: utf-8 -*-
"""
PostgresStagingWriter 單元測試（不連 DB，以 mock 取代 _run_sql / _run_query）

P1：批次 INSERT 0 rows（clinic_id 不符）→ 後驗應 raise ValueError
P2：_build_sql() 拋例外 → status 應標 failed
"""
import unittest
from unittest.mock import MagicMock, call, patch

from db_pipeline.datasets.models import DatasetBundle
from db_pipeline.storage import PostgresStagingWriter
from db_pipeline.validation.models import ValidationIssue, ValidationReport


def _valid_report():
    return ValidationReport(issues=[], dataset_counts={})


def _error_report():
    return ValidationReport(
        issues=[ValidationIssue(severity="error", dataset="t", code="e", message="x")],
        dataset_counts={},
    )


class TestStagingWriterP1(unittest.TestCase):
    """P1：後驗擋住跨診所批次（ON CONFLICT WHERE 0 rows 情境）"""

    def test_post_check_raises_when_batch_owned_by_other_clinic(self):
        """INSERT 後查到 clinic_id 不符 → raise ValueError，不進 main transaction"""
        writer = PostgresStagingWriter()
        with patch("db_pipeline.storage._run_sql") as mock_sql, \
             patch("db_pipeline.storage._run_query") as mock_q:

            # pre-check：batch 不存在
            # INSERT：成功（mock 不拋）
            # post-check：batch 屬於別的診所（clinic_id=99）
            mock_q.side_effect = ["", "99"]  # pre-check → empty, post-check → 99

            with self.assertRaises(ValueError) as ctx:
                writer.stage(
                    clinic_id=1,
                    batch_id="aaaaaaaa-0000-0000-0000-000000000001",
                    bundle=DatasetBundle(),
                    validation_report=_valid_report(),
                )

            self.assertIn("99", str(ctx.exception))
            # main transaction 不應被執行（_run_sql 只呼叫過 batch INSERT，共 1 次）
            mock_sql.assert_called_once()

    def test_post_check_passes_when_batch_owned_by_same_clinic(self):
        """INSERT 後查到正確 clinic_id → 繼續執行 main transaction"""
        writer = PostgresStagingWriter()
        with patch("db_pipeline.storage._run_sql") as mock_sql, \
             patch("db_pipeline.storage._run_query") as mock_q:

            mock_q.side_effect = ["", "1"]  # pre-check → empty, post-check → 1
            mock_sql.return_value = ""

            writer.stage(
                clinic_id=1,
                batch_id="aaaaaaaa-0000-0000-0000-000000000002",
                bundle=DatasetBundle(),
                validation_report=_valid_report(),
            )

            # batch INSERT + main transaction = 2 calls
            self.assertEqual(mock_sql.call_count, 2)


class TestStagingWriterP2(unittest.TestCase):
    """P2：_build_sql() 例外 → batch status 標 failed"""

    def test_build_sql_exception_marks_batch_failed(self):
        """_build_sql 拋 RuntimeError → except 應呼叫 UPDATE status='failed'"""
        writer = PostgresStagingWriter()
        writer._build_sql = MagicMock(side_effect=RuntimeError("build error"))

        with patch("db_pipeline.storage._run_sql") as mock_sql, \
             patch("db_pipeline.storage._run_query") as mock_q:

            mock_q.side_effect = ["", "5"]  # pre / post check
            mock_sql.return_value = ""

            with self.assertRaises(RuntimeError):
                writer.stage(
                    clinic_id=5,
                    batch_id="bbbbbbbb-0000-0000-0000-000000000001",
                    bundle=DatasetBundle(),
                    validation_report=_valid_report(),
                )

            # 應有一次 UPDATE status='failed'
            failed_calls = [
                c for c in mock_q.call_args_list
                if "failed" in str(c)
            ]
            self.assertEqual(len(failed_calls), 1, "應呼叫一次 UPDATE status='failed'")

    def test_run_sql_exception_marks_batch_failed(self):
        """_run_sql（main transaction）拋例外 → status 也應標 failed"""
        writer = PostgresStagingWriter()

        with patch("db_pipeline.storage._run_sql") as mock_sql, \
             patch("db_pipeline.storage._run_query") as mock_q:

            mock_q.side_effect = ["", "5"]
            mock_sql.side_effect = [None, RuntimeError("db error")]  # batch INSERT OK, main fails

            with self.assertRaises(RuntimeError):
                writer.stage(
                    clinic_id=5,
                    batch_id="bbbbbbbb-0000-0000-0000-000000000002",
                    bundle=DatasetBundle(),
                    validation_report=_valid_report(),
                )

            failed_calls = [c for c in mock_q.call_args_list if "failed" in str(c)]
            self.assertEqual(len(failed_calls), 1)


class TestStagingWriterIsValidGuard(unittest.TestCase):
    """驗證失敗 → stage() 應在連 DB 前就 raise"""

    def test_raises_on_error_report_without_db(self):
        writer = PostgresStagingWriter()
        with patch("db_pipeline.storage._run_sql") as mock_sql, \
             patch("db_pipeline.storage._run_query") as mock_q:

            with self.assertRaises(ValueError) as ctx:
                writer.stage(
                    clinic_id=1,
                    batch_id="cccccccc-0000-0000-0000-000000000001",
                    bundle=DatasetBundle(),
                    validation_report=_error_report(),
                )

            self.assertIn("驗證報告含錯誤", str(ctx.exception))
            mock_sql.assert_not_called()
            mock_q.assert_not_called()


if __name__ == "__main__":
    unittest.main()
