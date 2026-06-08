import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from db_pipeline.config.models import ClinicConfig
from db_pipeline.parsers.調和 import TiaoheParser
from db_pipeline.raw import collect_raw_sources
from db_pipeline.storage import PostgresStagingWriter


class TiaoheParserTests(unittest.TestCase):
    def test_non_roster_claims_and_selections_are_retained(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            roster_id = "A123456789"
            non_roster_id = "B223456789"

            roster = Workbook()
            roster_ws = roster.active
            roster_ws.append(["ID", "個案類別", "姓名"])
            roster_ws.append([roster_id, "1", "會員甲"])
            roster_ws.append([roster_id, "1", "會員甲"])
            roster.save(root / "家醫照護名單.xlsx")

            count_dir = root / "R11440次數"
            count_dir.mkdir()
            claims = Workbook()
            claims_ws = claims.active
            claims_ws.append(["身分證號", "次數", "申請金額"])
            claims_ws.append([roster_id, 2, 300])
            claims_ws.append([non_roster_id, 1, 150])
            claims.save(count_dir / "11501.xlsx")

            (root / "115年預選自選會員A115.CSV").write_text(
                "身分證號,姓名\n"
                f"{roster_id},會員甲\n"
                f"{non_roster_id},一般病患乙\n",
                encoding="utf-8-sig",
            )

            result = TiaoheParser().parse(
                root,
                ClinicConfig(
                    clinic_code="3531142830",
                    clinic_name="蘆洲大愛診所",
                    source_system="tiaohe",
                ),
                "batch-test",
            )
            raw_result = collect_raw_sources(root)
            result.bundle.raw_source_files = raw_result.source_files
            result.bundle.raw_source_rows = raw_result.rows

            self.assertEqual(len(result.bundle.members), 1)
            self.assertEqual(len(result.bundle.monthly_claims), 2)
            self.assertEqual(
                {record.person_id for record in result.bundle.monthly_claims},
                {roster_id, non_roster_id},
            )
            self.assertEqual(len(result.bundle.member_selections), 3)
            self.assertEqual(
                result.coverage.unlinked_rows,
                {"monthly_claims": 1, "member_selections": 1},
            )
            self.assertFalse(result.coverage.unmatched_rows)
            self.assertEqual(len(result.bundle.raw_source_files), 3)
            self.assertEqual(len(result.bundle.raw_source_rows), 9)
            sql = PostgresStagingWriter()._build_sql(
                clinic_id=1,
                batch_id="00000000-0000-0000-0000-000000000001",
                bundle=result.bundle,
                source_system="tiaohe",
                source_root=str(root),
                requested_by="test",
            )
            self.assertIn("INSERT INTO raw.uploaded_rows", sql)
            self.assertIn("COPY staging.claims", sql)
            self.assertTrue(
                all(
                    "資料已保留" in issue.message
                    for issue in result.issues
                    if issue.severity == "warning"
                )
            )


if __name__ == "__main__":
    unittest.main()
