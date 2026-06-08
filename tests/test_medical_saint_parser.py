import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from db_pipeline.config.models import load_clinic_config
from db_pipeline.parsers.醫聖 import MedicalSaintParser
from db_pipeline.validation.validator import validate_bundle


class MedicalSaintParserTests(unittest.TestCase):
    def _save_workbook(self, path: Path, title: str, rows: list[list[object]]) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = title
        for row in rows:
            ws.append(row)
        wb.save(path)

    def test_parses_junan_p4p_screening_health_and_non_roster_claims(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            temp_root = Path(tmp)
            root = temp_root / "source"
            root.mkdir()
            config_path = temp_root / "clinic.json"
            config_path.write_text(
                json.dumps(
                    {
                        "clinic_code": "3501186011",
                        "clinic_name": "鈞安診所",
                        "source_system": "medical_saint",
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            self._save_workbook(
                root / "114年-鈞安診所(ASCVD)需照護名單.xlsx",
                "照護名單",
                [
                    ["ID", "姓名", "生日", "個案類別", "ASCVD"],
                    ["A123456789", "王小明", "1980/01/02", "一般", "A"],
                ],
            )
            month_dir = root / "R11440"
            month_dir.mkdir()
            (month_dir / "P11501.txt").write_bytes(
                (
                    "病歷號,姓名,身分證,看診日,生日,申請金額\r\n"
                    "001,王小明,A123456789,115.01.10,69.01.02,300\r\n"
                    "002,李小華,B223456789,115.01.11,70.02.03,500\r\n"
                ).encode("cp950")
            )
            self._save_workbook(
                root / "複本 sm_3501186011_cliP4p收案.xlsx",
                "P4pCase",
                [
                    ["家醫收案會員ID", "P4P收案計畫", "收案狀態"],
                    ["A123456789", "糖尿病_DM", "本院收案"],
                ],
            )
            self._save_workbook(
                root / "成人.xlsx",
                "IndexNo",
                [
                    ["指標名稱", "ID", "生日", "姓名", "最後篩檢日期"],
                    ["成人預防保健檢查", "A123456789", "1980/01/02", "王小明", "2025/03/04"],
                    ["成人預防保健檢查", "B223456789", "1981/02/03", "李小華", "-"],
                ],
            )
            self._save_workbook(
                root / "鈞安個案健康管理列表.xlsx",
                "HealthCase",
                [
                    [
                        "家醫收案會員ID",
                        "最近一次HbA1c檢查結果(%)",
                        "最近一次HbA1c檢查日期",
                        "最近一次LDL檢查結果(mg/dL)",
                        "最近一次LDL檢查日期",
                        "最近一次UACR檢查結果(mg/gm)",
                        "最近一次UACR檢查日期",
                    ],
                    ["A123456789", "6.2", "2026/03/26", "77", "2026/03/26", "-", "-"],
                ],
            )

            result = MedicalSaintParser().parse(
                root,
                load_clinic_config(config_path),
                "batch-test",
            )
            report = validate_bundle(result.bundle)

            self.assertTrue(report.is_valid)
            self.assertEqual(result.coverage.discovered_files, 5)
            self.assertEqual(result.coverage.parsed_files, 5)
            expected_counts = {
                "members": 1,
                "monthly_claims": 2,
                "p4p_cases": 1,
                "p4p_tracks": 0,
                "lab_results": 2,
                "screenings": 1,
                "member_selections": 1,
            }
            actual_counts = result.bundle.counts()
            for dataset, expected in expected_counts.items():
                self.assertEqual(actual_counts[dataset], expected)
            self.assertEqual(result.coverage.unmatched_rows["monthly_claims"], 1)
            self.assertEqual(
                {record.test_code for record in result.bundle.lab_results},
                {"HbA1c", "LDL"},
            )
            self.assertEqual(
                result.bundle.screenings[0].screening_type,
                "成人健檢",
            )
            self.assertFalse(result.coverage.skipped_files)


if __name__ == "__main__":
    unittest.main()
