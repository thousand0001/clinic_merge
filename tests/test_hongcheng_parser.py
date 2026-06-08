from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from db_pipeline.config.models import ClinicConfig
from db_pipeline.parsers.宏誠 import HongchengParser, _parse_pdf_count_line


class HongchengParserTest(unittest.TestCase):
    def test_pdf_count_line(self) -> None:
        self.assertEqual(
            _parse_pdf_count_line("001234 王小明     7       3"),
            ("1234", "王小明", 7),
        )
        self.assertIsNone(_parse_pdf_count_line("病歷號碼 姓名 看診次數"))

    def test_csv_roster_and_non_roster_claim_are_retained(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "費用").mkdir()
            (root / "114年度需照護名單.csv").write_text(
                "院所ID,身分證號,生日,個案類別,疾病樣態\n"
                "'3501103076,A123456789,080/01/02,'2b,4\n",
                encoding="cp950",
            )
            wb = Workbook()
            ws = wb.active
            ws.append(["姓名", "病歷號", "看診日期", "申請金額", "身份證號"])
            ws.append(["王小明", "100", "114/01/03", 500, "B123456789"])
            wb.save(root / "費用" / "11401.xlsx")

            result = HongchengParser().parse(
                root,
                ClinicConfig(
                    clinic_code="3501103076",
                    clinic_name="德容聯合診所",
                    source_system="hongcheng",
                ),
                "batch",
            )

            self.assertEqual(len(result.bundle.members), 1)
            self.assertEqual(len(result.bundle.monthly_claims), 1)
            self.assertEqual(
                result.bundle.monthly_claims[0].person_id, "B123456789")
            self.assertEqual(
                result.coverage.unmatched_rows["monthly_claims"], 1)


if __name__ == "__main__":
    unittest.main()
