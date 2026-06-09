# -*- coding: utf-8 -*-
"""
member_builder 單元測試（不依賴 PostgreSQL）

驗證 build_from_bundle() 的核心欄位對應：
- AW:BJ（個案類別 ~ 高血糖）
- L/M/N/O（就診次數與費用平均）
- 名單旗標（designated / self_select / exclude_select）
- 檢驗、篩檢、P4P
"""
from __future__ import annotations

import datetime as dt
import unittest
from decimal import Decimal

from db_pipeline.datasets.models import (
    DatasetBundle,
    LabResultRecord,
    MemberRecord,
    MemberSelectionRecord,
    MonthlyClaimRecord,
    P4PCaseRecord,
    P4PTrackRecord,
    ScreeningRecord,
    SourceTrace,
)
from db_pipeline.output.member_builder import build_from_bundle


def _trace(pid: str = "A123456789") -> SourceTrace:
    return SourceTrace(
        clinic_code="0000000000",
        batch_id="test-batch",
        source_system="test",
        source_file="test.xlsx",
        source_sheet="Sheet1",
        source_row=1,
        raw_row_hash="abc",
    )


class TestAwBjFields(unittest.TestCase):
    def test_aw_bj_fields_mapped(self):
        bundle = DatasetBundle(
            members=[MemberRecord(
                trace=_trace("A123456789"),
                person_id="A123456789",
                name="王小明",
                birth_date=dt.date(1970, 1, 1),
                case_category="5",
                quality_roster="1",
                multi_chronic_65="0",
                high_visit="0",
                chronic_mark="1",
                non_chronic_mark="0",
                same_clinic_previous_year="1",
                disease_pattern="4",
                ascvd="b",
                three_highs="1",
                hypertension="0",
                hyperlipidemia="1",
                hyperglycemia="0",
            )],
            member_selections=[MemberSelectionRecord(
                trace=_trace("A123456789"),
                person_id="A123456789",
                selection_type="designated_114",
            )],
        )
        members = build_from_bundle(bundle)
        m = members["A123456789"]

        self.assertEqual(m["個案類別"], "5")
        self.assertEqual(m["論質名單"], "1")
        self.assertEqual(m["65歲以上多重慢性病註記"], "0")
        self.assertEqual(m["高診次註記"], "0")
        self.assertEqual(m["慢性病註記"], "1")
        self.assertEqual(m["非慢性病註記"], "0")
        self.assertEqual(m["與前一年家醫收案診所相同"], "1")
        self.assertEqual(m["疾病樣態"], "4")
        # disease_pattern="4" → "None"（非疾病）→ ASCVD 保留原始值 "b"
        self.assertEqual(m["ASCVD"], "b")
        self.assertEqual(m["三高"], "1")
        self.assertEqual(m["高血壓"], "0")
        self.assertEqual(m["高血脂"], "1")
        self.assertEqual(m["高血糖"], "0")


class TestLMNO(unittest.TestCase):
    def _claim(self, pid, year, month, count, amount):
        return MonthlyClaimRecord(
            trace=_trace(pid), person_id=pid,
            roc_year=year, month=month,
            visit_count=Decimal(str(count)),
            amount=Decimal(str(amount)),
        )

    def test_114_q1_and_115_count(self):
        bundle = DatasetBundle(
            monthly_claims=[
                self._claim("B123456789", 114, 1, 3, 30000),
                self._claim("B123456789", 114, 2, 4, 40000),
                self._claim("B123456789", 114, 3, 3, 30000),
                self._claim("B123456789", 114, 5, 2, 20000),   # Q2+
                self._claim("B123456789", 115, 1, 5, 50000),
                self._claim("B123456789", 115, 2, 3, 30000),
                self._claim("B123456789", 115, 3, 4, 40000),
            ],
        )
        members = build_from_bundle(bundle)
        m = members["B123456789"]

        # L：114年 Q1（月份 1-3）次數合計
        self.assertEqual(m["114_count_q1"], 10.0)
        # M：115年總次數
        self.assertEqual(m["115_count"], 12.0)
        # N：114全年總額（120000）/ 12
        self.assertAlmostEqual(m["114_avg_amount"], round(120000 / 12, 2))
        # O：115 全年總額（120000）/ distinct 115 月數（3）
        self.assertAlmostEqual(m["115_avg_amount"], round(120000 / 3, 2))

    def test_115_avg_includes_non_q1_months(self):
        """115_avg_amount 分子為全年總額（非僅 Q1），分母為有資料的月數。"""
        bundle = DatasetBundle(
            monthly_claims=[
                self._claim("Z123456789", 115, 1, 1, 100),  # Q1
                self._claim("Z123456789", 115, 5, 1, 100),  # 非 Q1
            ],
        )
        members = build_from_bundle(bundle)
        m = members["Z123456789"]
        # 全年 200 ÷ 2 月 = 100（非 Q1 僅算 Q1 的舊 bug 會得 50）
        self.assertAlmostEqual(m["115_avg_amount"], 100.0)

    def test_114_full_year_denominator(self):
        """114_avg_amount 分母固定為 12（全年），不限 Q1。"""
        claims = [
            MonthlyClaimRecord(
                trace=_trace(), person_id="C123456789",
                roc_year=114, month=m,
                visit_count=Decimal("1"), amount=Decimal("10000"),
            )
            for m in range(1, 13)
        ]
        bundle = DatasetBundle(monthly_claims=claims)
        members = build_from_bundle(bundle)
        m = members["C123456789"]
        self.assertAlmostEqual(m["114_avg_amount"], round(120000 / 12, 2))


class TestSelectionFlags(unittest.TestCase):
    def test_member_selection_flags(self):
        bundle = DatasetBundle(
            member_selections=[
                MemberSelectionRecord(trace=_trace("D123456789"),
                                      person_id="D123456789",
                                      selection_type="designated_114"),
                MemberSelectionRecord(trace=_trace("E123456789"),
                                      person_id="E123456789",
                                      selection_type="self_select"),
                MemberSelectionRecord(trace=_trace("F123456789"),
                                      person_id="F123456789",
                                      selection_type="exclude_select"),
            ],
        )
        members = build_from_bundle(bundle)
        self.assertEqual(members["D123456789"].get("designated"), "✔")
        self.assertEqual(members["D123456789"].get("is_114_member"), "✔")
        self.assertEqual(members["E123456789"].get("self_select"), "✔")
        self.assertEqual(members["F123456789"].get("exclude_select"), "✔")


class TestScreenings(unittest.TestCase):
    def test_screenings_mapped(self):
        bundle = DatasetBundle(
            screenings=[
                ScreeningRecord(trace=_trace(), person_id="G123456789",
                                screening_type="成人健檢",
                                screened_at=dt.date(2025, 3, 1)),
                ScreeningRecord(trace=_trace(), person_id="G123456789",
                                screening_type="老人流感",
                                screened_at=dt.date(2025, 10, 1)),
            ],
        )
        members = build_from_bundle(bundle)
        m = members["G123456789"]
        self.assertEqual(m["adult"], dt.date(2025, 3, 1))
        self.assertEqual(m["flu"], dt.date(2025, 10, 1))
        self.assertIsNone(m.get("pap"))


class TestLabResults(unittest.TestCase):
    def test_lab_results_mapped(self):
        bundle = DatasetBundle(
            lab_results=[
                LabResultRecord(trace=_trace(), person_id="H123456789",
                                test_code="HbA1c", result_value="7.2",
                                tested_at=dt.date(2025, 4, 1)),
                LabResultRecord(trace=_trace(), person_id="H123456789",
                                test_code="LDL", result_value="95",
                                tested_at=dt.date(2025, 5, 1)),
            ],
        )
        members = build_from_bundle(bundle)
        m = members["H123456789"]
        self.assertEqual(m["hba1c"], "7.2")
        self.assertEqual(m["ldl"], "95")
        self.assertIsNone(m.get("uacr"))


class TestP4P(unittest.TestCase):
    def test_p4p_case_and_track(self):
        bundle = DatasetBundle(
            p4p_cases=[P4PCaseRecord(
                trace=_trace(), person_id="I123456789",
                plan="DM", status="收案中",
                enrolled_at=dt.date(2025, 1, 1),
            )],
            p4p_tracks=[P4PTrackRecord(
                trace=_trace(), person_id="I123456789",
                plan="DM",
                last_tracked_at=dt.date(2025, 3, 1),
                next_track_at=dt.date(2025, 6, 1),
                overdue="",
            )],
        )
        members = build_from_bundle(bundle)
        m = members["I123456789"]
        self.assertEqual(m["p4p_plan"], "DM")
        self.assertEqual(m["p4p_status"], "收案中")
        self.assertEqual(m["p4p_enroll_date"], dt.date(2025, 1, 1))
        self.assertEqual(m["p4p_last_track"], dt.date(2025, 3, 1))
        self.assertEqual(m["p4p_next_track"], dt.date(2025, 6, 1))


class TestDiseaseCode(unittest.TestCase):
    def test_disease_code_inferred(self):
        bundle = DatasetBundle(
            members=[MemberRecord(
                trace=_trace(), person_id="J123456789",
                disease_pattern="1",  # → DM
            )],
        )
        members = build_from_bundle(bundle)
        m = members["J123456789"]
        self.assertEqual(m["disease_code"], "DM")
        # DM 且 disease_code != "None" → ASCVD 強制為 "1" → disease_class 帶 +ASCVD
        self.assertEqual(m["disease_class"], "DM+ASCVD")


class TestMerge(unittest.TestCase):
    def test_merge_does_not_overwrite(self):
        bundle = DatasetBundle(
            members=[
                MemberRecord(trace=_trace(), person_id="K123456789", name="第一筆"),
                MemberRecord(trace=_trace(), person_id="K123456789", name="第二筆"),
            ],
        )
        members = build_from_bundle(bundle)
        self.assertEqual(members["K123456789"]["name"], "第一筆")


if __name__ == "__main__":
    unittest.main()
