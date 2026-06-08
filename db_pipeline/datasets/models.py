from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Dict, List, Optional


@dataclass(frozen=True)
class SourceTrace:
    clinic_code: str
    batch_id: str
    source_system: str
    source_file: str
    source_sheet: str
    source_row: int
    raw_row_hash: str


@dataclass(frozen=True)
class MemberRecord:
    trace: SourceTrace
    person_id: str
    name: str = ""
    birth_date: Optional[dt.date] = None
    sex: str = ""
    phone: str = ""
    mobile: str = ""
    address: str = ""
    case_category: str = ""
    quality_roster: str = ""
    multi_chronic_65: str = ""
    high_visit: str = ""
    chronic_mark: str = ""
    non_chronic_mark: str = ""
    same_clinic_previous_year: str = ""
    disease_pattern: str = ""
    ascvd: str = ""
    three_highs: str = ""
    hypertension: str = ""
    hyperlipidemia: str = ""
    hyperglycemia: str = ""


@dataclass(frozen=True)
class MonthlyClaimRecord:
    trace: SourceTrace
    person_id: str
    roc_year: int
    month: int
    visit_count: Decimal = Decimal("0")
    amount: Decimal = Decimal("0")
    last_visit_date: Optional[dt.date] = None


@dataclass(frozen=True)
class P4PCaseRecord:
    trace: SourceTrace
    person_id: str
    plan: str = ""
    status: str = ""
    enrolled_at: Optional[dt.date] = None


@dataclass(frozen=True)
class P4PTrackRecord:
    trace: SourceTrace
    person_id: str
    plan: str = ""
    last_tracked_at: Optional[dt.date] = None
    next_track_at: Optional[dt.date] = None
    overdue: str = ""


@dataclass(frozen=True)
class LabResultRecord:
    trace: SourceTrace
    person_id: str
    test_code: str
    result_value: str = ""
    tested_at: Optional[dt.date] = None


@dataclass(frozen=True)
class ScreeningRecord:
    trace: SourceTrace
    person_id: str
    screening_type: str
    screened_at: Optional[dt.date] = None


@dataclass(frozen=True)
class MemberSelectionRecord:
    trace: SourceTrace
    person_id: str
    selection_type: str


@dataclass
class DatasetBundle:
    members: List[MemberRecord] = field(default_factory=list)
    monthly_claims: List[MonthlyClaimRecord] = field(default_factory=list)
    p4p_cases: List[P4PCaseRecord] = field(default_factory=list)
    p4p_tracks: List[P4PTrackRecord] = field(default_factory=list)
    lab_results: List[LabResultRecord] = field(default_factory=list)
    screenings: List[ScreeningRecord] = field(default_factory=list)
    member_selections: List[MemberSelectionRecord] = field(default_factory=list)

    def counts(self) -> Dict[str, int]:
        return {
            "members": len(self.members),
            "monthly_claims": len(self.monthly_claims),
            "p4p_cases": len(self.p4p_cases),
            "p4p_tracks": len(self.p4p_tracks),
            "lab_results": len(self.lab_results),
            "screenings": len(self.screenings),
            "member_selections": len(self.member_selections),
        }
