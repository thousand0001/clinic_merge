from __future__ import annotations

import unittest
from unittest.mock import patch

from db_pipeline.providers.postgres import PostgresDataProvider


class DataProviderTests(unittest.TestCase):
    def test_postgres_provider_reads_standard_bundle(self) -> None:
        responses = iter([
            "1|hongcheng",
            (
                '{"row_no":2,"patient_id_normalized":"A123456789",'
                '"name":"會員甲","birth_date":"1980-01-02","sex":"",'
                '"phone":"","mobile":"","address":"","member_type":"2b",'
                '"disease_code":"4","ascvd":"","raw_data":{},'
                '"row_hash":"member"}'
            ),
            (
                '{"row_no":3,"patient_id_normalized":"A123456789",'
                '"service_date":"2025-01-03","roc_year":114,"month":1,'
                '"visit_count":2,"claim_amount":500,"raw_data":{},'
                '"row_hash":"claim"}'
            ),
            (
                '{"patient_id_normalized":"A123456789",'
                '"flag_type":"designated_114","raw_data":{},'
                '"row_hash":"flag"}'
            ),
            "",
            "",
            "",
        ])

        with patch(
            "db_pipeline.providers.postgres._run_query",
            side_effect=lambda _sql: next(responses),
        ):
            bundle = PostgresDataProvider(
                clinic_code="3501103076",
                batch_id="00000000-0000-0000-0000-000000000001",
            ).load_bundle()

        self.assertEqual(len(bundle.members), 1)
        self.assertEqual(len(bundle.monthly_claims), 1)
        self.assertEqual(len(bundle.member_selections), 1)
        self.assertEqual(bundle.monthly_claims[0].amount, 500)


if __name__ == "__main__":
    unittest.main()
