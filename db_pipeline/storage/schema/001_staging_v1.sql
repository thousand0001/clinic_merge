BEGIN;

CREATE SCHEMA IF NOT EXISTS staging_v1;

CREATE TABLE IF NOT EXISTS staging_v1.members (
  staging_member_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  clinic_code TEXT NOT NULL,
  source_system TEXT NOT NULL,
  person_id TEXT NOT NULL,
  name TEXT,
  birth_date DATE,
  sex TEXT,
  phone TEXT,
  mobile TEXT,
  address TEXT,
  case_category TEXT,
  quality_roster TEXT,
  multi_chronic_65 TEXT,
  high_visit TEXT,
  chronic_mark TEXT,
  non_chronic_mark TEXT,
  same_clinic_previous_year TEXT,
  disease_pattern TEXT,
  ascvd TEXT,
  three_highs TEXT,
  hypertension TEXT,
  hyperlipidemia TEXT,
  hyperglycemia TEXT,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  raw_row_hash TEXT NOT NULL,
  staged_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (batch_id, source_file, source_sheet, source_row, person_id)
);

CREATE INDEX IF NOT EXISTS idx_staging_members_batch
  ON staging_v1.members(batch_id);
CREATE INDEX IF NOT EXISTS idx_staging_members_clinic_person
  ON staging_v1.members(clinic_id, person_id);

CREATE TABLE IF NOT EXISTS staging_v1.monthly_claims (
  staging_claim_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  clinic_code TEXT NOT NULL,
  source_system TEXT NOT NULL,
  person_id TEXT NOT NULL,
  roc_year SMALLINT NOT NULL CHECK (roc_year IN (114, 115)),
  month SMALLINT NOT NULL CHECK (month BETWEEN 1 AND 12),
  visit_count NUMERIC(14, 2) NOT NULL DEFAULT 0,
  amount NUMERIC(16, 2) NOT NULL DEFAULT 0,
  last_visit_date DATE,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  raw_row_hash TEXT NOT NULL,
  staged_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (batch_id, source_file, source_sheet, source_row, person_id, roc_year, month)
);

CREATE INDEX IF NOT EXISTS idx_staging_claims_batch
  ON staging_v1.monthly_claims(batch_id);
CREATE INDEX IF NOT EXISTS idx_staging_claims_clinic_person_month
  ON staging_v1.monthly_claims(clinic_id, person_id, roc_year, month);

CREATE TABLE IF NOT EXISTS staging_v1.p4p_cases (
  staging_p4p_case_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  clinic_code TEXT NOT NULL,
  source_system TEXT NOT NULL,
  person_id TEXT NOT NULL,
  plan TEXT,
  status TEXT,
  enrolled_at DATE,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  raw_row_hash TEXT NOT NULL,
  staged_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (batch_id, source_file, source_sheet, source_row, person_id)
);

CREATE INDEX IF NOT EXISTS idx_staging_p4p_cases_batch
  ON staging_v1.p4p_cases(batch_id);
CREATE INDEX IF NOT EXISTS idx_staging_p4p_cases_clinic_person
  ON staging_v1.p4p_cases(clinic_id, person_id);

CREATE TABLE IF NOT EXISTS staging_v1.p4p_tracks (
  staging_p4p_track_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  clinic_code TEXT NOT NULL,
  source_system TEXT NOT NULL,
  person_id TEXT NOT NULL,
  plan TEXT,
  last_tracked_at DATE,
  next_track_at DATE,
  overdue TEXT,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  raw_row_hash TEXT NOT NULL,
  staged_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (batch_id, source_file, source_sheet, source_row, person_id)
);

CREATE INDEX IF NOT EXISTS idx_staging_p4p_tracks_batch
  ON staging_v1.p4p_tracks(batch_id);
CREATE INDEX IF NOT EXISTS idx_staging_p4p_tracks_clinic_person
  ON staging_v1.p4p_tracks(clinic_id, person_id);

CREATE TABLE IF NOT EXISTS staging_v1.lab_results (
  staging_lab_result_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  clinic_code TEXT NOT NULL,
  source_system TEXT NOT NULL,
  person_id TEXT NOT NULL,
  test_code TEXT NOT NULL,
  result_value TEXT,
  tested_at DATE,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  raw_row_hash TEXT NOT NULL,
  staged_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (batch_id, source_file, source_sheet, source_row, person_id, test_code)
);

CREATE INDEX IF NOT EXISTS idx_staging_lab_results_batch
  ON staging_v1.lab_results(batch_id);
CREATE INDEX IF NOT EXISTS idx_staging_lab_results_clinic_person_test
  ON staging_v1.lab_results(clinic_id, person_id, test_code, tested_at);

CREATE TABLE IF NOT EXISTS staging_v1.screenings (
  staging_screening_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  clinic_code TEXT NOT NULL,
  source_system TEXT NOT NULL,
  person_id TEXT NOT NULL,
  screening_type TEXT NOT NULL,
  screened_at DATE,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  raw_row_hash TEXT NOT NULL,
  staged_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (
    batch_id,
    source_file,
    source_sheet,
    source_row,
    person_id,
    screening_type
  )
);

CREATE INDEX IF NOT EXISTS idx_staging_screenings_batch
  ON staging_v1.screenings(batch_id);
CREATE INDEX IF NOT EXISTS idx_staging_screenings_clinic_person_type
  ON staging_v1.screenings(clinic_id, person_id, screening_type, screened_at);

CREATE TABLE IF NOT EXISTS staging_v1.member_selections (
  staging_selection_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  clinic_code TEXT NOT NULL,
  source_system TEXT NOT NULL,
  person_id TEXT NOT NULL,
  selection_type TEXT NOT NULL CHECK (
    selection_type IN ('designated_114', 'self_selected_115', 'excluded_115')
  ),
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  raw_row_hash TEXT NOT NULL,
  staged_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (
    batch_id,
    source_file,
    source_sheet,
    source_row,
    person_id,
    selection_type
  )
);

CREATE INDEX IF NOT EXISTS idx_staging_member_selections_batch
  ON staging_v1.member_selections(batch_id);
CREATE INDEX IF NOT EXISTS idx_staging_member_selections_clinic_person
  ON staging_v1.member_selections(clinic_id, person_id, selection_type);

CREATE TABLE IF NOT EXISTS staging_v1.validation_issues (
  validation_issue_id BIGSERIAL PRIMARY KEY,
  batch_id UUID NOT NULL REFERENCES meta.import_batches(batch_id),
  clinic_id BIGINT NOT NULL REFERENCES meta.clinics(clinic_id),
  severity TEXT NOT NULL CHECK (severity IN ('error', 'warning', 'info')),
  dataset_name TEXT NOT NULL,
  issue_code TEXT NOT NULL,
  message TEXT NOT NULL,
  source_file TEXT,
  source_sheet TEXT,
  source_row INTEGER,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_staging_validation_issues_batch
  ON staging_v1.validation_issues(batch_id, severity);

COMMIT;
