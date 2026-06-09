-- Migration 002: batch_source_files
-- 記錄每個批次引用的來源檔案（不可變追溯）
-- 2026-06-09

CREATE TABLE IF NOT EXISTS meta.batch_source_files (
    batch_id       UUID   NOT NULL REFERENCES meta.import_batches(batch_id) ON DELETE CASCADE,
    source_file_id BIGINT NOT NULL REFERENCES meta.source_files(source_file_id) ON DELETE CASCADE,
    PRIMARY KEY (batch_id, source_file_id)
);

COMMENT ON TABLE meta.batch_source_files IS
  '批次與來源檔案的多對多關聯。meta.source_files.batch_id 不再更新（記錄首次引入的批次）；'
  '此表讓每個批次都能查到它引用了哪些來源檔。';

-- 從現有 raw.uploaded_rows 回填歷史資料
-- （uploaded_rows 保有各自的 batch_id + source_file_id，比 source_files.batch_id 更可靠）
INSERT INTO meta.batch_source_files (batch_id, source_file_id)
SELECT DISTINCT ur.batch_id, ur.source_file_id
FROM raw.uploaded_rows ur
JOIN meta.import_batches ib ON ib.batch_id = ur.batch_id
WHERE ur.source_file_id IS NOT NULL
ON CONFLICT (batch_id, source_file_id) DO NOTHING;
