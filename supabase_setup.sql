-- Supabase 데이터베이스 스키마 설정
-- Excel 파일 처리 시스템을 위한 테이블 생성

-- 1. Excel 파일 정보 테이블
CREATE TABLE IF NOT EXISTS excel_files (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    filename VARCHAR(255) NOT NULL,
    file_size INTEGER NOT NULL,
    file_data TEXT NOT NULL, -- Base64 인코딩된 파일 데이터
    project_name VARCHAR(100) DEFAULT 'default',
    upload_time TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    status VARCHAR(50) DEFAULT 'uploaded',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 2. 파일 처리 작업 테이블
CREATE TABLE IF NOT EXISTS process_excel (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    file_id UUID REFERENCES excel_files(id) ON DELETE CASCADE,
    processing_options JSONB DEFAULT '{}',
    status VARCHAR(50) DEFAULT 'pending', -- pending, processing, completed, failed
    start_time TIMESTAMP WITH TIME ZONE,
    end_time TIMESTAMP WITH TIME ZONE,
    result_file_data TEXT, -- Base64 인코딩된 결과 파일
    error_message TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 3. 처리 결과 메타데이터 테이블
CREATE TABLE IF NOT EXISTS process_metadata (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    process_id UUID REFERENCES process_excel(id) ON DELETE CASCADE,
    total_rows INTEGER,
    total_columns INTEGER,
    processed_sheets INTEGER,
    processing_time_seconds INTEGER,
    metadata JSONB DEFAULT '{}',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 4. 인덱스 생성 (성능 최적화)
CREATE INDEX IF NOT EXISTS idx_excel_files_project_name ON excel_files(project_name);
CREATE INDEX IF NOT EXISTS idx_excel_files_status ON excel_files(status);
CREATE INDEX IF NOT EXISTS idx_process_excel_file_id ON process_excel(file_id);
CREATE INDEX IF NOT EXISTS idx_process_excel_status ON process_excel(status);
CREATE INDEX IF NOT EXISTS idx_process_metadata_process_id ON process_metadata(process_id);

-- 5. RLS (Row Level Security) 정책 설정
ALTER TABLE excel_files ENABLE ROW LEVEL SECURITY;
ALTER TABLE process_excel ENABLE ROW LEVEL SECURITY;
ALTER TABLE process_metadata ENABLE ROW LEVEL SECURITY;

-- 6. 기본 RLS 정책 (모든 사용자가 자신의 데이터만 접근 가능)
CREATE POLICY "Users can view their own files" ON excel_files
    FOR SELECT USING (auth.uid()::text = project_name);

CREATE POLICY "Users can insert their own files" ON excel_files
    FOR INSERT WITH CHECK (auth.uid()::text = project_name);

CREATE POLICY "Users can update their own files" ON excel_files
    FOR UPDATE USING (auth.uid()::text = project_name);

CREATE POLICY "Users can delete their own files" ON excel_files
    FOR DELETE USING (auth.uid()::text = project_name);

-- 7. 함수 생성 (자동 업데이트 시간)
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ language 'plpgsql';

-- 8. 트리거 생성
CREATE TRIGGER update_excel_files_updated_at 
    BEFORE UPDATE ON excel_files 
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_process_excel_updated_at 
    BEFORE UPDATE ON process_excel 
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- 9. 뷰 생성 (처리 상태 조회용)
CREATE OR REPLACE VIEW process_status_view AS
SELECT 
    pe.id as process_id,
    ef.filename,
    ef.project_name,
    pe.status,
    pe.start_time,
    pe.end_time,
    CASE 
        WHEN pe.status = 'completed' THEN 
            EXTRACT(EPOCH FROM (pe.end_time - pe.start_time))::INTEGER
        ELSE NULL 
    END as processing_time_seconds,
    pm.total_rows,
    pm.total_columns,
    pm.processed_sheets
FROM process_excel pe
JOIN excel_files ef ON pe.file_id = ef.id
LEFT JOIN process_metadata pm ON pe.id = pm.process_id;

-- 10. 샘플 데이터 삽입 (테스트용)
INSERT INTO excel_files (filename, file_size, file_data, project_name, status) 
VALUES 
    ('sample1.xlsx', 1024, 'dGVzdCBkYXRh', 'test_project', 'uploaded'),
    ('sample2.xlsx', 2048, 'dGVzdCBkYXRh', 'test_project', 'uploaded')
ON CONFLICT DO NOTHING;

-- 11. 통계 함수 생성
CREATE OR REPLACE FUNCTION get_project_stats(project_name_param VARCHAR)
RETURNS TABLE (
    total_files BIGINT,
    total_size BIGINT,
    completed_processes BIGINT,
    failed_processes BIGINT
) AS $$
BEGIN
    RETURN QUERY
    SELECT 
        COUNT(ef.id) as total_files,
        COALESCE(SUM(ef.file_size), 0) as total_size,
        COUNT(CASE WHEN pe.status = 'completed' THEN 1 END) as completed_processes,
        COUNT(CASE WHEN pe.status = 'failed' THEN 1 END) as failed_processes
    FROM excel_files ef
    LEFT JOIN process_excel pe ON ef.id = pe.file_id
    WHERE ef.project_name = project_name_param;
END;
$$ LANGUAGE plpgsql;

-- 12. 정리 함수 생성 (오래된 파일 삭제)
CREATE OR REPLACE FUNCTION cleanup_old_files(days_old INTEGER DEFAULT 30)
RETURNS INTEGER AS $$
DECLARE
    deleted_count INTEGER;
BEGIN
    DELETE FROM excel_files 
    WHERE created_at < NOW() - INTERVAL '1 day' * days_old
    AND status = 'completed';
    
    GET DIAGNOSTICS deleted_count = ROW_COUNT;
    RETURN deleted_count;
END;
$$ LANGUAGE plpgsql;

-- 13. 권한 설정
GRANT USAGE ON SCHEMA public TO anon, authenticated;
GRANT ALL ON ALL TABLES IN SCHEMA public TO anon, authenticated;
GRANT ALL ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;
GRANT EXECUTE ON ALL FUNCTIONS IN SCHEMA public TO anon, authenticated;

-- 14. 완료 메시지
DO $$
BEGIN
    RAISE NOTICE 'Supabase 데이터베이스 스키마 설정이 완료되었습니다!';
    RAISE NOTICE 'Excel 파일 처리 시스템을 사용할 준비가 되었습니다.';
END $$;
