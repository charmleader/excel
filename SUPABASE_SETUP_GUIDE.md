# 🌐 Supabase 설정 가이드

## 1단계: Supabase 계정 생성

### 1.1 Supabase 웹사이트 접속
- 브라우저에서 https://supabase.com 접속
- "Start your project" 또는 "Sign up" 클릭

### 1.2 계정 생성
- GitHub, Google, 또는 이메일로 계정 생성
- 이메일 인증 완료

## 2단계: 새 프로젝트 생성

### 2.1 프로젝트 생성
- 대시보드에서 "New Project" 클릭
- 프로젝트 이름: `excel-processor` (또는 원하는 이름)
- 데이터베이스 비밀번호 설정 (안전한 비밀번호 사용)
- 지역 선택: `Northeast Asia (Seoul)` (한국에서 가장 빠름)

### 2.2 프로젝트 생성 완료 대기
- 프로젝트 생성에 2-3분 소요
- 생성 완료 후 대시보드 접속

## 3단계: API 키 및 URL 확인

### 3.1 Settings 메뉴 접속
- 왼쪽 사이드바에서 "Settings" 클릭
- "API" 탭 선택

### 3.2 필요한 정보 복사
- **Project URL**: `https://your-project-id.supabase.co`
- **anon public key**: `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...`

## 4단계: 데이터베이스 스키마 설정

### 4.1 SQL Editor 접속
- 왼쪽 사이드바에서 "SQL Editor" 클릭
- "New query" 클릭

### 4.2 스키마 실행
- `supabase_setup.sql` 파일의 내용을 복사
- SQL Editor에 붙여넣기
- "Run" 버튼 클릭하여 실행

### 4.3 테이블 생성 확인
- 왼쪽 사이드바에서 "Table Editor" 클릭
- 다음 테이블들이 생성되었는지 확인:
  - `excel_files`
  - `process_excel`
  - `process_metadata`

## 5단계: 설정 파일 생성

### 5.1 supabase_config.json 생성
```json
{
  "supabase_url": "https://your-project-id.supabase.co",
  "supabase_key": "your-anon-key-here",
  "project_name": "excel_processor"
}
```

### 5.2 실제 값으로 교체
- `your-project-id`를 실제 프로젝트 ID로 교체
- `your-anon-key-here`를 실제 anon key로 교체

## 6단계: 연결 테스트

### 6.1 테스트 실행
```bash
python test_supabase_connection.py
```

### 6.2 성공 메시지 확인
- "✅ Supabase 연결 성공!" 메시지 확인
- 연결 실패 시 설정 재확인

## 7단계: 웹 인터페이스 실행

### 7.1 웹 앱 실행
```bash
python cloud_excel_launcher.py
```

### 7.2 브라우저 접속
- http://localhost:8501 접속
- Excel 파일 업로드 및 처리 테스트

## 🔧 문제 해결

### 연결 오류 시
1. URL과 Key가 정확한지 확인
2. 프로젝트가 활성화되어 있는지 확인
3. 방화벽 설정 확인

### 권한 오류 시
1. RLS 정책 확인
2. API 키 권한 확인
3. 테이블 생성 완료 확인

## 📞 지원

- 이메일: charmleader@gmail.com
- 문제 발생 시 설정 정보와 함께 문의

