#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Supabase 연결 테스트 스크립트
설정이 올바른지 확인하고 연결을 테스트합니다.
"""

import json
import requests
from pathlib import Path

def test_supabase_connection():
    """Supabase 연결 테스트"""
    print("=" * 60)
    print("🔍 Supabase 연결 테스트")
    print("=" * 60)
    
    # 1. 설정 파일 확인
    config_file = Path("supabase_config.json")
    
    if not config_file.exists():
        print("❌ supabase_config.json 파일이 없습니다.")
        print("📝 설정 파일을 생성합니다...")
        
        # 기본 설정 파일 생성
        default_config = {
            "supabase_url": "https://your-project-id.supabase.co",
            "supabase_key": "your-anon-key-here",
            "project_name": "excel_processor"
        }
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=2, ensure_ascii=False)
        
        print(f"✅ 설정 파일 생성: {config_file}")
        print("📋 설정 파일을 수정한 후 다시 실행해주세요.")
        return False
    
    # 2. 설정 파일 읽기
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        print("✅ 설정 파일 읽기 성공")
        print(f"📁 프로젝트: {config.get('project_name', 'N/A')}")
        
    except Exception as e:
        print(f"❌ 설정 파일 읽기 실패: {e}")
        return False
    
    # 3. 설정 값 검증
    supabase_url = config.get("supabase_url", "")
    supabase_key = config.get("supabase_key", "")
    
    if not supabase_url or supabase_url == "https://your-project-id.supabase.co":
        print("❌ Supabase URL이 설정되지 않았습니다.")
        print("📝 supabase_config.json에서 supabase_url을 수정해주세요.")
        return False
    
    if not supabase_key or supabase_key == "your-anon-key-here":
        print("❌ Supabase Key가 설정되지 않았습니다.")
        print("📝 supabase_config.json에서 supabase_key를 수정해주세요.")
        return False
    
    print("✅ 설정 값 검증 완료")
    
    # 4. Supabase 연결 테스트
    print("\n🔗 Supabase 연결 테스트 중...")
    
    headers = {
        "apikey": supabase_key,
        "Authorization": f"Bearer {supabase_key}",
        "Content-Type": "application/json"
    }
    
    try:
        # 간단한 연결 테스트 (테이블 목록 조회)
        response = requests.get(
            f"{supabase_url}/rest/v1/excel_files?select=id&limit=1",
            headers=headers,
            timeout=10
        )
        
        if response.status_code == 200:
            print("✅ Supabase 연결 성공!")
            print("🎉 데이터베이스에 정상적으로 접근할 수 있습니다.")
            
            # 테이블 존재 확인
            if response.json():
                print("✅ excel_files 테이블이 존재합니다.")
            else:
                print("⚠️  excel_files 테이블이 비어있습니다. (정상)")
            
            return True
            
        elif response.status_code == 401:
            print("❌ 인증 실패: API 키가 올바르지 않습니다.")
            print("📝 supabase_config.json에서 supabase_key를 확인해주세요.")
            return False
            
        elif response.status_code == 404:
            print("❌ 테이블을 찾을 수 없습니다.")
            print("📝 데이터베이스 스키마를 먼저 설정해주세요.")
            print("   supabase_setup.sql 파일을 SQL Editor에서 실행하세요.")
            return False
            
        else:
            print(f"❌ 연결 실패: HTTP {response.status_code}")
            print(f"📝 응답: {response.text}")
            return False
            
    except requests.exceptions.Timeout:
        print("❌ 연결 시간 초과: 네트워크 연결을 확인해주세요.")
        return False
        
    except requests.exceptions.ConnectionError:
        print("❌ 연결 오류: Supabase URL을 확인해주세요.")
        return False
        
    except Exception as e:
        print(f"❌ 예상치 못한 오류: {e}")
        return False

def create_sample_data():
    """샘플 데이터 생성 (선택사항)"""
    print("\n📊 샘플 데이터 생성 중...")
    
    try:
        with open("supabase_config.json", 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        supabase_url = config["supabase_url"]
        supabase_key = config["supabase_key"]
        
        headers = {
            "apikey": supabase_key,
            "Authorization": f"Bearer {supabase_key}",
            "Content-Type": "application/json"
        }
        
        # 샘플 파일 데이터 생성
        sample_data = {
            "filename": "test_sample.xlsx",
            "file_size": 1024,
            "file_data": "dGVzdCBkYXRh",  # "test data" in base64
            "project_name": "test_project",
            "status": "uploaded"
        }
        
        response = requests.post(
            f"{supabase_url}/rest/v1/excel_files",
            headers=headers,
            json=sample_data
        )
        
        if response.status_code == 201:
            print("✅ 샘플 데이터 생성 성공!")
            return True
        else:
            print(f"⚠️  샘플 데이터 생성 실패: {response.text}")
            return False
            
    except Exception as e:
        print(f"⚠️  샘플 데이터 생성 오류: {e}")
        return False

def main():
    """메인 실행 함수"""
    print("🌐 Supabase Excel 처리 시스템 연결 테스트")
    print("📧 제작자: charmleader@gmail.com")
    print()
    
    # 연결 테스트
    if test_supabase_connection():
        print("\n🎉 모든 테스트 통과!")
        print("✅ Supabase 설정이 완료되었습니다.")
        print("🚀 이제 cloud_excel_launcher.py를 실행할 수 있습니다.")
        
        # 샘플 데이터 생성 여부 확인
        create_sample = input("\n📊 샘플 데이터를 생성하시겠습니까? (y/n): ").strip().lower()
        if create_sample == 'y':
            create_sample_data()
    else:
        print("\n❌ 연결 테스트 실패!")
        print("📋 다음 단계를 확인해주세요:")
        print("1. Supabase 계정 생성 및 프로젝트 생성")
        print("2. supabase_config.json 파일 수정")
        print("3. 데이터베이스 스키마 설정 (supabase_setup.sql 실행)")
        print("4. 네트워크 연결 확인")
    
    print("\n" + "=" * 60)
    input("✨ Enter 키를 눌러 종료하세요...")

if __name__ == "__main__":
    main()

