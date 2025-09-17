#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
클라우드 기반 Excel 처리 시스템 런처
Supabase 설정 및 웹 인터페이스 제공
"""

import os
import sys
import json
import webbrowser
import time
from pathlib import Path
import subprocess
import streamlit as st

def check_supabase_config():
    """Supabase 설정 확인"""
    config_file = Path("supabase_config.json")
    
    if not config_file.exists():
        print("❌ Supabase 설정 파일이 없습니다.")
        print("📋 설정 방법:")
        print("1. https://supabase.com 에서 계정 생성")
        print("2. 새 프로젝트 생성")
        print("3. Settings > API에서 URL과 Key 복사")
        print("4. supabase_config.json 파일 생성")
        
        # 설정 파일 템플릿 생성
        template = {
            "supabase_url": "https://your-project.supabase.co",
            "supabase_key": "your-anon-key",
            "project_name": "excel_processor"
        }
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(template, f, indent=2, ensure_ascii=False)
        
        print(f"📝 설정 템플릿이 생성되었습니다: {config_file}")
        print("설정을 완료한 후 다시 실행해주세요.")
        return False
    
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        if config["supabase_url"] == "https://your-project.supabase.co":
            print("❌ Supabase 설정이 완료되지 않았습니다.")
            print(f"📝 설정 파일을 수정해주세요: {config_file}")
            return False
        
        print("✅ Supabase 설정 확인 완료")
        return True
        
    except Exception as e:
        print(f"❌ 설정 파일 읽기 오류: {e}")
        return False

def create_web_interface():
    """웹 인터페이스 생성"""
    web_content = '''
import streamlit as st
import json
import base64
import pandas as pd
from excel_cloud_processor import ExcelCloudProcessor
from pathlib import Path

# 페이지 설정
st.set_page_config(
    page_title="🌐 클라우드 Excel 처리기",
    page_icon="🌐",
    layout="wide"
)

# 제목
st.title("🌐 클라우드 기반 Excel 파일 처리 시스템")
st.markdown("---")

# Supabase 설정 로드
@st.cache_data
def load_config():
    try:
        with open("supabase_config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return None

config = load_config()

if not config:
    st.error("❌ Supabase 설정 파일을 찾을 수 없습니다.")
    st.info("supabase_config.json 파일을 생성하고 설정을 완료해주세요.")
    st.stop()

# 클라이언트 초기화
@st.cache_resource
def get_processor():
    return ExcelCloudProcessor(config["supabase_url"], config["supabase_key"])

processor = get_processor()

# 사이드바
with st.sidebar:
    st.header("📋 메뉴")
    menu = st.selectbox(
        "기능 선택",
        ["📤 파일 업로드", "🔄 파일 처리", "📥 결과 다운로드", "📊 프로젝트 현황"]
    )

# 메인 컨텐츠
if menu == "📤 파일 업로드":
    st.header("📤 Excel 파일 업로드")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_file = st.file_uploader(
            "Excel 파일을 선택하세요",
            type=['xlsx', 'xls'],
            help="처리할 Excel 파일을 업로드합니다."
        )
    
    with col2:
        project_name = st.text_input(
            "프로젝트 이름",
            value=config.get("project_name", "default"),
            help="파일을 그룹화할 프로젝트 이름"
        )
    
    if uploaded_file and st.button("🚀 업로드", type="primary"):
        with st.spinner("파일을 업로드하는 중..."):
            # 임시 파일로 저장
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # 클라우드에 업로드
            result = processor.upload_excel_file(temp_path, project_name)
            
            # 임시 파일 삭제
            os.remove(temp_path)
            
            if result["success"]:
                st.success(f"✅ 업로드 완료! 파일 ID: {result['file_id']}")
                st.session_state['uploaded_file_id'] = result['file_id']
            else:
                st.error(f"❌ 업로드 실패: {result['error']}")

elif menu == "🔄 파일 처리":
    st.header("🔄 Excel 파일 처리")
    
    file_id = st.text_input(
        "파일 ID",
        value=st.session_state.get('uploaded_file_id', ''),
        help="처리할 파일의 ID를 입력하세요"
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("처리 옵션")
        merge_sheets = st.checkbox("시트 통합", value=True)
        remove_empty = st.checkbox("빈 행 제거", value=True)
        add_metadata = st.checkbox("메타데이터 추가", value=True)
    
    with col2:
        st.subheader("정렬 옵션")
        sort_data = st.checkbox("데이터 정렬", value=True)
        sort_by = st.selectbox("정렬 기준", ["파일명", "시트명", "행수"])
    
    if st.button("🔄 처리 시작", type="primary"):
        if not file_id:
            st.error("❌ 파일 ID를 입력해주세요.")
        else:
            processing_options = {
                "merge_sheets": merge_sheets,
                "remove_empty_rows": remove_empty,
                "add_metadata": add_metadata,
                "sort_data": sort_data,
                "sort_by": sort_by
            }
            
            with st.spinner("파일을 처리하는 중..."):
                result = processor.process_excel_file(file_id, processing_options)
                
                if result["success"]:
                    st.success(f"✅ 처리 시작! 처리 ID: {result['process_id']}")
                    st.session_state['process_id'] = result['process_id']
                else:
                    st.error(f"❌ 처리 실패: {result['error']}")

elif menu == "📥 결과 다운로드":
    st.header("📥 처리 결과 다운로드")
    
    process_id = st.text_input(
        "처리 ID",
        value=st.session_state.get('process_id', ''),
        help="다운로드할 처리 결과의 ID를 입력하세요"
    )
    
    output_filename = st.text_input(
        "저장할 파일명",
        value="processed_result.xlsx",
        help="다운로드할 파일의 이름"
    )
    
    if st.button("📥 다운로드", type="primary"):
        if not process_id:
            st.error("❌ 처리 ID를 입력해주세요.")
        else:
            with st.spinner("결과를 다운로드하는 중..."):
                result = processor.download_result(process_id, output_filename)
                
                if result["success"]:
                    st.success(f"✅ 다운로드 완료: {result['file_path']}")
                    
                    # 파일 다운로드 버튼
                    with open(result['file_path'], 'rb') as f:
                        st.download_button(
                            label="💾 파일 다운로드",
                            data=f.read(),
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error(f"❌ 다운로드 실패: {result['error']}")

elif menu == "📊 프로젝트 현황":
    st.header("📊 프로젝트 현황")
    
    if st.button("🔄 새로고침"):
        st.rerun()
    
    # 프로젝트 목록 조회
    result = processor.list_projects()
    
    if result["success"]:
        projects = result["projects"]
        
        if projects:
            st.subheader("📁 프로젝트 목록")
            for project in projects:
                with st.expander(f"📂 {project}"):
                    st.write(f"프로젝트: {project}")
                    # 여기에 프로젝트별 통계 추가 가능
        else:
            st.info("📭 등록된 프로젝트가 없습니다.")
    else:
        st.error(f"❌ 프로젝트 조회 실패: {result['error']}")

# 푸터
st.markdown("---")
st.markdown("💌 문의: charmleader@gmail.com | 🌐 클라우드 기반 Excel 처리 시스템")
'''
    
    with open("cloud_excel_app.py", "w", encoding="utf-8") as f:
        f.write(web_content)
    
    print("✅ 웹 인터페이스 생성 완료: cloud_excel_app.py")

def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("🌐 클라우드 기반 Excel 처리 시스템 런처")
    print("📧 제작자: charmleader@gmail.com")
    print("=" * 60)
    
    # 1. Supabase 설정 확인
    if not check_supabase_config():
        return
    
    # 2. 웹 인터페이스 생성
    create_web_interface()
    
    # 3. Streamlit 앱 실행
    print("\n🚀 웹 인터페이스를 시작합니다...")
    print("📱 브라우저에서 http://localhost:8501 로 접속하세요")
    print("⏹️  종료하려면 Ctrl+C를 누르세요")
    print("-" * 60)
    
    try:
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", 
            "cloud_excel_app.py", 
            "--server.port", "8501",
            "--server.headless", "true"
        ])
    except KeyboardInterrupt:
        print("\n👋 프로그램을 종료합니다.")
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")

if __name__ == "__main__":
    main()
