#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
새로운 엑셀 통합기 - 개선된 버전
"""

import streamlit as st
import pandas as pd
import os
import glob
from pathlib import Path
import re
from datetime import datetime
import warnings
import zipfile
import tempfile
import shutil
import subprocess
import sys

# 경고 메시지 숨기기
warnings.filterwarnings('ignore')

def main():
    st.set_page_config(
        page_title="🚀 스마트 엑셀 파일 통합기",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # 메인 타이틀
    st.title("🚀 스마트 엑셀 파일 통합기")
    st.markdown("---")
    
    # 사이드바
    with st.sidebar:
        st.header("⚙️ 설정")
        
        # 통합 옵션
        st.subheader("📋 통합 옵션")
        add_filename = st.checkbox("파일명 열 추가", value=True)
        add_folder = st.checkbox("폴더명 열 추가", value=False)
        
        # 출력 형식
        st.subheader("📤 출력 형식")
        output_format = st.selectbox(
            "다운로드 형식",
            ["Excel (.xlsx)", "CSV (.csv)"],
            index=0
        )
        
        # 시트 선택
        st.subheader("📄 시트 선택")
        select_sheets = st.checkbox("특정 시트만 선택", value=False)
    
    # 메인 컨텐츠
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 파일 업로드")
        
        # 파일 업로드
        uploaded_files = st.file_uploader(
            "엑셀/CSV 파일들을 선택하세요",
            type=['xlsx', 'xls', 'csv'],
            accept_multiple_files=True,
            help="여러 파일을 동시에 선택할 수 있습니다"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)}개 파일이 업로드되었습니다!")
            
            # 파일 목록 표시
            with st.expander("📋 업로드된 파일 목록", expanded=True):
                for i, file in enumerate(uploaded_files, 1):
                    st.write(f"{i}. {file.name} ({file.size:,} bytes)")
            
            # 통합 실행 버튼
            if st.button("🔄 파일 통합하기", type="primary", use_container_width=True):
                with st.spinner("파일 통합 중..."):
                    try:
                        # 임시 폴더 생성
                        temp_dir = tempfile.mkdtemp()
                        
                        # 파일들 저장
                        file_paths = []
                        for file in uploaded_files:
                            file_path = os.path.join(temp_dir, file.name)
                            with open(file_path, "wb") as f:
                                f.write(file.getbuffer())
                            file_paths.append(file_path)
                        
                        # 통합 실행
                        result_df = merge_files(file_paths, add_filename, add_folder)
                        
                        # 결과 표시
                        st.success("✅ 파일 통합이 완료되었습니다!")
                        
                        # 미리보기
                        st.subheader("📊 통합 결과 미리보기")
                        st.dataframe(result_df.head(10), use_container_width=True)
                        
                        # 다운로드 버튼
                        if output_format == "Excel (.xlsx)":
                            output = result_df.to_excel(index=False)
                            st.download_button(
                                label="📥 Excel 파일 다운로드",
                                data=output,
                                file_name=f"통합파일_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            output = result_df.to_csv(index=False, encoding='utf-8-sig')
                            st.download_button(
                                label="📥 CSV 파일 다운로드",
                                data=output,
                                file_name=f"통합파일_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                mime="text/csv"
                            )
                        
                        # 임시 폴더 정리
                        shutil.rmtree(temp_dir)
                        
                    except Exception as e:
                        st.error(f"❌ 오류가 발생했습니다: {str(e)}")
    
    with col2:
        st.header("📊 통계")
        
        if uploaded_files:
            # 파일 통계
            total_size = sum(file.size for file in uploaded_files)
            st.metric("총 파일 수", len(uploaded_files))
            st.metric("총 크기", f"{total_size:,} bytes")
            
            # 파일 형식별 통계
            file_types = {}
            for file in uploaded_files:
                ext = file.name.split('.')[-1].lower()
                file_types[ext] = file_types.get(ext, 0) + 1
            
            st.subheader("📈 파일 형식별 분포")
            for ext, count in file_types.items():
                st.write(f"• .{ext}: {count}개")
        
        st.header("ℹ️ 사용법")
        st.info("""
        1. **파일 선택**: 통합할 엑셀/CSV 파일들을 선택
        2. **옵션 설정**: 사이드바에서 통합 옵션 선택
        3. **통합 실행**: '파일 통합하기' 버튼 클릭
        4. **다운로드**: 통합된 파일을 다운로드
        """)

def merge_files(file_paths, add_filename=False, add_folder=False):
    """파일들을 통합하는 함수"""
    all_data = []
    
    for file_path in file_paths:
        try:
            # 파일 확장자 확인
            file_ext = file_path.split('.')[-1].lower()
            
            if file_ext == 'csv':
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file_path)
            
            # 파일명 열 추가
            if add_filename:
                df['파일명'] = os.path.basename(file_path)
            
            # 폴더명 열 추가
            if add_folder:
                df['폴더명'] = os.path.dirname(file_path)
            
            all_data.append(df)
            
        except Exception as e:
            st.warning(f"⚠️ {file_path} 파일을 읽는 중 오류: {str(e)}")
            continue
    
    if not all_data:
        raise Exception("읽을 수 있는 파일이 없습니다.")
    
    # 모든 데이터 통합
    result_df = pd.concat(all_data, ignore_index=True)
    
    return result_df

if __name__ == "__main__":
    main()
