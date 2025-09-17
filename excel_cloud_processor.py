#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
클라우드 데이터베이스 기반 Excel 처리 시스템
Supabase PostgreSQL을 활용한 안전한 Excel 파일 처리
"""

import os
import json
import base64
import pandas as pd
import requests
from datetime import datetime
from pathlib import Path
import io
import zipfile

class ExcelCloudProcessor:
    def __init__(self, supabase_url, supabase_key):
        """
        Supabase 클라이언트 초기화
        
        Args:
            supabase_url: Supabase 프로젝트 URL
            supabase_key: Supabase API 키
        """
        self.supabase_url = supabase_url
        self.supabase_key = supabase_key
        self.headers = {
            "apikey": supabase_key,
            "Authorization": f"Bearer {supabase_key}",
            "Content-Type": "application/json"
        }
        
    def upload_excel_file(self, file_path, project_name="default"):
        """
        Excel 파일을 클라우드에 업로드
        
        Args:
            file_path: 업로드할 Excel 파일 경로
            project_name: 프로젝트 이름
            
        Returns:
            dict: 업로드 결과 (file_id, status)
        """
        try:
            # 파일 읽기
            with open(file_path, 'rb') as f:
                file_data = f.read()
            
            # Base64 인코딩
            file_base64 = base64.b64encode(file_data).decode('utf-8')
            
            # 파일 정보
            file_info = {
                "filename": os.path.basename(file_path),
                "file_size": len(file_data),
                "file_data": file_base64,
                "project_name": project_name,
                "upload_time": datetime.now().isoformat(),
                "status": "uploaded"
            }
            
            # Supabase에 업로드
            response = requests.post(
                f"{self.supabase_url}/rest/v1/excel_files",
                headers=self.headers,
                json=file_info
            )
            
            if response.status_code == 201:
                result = response.json()
                print(f"✅ 파일 업로드 성공: {file_info['filename']}")
                print(f"📁 파일 ID: {result[0]['id']}")
                return {"success": True, "file_id": result[0]['id'], "data": result[0]}
            else:
                print(f"❌ 업로드 실패: {response.text}")
                return {"success": False, "error": response.text}
                
        except Exception as e:
            print(f"❌ 업로드 오류: {e}")
            return {"success": False, "error": str(e)}
    
    def process_excel_file(self, file_id, processing_options=None):
        """
        클라우드에서 Excel 파일 처리
        
        Args:
            file_id: 처리할 파일 ID
            processing_options: 처리 옵션
            
        Returns:
            dict: 처리 결과
        """
        try:
            if processing_options is None:
                processing_options = {
                    "merge_sheets": True,
                    "remove_empty_rows": True,
                    "add_metadata": True,
                    "sort_data": True
                }
            
            # 처리 요청
            process_data = {
                "file_id": file_id,
                "processing_options": processing_options,
                "status": "processing",
                "start_time": datetime.now().isoformat()
            }
            
            response = requests.post(
                f"{self.supabase_url}/rest/v1/process_excel",
                headers=self.headers,
                json=process_data
            )
            
            if response.status_code == 201:
                result = response.json()
                print(f"🔄 파일 처리 시작: {file_id}")
                return {"success": True, "process_id": result[0]['id']}
            else:
                print(f"❌ 처리 요청 실패: {response.text}")
                return {"success": False, "error": response.text}
                
        except Exception as e:
            print(f"❌ 처리 오류: {e}")
            return {"success": False, "error": str(e)}
    
    def get_processing_status(self, process_id):
        """
        처리 상태 확인
        
        Args:
            process_id: 처리 ID
            
        Returns:
            dict: 처리 상태
        """
        try:
            response = requests.get(
                f"{self.supabase_url}/rest/v1/process_excel?id=eq.{process_id}",
                headers=self.headers
            )
            
            if response.status_code == 200:
                result = response.json()
                if result:
                    return {"success": True, "status": result[0]}
                else:
                    return {"success": False, "error": "처리 정보를 찾을 수 없습니다"}
            else:
                return {"success": False, "error": response.text}
                
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def download_result(self, process_id, output_path="result.xlsx"):
        """
        처리 결과 다운로드
        
        Args:
            process_id: 처리 ID
            output_path: 저장할 파일 경로
            
        Returns:
            dict: 다운로드 결과
        """
        try:
            # 처리 상태 확인
            status = self.get_processing_status(process_id)
            if not status["success"]:
                return status
            
            if status["status"]["status"] != "completed":
                return {"success": False, "error": "처리가 아직 완료되지 않았습니다"}
            
            # 결과 파일 다운로드
            response = requests.get(
                f"{self.supabase_url}/rest/v1/process_excel?id=eq.{process_id}&select=result_file_data",
                headers=self.headers
            )
            
            if response.status_code == 200:
                result = response.json()
                if result and result[0]["result_file_data"]:
                    # Base64 디코딩
                    file_data = base64.b64decode(result[0]["result_file_data"])
                    
                    # 파일 저장
                    with open(output_path, 'wb') as f:
                        f.write(file_data)
                    
                    print(f"✅ 결과 다운로드 완료: {output_path}")
                    return {"success": True, "file_path": output_path}
                else:
                    return {"success": False, "error": "결과 파일이 없습니다"}
            else:
                return {"success": False, "error": response.text}
                
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def list_projects(self):
        """
        프로젝트 목록 조회
        
        Returns:
            list: 프로젝트 목록
        """
        try:
            response = requests.get(
                f"{self.supabase_url}/rest/v1/excel_files?select=project_name&distinct=true",
                headers=self.headers
            )
            
            if response.status_code == 200:
                projects = [item["project_name"] for item in response.json()]
                return {"success": True, "projects": projects}
            else:
                return {"success": False, "error": response.text}
                
        except Exception as e:
            return {"success": False, "error": str(e)}

def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("🌐 클라우드 기반 Excel 처리 시스템")
    print("📧 제작자: charmleader@gmail.com")
    print("=" * 60)
    
    # Supabase 설정 (실제 값으로 변경 필요)
    SUPABASE_URL = "https://your-project.supabase.co"
    SUPABASE_KEY = "your-anon-key"
    
    # 클라이언트 초기화
    processor = ExcelCloudProcessor(SUPABASE_URL, SUPABASE_KEY)
    
    print("📋 사용 가능한 기능:")
    print("1. Excel 파일 업로드")
    print("2. 파일 처리")
    print("3. 결과 다운로드")
    print("4. 프로젝트 목록 조회")
    
    while True:
        print("\n" + "-" * 40)
        choice = input("선택하세요 (1-4, q: 종료): ").strip()
        
        if choice == 'q':
            break
        elif choice == '1':
            file_path = input("업로드할 Excel 파일 경로: ").strip()
            project_name = input("프로젝트 이름 (기본값: default): ").strip() or "default"
            
            if os.path.exists(file_path):
                result = processor.upload_excel_file(file_path, project_name)
                if result["success"]:
                    print(f"✅ 업로드 성공! 파일 ID: {result['file_id']}")
                else:
                    print(f"❌ 업로드 실패: {result['error']}")
            else:
                print("❌ 파일을 찾을 수 없습니다.")
        
        elif choice == '2':
            file_id = input("처리할 파일 ID: ").strip()
            result = processor.process_excel_file(file_id)
            if result["success"]:
                print(f"✅ 처리 시작! 처리 ID: {result['process_id']}")
            else:
                print(f"❌ 처리 실패: {result['error']}")
        
        elif choice == '3':
            process_id = input("다운로드할 처리 ID: ").strip()
            output_path = input("저장할 파일 경로 (기본값: result.xlsx): ").strip() or "result.xlsx"
            
            result = processor.download_result(process_id, output_path)
            if result["success"]:
                print(f"✅ 다운로드 완료: {result['file_path']}")
            else:
                print(f"❌ 다운로드 실패: {result['error']}")
        
        elif choice == '4':
            result = processor.list_projects()
            if result["success"]:
                print("📁 프로젝트 목록:")
                for project in result["projects"]:
                    print(f"  - {project}")
            else:
                print(f"❌ 조회 실패: {result['error']}")
        
        else:
            print("❌ 잘못된 선택입니다.")

if __name__ == "__main__":
    main()
