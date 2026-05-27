#!/usr/bin/env python3
"""
HPPK KRS01 — 엔지니어팀 정기 점검 Task Registry 자동 업데이트 스크립트
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

기능:
  • Google Drive 상의 Task Registry 자동 읽기
  • 다음 예정일 자동 계산 (마지막 완료일 + 주기)
  • 상태 자동 판정 (🟢 정상 / 🟡 임박 / 🔴 연체)
  • 대시보드 KPI 자동 집계
  • Google Drive 자동 쓰기
  • 콘솔 로그 출력 (연체/임박 항목 강조)

실행:
  python pm_task_auto_updater.py

의존성:
  pip install -r requirements.txt
"""

import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import json
import os
import sys
from typing import List, Dict, Tuple

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SPREADSHEET_ID = "1fK_NQQXD7T4p80rtPAAOUkEWNK-5GyPw"
SHEET_CHECKLIST = "점검 목록"
SHEET_DASHBOARD = "📊 대시보드"
SHEET_HISTORY = "완료 이력"

# Google 서비스 계정 인증 파일 경로
CREDENTIALS_PATH = os.getenv("GOOGLE_CREDENTIALS_PATH", "./possible-fabric-203813-e187d8cadfce.json")

# D-day 임계값
THRESHOLD_OVERDUE = 0  # D-day < 0 → 연체
THRESHOLD_IMMINENT = 7  # 0 <= D-day <= 7 → 임박

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 유틸리티 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def parse_date(date_val) -> datetime:
    """날짜 값 파싱 (str, datetime, None 모두 처리)"""
    if date_val is None:
        return None
    if isinstance(date_val, datetime):
        return date_val
    if isinstance(date_val, str):
        # 여러 날짜 형식 시도
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"]:
            try:
                return datetime.strptime(date_val, fmt)
            except ValueError:
                continue
        return None
    return None

def get_next_scheduled_date(last_completed: datetime, cycle: int, unit: str) -> datetime:
    """다음 예정일 계산: 마지막 완료일 + 주기"""
    if last_completed is None:
        return None
    
    try:
        if unit == "주":
            return last_completed + timedelta(weeks=cycle)
        elif unit == "월":
            # 월 단위 계산 (간단한 방식: 30일 근사)
            return last_completed + timedelta(days=cycle * 30)
        elif unit == "반기":
            return last_completed + timedelta(days=cycle * 180)
        elif unit == "연":
            return last_completed + timedelta(days=cycle * 365)
        else:
            return None
    except Exception as e:
        print(f"⚠️  주기 계산 오류: {e}")
        return None

def judge_status(next_scheduled: datetime, current_date: datetime = None) -> str:
    """상태 판정: 🟢 정상 / 🟡 임박 / 🔴 연체"""
    if next_scheduled is None:
        return "⚪ 비활성"
    
    if current_date is None:
        current_date = datetime.now()
    
    days_left = (next_scheduled - current_date).days
    
    if days_left < THRESHOLD_OVERDUE:
        return "🔴 연체"
    elif days_left <= THRESHOLD_IMMINENT:
        return "🟡 임박"
    else:
        return "🟢 정상"

def format_date_for_sheet(date_obj: datetime) -> str:
    """Google Sheet에 입력할 날짜 포맷"""
    if date_obj is None:
        return ""
    return date_obj.strftime("%Y-%m-%d")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Google Drive API
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def authenticate_google_drive() -> gspread.Spreadsheet:
    """Google Drive API 인증 및 Spreadsheet 접근"""
    try:
        if not os.path.exists(CREDENTIALS_PATH):
            raise FileNotFoundError(f"❌ 서비스 계정 JSON 파일을 찾을 수 없습니다: {CREDENTIALS_PATH}")
        
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scopes)
        client = gspread.authorize(creds)
        
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        print(f"✅ Google Drive 인증 성공: {spreadsheet.title}")
        return spreadsheet
    except Exception as e:
        print(f"❌ Google Drive 인증 실패: {e}")
        sys.exit(1)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Task Registry 처리
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def read_task_registry(worksheet) -> List[Dict]:
    """Task Registry 읽기 (행 2: 헤더, 행 3부터: 데이터)"""
    try:
        all_values = worksheet.get_all_values()
        
        if len(all_values) < 3:
            print("❌ Task Registry 데이터가 없습니다.")
            return []
        
        headers = all_values[1]  # 행 2: 헤더
        
        tasks = []
        for row_idx, row in enumerate(all_values[2:], start=3):  # 행 3부터
            if not row or not row[0]:  # 빈 행 스킵
                continue
            
            task = {
                "row_index": row_idx,
                "Task ID": row[0] if len(row) > 0 else "",
                "구분": row[1] if len(row) > 1 else "",
                "업무명": row[2] if len(row) > 2 else "",
                "주기": int(row[3]) if len(row) > 3 and row[3] else None,
                "단위": row[4] if len(row) > 4 else "",
                "일정 기준": row[5] if len(row) > 5 else "",
                "마지막 완료일": parse_date(row[6]) if len(row) > 6 else None,
                "다음 예정일": parse_date(row[7]) if len(row) > 7 else None,
                "담당자": row[8] if len(row) > 8 else "",
                "상태": row[9] if len(row) > 9 else "",
                "비고": row[10] if len(row) > 10 else "",
            }
            
            tasks.append(task)
        
        print(f"✅ Task Registry 로드 완료: {len(tasks)}개 항목")
        return tasks
    except Exception as e:
        print(f"❌ Task Registry 읽기 오류: {e}")
        return []

def update_task_registry(worksheet, tasks: List[Dict], current_date: datetime = None) -> Tuple[List, List]:
    """Task Registry 업데이트: 다음 예정일 & 상태 자동 계산"""
    if current_date is None:
        current_date = datetime.now()
    
    overdue_tasks = []
    imminent_tasks = []
    
    updates = []  # Google Sheet에 반영할 업데이트 목록
    
    for task in tasks:
        last_completed = task["마지막 완료일"]
        cycle = task["주기"]
        unit = task["단위"]
        
        # 다음 예정일 계산
        if last_completed and cycle and unit:
            next_scheduled = get_next_scheduled_date(last_completed, cycle, unit)
            task["다음 예정일"] = next_scheduled
        else:
            next_scheduled = task["다음 예정일"]
        
        # 상태 판정
        status = judge_status(next_scheduled, current_date)
        task["상태"] = status
        
        # 연체/임박 분류
        if status == "🔴 연체":
            overdue_tasks.append(task)
        elif status == "🟡 임박":
            imminent_tasks.append(task)
        
        # Google Sheet 업데이트 준비
        if next_scheduled:
            updates.append({
                "row": task["row_index"],
                "col_next_date": 8,  # 다음 예정일 (H 컬럼)
                "value_next_date": format_date_for_sheet(next_scheduled),
                "col_status": 10,  # 상태 (J 컬럼)
                "value_status": status
            })
    
    return overdue_tasks, imminent_tasks, updates

def write_updates_to_sheet(worksheet, updates: List[Dict]):
    """Google Sheet에 업데이트 반영 (배치 처리)"""
    if not updates:
        print("📝 업데이트할 항목이 없습니다.")
        return
    
    try:
        # 다음 예정일 업데이트
        for update in updates:
            row = update["row"]
            col_next_date = update["col_next_date"]
            value_next_date = update["value_next_date"]
            
            if value_next_date:
                cell_ref = gspread.utils.rowcol_to_a1(row, col_next_date)
                worksheet.update(f"{cell_ref}", value_next_date)
        
        # 상태 업데이트
        for update in updates:
            row = update["row"]
            col_status = update["col_status"]
            value_status = update["value_status"]
            
            cell_ref = gspread.utils.rowcol_to_a1(row, col_status)
            worksheet.update(f"{cell_ref}", value_status)
        
        print(f"✅ Google Sheet 업데이트 완료: {len(updates)}개 행")
    except Exception as e:
        print(f"⚠️  Google Sheet 업데이트 오류: {e}")

def update_dashboard_date(worksheet, current_date: datetime):
    """대시보드 '기준일' 업데이트"""
    try:
        # 행 2, 컬럼 A: 기준일 갱신
        date_str = current_date.strftime("%Y년 %m월 %d일 (%a)")
        date_str_kr = date_str.replace("Mon", "월").replace("Tue", "화").replace("Wed", "수").replace("Thu", "목").replace("Fri", "금").replace("Sat", "토").replace("Sun", "일")
        
        dashboard_text = f"📅  기준일:  {date_str_kr}     |     매주 월요일 09:00 KST  Claude 자동 업데이트"
        worksheet.update("A2", dashboard_text)
        
        print(f"✅ 대시보드 기준일 업데이트: {date_str_kr}")
    except Exception as e:
        print(f"⚠️  대시보드 업데이트 오류: {e}")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 콘솔 로그
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def print_report(current_date: datetime, overdue_tasks: List[Dict], imminent_tasks: List[Dict]):
    """콘솔 보고서 출력"""
    print("\n" + "=" * 100)
    print("🏢 HPPK KRS01 — 엔지니어팀 정기 점검 대시보드 (자동 업데이트)")
    print("=" * 100)
    
    print(f"\n📅 기준일: {current_date.strftime('%Y년 %m월 %d일 (%a)').replace('Mon', '월').replace('Tue', '화').replace('Wed', '수').replace('Thu', '목').replace('Fri', '금').replace('Sat', '토').replace('Sun', '일')}")
    print(f"⏰ 실행 시각: {current_date.strftime('%H:%M:%S')} KST\n")
    
    # 연체 항목
    if overdue_tasks:
        print(f"🔴 【연체】 {len(overdue_tasks)}개 항목")
        print("-" * 100)
        for task in sorted(overdue_tasks, key=lambda x: x.get("다음 예정일") or datetime.max):
            next_date = task.get("다음 예정일")
            days_overdue = (current_date - next_date).days if next_date else 0
            print(f"  • [{task['Task ID']}] {task['업무명']}")
            print(f"    └─ 담당자: {task['담당자']} | 예정일: {format_date_for_sheet(next_date)} (D-{days_overdue})")
        print()
    
    # 임박 항목
    if imminent_tasks:
        print(f"🟡 【임박】 {len(imminent_tasks)}개 항목 (D-7 이내)")
        print("-" * 100)
        for task in sorted(imminent_tasks, key=lambda x: x.get("다음 예정일") or datetime.max):
            next_date = task.get("다음 예정일")
            days_left = (next_date - current_date).days if next_date else 0
            print(f"  • [{task['Task ID']}] {task['업무명']}")
            print(f"    └─ 담당자: {task['담당자']} | 예정일: {format_date_for_sheet(next_date)} (D-{days_left})")
        print()
    
    if not overdue_tasks and not imminent_tasks:
        print("✅ 연체 및 임박 항목 없음 (모든 점검 정상)")
        print()
    
    print("=" * 100 + "\n")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 메인
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def main():
    print("\n🚀 HPPK KRS01 PM Task Auto Updater 시작...\n")
    
    current_date = datetime.now()
    
    # 1️⃣ Google Drive 인증
    spreadsheet = authenticate_google_drive()
    
    # 2️⃣ Task Registry 읽기
    worksheet_checklist = spreadsheet.worksheet(SHEET_CHECKLIST)
    tasks = read_task_registry(worksheet_checklist)
    
    if not tasks:
        print("❌ Task를 읽을 수 없습니다.")
        sys.exit(1)
    
    # 3️⃣ 자동 계산 및 업데이트
    overdue_tasks, imminent_tasks, updates = update_task_registry(tasks, current_date)
    
    # 4️⃣ Google Sheet 업데이트
    write_updates_to_sheet(worksheet_checklist, updates)
    
    # 5️⃣ 대시보드 기준일 업데이트
    worksheet_dashboard = spreadsheet.worksheet(SHEET_DASHBOARD)
    update_dashboard_date(worksheet_dashboard, current_date)
    
    # 6️⃣ 콘솔 보고서 출력
    print_report(current_date, overdue_tasks, imminent_tasks)
    
    print("✅ 작업 완료!\n")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 예상치 못한 오류: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
