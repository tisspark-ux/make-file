"""가계부 Excel 생성 메인 모듈"""

from openpyxl import Workbook
from sheets import (config_sheet, budget_sheet, banksalad_sheet,
                    monthly_sheet, summary_sheet, asset_sheet,
                    dashboard_sheet, payment_sheet, analysis_sheet,
                    loan_sheet)


def create(year: int = 2026, output_path: str = None) -> str:
    if output_path is None:
        output_path = f"data/가계부_{year}.xlsx"

    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    # 시트 생성 순서:
    # 설정 → 수입&예산 → 뱅샐입력 → 대시보드 → 요약
    # → 01~12월 → 자산현황 → 결제수단 → 분석
    config_sheet.build(wb, year)       # 0: 설정
    budget_sheet.build(wb, year)       # 1: 수입&예산
    banksalad_sheet.build(wb)          # 2: 뱅샐입력
    dashboard_sheet.build(wb)          # 3: 대시보드  (파일 열 때 이 탭)
    summary_sheet.build(wb, year)      # 4: 요약
    monthly_sheet.build_all(wb, year)  # 5~16: 01~12월
    asset_sheet.build(wb)              # 17: 자산현황
    payment_sheet.build(wb)            # 18: 결제수단
    analysis_sheet.build(wb, year)     # 19: 분석
    loan_sheet.build(wb)               # 20: 대출상환

    wb.save(output_path)
    return output_path
