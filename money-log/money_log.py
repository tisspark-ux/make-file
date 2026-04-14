"""가계부 Excel 생성 메인 모듈"""

from openpyxl import Workbook
from sheets import (config_sheet, budget_sheet, banksalad_sheet,
                    monthly_sheet, summary_sheet, asset_sheet)


def create(year: int = 2026, output_path: str = None) -> str:
    if output_path is None:
        output_path = f"data/가계부_{year}.xlsx"

    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    # 시트 생성 순서: 설정 → 수입&예산 → 뱅샐입력 → 요약 → 01~12월 → 자산현황
    config_sheet.build(wb, year)
    budget_sheet.build(wb, year)
    banksalad_sheet.build(wb)
    summary_sheet.build(wb, year)
    monthly_sheet.build_all(wb, year)
    asset_sheet.build(wb)

    wb.save(output_path)
    return output_path
