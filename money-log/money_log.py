from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date


CATEGORIES = [
    "식비",
    "교통",
    "주거/통신",
    "의료/건강",
    "문화/여가",
    "쇼핑",
    "교육",
    "기타",
]

HEADERS = ["날짜", "분류", "내용", "수입", "지출", "잔액", "메모"]


class MoneyLog:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self._build_monthly_sheets()
        self._build_summary_sheet()

    def _header_fill(self):
        return PatternFill("solid", fgColor="4472C4")

    def _thin_border(self):
        side = Side(style="thin", color="BBBBBB")
        return Border(left=side, right=side, top=side, bottom=side)

    def _apply_header(self, ws, headers, row=1):
        fill = self._header_fill()
        font = Font(bold=True, color="FFFFFF")
        align = Alignment(horizontal="center", vertical="center")
        for col, title in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=title)
            cell.fill = fill
            cell.font = font
            cell.alignment = align
            cell.border = self._thin_border()

    def _build_monthly_sheets(self):
        today = date.today()
        year = today.year
        for month in range(1, 13):
            ws = self.wb.create_sheet(title=f"{month}월")
            self._apply_header(ws, HEADERS)

            # 입력 예시 행 (첫 번째 데이터 행)
            example_row = [
                f"{year}-{month:02d}-01",
                "식비",
                "점심",
                "",
                12000,
                f"=IF(ROW()=2, D2-E2, F1+D2-E2)",
                "",
            ]
            for col, val in enumerate(example_row, start=1):
                cell = ws.cell(row=2, column=col, value=val)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = self._thin_border()

            # 잔액 수식을 일반 텍스트가 아닌 실제 수식으로 교체
            ws.cell(row=2, column=6).value = "=D2-E2"

            # 열 너비 설정
            col_widths = [14, 14, 20, 12, 12, 12, 20]
            for i, width in enumerate(col_widths, start=1):
                ws.column_dimensions[get_column_letter(i)].width = width

            ws.row_dimensions[1].height = 22

    def _build_summary_sheet(self):
        ws = self.wb.create_sheet(title="요약", index=0)
        ws.title = "요약"

        summary_headers = ["월", "총 수입", "총 지출", "순이익"]
        self._apply_header(ws, summary_headers)

        for i, month in enumerate(range(1, 13), start=2):
            sheet_name = f"{month}월"
            ws.cell(row=i, column=1, value=f"{month}월").alignment = Alignment(horizontal="center")
            ws.cell(row=i, column=2, value=f"=SUM('{sheet_name}'!D:D)").alignment = Alignment(horizontal="center")
            ws.cell(row=i, column=3, value=f"=SUM('{sheet_name}'!E:E)").alignment = Alignment(horizontal="center")
            ws.cell(row=i, column=4, value=f"=B{i}-C{i}").alignment = Alignment(horizontal="center")
            for col in range(1, 5):
                ws.cell(row=i, column=col).border = self._thin_border()

        col_widths = [10, 14, 14, 14]
        for i, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.row_dimensions[1].height = 22

    def save(self, path: str):
        self.wb.save(path)
