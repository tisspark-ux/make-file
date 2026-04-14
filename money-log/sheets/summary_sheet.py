"""요약 시트 생성 모듈"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import get_column_letter
from .budget_sheet import BUDGET_ITEMS


def _side():
    return Side(style="thin", color="CCCCCC")

def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, size=11, color="000000"):
    return Font(bold=bold, size=size, color=color)

def _align(h="center", v="center"):
    return Alignment(horizontal=h, vertical=v)


def build(wb, year: int = 2026):
    ws = wb.create_sheet("요약", 3)  # 설정, 수입&예산, 뱅샐입력 다음
    ws.sheet_view.showGridLines = False

    col_widths = {1: 6, 2: 14, 3: 16, 4: 16, 5: 16, 6: 16, 7: 16, 8: 6}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    for r in range(1, 50):
        ws.row_dimensions[r].height = 20

    # ── 제목 ─────────────────────────────────────────────────
    title = ws.cell(row=1, column=2,
                    value=f'=설정!C3&"년 가계부 요약"')
    title.font = _font(bold=True, size=16, color="1F4E79")
    title.alignment = _align(h="left")
    ws.merge_cells("B1:G1")
    ws.row_dimensions[1].height = 30

    # ── 월별 수입/지출/잔액 테이블 ───────────────────────────
    section_row = 3
    _section_title(ws, section_row, "월별 수입 / 지출 / 잔액")

    headers = ["월", "수입", "지출", "잔액", "예산", "예산대비"]
    header_row = section_row + 1
    header_colors = ["1F4E79", "375623", "C00000", "1F4E79", "7030A0", "7030A0"]
    for ci, (h, hc) in enumerate(zip(headers, header_colors), start=2):
        c = ws.cell(row=header_row, column=ci, value=h)
        c.fill = _fill(hc)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()

    data_start = header_row + 1
    for month in range(1, 13):
        r = data_start + month - 1
        ws.row_dimensions[r].height = 20
        sheet_name = f"{month:02d}월"
        fill_color = "F5F9FF" if month % 2 == 0 else "FFFFFF"

        ws.cell(row=r, column=2, value=f"{month}월").alignment = _align()
        ws.cell(row=r, column=2).fill = _fill(fill_color)
        ws.cell(row=r, column=2).border = _border()
        ws.cell(row=r, column=2).font = _font(bold=True)

        ws.cell(row=r, column=3, value=f"='{sheet_name}'!B2").number_format = '#,##0'
        ws.cell(row=r, column=4, value=f"='{sheet_name}'!D2").number_format = '#,##0'
        ws.cell(row=r, column=5, value=f"='{sheet_name}'!F2").number_format = '#,##0'
        # 예산 합계: 해당 월 예산합계 행
        budget_total_row = 4 + len(BUDGET_ITEMS) + 1
        ws.cell(row=r, column=6,
                value=f"=IFERROR('{sheet_name}'!E{budget_total_row},0)").number_format = '#,##0'
        ws.cell(row=r, column=7,
                value=f'=IFERROR(D{r}/F{r},0)').number_format = '0%'

        for ci in range(3, 8):
            ws.cell(row=r, column=ci).fill = _fill(fill_color)
            ws.cell(row=r, column=ci).alignment = _align()
            ws.cell(row=r, column=ci).border = _border()

    # 연간 합계 행
    total_r = data_start + 12
    ws.row_dimensions[total_r].height = 22
    ws.cell(row=total_r, column=2, value="연간 합계")
    ws.merge_cells(f"B{total_r}:B{total_r}")
    for ci, col_l in [(3, "C"), (4, "D"), (5, "E")]:
        ws.cell(row=total_r, column=ci,
                value=f"=SUM({col_l}{data_start}:{col_l}{total_r-1})").number_format = '#,##0'
    for ci in range(2, 8):
        c = ws.cell(row=total_r, column=ci)
        c.fill = _fill("1F4E79")
        c.font = _font(bold=True, color="FFFFFF")
        c.alignment = _align()
        c.border = _border()

    # ── 차트: 월별 수입 vs 지출 ──────────────────────────────
    chart_row = section_row + 17
    _add_income_expense_chart(ws, data_start, chart_row)

    return ws


def _section_title(ws, row, title):
    c = ws.cell(row=row, column=2, value=title)
    c.font = _font(bold=True, size=12, color="FFFFFF")
    c.fill = _fill("1F4E79")
    c.alignment = _align(h="left")
    c.border = _border()
    ws.merge_cells(f"B{row}:G{row}")
    ws.row_dimensions[row].height = 24


def _add_income_expense_chart(ws, data_start, anchor_row):
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "월별 수입 / 지출"
    chart.y_axis.title = "금액 (원)"
    chart.x_axis.title = "월"
    chart.style = 10
    chart.width = 20
    chart.height = 12

    # 수입 데이터 (C열)
    income_ref = Reference(ws, min_col=3, min_row=data_start,
                           max_row=data_start + 11)
    # 지출 데이터 (D열)
    expense_ref = Reference(ws, min_col=4, min_row=data_start,
                            max_row=data_start + 11)
    # 월 라벨 (B열)
    cats = Reference(ws, min_col=2, min_row=data_start,
                     max_row=data_start + 11)

    chart.add_data(income_ref, titles_from_data=False)
    chart.add_data(expense_ref, titles_from_data=False)
    chart.set_categories(cats)
    chart.series[0].title = SeriesLabel(v="수입")
    chart.series[1].title = SeriesLabel(v="지출")

    ws.add_chart(chart, f"B{anchor_row}")
