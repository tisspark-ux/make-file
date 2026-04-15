"""분석 시트 생성 모듈 — 연간 카테고리 지출 분석 + 월별 히트맵"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from .budget_sheet import BUDGET_ITEMS
from .banksalad_sheet import DATA_LAST_ROW


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
    return Alignment(horizontal=h, vertical=v, wrap_text=False)


def build(wb, year: int = 2026):
    ws = wb.create_sheet("분석")
    ws.sheet_view.showGridLines = False

    # 컬럼 레이아웃
    # A(1):여백  B(2):대분류  C(3):중분류  D(4):연간예산/1월  E(5):연간실제/2월  F(6):차액/3월
    # G(7):달성율/4월  H(8):5월  I(9):6월  J(10):7월  K(11):8월  L(12):9월
    # M(13):10월  N(14):11월  O(15):12월  P(16):연간합계  Q(17):여백
    col_widths = {1: 4, 2: 12, 3: 16, 4: 13, 5: 13, 6: 12, 7: 10,
                  8: 10, 9: 10, 10: 10, 11: 10, 12: 10,
                  13: 10, 14: 10, 15: 10, 16: 13, 17: 4}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    row = 1

    # ════════════════════════════════════════════════════════
    # Section 1: 카테고리별 연간 지출 분석
    # ════════════════════════════════════════════════════════
    ws.row_dimensions[row].height = 30
    title = ws.cell(row=row, column=2,
                    value='=설정!C3&"년 카테고리별 지출 분석"')
    title.font = _font(bold=True, size=16, color="17375E")
    title.alignment = _align(h="left")
    ws.merge_cells("B1:P1")

    row = 2
    ws.row_dimensions[row].height = 14
    sub = ws.cell(row=row, column=2,
                  value="※ 뱅샐입력 데이터 기준 자동 집계. 연간 예산 대비 실제 지출을 분석합니다.")
    sub.font = Font(size=9, color="888888", italic=True)
    sub.alignment = _align(h="left")
    ws.merge_cells("B2:P2")

    row = 4
    _section_title(ws, row, "① 카테고리별 연간 지출 요약", end_col=7)
    row += 1

    # 헤더
    ann_header_row = row
    ws.row_dimensions[row].height = 22
    ann_headers = ["대분류", "중분류", "연간예산", "연간실제지출", "차액", "달성율"]
    ann_colors  = ["17375E", "17375E", "2D6A4F", "922B21", "1B6CA8", "1B6CA8"]
    for ci, (h, hc) in enumerate(zip(ann_headers, ann_colors), start=2):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill(hc)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()
    row += 1

    # 데이터 행
    ann_data_start = row
    for i, (main_cat, sub_cat, *_) in enumerate(BUDGET_ITEMS):
        fill_c = "F8FAFC" if i % 2 == 0 else "FFFFFF"
        ws.row_dimensions[row].height = 20

        for ci, val in [(2, main_cat), (3, sub_cat)]:
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(fill_c)
            c.alignment = _align()
            c.border = _border()

        # D: 연간예산 = 월예산(수입&예산 H열) × 10000 × 12
        ws.cell(row=row, column=4,
                value=(f'=IFERROR(SUMIFS(수입&예산!H$15:H$50,'
                       f'수입&예산!B$15:B$50,B{row},'
                       f'수입&예산!C$15:C$50,C{row})*10000*12,0)'
                       )).number_format = '#,##0"원"'

        # E: 연간실제지출 = 뱅샐입력에서 해당연도 해당 대분류 지출 합계
        ws.cell(row=row, column=5,
                value=(f'=SUMPRODUCT('
                       f'(YEAR(뱅샐입력!A$4:A${DATA_LAST_ROW})=설정!C3)*'
                       f'(뱅샐입력!C$4:C${DATA_LAST_ROW}="지출")*'
                       f'(뱅샐입력!D$4:D${DATA_LAST_ROW}=B{row})*'
                       f'ABS(뱅샐입력!G$4:G${DATA_LAST_ROW}))'
                       )).number_format = '#,##0"원"'

        # F: 차액
        ws.cell(row=row, column=6,
                value=f"=D{row}-E{row}").number_format = '#,##0"원"'

        # G: 달성율
        ws.cell(row=row, column=7,
                value=f"=IFERROR(E{row}/D{row},0)").number_format = '0%'

        for ci in range(4, 8):
            ws.cell(row=row, column=ci).fill = _fill(fill_c)
            ws.cell(row=row, column=ci).alignment = _align()
            ws.cell(row=row, column=ci).border = _border()

        row += 1

    ann_data_end = row - 1

    # 합계 행
    total_r = row
    ws.row_dimensions[row].height = 22
    for ci in range(2, 8):
        c = ws.cell(row=row, column=ci)
        c.fill = _fill("17375E")
        c.font = _font(bold=True, color="FFFFFF")
        c.alignment = _align()
        c.border = _border()
    ws.cell(row=row, column=2, value="합계")
    ws.merge_cells(f"B{row}:C{row}")
    for ci, col_l in [(4, "D"), (5, "E"), (6, "F")]:
        ws.cell(row=row, column=ci,
                value=f"=SUM({col_l}{ann_data_start}:{col_l}{ann_data_end})").number_format = '#,##0"원"'
    ws.cell(row=row, column=7,
            value=f"=IFERROR(E{row}/D{row},0)").number_format = '0%'
    row += 2

    # 연간 분석 차트 (BarChart: 예산 vs 실제)
    chart_anchor = row
    _add_annual_chart(ws, ann_header_row, ann_data_start, ann_data_end, chart_anchor)
    row += 20  # 차트 높이(~17행) + 여백

    # ════════════════════════════════════════════════════════
    # Section 2: 월별 카테고리 지출 히트맵
    # ════════════════════════════════════════════════════════
    _section_title(ws, row, "② 월별 카테고리 지출 히트맵  (진할수록 지출 많음)", end_col=16)
    row += 1

    hm_hint = ws.cell(row=row, column=2,
                      value="※ 조건부 서식으로 색상 강도를 표시합니다. 흰색=0원, 진한 빨강=해당 월 최대 지출.")
    hm_hint.font = Font(size=9, color="888888", italic=True)
    hm_hint.alignment = _align(h="left")
    ws.merge_cells(f"B{row}:P{row}")
    ws.row_dimensions[row].height = 14
    row += 1

    # 히트맵 헤더 (대분류 | 중분류 | 1월~12월 | 연간합계)
    hm_header_row = row
    ws.row_dimensions[row].height = 22
    hm_headers = ["대분류", "중분류"] + [f"{m}월" for m in range(1, 13)] + ["연간합계"]
    hm_colors  = ["17375E", "17375E"] + ["2D6A4F"] * 12 + ["922B21"]
    for ci, (h, hc) in enumerate(zip(hm_headers, hm_colors), start=2):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill(hc)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()
    row += 1

    # 히트맵 데이터 행
    hm_data_start = row
    for i, (main_cat, sub_cat, *_) in enumerate(BUDGET_ITEMS):
        fill_c = "FFFFFF"
        ws.row_dimensions[row].height = 20

        for ci, val in [(2, main_cat), (3, sub_cat)]:
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill("F8FAFC" if i % 2 == 0 else "FFFFFF")
            c.alignment = _align()
            c.border = _border()

        # D~O: 1월~12월 지출 SUMPRODUCT
        for m in range(1, 13):
            col = m + 3  # D=4, E=5, ..., O=15
            ws.cell(row=row, column=col,
                    value=(f'=SUMPRODUCT('
                           f'(MONTH(뱅샐입력!A$4:A${DATA_LAST_ROW})={m})*'
                           f'(YEAR(뱅샐입력!A$4:A${DATA_LAST_ROW})=설정!C3)*'
                           f'(뱅샐입력!C$4:C${DATA_LAST_ROW}="지출")*'
                           f'(뱅샐입력!D$4:D${DATA_LAST_ROW}=B{row})*'
                           f'ABS(뱅샐입력!G$4:G${DATA_LAST_ROW}))'
                           )).number_format = '#,##0"원"'
            ws.cell(row=row, column=col).alignment = _align()
            ws.cell(row=row, column=col).border = _border()

        # P: 연간합계
        ws.cell(row=row, column=16,
                value=f"=SUM(D{row}:O{row})").number_format = '#,##0"원"'
        ws.cell(row=row, column=16).font = _font(bold=True)
        ws.cell(row=row, column=16).alignment = _align()
        ws.cell(row=row, column=16).border = _border()

        row += 1

    hm_data_end = row - 1

    # 합계 행
    ws.row_dimensions[row].height = 22
    for ci in range(2, 17):
        c = ws.cell(row=row, column=ci)
        c.fill = _fill("17375E")
        c.font = _font(bold=True, color="FFFFFF")
        c.alignment = _align()
        c.border = _border()
    ws.cell(row=row, column=2, value="합계")
    ws.merge_cells(f"B{row}:C{row}")
    for ci in range(4, 17):
        col_l = get_column_letter(ci)
        ws.cell(row=row, column=ci,
                value=f"=SUM({col_l}{hm_data_start}:{col_l}{hm_data_end})").number_format = '#,##0"원"'

    # ── ColorScale 조건부 서식 (히트맵 효과) ─────────────────
    hm_range = f"D{hm_data_start}:O{hm_data_end}"
    ws.conditional_formatting.add(
        hm_range,
        ColorScaleRule(
            start_type="num",   start_value=0,        start_color="FFFFFF",
            mid_type="percentile", mid_value=50,       mid_color="FFEB84",
            end_type="percentile", end_value=100,      end_color="922B21"
        )
    )

    ws.freeze_panes = "D5"
    return ws


def _add_annual_chart(ws, header_row, data_start, data_end, anchor_row):
    """예산 vs 실제지출 클러스터 막대 차트"""
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.title = "카테고리별 연간 예산 vs 실제지출"
    bar.y_axis.title = "금액 (원)"
    bar.x_axis.title = "카테고리"
    bar.style = 10
    bar.width = 26
    bar.height = 14

    # 연간예산(D열=4), 연간실제(E열=5)
    budget_ref = Reference(ws, min_col=4, min_row=data_start, max_row=data_end)
    actual_ref = Reference(ws, min_col=5, min_row=data_start, max_row=data_end)
    bar.add_data(budget_ref, titles_from_data=False)
    bar.add_data(actual_ref, titles_from_data=False)
    bar.series[0].title = SeriesLabel(v="연간예산")
    bar.series[1].title = SeriesLabel(v="연간실제지출")
    bar.series[0].graphicalProperties.solidFill = "1B6CA8"
    bar.series[1].graphicalProperties.solidFill = "922B21"

    # X축: 중분류(C열=3) 라벨
    cats = Reference(ws, min_col=3, min_row=data_start, max_row=data_end)
    bar.set_categories(cats)

    ws.add_chart(bar, f"B{anchor_row}")


def _section_title(ws, row, title, end_col=16):
    c = ws.cell(row=row, column=2, value=title)
    c.font = _font(bold=True, size=12, color="FFFFFF")
    c.fill = _fill("17375E")
    c.alignment = _align(h="left")
    c.border = _border()
    ws.merge_cells(f"B{row}:{get_column_letter(end_col)}{row}")
    ws.row_dimensions[row].height = 24
