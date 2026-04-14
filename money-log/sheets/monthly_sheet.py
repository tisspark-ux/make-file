"""월별 시트 생성 모듈 (01월 ~ 12월)"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from .budget_sheet import BUDGET_ITEMS


def _side(color="CCCCCC"):
    return Side(style="thin", color=color)

def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, size=11, color="000000"):
    return Font(bold=bold, size=size, color=color)

def _align(h="center", v="center"):
    return Alignment(horizontal=h, vertical=v, wrap_text=False)


# 월별 시트 컬럼 정의
# A:대분류 B:중분류 C:예산기본 D:예산추가 E:예산합계 F:실제지출 G:차액 H:달성율
COL_MAIN   = 1  # A: 대분류
COL_SUB    = 2  # B: 중분류
COL_BASE   = 3  # C: 예산기본 (수입&예산에서 자동)
COL_EXTRA  = 4  # D: 예산추가 (수동입력)
COL_TOTAL  = 5  # E: 예산합계
COL_ACTUAL = 6  # F: 실제지출 (뱅샐입력에서 자동)
COL_DIFF   = 7  # G: 차액
COL_RATE   = 8  # H: 달성율

DETAIL_COLS = ["날짜", "타입", "대분류", "소분류", "내용", "금액", "결제수단", "메모"]


def build_all(wb, year: int = 2026):
    for month in range(1, 13):
        _build_month(wb, year, month)


def _build_month(wb, year: int, month: int):
    sheet_name = f"{month:02d}월"
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    col_widths = {1: 14, 2: 16, 3: 14, 4: 12, 5: 14, 6: 14, 7: 14, 8: 10}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    for r in range(1, 5):
        ws.row_dimensions[r].height = 22

    # ── 제목 ─────────────────────────────────────────────────
    title = ws.cell(row=1, column=1,
                    value=f'=TEXT(DATE(설정!C3,{month},1),"YYYY년 MM월 가계부")')
    title.font = _font(bold=True, size=14, color="1F4E79")
    title.alignment = _align(h="left")
    ws.merge_cells("A1:H1")

    # ── 요약 바 (수입 / 지출 / 잔액) ─────────────────────────
    summary_labels = ["이번달 수입", "이번달 지출", "잔액"]
    summary_colors = ["375623", "C00000", "1F4E79"]
    for i, (label, color) in enumerate(zip(summary_labels, summary_colors)):
        col = i * 2 + 1
        lc = ws.cell(row=2, column=col, value=label)
        lc.font = _font(bold=True, color="FFFFFF", size=10)
        lc.fill = _fill(color)
        lc.alignment = _align()
        lc.border = _border()

        vc = ws.cell(row=2, column=col + 1)
        vc.fill = _fill("F2F2F2")
        vc.font = _font(bold=True, size=11)
        vc.alignment = _align()
        vc.border = _border()
        vc.number_format = '#,##0"원"'

    # 수입/지출/잔액 수식 — 뱅샐입력 시트에서 집계
    ws["B2"] = (f"=SUMPRODUCT((MONTH(뱅샐입력!A$4:A$1003)={month})*"
                f"(YEAR(뱅샐입력!A$4:A$1003)=설정!C3)*"
                f"(뱅샐입력!C$4:C$1003=\"수입\")*뱅샐입력!G$4:G$1003)")
    ws["D2"] = (f"=SUMPRODUCT((MONTH(뱅샐입력!A$4:A$1003)={month})*"
                f"(YEAR(뱅샐입력!A$4:A$1003)=설정!C3)*"
                f"(뱅샐입력!C$4:C$1003=\"지출\")*ABS(뱅샐입력!G$4:G$1003))")
    ws["F2"] = "=B2-D2"
    ws["B2"].number_format = '#,##0"원"'
    ws["D2"].number_format = '#,##0"원"'
    ws["F2"].number_format = '#,##0"원"'

    # ── 예산 vs 실제 헤더 ────────────────────────────────────
    budget_headers = ["대분류", "중분류", "예산기본", "예산추가", "예산합계", "실제지출", "차액", "달성율"]
    header_row = 4
    header_colors = ["1F4E79"] * 2 + ["375623"] * 3 + ["C00000"] + ["7030A0"] * 2
    for ci, (h, hc) in enumerate(zip(budget_headers, header_colors), start=1):
        c = ws.cell(row=header_row, column=ci, value=h)
        c.fill = _fill(hc)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()
        ws.row_dimensions[header_row].height = 22

    # ── 예산 항목 행 ─────────────────────────────────────────
    data_start = 5
    for i, (main_cat, sub_cat, *_) in enumerate(BUDGET_ITEMS):
        r = data_start + i
        fill_color = "F2F2F2" if i % 2 == 0 else "FFFFFF"
        ws.row_dimensions[r].height = 20

        # 대분류 / 중분류
        for ci, val in [(1, main_cat), (2, sub_cat)]:
            c = ws.cell(row=r, column=ci, value=val)
            c.fill = _fill(fill_color)
            c.alignment = _align()
            c.border = _border()

        # C: 예산기본 — 수입&예산 시트에서 SUMIFS
        ws.cell(row=r, column=COL_BASE,
                value=(f'=IFERROR(SUMIFS(수입&예산!H$15:H$50,'
                       f'수입&예산!B$15:B$50,A{r},'
                       f'수입&예산!C$15:C$50,B{r})*10000,0)')).number_format = '#,##0"원"'

        # D: 예산추가 — 수동입력
        ws.cell(row=r, column=COL_EXTRA, value=0).number_format = '#,##0"원"'

        # E: 예산합계
        ws.cell(row=r, column=COL_TOTAL,
                value=f'=C{r}+D{r}').number_format = '#,##0"원"'

        # F: 실제지출 — 뱅샐입력에서 SUMPRODUCT
        ws.cell(row=r, column=COL_ACTUAL,
                value=(f'=SUMPRODUCT('
                       f'(MONTH(뱅샐입력!A$4:A$1003)={month})*'
                       f'(YEAR(뱅샐입력!A$4:A$1003)=설정!C3)*'
                       f'(뱅샐입력!C$4:C$1003="지출")*'
                       f'(뱅샐입력!D$4:D$1003=A{r})*'
                       f'ABS(뱅샐입력!G$4:G$1003))')).number_format = '#,##0"원"'

        # G: 차액 (예산 - 실제)
        ws.cell(row=r, column=COL_DIFF,
                value=f'=E{r}-F{r}').number_format = '#,##0"원"'

        # H: 달성율
        ws.cell(row=r, column=COL_RATE,
                value=f'=IFERROR(F{r}/E{r},0)').number_format = '0%'

        for ci in range(COL_BASE, COL_RATE + 1):
            ws.cell(row=r, column=ci).fill = _fill(fill_color)
            ws.cell(row=r, column=ci).alignment = _align()
            ws.cell(row=r, column=ci).border = _border()

    # ── 합계 행 ──────────────────────────────────────────────
    total_row = data_start + len(BUDGET_ITEMS)
    ws.row_dimensions[total_row].height = 22
    for ci in range(1, 9):
        c = ws.cell(row=total_row, column=ci)
        c.fill = _fill("1F4E79")
        c.font = _font(bold=True, color="FFFFFF")
        c.alignment = _align()
        c.border = _border()
    ws.cell(row=total_row, column=1, value="합계")
    ws.merge_cells(f"A{total_row}:B{total_row}")
    for ci, col_letter in [(COL_BASE, "C"), (COL_EXTRA, "D"), (COL_TOTAL, "E"),
                           (COL_ACTUAL, "F"), (COL_DIFF, "G")]:
        ws.cell(row=total_row, column=ci,
                value=f'=SUM({col_letter}{data_start}:{col_letter}{total_row-1})').number_format = '#,##0"원"'
    ws.cell(row=total_row, column=COL_RATE,
            value=f'=IFERROR(F{total_row}/E{total_row},0)').number_format = '0%'

    # ── 거래 상세 내역 ────────────────────────────────────────
    detail_start = total_row + 2
    ws.row_dimensions[detail_start].height = 22

    # 상세 헤더
    detail_col_widths = [14, 8, 14, 14, 24, 14, 22, 20]
    for ci, (h, w) in enumerate(zip(DETAIL_COLS, detail_col_widths), start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
        c = ws.cell(row=detail_start, column=ci, value=h)
        c.fill = _fill("404040")
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()

    # 상세 데이터: FILTER로 해당 월 + 해당 연도 + 이체 제외
    detail_data_row = detail_start + 1
    ws.cell(row=detail_data_row, column=1,
            value=(f'=IFERROR(FILTER('
                   f'CHOOSE({{1,2,3,4,5,6,7,8}},'
                   f'뱅샐입력!A$4:A$1003,뱅샐입력!C$4:C$1003,'
                   f'뱅샐입력!D$4:D$1003,뱅샐입력!E$4:E$1003,'
                   f'뱅샐입력!F$4:F$1003,뱅샐입력!G$4:G$1003,'
                   f'뱅샐입력!I$4:I$1003,뱅샐입력!J$4:J$1003),'
                   f'(MONTH(뱅샐입력!A$4:A$1003)={month})*'
                   f'(YEAR(뱅샐입력!A$4:A$1003)=설정!C3)*'
                   f'(뱅샐입력!C$4:C$1003<>"이체")),'
                   f'"데이터 없음")'))
    ws.cell(row=detail_data_row, column=1).number_format = "YYYY-MM-DD"

    # 틀 고정
    ws.freeze_panes = f"A{data_start}"

    return ws
