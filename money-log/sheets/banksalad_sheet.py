"""뱅샐입력 시트 생성 모듈"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def _side():
    return Side(style="thin", color="E2E8F0")

def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, size=11, color="000000"):
    return Font(name="Calibri", bold=bold, size=size, color=color)

def _align(h="center", v="center"):
    return Alignment(horizontal=h, vertical=v, wrap_text=False)


HEADERS = ["날짜", "시간", "타입", "대분류", "소분류", "내용", "금액", "화폐", "결제수단", "메모"]

# 수식 범위 상한 (연간 최대 거래 건수 여유 포함)
DATA_LAST_ROW = 3003


def build(wb):
    ws = wb.create_sheet("뱅샐입력")
    ws.sheet_view.showGridLines = True

    col_widths = {1: 14, 2: 12, 3: 8, 4: 14, 5: 14,
                  6: 22, 7: 14, 8: 8, 9: 22, 10: 20}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 40

    # ── 안내 문구 ─────────────────────────────────────────────
    guide = ws.cell(row=1, column=1,
                    value="▶  뱅크샐러드 앱 → 내보내기 → '가계부 내역' 시트를 복사하여 3행부터 붙여넣기 하세요.")
    guide.font = Font(bold=True, size=10, color="C0392B")
    guide.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("A1:J1")

    guide2 = ws.cell(row=2, column=1,
                     value="※ 헤더(1행)는 그대로 두고, 데이터만 붙여넣기. '이체' 타입은 집계에서 자동 제외됩니다.")
    guide2.font = Font(size=9, color="888888", italic=True)
    guide2.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("A2:J2")

    # ── 헤더 행 ───────────────────────────────────────────────
    header_row = 3
    for ci, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=header_row, column=ci, value=h)
        c.fill = _fill("1A2B4A")
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()

    # ── 데이터 영역 서식 (4행~1003행 미리 지정) ──────────────
    date_fill_even = _fill("F4F8FC")
    date_fill_odd  = _fill("FFFFFF")
    for r in range(4, 1004):
        fill = date_fill_even if r % 2 == 0 else date_fill_odd
        for ci in range(1, 11):
            c = ws.cell(row=r, column=ci)
            c.fill = fill
            c.alignment = _align()
            c.border = _border()
            if ci == 1:   # 날짜
                c.number_format = "YYYY-MM-DD"
            elif ci == 7: # 금액
                c.number_format = '#,##0'

    # 틀 고정 (헤더 행 아래)
    ws.freeze_panes = "A4"

    return ws
