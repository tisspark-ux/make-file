"""자산현황 시트 생성 모듈"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


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


# (항목명, 기본값, 비고, 카테고리)  카테고리: asset/liability/auto
ASSET_ITEMS = [
    # ── 자산 ──────────────────────────────────────────────────
    ("부동산",      900_000_000, "기흥역센트럴푸르지오",     "asset"),
    ("계좌(KB)",              0, "직접 입력",                "asset"),
    ("계좌(토스)",            0, "직접 입력",                "asset"),
    ("계좌(기타)",            0, "직접 입력",                "asset"),
    ("주식",                  0, "삼성증권 등",              "asset"),
    ("코인",                  0, "직접 입력",                "asset"),
    ("퇴직금(예상)",          0, "직접 입력",                "asset"),
    ("현금",                  0, "직접 입력",                "asset"),
    # ── 부채 ──────────────────────────────────────────────────
    ("대출잔액",              0, "디딤돌대출 등",            "liability"),
    ("기타부채",              0, "직접 입력",                "liability"),
]


def build(wb):
    ws = wb.create_sheet("자산현황")
    ws.sheet_view.showGridLines = False

    col_widths = {1: 4, 2: 20, 3: 18, 4: 24, 5: 4}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    for r in range(1, 40):
        ws.row_dimensions[r].height = 22

    # ── 제목 ─────────────────────────────────────────────────
    title = ws.cell(row=1, column=2, value="자산 현황")
    title.font = _font(bold=True, size=14, color="1F4E79")
    title.alignment = _align(h="left")
    ws.merge_cells("B1:D1")
    ws.row_dimensions[1].height = 30

    note = ws.cell(row=2, column=2,
                   value="※ 계좌 잔액, 주식, 코인, 퇴직금은 직접 입력하세요.")
    note.font = Font(size=9, color="888888", italic=True)
    note.alignment = _align(h="left")
    ws.merge_cells("B2:D2")

    # ── 헤더 ─────────────────────────────────────────────────
    header_row = 4
    for ci, (h, hc) in enumerate(zip(["항목", "금액", "비고"],
                                     ["1F4E79", "375623", "1F4E79"]), start=2):
        c = ws.cell(row=header_row, column=ci, value=h)
        c.fill = _fill(hc)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()

    # ── 자산 항목 ─────────────────────────────────────────────
    row = 5
    asset_rows, liability_rows = [], []

    _sub_header(ws, row, "[ 자산 ]", "375623")
    row += 1
    for item, default, note_text, cat in ASSET_ITEMS:
        if cat != "asset":
            continue
        c_item = ws.cell(row=row, column=2, value=item)
        c_amt  = ws.cell(row=row, column=3, value=default)
        c_note = ws.cell(row=row, column=4, value=note_text)

        fill_c = "E2EFDA" if len(asset_rows) % 2 == 0 else "FFFFFF"
        for c in (c_item, c_amt, c_note):
            c.fill = _fill(fill_c)
            c.alignment = _align()
            c.border = _border()
        c_amt.number_format = '#,##0"원"'
        asset_rows.append(row)
        row += 1

    total_asset_row = row
    _total_row(ws, row, asset_rows, "총 자산", "375623")
    row += 2

    _sub_header(ws, row, "[ 부채 ]", "C00000")
    row += 1
    for item, default, note_text, cat in ASSET_ITEMS:
        if cat != "liability":
            continue
        c_item = ws.cell(row=row, column=2, value=item)
        c_amt  = ws.cell(row=row, column=3, value=default)
        c_note = ws.cell(row=row, column=4, value=note_text)

        fill_c = "FCE4D6" if len(liability_rows) % 2 == 0 else "FFFFFF"
        for c in (c_item, c_amt, c_note):
            c.fill = _fill(fill_c)
            c.alignment = _align()
            c.border = _border()
        c_amt.number_format = '#,##0"원"'
        liability_rows.append(row)
        row += 1

    total_liability_row = row
    _total_row(ws, row, liability_rows, "총 부채", "C00000")
    row += 2

    # ── 순자산 ───────────────────────────────────────────────
    for ci in range(2, 5):
        c = ws.cell(row=row, column=ci)
        c.fill = _fill("1F4E79")
        c.font = _font(bold=True, color="FFFFFF", size=12)
        c.alignment = _align()
        c.border = _border()
    ws.cell(row=row, column=2, value="순 자산 (자산 - 부채)")
    ws.merge_cells(f"B{row}:B{row}")
    net = ws.cell(row=row, column=3,
                  value=f"=C{total_asset_row}-C{total_liability_row}")
    net.number_format = '#,##0"원"'
    net.font = _font(bold=True, size=12, color="FFFFFF")

    return ws


def _sub_header(ws, row, title, color):
    c = ws.cell(row=row, column=2, value=title)
    c.font = _font(bold=True, color="FFFFFF")
    c.fill = _fill(color)
    c.alignment = _align(h="left")
    c.border = _border()
    ws.merge_cells(f"B{row}:D{row}")


def _total_row(ws, row, data_rows, label, color):
    c_label = ws.cell(row=row, column=2, value=label)
    c_label.font = _font(bold=True, color="FFFFFF")
    c_label.fill = _fill(color)
    c_label.alignment = _align()
    c_label.border = _border()

    row_refs = "+".join(f"C{r}" for r in data_rows)
    c_total = ws.cell(row=row, column=3, value=f"={row_refs}")
    c_total.number_format = '#,##0"원"'
    c_total.font = _font(bold=True)
    c_total.fill = _fill("D6E4F0")
    c_total.alignment = _align()
    c_total.border = _border()

    c_note = ws.cell(row=row, column=4, value="")
    c_note.fill = _fill("D6E4F0")
    c_note.border = _border()
