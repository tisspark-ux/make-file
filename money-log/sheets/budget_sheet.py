"""수입&예산 시트 생성 모듈"""

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
    return Alignment(horizontal=h, vertical=v, wrap_text=True)


# ── 수입 항목 ────────────────────────────────────────────────
# (항목명, 비고, Tiss금액_만원, Tiss주기, Tiss대상월, JM금액_만원, JM주기, JM대상월)
INCOME_ITEMS = [
    ("월급",        "",         554.3,  "M", "",       0,    "",  ""),
    ("기타상여금",  "통신지원금",  5.52,  "M", "",       0,    "",  ""),
    ("기타상여금",  "인센티브",    0,     "Y", "",       0,    "",  "연초 확정 후 입력"),
    ("기타상여금",  "명절-설",     0,     "Y", "01",     0,    "",  ""),
    ("기타상여금",  "명절-추석",   0,     "Y", "09",     0,    "",  ""),
    ("기타상여금",  "하계휴가비",  0,     "Y", "07",     0,    "",  ""),
    ("복지포인트",  "",           0,     "Y", "",       0,    "",  ""),
    ("금융수입",    "",           0,     "M", "",       0,    "",  "이자/배당 등"),
]

# ── 예산(소비) 항목 ──────────────────────────────────────────
# (대분류, 중분류, 공동_만원, Tiss_만원, JM_만원, 주기, 납부월, 비고)
BUDGET_ITEMS = [
    ("고정비",    "교통비",        0,    8,    8,   "M", "",   ""),
    ("고정비",    "휴대폰/인터넷", 0,    2.2,  2.2, "M", "",   ""),
    ("고정비",    "보험/연금",     0,    23.1, 0,   "M", "",   "동양생명 등"),
    ("고정비",    "부모님",        0,    0,    0,   "M", "",   ""),
    ("고정비",    "기흥모임",      13,   0,    0,   "M", "",   ""),
    ("고정비",    "구독형서비스",  0,    0,    0,   "M", "",   "넷플릭스 등"),
    ("고정비",    "용돈",          0,    10,   8.5, "M", "",   ""),
    ("고정비",    "아파트관리비",  33,   0,    0,   "M", "",   ""),
    ("생활비",    "생활비",        35,   0,    0,   "M", "",   ""),
    ("의료/건강", "의료/건강",     0,    0,    0,   "M", "",   ""),
    ("모임",      "모임",          0,    0,    0,   "M", "",   ""),
    ("Tiss",      "개인",          0,    0,    0,   "M", "",   ""),
    ("JM",        "개인",          0,    0,    0,   "M", "",   ""),
    ("세금",      "재산세(주택)", 30,    0,    0,   "Y", "07", ""),
    ("세금",      "종합소득세",    0,    0,    0,   "Y", "05", ""),
    ("대출상환",  "디딤돌대출",    0,    0,    0,   "M", "",   "자동이체"),
    ("부동산",    "부동산",        0,    0,    0,   "Y", "",   ""),
    ("기타",      "기타",          0,    0,    0,   "M", "",   ""),
]

# ── 저축 목표 항목 ───────────────────────────────────────────
# (항목명, 금액_만원, 비고)
SAVINGS_ITEMS = [
    ("청약저축",    20,  ""),
    ("정기적금",     0,  ""),
    ("비상금 적립",  0,  ""),
    ("기타저축",     0,  ""),
]


def build(wb, year: int = 2026):
    ws = wb.create_sheet("수입&예산")
    ws.sheet_view.showGridLines = False

    col_widths = {1: 3, 2: 14, 3: 16, 4: 10, 5: 10, 6: 10, 7: 6,
                  8: 8, 9: 10, 10: 10, 11: 6, 12: 8, 13: 6, 14: 20}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    for r in range(1, 100):
        ws.row_dimensions[r].height = 20

    row = 2

    # ════════════════════════════════
    # [수입] 섹션
    # ════════════════════════════════
    row = _section_header(ws, row, "수입", year)
    row += 1

    income_headers = ["항목", "비고", "Tiss 금액", "주기", "대상월",
                      "JM 금액", "주기", "대상월", "", "월수입(Tiss)", "월수입(JM)", "월수입(합계)", "연수입(합계)"]
    for ci, h in enumerate(income_headers, start=2):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill("5B7490")
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()
    row += 1

    income_data_start = row
    for i, (item, note, tiss_amt, tiss_cycle, tiss_month,
            jm_amt, jm_cycle, jm_month) in enumerate(INCOME_ITEMS):
        fill_color = "F0F4F8" if i % 2 == 0 else "FFFFFF"
        vals = [item, note, tiss_amt or "", tiss_cycle, tiss_month,
                jm_amt or "", jm_cycle, jm_month]
        for ci, val in enumerate(vals, start=2):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(fill_color)
            c.alignment = _align()
            c.border = _border()
            if ci in (4, 7) and val:
                c.number_format = '#,##0.0"만"'
        r = row
        ws.cell(row=r, column=11,
                value=f'=IF(E{r}="M",D{r},0)').number_format = '#,##0.0"만"'
        ws.cell(row=r, column=12,
                value=f'=IF(H{r}="M",G{r},0)').number_format = '#,##0.0"만"'
        ws.cell(row=r, column=13,
                value=f'=K{r}+L{r}').number_format = '#,##0.0"만"'
        ws.cell(row=r, column=14,
                value=f'=IF(E{r}="M",D{r}*12,D{r})+IF(H{r}="M",G{r}*12,G{r})').number_format = '#,##0.0"만"'
        for ci in (11, 12, 13, 14):
            ws.cell(row=r, column=ci).fill = _fill(fill_color)
            ws.cell(row=r, column=ci).alignment = _align()
            ws.cell(row=r, column=ci).border = _border()
        row += 1

    income_data_end = row - 1
    _total_row(ws, row, income_data_start, income_data_end,
               sum_cols=[4, 7, 11, 12, 13, 14], label="월수입 합계")
    income_total_row = row
    row += 2

    # ════════════════════════════════
    # [예산(소비)] 섹션
    # ════════════════════════════════
    row = _section_header(ws, row, "예산(소비)", year)
    row += 1

    budget_headers = ["대분류", "중분류", "공동", "Tiss", "JM",
                      "주기", "합계(월)", "납부월", "비고"]
    for ci, h in enumerate(budget_headers, start=2):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill("4D7063")
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()
    row += 1

    budget_data_start = row
    for i, (main_cat, sub_cat, common, tiss, jm, cycle, pay_month, note) in enumerate(BUDGET_ITEMS):
        fill_color = "EBF2EE" if i % 2 == 0 else "FFFFFF"
        vals = [main_cat, sub_cat, common or "", tiss or "", jm or "",
                cycle, "", pay_month, note]
        for ci, val in enumerate(vals, start=2):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(fill_color)
            c.alignment = _align()
            c.border = _border()
            if ci in (4, 5, 6) and val:
                c.number_format = '#,##0.0"만"'
        r = row
        ws.cell(row=r, column=8,
                value=f'=IF(G{r}="M",D{r}+E{r}+F{r},(D{r}+E{r}+F{r})/12)').number_format = '#,##0.0"만"'
        ws.cell(row=r, column=8).fill = _fill(fill_color)
        ws.cell(row=r, column=8).alignment = _align()
        ws.cell(row=r, column=8).border = _border()
        row += 1

    budget_data_end = row - 1
    _total_row(ws, row, budget_data_start, budget_data_end,
               sum_cols=[4, 5, 6, 8], label="월지출 예산 합계")
    budget_total_row = row
    row += 2

    # ── 순수입 요약 ───────────────────────────────────────────
    순수입_row = row
    c = ws.cell(row=row, column=2, value="월 순수입 (수입-지출)")
    c.font = _font(bold=True, color="FFFFFF")
    c.fill = _fill("8C5858")
    c.alignment = _align()
    c.border = _border()
    ws.merge_cells(f"B{row}:G{row}")
    c2 = ws.cell(row=row, column=8,
                 value=f"=K{income_total_row}-H{budget_total_row}")
    c2.number_format = '#,##0.0"만"'
    c2.font = _font(bold=True)
    c2.fill = _fill("F5ECEC")
    c2.alignment = _align()
    c2.border = _border()
    row += 3

    # ════════════════════════════════
    # ③ [저축 목표] 섹션
    # ════════════════════════════════
    row = _section_header(ws, row, "저축 목표", year)
    row += 1

    savings_headers = ["항목", "금액(만원)", "비고"]
    for ci, h in enumerate(savings_headers, start=2):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill("6B7B8D")
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()
    row += 1

    savings_data_start = row
    for i, (item, amount, note) in enumerate(SAVINGS_ITEMS):
        fill_color = "F5F2EA" if i % 2 == 0 else "FFFFFF"
        for ci, val in [(2, item), (3, amount or 0), (4, note)]:
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(fill_color)
            c.alignment = _align()
            c.border = _border()
            if ci == 3:
                c.number_format = '#,##0.0"만"'
        row += 1

    savings_data_end = row - 1

    # 저축 합계 행
    for ci in range(2, 12):
        c = ws.cell(row=row, column=ci)
        c.fill = _fill("6B7B8D")
        c.font = _font(bold=True, color="FFFFFF")
        c.alignment = _align()
        c.border = _border()
    ws.cell(row=row, column=2, value="월 저축 합계")
    ws.merge_cells(f"B{row}:C{row}")
    savings_total_cell = ws.cell(row=row, column=4,
                                 value=f"=SUM(C{savings_data_start}:C{savings_data_end})")
    savings_total_cell.number_format = '#,##0.0"만"'
    savings_total_cell.font = _font(bold=True)
    savings_total_cell.fill = _fill("EDE8D5")
    savings_total_cell.alignment = _align()
    savings_total_cell.border = _border()
    savings_total_row = row
    row += 2

    # 월 순잉여금 (순수입 - 저축)
    c = ws.cell(row=row, column=2, value="월 순잉여금 (저축 후 남는 돈)")
    c.font = _font(bold=True, color="FFFFFF")
    c.fill = _fill("3D5470")
    c.alignment = _align()
    c.border = _border()
    ws.merge_cells(f"B{row}:G{row}")
    c2 = ws.cell(row=row, column=8,
                 value=f"=H{순수입_row}-D{savings_total_row}")
    c2.number_format = '#,##0.0"만"'
    c2.font = _font(bold=True)
    c2.fill = _fill("DAE4EF")
    c2.alignment = _align()
    c2.border = _border()

    return ws


# ── 내부 헬퍼 ─────────────────────────────────────────────────
def _section_header(ws, row, title, year):
    c = ws.cell(row=row, column=2, value=f"[ {title} ]  {year}년  (단위: 만원)")
    c.font = _font(bold=True, size=13, color="FFFFFF")
    c.fill = _fill("3D5470")
    c.alignment = _align(h="left")
    c.border = _border()
    ws.merge_cells(f"B{row}:N{row}")
    ws.row_dimensions[row].height = 24
    return row


def _total_row(ws, row, start, end, sum_cols, label):
    c = ws.cell(row=row, column=2, value=label)
    c.font = _font(bold=True, color="FFFFFF")
    c.fill = _fill("3D5470")
    c.alignment = _align()
    c.border = _border()
    ws.merge_cells(f"B{row}:C{row}")
    for ci in range(4, 15):
        if ci in sum_cols:
            cell = ws.cell(row=row, column=ci,
                           value=f"=SUM({get_column_letter(ci)}{start}:{get_column_letter(ci)}{end})")
            cell.number_format = '#,##0.0"만"'
            cell.font = _font(bold=True)
            cell.fill = _fill("DAE4EF")
        else:
            cell = ws.cell(row=row, column=ci, value="")
            cell.fill = _fill("DAE4EF")
        cell.alignment = _align()
        cell.border = _border()
