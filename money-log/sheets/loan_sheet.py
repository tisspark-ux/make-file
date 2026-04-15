"""대출 상환 추적 시트 생성 모듈

대출 기본 정보:
- 대출금: 480,000,000원 / 실행일: 2026-03-30 / 만기: 30년(360회)
- 상환: 원리금균등 / 현재 금리: 4.07% (6개월 변동)
"""

from datetime import date
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# ── 대출 기본 상수 ────────────────────────────────────────────
LOAN_AMT      = 480_000_000
INITIAL_RATE  = 0.0407
BASE_RATE     = 0.0247   # 기준금리 (참고)
SPREAD_RATE   = 0.0160   # 가산금리 (참고, 고정)
LOAN_DATE     = date(2026, 3, 30)
LOAN_MONTHS   = 360

# ── 행 번호 상수 ──────────────────────────────────────────────
INFO_AMT_ROW   = 5     # B5: 대출금액
INFO_DATE_ROW  = 6     # B6: 대출실행일
INFO_MONTHS_ROW= 7     # B7: 만기
INFO_METHOD_ROW= 8     # B8: 상환방법
INFO_SPREAD_ROW= 9     # B9: 가산금리

RATE_SEC_ROW   = 11    # 금리 변동 이력 섹션 헤더
RATE_HINT_ROW  = 12
RATE_HDR_ROW   = 13    # 금리 테이블 컬럼 헤더
RATE_DATA_START= 14    # 금리 데이터 시작 (행 14 = 초기 금리)
RATE_DATA_END  = 53    # 금리 데이터 끝  (최대 40개 항목)

SUMM_SEC_ROW   = 55    # 요약 섹션 헤더
SUMM_HINT_ROW  = 56
SUMM_DATA_ROW  = 57    # 요약 지표 시작 (4행)

SCHED_SEC_ROW  = 63    # 스케줄 섹션 헤더
SCHED_HINT_ROW = 64
SCHED_HDR_ROW  = 66    # 스케줄 컬럼 헤더
SCHED_START    = 67    # 스케줄 데이터 시작 (회차 1)
SCHED_END      = SCHED_START + LOAN_MONTHS - 1  # = 426


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


# 금리 테이블 열 배치: A(col1)=날짜, B(col2)=금리, C:J=비고
# 기본 정보 열 배치: A:B(merged)=레이블, C:D(merged)=값, E:J=비고
_RATE_DATE_RNG = f"$A${RATE_DATA_START}:$A${RATE_DATA_END}"  # 날짜 열 = A
_RATE_VAL_RNG  = f"$B${RATE_DATA_START}:$B${RATE_DATA_END}"  # 금리 열 = B
_INFO_AMT_CELL = f"$C${INFO_AMT_ROW}"    # 대출금액 값 셀 = C5
_INFO_DATE_CELL= f"$C${INFO_DATE_ROW}"   # 실행일 값 셀   = C6


def build(wb):
    ws = wb.create_sheet("대출상환")
    ws.sheet_view.showGridLines = False

    # ── 컬럼 너비 ─────────────────────────────────────────────
    col_widths = {
        1: 8,   # A: 회차
        2: 16,  # B: 납입일 / 레이블
        3: 14,  # C: 적용금리 / 값
        4: 16,  # D: 월납입액 / 값(우측)
        5: 14,  # E: 이자 / 레이블2
        6: 14,  # F: 원금상환 / 값2
        7: 14,  # G: 중도상환 (입력)
        8: 18,  # H: 잔액
        9: 8,   # I: 납입 (O 입력)
        10: 24, # J: 비고
    }
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    row = 1

    # ════════════════════════════════════════════════════════
    # ① 제목
    # ════════════════════════════════════════════════════════
    ws.row_dimensions[row].height = 32
    t = ws.cell(row=row, column=1, value="대출 상환 추적표")
    t.font = Font(name="Calibri Light", bold=False, size=16, color="1A2B4A")
    t.alignment = _align(h="left")
    ws.merge_cells("A1:J1")

    row = 2
    ws.row_dimensions[row].height = 14
    sub = ws.cell(row=row, column=1,
                  value=f"기흥역센트럴푸르지오 / 주택자금대출(디딤돌) / {LOAN_DATE.strftime('%Y-%m-%d')} 실행")
    sub.font = Font(size=9, color="888888", italic=True)
    sub.alignment = _align(h="left")
    ws.merge_cells("A2:J2")

    # ════════════════════════════════════════════════════════
    # ② 기본 정보
    # ════════════════════════════════════════════════════════
    _section_title(ws, 4, "기본 정보")

    info_rows = [
        (INFO_AMT_ROW,    "대출금액",      LOAN_AMT,     '#,##0"원"',   "대출원금"),
        (INFO_DATE_ROW,   "대출실행일",    LOAN_DATE,    "YYYY-MM-DD",  "첫 납입: 2026-04-30"),
        (INFO_MONTHS_ROW, "만기",          f"{LOAN_MONTHS//12}년 ({LOAN_MONTHS}회)", None, ""),
        (INFO_METHOD_ROW, "상환방법",      "원리금균등",  None,          "고정납입액 (금리 변동 시 재계산)"),
        (INFO_SPREAD_ROW, "가산금리 (참고)", SPREAD_RATE, '0.00%',      f"기준금리 현재 {BASE_RATE:.2%} + 가산 {SPREAD_RATE:.2%} = {INITIAL_RATE:.2%}"),
    ]
    for r, label, val, fmt, note in info_rows:
        ws.row_dimensions[r].height = 22
        lc = ws.cell(row=r, column=1, value=label)
        lc.fill = _fill("2471A3")
        lc.font = _font(bold=True, color="FFFFFF", size=10)
        lc.alignment = _align(h="left")
        lc.border = _border()
        ws.merge_cells(f"A{r}:B{r}")

        vc = ws.cell(row=r, column=3, value=val)
        vc.fill = _fill("F4F8FC")
        vc.font = _font(bold=True, size=11)
        vc.alignment = _align()
        vc.border = _border()
        if fmt:
            vc.number_format = fmt
        ws.merge_cells(f"C{r}:D{r}")

        nc = ws.cell(row=r, column=5, value=note)
        nc.fill = _fill("F4F8FC")
        nc.font = Font(size=9, color="888888", italic=True)
        nc.alignment = _align(h="left")
        nc.border = _border()
        ws.merge_cells(f"E{r}:J{r}")

    # ════════════════════════════════════════════════════════
    # ③ 금리 변동 이력 (사용자 입력)
    # ════════════════════════════════════════════════════════
    _section_title(ws, RATE_SEC_ROW, "금리 변동 이력  (변동 시 행 추가 — 날짜 오름차순 유지)")

    ws.row_dimensions[RATE_HINT_ROW].height = 14
    hint1 = ws.cell(row=RATE_HINT_ROW, column=1,
                    value="※ 총금리를 직접 입력하세요. 날짜는 오름차순 정렬 필수. "
                          "과거 회차는 소급 변경되지 않습니다.")
    hint1.font = Font(size=9, color="888888", italic=True)
    hint1.alignment = _align(h="left")
    ws.merge_cells(f"A{RATE_HINT_ROW}:J{RATE_HINT_ROW}")

    # 금리 테이블 헤더
    ws.row_dimensions[RATE_HDR_ROW].height = 22
    rate_headers = ["변동일", "총금리 (%)", "비고 (기준금리·변동 사유 등)"]
    rate_hcolors = ["1A2B4A", "1E8449", "1A2B4A"]
    rate_spans   = ["A", "B", "C:J"]  # 열 병합 범위
    for (col_start, span_end), h, hc in zip(
            [("A","A"),("B","B"),("C","J")], rate_headers, rate_hcolors):
        c = ws.cell(row=RATE_HDR_ROW, column=ord(col_start)-64, value=h)
        c.fill = _fill(hc)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()
        if col_start != span_end:
            ws.merge_cells(f"{col_start}{RATE_HDR_ROW}:{span_end}{RATE_HDR_ROW}")

    # 금리 데이터 행 (행 14 = 초기 금리, 행 15-53 = 미래 입력용)
    for i in range(RATE_DATA_END - RATE_DATA_START + 1):
        r = RATE_DATA_START + i
        ws.row_dimensions[r].height = 20
        fill_c = "E8F8F0" if i == 0 else ("F8FAFB" if i % 2 == 0 else "FFFFFF")

        # 날짜 열
        date_cell = ws.cell(row=r, column=1,
                            value=LOAN_DATE if i == 0 else None)
        date_cell.fill = _fill(fill_c)
        date_cell.alignment = _align()
        date_cell.border = _border()
        if i == 0:
            date_cell.number_format = "YYYY-MM-DD"

        # 금리 열
        rate_cell = ws.cell(row=r, column=2,
                            value=INITIAL_RATE if i == 0 else None)
        rate_cell.fill = _fill(fill_c)
        rate_cell.alignment = _align()
        rate_cell.border = _border()
        rate_cell.number_format = '0.00%'

        # 비고 열
        note_val = (f"초기 설정 (기준 {BASE_RATE:.2%} + 가산 {SPREAD_RATE:.2%})"
                    if i == 0 else None)
        note_cell = ws.cell(row=r, column=3, value=note_val)
        note_cell.fill = _fill(fill_c)
        note_cell.alignment = _align(h="left")
        note_cell.border = _border()
        ws.merge_cells(f"C{r}:J{r}")

    # ════════════════════════════════════════════════════════
    # ④ 현재 상환 현황 요약
    # ════════════════════════════════════════════════════════
    _section_title(ws, SUMM_SEC_ROW, "현재 상환 현황 요약")

    ws.row_dimensions[SUMM_HINT_ROW].height = 14
    hint2 = ws.cell(row=SUMM_HINT_ROW, column=1,
                    value="※ I열에 'O' 입력 시 납입 완료 처리. G열에 금액 입력 시 중도상환 반영.")
    hint2.font = Font(size=9, color="888888", italic=True)
    hint2.alignment = _align(h="left")
    ws.merge_cells(f"A{SUMM_HINT_ROW}:J{SUMM_HINT_ROW}")

    _CNTIF = f'COUNTIF($I${SCHED_START}:$I${SCHED_END},"O")'
    summary_metrics = [
        # (row_offset, label1, formula1, fmt1, label2, formula2, fmt2)
        (0, "납입 완료",    f"={_CNTIF}",          '0"회 / 360회"',
             "남은 회차",   f"=360-{_CNTIF}",       '0"회"'),
        (1, "현재 잔여 원금",
             f"=IF({_CNTIF}=0,$C${INFO_AMT_ROW},"
             f"OFFSET($H${SCHED_START},{_CNTIF}-1,0))",
             '#,##0"원"',
             "총 이자 납부",
             f"=SUMPRODUCT(($I${SCHED_START}:$I${SCHED_END}=\"O\")*$E${SCHED_START}:$E${SCHED_END})",
             '#,##0"원"'),
        (2, "완료율",       f"={_CNTIF}/360",        '0.0%',
             "현재 적용 금리",
             f"=IFERROR(INDEX({_RATE_VAL_RNG},"
             f"MATCH(MAXIFS({_RATE_DATE_RNG},{_RATE_DATE_RNG},"
             f"\"<=\"&TODAY()),{_RATE_DATE_RNG},0)),"
             f"$B${RATE_DATA_START})",
             '0.00%'),
        (3, "총 중도상환",  f"=SUM($G${SCHED_START}:$G${SCHED_END})", '#,##0"원"',
             "완납 예정일",
             f'=IFERROR(INDEX($B${SCHED_START}:$B${SCHED_END},'
             f'MATCH("",$I${SCHED_START}:$I${SCHED_END},0)),"완납")',
             "YYYY-MM-DD"),
    ]

    for offset, l1, f1, fmt1, l2, f2, fmt2 in summary_metrics:
        r = SUMM_DATA_ROW + offset
        ws.row_dimensions[r].height = 26

        # 레이블1
        lc1 = ws.cell(row=r, column=1, value=l1)
        lc1.fill = _fill("1A2B4A")
        lc1.font = _font(bold=True, color="FFFFFF", size=10)
        lc1.alignment = _align()
        lc1.border = _border()
        ws.merge_cells(f"A{r}:B{r}")

        # 값1
        vc1 = ws.cell(row=r, column=3, value=f1)
        vc1.fill = _fill("D6EAF8")
        vc1.font = _font(bold=True, size=12)
        vc1.alignment = _align()
        vc1.border = _border()
        vc1.number_format = fmt1
        ws.merge_cells(f"C{r}:D{r}")

        # 레이블2
        lc2 = ws.cell(row=r, column=5, value=l2)
        lc2.fill = _fill("1E8449")
        lc2.font = _font(bold=True, color="FFFFFF", size=10)
        lc2.alignment = _align()
        lc2.border = _border()
        ws.merge_cells(f"E{r}:F{r}")

        # 값2
        vc2 = ws.cell(row=r, column=7, value=f2)
        vc2.fill = _fill("E8F8F0")
        vc2.font = _font(bold=True, size=12)
        vc2.alignment = _align()
        vc2.border = _border()
        vc2.number_format = fmt2
        ws.merge_cells(f"G{r}:J{r}")

    # ════════════════════════════════════════════════════════
    # ⑤ 월별 상환 스케줄
    # ════════════════════════════════════════════════════════
    _section_title(ws, SCHED_SEC_ROW, "월별 상환 스케줄  (360회)")

    ws.row_dimensions[SCHED_HINT_ROW].height = 14
    hint3 = ws.cell(row=SCHED_HINT_ROW, column=1,
                    value="※ C열(적용금리)은 금리 이력에서 자동 조회. "
                          "G열에 중도상환 금액 직접 입력. I열에 'O' 입력 시 납입 완료.")
    hint3.font = Font(size=9, color="888888", italic=True)
    hint3.alignment = _align(h="left")
    ws.merge_cells(f"A{SCHED_HINT_ROW}:J{SCHED_HINT_ROW}")

    ws.row_dimensions[65].height = 4  # 헤더 위 소 여백 행

    # 스케줄 컬럼 헤더
    ws.row_dimensions[SCHED_HDR_ROW].height = 22
    sched_headers = ["회차", "납입일", "적용금리", "월납입액", "이자",
                     "원금상환", "중도상환", "잔액", "납입", "비고"]
    sched_hcolors = ["1A2B4A","1A2B4A","1E8449","1E8449","2471A3",
                     "2471A3","C0392B","1A2B4A","2471A3","4A5568"]
    for ci, (h, hc) in enumerate(zip(sched_headers, sched_hcolors), start=1):
        c = ws.cell(row=SCHED_HDR_ROW, column=ci, value=h)
        c.fill = _fill(hc)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _align()
        c.border = _border()

    # ── 데이터 행 생성 ──────────────────────────────────────
    for i in range(LOAN_MONTHS):
        r = SCHED_START + i
        fill_c = "F8FAFB" if i % 2 == 0 else "FFFFFF"
        ws.row_dimensions[r].height = 18

        # A: 회차
        ws.cell(row=r, column=1, value=i + 1).fill = _fill(fill_c)
        ws.cell(row=r, column=1).alignment = _align()
        ws.cell(row=r, column=1).border = _border()

        # B: 납입일 (실행일 기준 n번째 달의 30일, 없으면 말일)
        date_f = (f"=DATE(YEAR($C${INFO_DATE_ROW}),"
                  f"MONTH($C${INFO_DATE_ROW})+A{r},"
                  f"MIN(30,DAY(EOMONTH(DATE(YEAR($C${INFO_DATE_ROW}),"
                  f"MONTH($C${INFO_DATE_ROW})+A{r},1),0))))")
        bc = ws.cell(row=r, column=2, value=date_f)
        bc.number_format = "YYYY-MM-DD"
        bc.fill = _fill(fill_c)
        bc.alignment = _align()
        bc.border = _border()

        # C: 적용금리 (금리 이력에서 자동 조회)
        rate_f = (f"=IFERROR(INDEX({_RATE_VAL_RNG},"
                  f"MATCH(MAXIFS({_RATE_DATE_RNG},{_RATE_DATE_RNG},"
                  f"\"<=\"&B{r}),{_RATE_DATE_RNG},0)),"
                  f"$B${RATE_DATA_START})")
        cc = ws.cell(row=r, column=3, value=rate_f)
        cc.number_format = '0.00%'
        cc.fill = _fill(fill_c)
        cc.alignment = _align()
        cc.border = _border()

        # 이전 잔액 참조 (row 1: 대출금액, 그 이후: 이전 행 H열)
        prev_bal = f"$C${INFO_AMT_ROW}" if i == 0 else f"H{r - 1}"

        # D: 월납입액 PMT (남은 회차·이전 잔액 기준 재계산)
        remaining = LOAN_MONTHS - i  # 첫 행: 360, 마지막 행: 1
        pmt_f = (f"=IFERROR(IF({prev_bal}<=0,0,"
                 f"ABS(PMT(C{r}/12,{remaining},-{prev_bal}))),0)")
        dc = ws.cell(row=r, column=4, value=pmt_f)
        dc.number_format = '#,##0"원"'
        dc.fill = _fill(fill_c)
        dc.alignment = _align()
        dc.border = _border()

        # E: 이자 = 이전잔액 × 금리 / 12
        ec = ws.cell(row=r, column=5,
                     value=f"=IF({prev_bal}<=0,0,{prev_bal}*C{r}/12)")
        ec.number_format = '#,##0"원"'
        ec.fill = _fill(fill_c)
        ec.alignment = _align()
        ec.border = _border()

        # F: 원금상환 = 월납입액 - 이자
        fc = ws.cell(row=r, column=6,
                     value=f"=IF(D{r}=0,0,D{r}-E{r})")
        fc.number_format = '#,##0"원"'
        fc.fill = _fill(fill_c)
        fc.alignment = _align()
        fc.border = _border()

        # G: 중도상환 (사용자 직접 입력, 기본 공백)
        gc = ws.cell(row=r, column=7, value=None)
        gc.fill = _fill("FFFFFF")  # 연한 노랑 = 입력 유도
        gc.alignment = _align()
        gc.border = _border()
        gc.number_format = '#,##0"원"'

        # H: 잔액 = max(이전잔액 - 원금상환 - 중도상환, 0)
        hc = ws.cell(row=r, column=8,
                     value=f"=MAX({prev_bal}-F{r}-IFERROR(G{r},0),0)")
        hc.number_format = '#,##0"원"'
        hc.font = _font(bold=True)
        hc.fill = _fill(fill_c)
        hc.alignment = _align()
        hc.border = _border()

        # I: 납입 완료 (사용자 입력: "O" 또는 공백)
        ic = ws.cell(row=r, column=9, value=None)
        ic.fill = _fill(fill_c)
        ic.alignment = _align()
        ic.border = _border()

        # J: 비고 (사용자 입력)
        jc = ws.cell(row=r, column=10, value=None)
        jc.fill = _fill(fill_c)
        jc.alignment = _align(h="left")
        jc.border = _border()

    # ── 조건부 서식 ──────────────────────────────────────────
    sched_range = f"A{SCHED_START}:J{SCHED_END}"

    # 납입 완료 (I열 = "O") → 녹색
    paid_fill = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
    ws.conditional_formatting.add(
        sched_range,
        FormulaRule(formula=[f'=$I{SCHED_START}="O"'], fill=paid_fill)
    )

    # 미납 경고 (납입일이 오늘 이전 + 미납입) → 연한 빨강
    overdue_fill = PatternFill(start_color="FFD0D0", end_color="FFD0D0", fill_type="solid")
    ws.conditional_formatting.add(
        sched_range,
        FormulaRule(
            formula=[f'=AND(B{SCHED_START}<TODAY(),$I{SCHED_START}<>"O")'],
            fill=overdue_fill
        )
    )

    ws.freeze_panes = f"A{SCHED_START}"
    return ws


def _section_title(ws, row, title):
    c = ws.cell(row=row, column=1, value=title)
    c.font = _font(bold=True, size=12, color="FFFFFF")
    c.fill = _fill("1A2B4A")
    c.alignment = _align(h="left")
    c.border = _border()
    ws.merge_cells(f"A{row}:J{row}")
    ws.row_dimensions[row].height = 24
