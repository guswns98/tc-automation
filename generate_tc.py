import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
import os, sys, importlib.util

# ──────────────────────────────────────────
# tc_data 파일 로드
# ──────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python generate_tc.py tc_data/<화면명>.py")
    sys.exit(1)

tc_file = sys.argv[1]
spec = importlib.util.spec_from_file_location("tc_module", tc_file)
mod  = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

screen_name = mod.screen_name
tc_data     = mod.tc_data

# ──────────────────────────────────────────
# 스타일 헬퍼
# ──────────────────────────────────────────
def fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def border_thin():
    s = Side(border_style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

FILL_HEADER = fill("BFBFBF")
FILL_NONE   = fill("FFFFFF")

FONT_TITLE  = Font(name="Arial", size=18, bold=True, color="002060")
FONT_HEADER = Font(name="Arial", size=9,  bold=True)
FONT_DATA   = Font(name="Arial", size=9)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ──────────────────────────────────────────
# 워크북 생성
# ──────────────────────────────────────────
output_dir = r"C:\Users\rkd98\QA\TC_Automation\output"
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, f"QA_{screen_name}_TC.xlsx")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = f"{screen_name} TC"

# ── 컬럼 너비 ──
col_widths = {
    "A": 13.75, "B": 20.25, "C": 22.13, "D": 22.75,
    "E": 24.0,  "F": 22.13, "G": 8.25,  "H": 32.38,
    "I": 12.63, "J": 10.0,  "K": 10.0,  "L": 13.63,
    "M": 43.5,
}
for col_letter, width in col_widths.items():
    ws.column_dimensions[col_letter].width = width

# ──────────────────────────────────────────
# 행 1: 우상단 요약 헤더
# ──────────────────────────────────────────
for col, val in [(18, "Priority"), (19, "개수")]:
    c = ws.cell(row=1, column=col, value=val)
    c.fill = fill("D9D9D9")
    c.font = Font(name="Arial", size=9, bold=True)
    c.alignment = ALIGN_CENTER
    c.border = border_thin()

# ──────────────────────────────────────────
# 행 2: 제목 + 우상단 P1 집계
# ──────────────────────────────────────────
ws.merge_cells("A2:M2")
c = ws.cell(row=2, column=1, value=screen_name)
c.font = FONT_TITLE
c.alignment = ALIGN_LEFT

for col, val in [(18, "P1"), (19, "=COUNTIF(G:G,R2)")]:
    c = ws.cell(row=2, column=col, value=val)
    c.font = FONT_DATA
    c.alignment = ALIGN_CENTER
    c.border = border_thin()

# ──────────────────────────────────────────
# 행 3~4: 컬럼 헤더
# ──────────────────────────────────────────
single_headers = [
    ("A", "1 Depth"), ("B", "2 Depth"), ("C", "3 Depth"), ("D", "4 Depth"),
    ("E", "5 Depth"), ("F", "6 Depth"), ("G", "Priority"),
    ("H", "Expected Result"), ("M", "합계"),
]
for col_letter, val in single_headers:
    col = ord(col_letter) - ord("A") + 1
    ws.merge_cells(f"{col_letter}3:{col_letter}4")
    c = ws.cell(row=3, column=col, value=val)
    c.fill = FILL_HEADER
    c.font = FONT_HEADER
    c.alignment = ALIGN_CENTER
    c.border = border_thin()

ws.merge_cells("I3:L3")
c = ws.cell(row=3, column=9, value="Test Result")
c.fill = FILL_HEADER
c.font = FONT_HEADER
c.alignment = ALIGN_CENTER
c.border = border_thin()

sub_headers = [(9, "Android"), (10, "iOS"), (11, "Date"), (12, "BTS ID")]
for col, val in sub_headers:
    c = ws.cell(row=4, column=col, value=val)
    c.fill = FILL_HEADER
    c.font = FONT_HEADER
    c.alignment = ALIGN_CENTER
    c.border = border_thin()

for row, p_val, cnt_val in [(3, "P2", "=COUNTIF(G:G,R3)"), (4, "P3", "=COUNTIF(G:G,R4)")]:
    ws.cell(row=row, column=18, value=p_val).font = FONT_DATA
    ws.cell(row=row, column=19, value=cnt_val).font = FONT_DATA
    for col in (18, 19):
        ws.cell(row=row, column=col).alignment = ALIGN_CENTER
        ws.cell(row=row, column=col).border = border_thin()

ws.cell(row=3, column=17, value="=SUM(S2:S4)").font = FONT_DATA
ws.freeze_panes = "A5"

# ──────────────────────────────────────────
# 데이터 행 (행 5부터)
# ──────────────────────────────────────────
for row_idx, row in enumerate(tc_data, start=5):
    d1, d2, d3, d4, d5, d6, priority, expected = row

    for col_idx, val in enumerate([d1, d2, d3, d4, d5, d6, priority, expected], start=1):
        c = ws.cell(row=row_idx, column=col_idx, value=val if val else None)
        c.font = FONT_DATA
        c.alignment = ALIGN_CENTER if col_idx == 7 else ALIGN_LEFT
        c.border = border_thin()

    for col_idx in range(9, 13):
        c = ws.cell(row=row_idx, column=col_idx)
        c.font = FONT_DATA
        c.border = border_thin()

    ws.row_dimensions[row_idx].height = 30

ws.auto_filter.ref = "A4:L4"

# ──────────────────────────────────────────
# A~D열 (1~4 Depth) 연속 동일 값 셀 병합
# 상위 Depth가 같을 때만 하위 Depth 병합
# ──────────────────────────────────────────
data_start = 5
data_end   = 4 + len(tc_data)
col_letters = ["A", "B", "C", "D"]  # 1~4 Depth

for col_idx, col_letter in enumerate(col_letters, start=1):
    merge_start = data_start
    for r in range(data_start + 1, data_end + 2):
        # 현재 행과 병합 시작 행의 1~col_idx 값이 모두 같아야 병합 유지
        same_group = all(
            ws.cell(row=r, column=c).value == ws.cell(row=merge_start, column=c).value
            for c in range(1, col_idx + 1)
        )

        if not same_group or r == data_end + 1:
            if r - 1 > merge_start:
                ws.merge_cells(f"{col_letter}{merge_start}:{col_letter}{r - 1}")
                c = ws.cell(row=merge_start, column=col_idx)
                c.alignment = ALIGN_CENTER
                c.font      = FONT_DATA
                c.border    = border_thin()
            merge_start = r

# ──────────────────────────────────────────
# 데이터 유효성 검사: Android(I), iOS(J) 컬럼 → P / F / NA / NT
# ──────────────────────────────────────────
last_row = 4 + len(tc_data)
dv = DataValidation(
    type="list",
    formula1='"P,F,NA,NT"',
    allow_blank=True,
    showDropDown=False,
)
dv.sqref = f"I5:J{last_row}"
ws.add_data_validation(dv)

# ──────────────────────────────────────────
# 조건부 서식: F 값 → 연한 빨간색 배경
# ──────────────────────────────────────────
red_fill = PatternFill(fill_type="solid", fgColor="FFCCCC")
ws.conditional_formatting.add(
    f"I5:J{last_row}",
    CellIsRule(operator="equal", formula=['"F"'], fill=red_fill),
)

wb.save(file_path)
print(f"저장 완료: {file_path}")
print(f"총 TC: {len(tc_data)}개")
