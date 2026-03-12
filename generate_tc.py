import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, copy

# ──────────────────────────────────────────
# TC 데이터 (1Depth~6Depth, Priority, Expected Result)
# ──────────────────────────────────────────
# (1D, 2D, 3D, 4D, 5D, 6D, Priority, Expected Result)
tc_data = [
    # ── 사용권 업그레이드_기간권 이용 ──
    ("사용권 업그레이드", "기간권 이용", "화면 진입", "", "", "", "P1", "사용권 업그레이드 화면이 정상 노출된다"),
    ("", "", "Top Navigation", "뒤로가기 버튼 탭", "", "", "P1", "이전 화면으로 이동한다"),
    ("", "", "Tabs", "주식 실전 투자권 탭 활성 상태 확인", "", "", "P2", "주식 실전 투자권 탭이 활성(볼드·언더라인) 상태로 노출된다"),
    ("", "", "", "코인 실전 투자권 탭 탭", "", "", "P2", "코인 실전 투자권 탭으로 전환되며 해당 탭이 활성화된다"),
    ("", "", "내 투자권 카드", "카드 정보 노출 확인", "", "", "P2", "'내 주식 실전 투자권 (기간권)' 타이틀과 '주식 프리미엄 1개월권 · 32일 후 만료' 텍스트가 노출된다"),
    ("", "", "", "사용권함 이동 버튼 탭", "", "", "P1", "사용권함 화면으로 이동한다"),
    ("", "", "안내 문구", "업그레이드 조건 주의문구 노출 확인", "", "", "P3", "'* 이용 중인 투자권보다 높은 금액의 사용권으로만 업그레이드 가능...' 주의문구가 노출된다"),
    ("", "", "서브탭", "1개월권 탭 선택 상태 확인", "", "", "P2", "1개월권 탭이 선택(흰 배경) 상태로 노출되고 해당 상품 목록이 표시된다"),
    ("", "", "", "1년권 탭 탭", "", "", "P2", "1년권 탭이 활성화되며 1년권 상품 목록으로 전환된다"),
    ("", "", "업그레이드 투자권 목록", "목록 노출 확인 (기간권 계좌 2개 이용 중)", "", "", "P1", "계좌 2개·3개·4개·5개 투자권 카드가 순서대로 노출된다"),
    ("", "", "", "60% 할인 배지 노출 확인", "", "", "P2", "녹색 배경의 '60%' 배지가 각 카드에 노출된다"),
    ("", "", "", "정가/할인가 표시 확인", "", "", "P2", "정가 50,000이 취소선으로, 할인가 33,000이 강조 표시된다"),
    ("", "", "", "카드 내 투자 가능 정보 노출 확인", "", "", "P2", "각 카드 하단에 '주식 계좌 N개 투자 가능' 텍스트가 노출된다"),
    ("", "", "", "최고 할인 배지 확인 (계좌 5개 카드)", "", "", "P2", "계좌 5개 카드에 그라디언트 배경과 '최고 할인' 텍스트 및 fire 아이콘이 노출된다"),
    ("", "", "업그레이드 투자권 카드", "카드 탭 → 구매 바텀시트 진입", "", "", "P1", "구매 바텀시트가 하단에서 슬라이드업 되어 노출된다"),

    # ── 사용권 업그레이드_기간권 이용_최대 케이스 ──
    ("", "기간권 이용_최대 케이스", "화면 진입", "1년권 최상위 이용 중 상태에서 업그레이드 화면 진입", "", "", "P1", "사용권 업그레이드 화면이 정상 노출된다"),
    ("", "", "내 투자권 카드", "1년권 카드 정보 노출 확인", "", "", "P2", "'내 주식 실전 투자권 (기간권)' 타이틀과 '주식 프리미엄 1년권 · 32일 후 만료' 텍스트가 노출된다"),
    ("", "", "", "사용권함 이동 버튼 탭", "", "", "P2", "사용권함 화면으로 이동한다"),
    ("", "", "서브탭", "1개월권 탭 기본 선택 상태 확인", "", "", "P2", "1개월권 탭이 기본 선택된 상태로 노출된다"),
    ("", "", "업그레이드 가능 투자권 없음", "1개월권 탭 빈 상태 확인", "", "", "P1", "빈 상태 아이콘 및 '업그레이드 가능한 사용권이 없어요' 안내 텍스트가 노출된다"),
    ("", "", "", "빈 상태 서브 안내 문구 확인", "", "", "P3", "'이용 중인 기간권보다 높은 금액의 사용권으로만 업그레이드 가능해요' 텍스트가 노출된다"),
    ("", "", "", "1년권 탭 탭 시 빈 상태 확인", "", "", "P1", "1년권 탭에서도 동일하게 빈 상태 아이콘 및 안내 텍스트가 노출된다"),

    # ── 사용권 업그레이드_구독권 이용 ──
    ("", "구독권 이용", "화면 진입", "구독권 이용 중 상태에서 업그레이드 화면 진입", "", "", "P1", "사용권 업그레이드 화면이 정상 노출된다"),
    ("", "", "구독권 사용 중 카드", "카드 노출 확인", "", "", "P1", "'구독권 사용 중' 타이틀과 '볼트 업그레이드는 기간권만 가능해요' 안내문이 노출된다"),
    ("", "", "", "구독권 업그레이드 이동 버튼 탭", "", "", "P1", "구독권 업그레이드 화면으로 이동한다"),
    ("", "", "업그레이드 가능 투자권 없음", "빈 상태 노출 확인", "", "", "P1", "빈 상태 아이콘 및 '업그레이드 가능한 사용권이 없어요' 텍스트가 노출된다"),
    ("", "", "", "빈 상태 서브 안내 문구 확인", "", "", "P3", "'이용 중인 기간권보다 높은 금액의 사용권으로만 업그레이드 가능해요' 텍스트가 노출된다"),

    # ── 사용권 업그레이드_사용 중인 투자권 없음 ──
    ("", "사용 중인 투자권 없음", "화면 진입", "기간권 미보유 상태에서 업그레이드 화면 진입", "", "", "P1", "사용권 업그레이드 화면이 정상 노출된다"),
    ("", "", "기간권 없음 카드", "카드 노출 확인", "", "", "P1", "'사용 중인 기간권 없음' 타이틀이 노출된다"),
    ("", "", "", "카드 상세 안내문 확인", "", "", "P3", "'사용 중인 기간권이 있어야 업그레이드할 수 있습니다. 기간권은 볼트 스토어에서 구매할 수 있습니다.' 문구가 노출된다"),
    ("", "", "", "볼트 스토어 이동 버튼 탭", "", "", "P1", "볼트 스토어 화면으로 이동한다"),
    ("", "", "", "사용권함 이동 버튼 탭", "", "", "P2", "사용권함 화면으로 이동한다"),
    ("", "", "업그레이드 가능 투자권 없음", "빈 상태 노출 확인", "", "", "P1", "빈 상태 아이콘 및 '업그레이드 가능한 사용권이 없어요' 텍스트가 노출된다"),

    # ── 내 정보_마일리지_스토어_구매 (바텀시트 정상) ──
    ("스토어_구매", "바텀시트_정상", "구매 바텀시트", "투자권 카드 탭 시 바텀시트 노출", "", "", "P1", "구매 바텀시트가 하단에서 슬라이드업 되어 노출된다"),
    ("", "", "", "선택 상품명 노출 확인", "", "", "P2", "'주식 실전 투자권 1계좌권 1개월권' 텍스트가 노출된다"),
    ("", "", "", "상품 가격(볼트) 표시 확인", "", "", "P2", "볼트 아이콘(⚡)과 함께 상품 가격 8,000이 노출된다"),
    ("", "", "", "가격 계산 상세 표시 확인", "", "", "P2", "선택한 사용권 정가 18,000 / 내 사용권 정가 차감 -9,000이 각각 노출된다"),
    ("", "", "", "내 마일리지 잔액 노출 확인", "", "", "P2", "볼트 아이콘과 함께 내 마일리지 10,000이 노출된다"),
    ("", "", "안내 문구", "환불 불가 안내 문구 노출 확인", "", "", "P3", "'*상품 구매 이후에는 이용과 상관없이 환불이 불가합니다.' 문구가 노출된다"),
    ("", "", "바로 업그레이드 버튼", "마일리지 충분 시 버튼 활성 확인", "", "", "P1", "검은색 배경의 '바로 업그레이드' 버튼이 활성화된 상태로 노출된다"),
    ("", "", "", "바로 업그레이드 버튼 탭", "", "", "P1", "구매가 진행되며 버튼 영역이 로딩 상태(점 3개)로 전환된다"),

    # ── 내 정보_마일리지_스토어_구매 (로딩) ──
    ("", "바텀시트_로딩", "구매 로딩", "버튼 탭 후 로딩 인디케이터 전환 확인", "", "", "P1", "버튼 영역에 로딩 인디케이터(점 3개)가 노출된다"),
    ("", "", "", "내 볼트 잔액 노출 확인", "", "", "P2", "볼트 아이콘과 함께 내 볼트 10,000이 노출된다"),
    ("", "", "", "로딩 중 버튼 중복 탭 방지 확인", "", "", "P1", "버튼이 비활성화되어 중복 탭이 되지 않는다"),

    # ── 내 정보_마일리지_스토어_구매_실패 ──
    ("", "바텀시트_구매 실패", "실패 토스트", "업그레이드 실패 시 토스트 메시지 노출", "", "", "P1", "'업그레이드에 실패했어요.' 토스트 메시지가 화면 상단에 노출된다"),
    ("", "", "", "실패 후 화면 상태 유지 확인", "", "", "P2", "스토어 목록 화면이 그대로 유지되고 다시 구매 시도 가능하다"),
]

# ──────────────────────────────────────────
# 스타일 헬퍼
# ──────────────────────────────────────────
def fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def border_thin():
    s = Side(border_style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

FILL_HEADER = fill("BFBFBF")
FILL_TITLE  = fill("FFFFFF")
FILL_NONE   = fill("FFFFFF")

FONT_TITLE  = Font(name="Arial", size=18, bold=True, color="002060")
FONT_HEADER = Font(name="Arial", size=9,  bold=True)
FONT_DATA   = Font(name="Arial", size=9)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ──────────────────────────────────────────
# 워크북 생성
# ──────────────────────────────────────────
output_dir = r"C:\Users\rkd98\TC_Automation\output"
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, "QA_사용권업그레이드_TC.xlsx")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "사용권 업그레이드 TC"

# ── 컬럼 너비 (원본 양식 기준) ──
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
c = ws.cell(row=2, column=1, value="사용권 업그레이드")
c.font = FONT_TITLE
c.alignment = ALIGN_LEFT

for col, val in [(18, "P1"), (19, "=COUNTIF(G:G,R2)")]:
    c = ws.cell(row=2, column=col, value=val)
    c.font = FONT_DATA
    c.alignment = ALIGN_CENTER
    c.border = border_thin()

# ──────────────────────────────────────────
# 행 3~4: 컬럼 헤더 (원본과 동일한 병합 구조)
# ──────────────────────────────────────────
# 단일 컬럼 헤더 (행3-4 병합)
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

# Test Result (I3:L3 병합)
ws.merge_cells("I3:L3")
c = ws.cell(row=3, column=9, value="Test Result")
c.fill = FILL_HEADER
c.font = FONT_HEADER
c.alignment = ALIGN_CENTER
c.border = border_thin()

# 행 4: Test Result 서브 헤더
sub_headers = [
    (9,  "Android"),
    (10, "iOS"),
    (11, "Date"),
    (12, "BTS ID"),
]
for col, val in sub_headers:
    c = ws.cell(row=4, column=col, value=val)
    c.fill = FILL_HEADER
    c.font = FONT_HEADER
    c.alignment = ALIGN_CENTER
    c.border = border_thin()

# 우상단 행3~4: P2/P3
for row, p_val, cnt_val in [(3, "P2", "=COUNTIF(G:G,R3)"), (4, "P3", "=COUNTIF(G:G,R4)")]:
    ws.cell(row=row, column=18, value=p_val).font = FONT_DATA
    ws.cell(row=row, column=19, value=cnt_val).font = FONT_DATA
    for col in (18, 19):
        ws.cell(row=row, column=col).alignment = ALIGN_CENTER
        ws.cell(row=row, column=col).border = border_thin()

# 합계 행3 우측
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

    # Test Result 빈 칸 (I~L)
    for col_idx in range(9, 13):
        c = ws.cell(row=row_idx, column=col_idx)
        c.font = FONT_DATA
        c.border = border_thin()

    ws.row_dimensions[row_idx].height = 30

# ──────────────────────────────────────────
# 자동 필터 (헤더 행4 기준)
# ──────────────────────────────────────────
ws.auto_filter.ref = f"A4:L4"

wb.save(file_path)
print(f"저장 완료: {file_path}")
print(f"총 TC: {len(tc_data)}개")
