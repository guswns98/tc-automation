# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Quantus QA Automation Rules

## 역할
Figma 디자인을 분석하여 QA 테스트 케이스를 자동 생성하고 xlsx 파일로 저장하는 자동화 도구.

---

## 아키텍처

```
TC_Automation/
├── generate_tc.py        # xlsx 생성 엔진 (openpyxl 기반)
├── tc_data/<화면명>.py   # 화면별 TC 데이터 (screen_name + tc_data 변수)
└── output/               # 생성된 QA_[화면명]_TC.xlsx 저장 위치
```

**엔진/데이터 분리 구조**: `generate_tc.py`는 스타일·레이아웃·유효성 검사 등 xlsx 포맷만 담당. TC 내용은 `tc_data/*.py`에만 작성. 포맷 변경 시에만 엔진을 수정한다.

### tc_data 파일 형식
```python
screen_name = "화면명"   # xlsx 제목 및 파일명에 사용

tc_data = [
    # (1Depth, 2Depth, 3Depth, 4Depth, 5Depth, 6Depth, Priority, Expected Result)
    ("화면명", "시나리오", "액션", "", "", "", "P1", "기대결과가 노출된다"),
    ("", "", "다음 액션", "", "", "", "P2", "다음 기대결과가 노출된다"),
]
```

---

## 스크립트 실행

```bash
python C:/Users/rkd98/TC_Automation/generate_tc.py tc_data/<화면명>.py
```

출력: `C:/Users/rkd98/TC_Automation/output/QA_[화면명]_TC.xlsx`

---

## 워크플로우

### 1. Figma 디자인 읽기

> **중요**: 전체 섹션/파일 단위 호출은 데이터 과다로 **실사용 불가**. 반드시 **개별 프레임의 node-id 단위**로 호출한다.

- URL에서 node-id 추출: `?node-id=123-456` → nodeId = `123:456` (하이픈을 콜론으로 변환)
- node-id가 없으면 사용자에게 요청한다
- `get_design_context(fileKey, nodeId)` 프레임 단위 호출
- URL이 2개 이상이면 **모든 프레임을 병렬 호출**하고 완료 후 병합

> **필수**: `get_design_context` 응답에 포함된 **스크린샷을 반드시 확인**한 뒤 TC를 작성한다.
> XML 구조만으로는 `<instance>` 컴포넌트의 레이블·위치를 정확히 파악할 수 없다.
> 버튼 텍스트, 위치(상단/하단), 노출 조건은 **스크린샷 기준으로 확정**한다.

### 2. TC 작성 기준

| 컬럼 | 설명 |
|------|------|
| 1~6 Depth | 계층형 시나리오 (행위 중심: 동사 + 목적어) |
| Priority | P1 / P2 / P3 |
| Expected Result | 사용자 관점 ("~가 노출된다", "~으로 이동한다") |

**우선순위**
- **P1**: 화면 진입, 핵심 기능 동작, 구매 플로우, 에러 상태 처리
- **P2**: UI 노출 검증, 배지/텍스트 정확성, 스크롤
- **P3**: 안내 문구 세부 텍스트, 엣지 케이스

**커버리지 체크리스트** (화면당 필수)
- [ ] 화면 정상 진입
- [ ] 뒤로가기 / 네비게이션
- [ ] 건너뛰기(Skip) 버튼 — 온보딩·가입 플로우에서 반드시 확인
- [ ] 탭/필터 전환
- [ ] 각 카드/컴포넌트 정보 노출
- [ ] CTA 버튼 탭 → 다음 플로우 진입
- [ ] 선택형 컴포넌트(카드/탭/라디오)는 **각 항목별 개별 선택 TC** 작성 + 선택값이 다음 화면까지 유지되는지 확인
- [ ] 예외 상태 (부족, 품절, 중복 등)
- [ ] 안내 텍스트 영역

**Figma `<instance>` 컴포넌트 처리 규칙**
- `<instance>` 노드는 텍스트 레이블이 XML에 직접 노출되지 않는다.
- 같은 프레임 내 `<instance>` 가 2개 이상이면 → 버튼 종류(주 CTA / 건너뛰기 등)를 **추정하지 말고** 사용자에게 레이블을 확인한다.
- 확인 전까지 TC 2Depth를 "주 CTA 버튼", "건너뛰기 버튼" 등 **구체적 명칭으로 분리**하여 작성하고, 불확실한 항목은 `[확인 필요]` 표시를 남긴다.

---

## xlsx 출력 스펙

### 컬럼 구조 (A~M)

| 열 | 헤더 | 비고 |
|----|------|------|
| A~F | 1~6 Depth | |
| G | Priority | |
| H | Expected Result | |
| I | Android | 드롭다운: P / F / NA / NT |
| J | iOS | 드롭다운: P / F / NA / NT |
| K | Date | |
| L | BTS ID | |
| M | 합계 | |

- **I, J 컬럼**: 데이터 유효성 검사 드롭다운 (P / F / NA / NT), `F` 입력 시 배경색 `#FFCCCC` (연한 빨간색) 자동 적용

### 헤더 구조
- 행 1: 우상단 Priority/개수 요약 (col 18~19, 배경 `#D9D9D9`)
- 행 2: 제목 (A2:M2 병합, Arial 18pt Bold, `#002060`) + 우상단 P1 COUNTIF
- 행 3: I3:L3 병합 → "Test Result"
- 행 4: 서브헤더 (Android / iOS / Date / BTS ID), 자동 필터 A4:L4
- 행 5~: 데이터 (높이 30pt), freeze_panes = A5

### 스타일
- 폰트: Arial (제목 18pt Bold, 헤더·데이터 9pt)
- 헤더 배경: `#BFBFBF`, Bold
- 컬럼 너비: A=13.75, B=20.25, C=22.13, D=22.75, E=24.0, F=22.13, G=8.25, H=32.38, I=12.63, J=10.0, K=10.0, L=13.63, M=43.5

---

## 명령어 트리거

| 트리거 | 액션 |
|--------|------|
| "TC 작성해줘" | Figma MCP로 화면 읽고 TC 생성 + xlsx 저장 |
| "TC xlsx로 뽑아줘" | xlsx 파일로 출력 |
| "커버리지 확인해줘" | 체크리스트 기준으로 누락 TC 보고 |
| "P1만 뽑아줘" | 우선순위 P1 TC만 필터링하여 출력 |

TC 작성 완료 시 **항상 자동으로 xlsx를 생성하고 저장**한다. 별도 요청 불필요.
