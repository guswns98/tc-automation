# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Quantus QA Automation Rules

## 역할
이 프로젝트에서 Claude는 Figma 디자인을 분석하여 QA 테스트 케이스를 자동 생성하는 역할을 한다.

---

## 스크립트 실행

```bash
python C:/Users/rkd98/TC_Automation/generate_tc.py
```

출력 파일은 `C:/Users/rkd98/TC_Automation/output/` 에 저장된다.

---

## 워크플로우

### 1. Figma 디자인 읽기

> **중요**: 전체 섹션/파일 단위 읽기는 데이터가 너무 많아 실제 사용 불가.
> 반드시 **개별 프레임의 node-id** 단위로 호출한다.

- Figma URL에서 `node-id` 파라미터를 추출하여 개별 프레임 단위로 읽는다
  - URL 예시: `figma.com/design/:fileKey/...?node-id=123-456` → nodeId = `123:456`
  - node-id가 없으면 사용자에게 특정 프레임의 node-id를 요청한다
- 프레임 하나씩 `get_design_context(fileKey, nodeId)` 호출
- 화면명, 컴포넌트 목록, 인터랙션 요소를 추출한다
- 버튼, 탭, 입력값, 상태(활성/비활성/에러) 등을 식별한다

#### 다중 URL 병렬 처리
- URL이 2개 이상이면 **모든 프레임을 동시에 병렬 호출**한다
- 각 프레임의 TC는 독립적으로 생성한 뒤 순서대로 병합한다
- xlsx는 프레임별 시트로 분리하거나, 단일 시트에 화면명으로 구분하여 저장한다
- 병렬 처리 완료 후 "총 N개 화면 처리 완료" 요약을 출력한다

### 2. TC 작성 기준
테스트 케이스는 아래 컬럼 구조를 반드시 따른다:

| 컬럼 | 설명 |
|------|------|
| 1~6 Depth | 계층형 테스트 시나리오 (행위 중심, 동사 + 목적어) |
| Priority | P1 / P2 / P3 |
| Expected Result | 사용자 관점의 예상 결과 ("~가 노출된다", "~으로 이동한다") |

### 3. 우선순위 기준
- **P1**: 화면 진입, 핵심 기능 동작, 구매 플로우, 에러 상태 처리
- **P2**: UI 노출 검증, 배지/텍스트 정확성, 스크롤
- **P3**: 안내 문구 세부 텍스트, 엣지 케이스

### 4. TC 커버리지 체크리스트
Figma 화면당 아래 항목을 반드시 포함한다:
- [ ] 화면 정상 진입
- [ ] 뒤로가기 / 네비게이션
- [ ] 탭/필터 전환
- [ ] 각 카드/컴포넌트 정보 노출
- [ ] CTA 버튼 탭 → 다음 플로우 진입
- [ ] 예외 상태 (부족, 품절, 중복 등)
- [ ] 안내 텍스트 영역

---

## 출력 형식

### TC 생성 후 자동 xlsx 저장 (필수)
TC 작성이 완료되면 **항상 자동으로 xlsx 파일을 생성하고 저장**한다. 별도 요청 불필요.

- 파일명: `QA_[화면명]_TC.xlsx`
- 저장 위치: `C:/Users/rkd98/TC_Automation/output/`
- 저장 완료 후 파일 경로를 사용자에게 알린다

### xlsx 컬럼 구조 (A~M)

| 열 | 헤더 |
|----|------|
| A | 1 Depth |
| B | 2 Depth |
| C | 3 Depth |
| D | 4 Depth |
| E | 5 Depth |
| F | 6 Depth |
| G | Priority |
| H | Expected Result |
| I | Android (전체) |
| J | iOS |
| K | Date |
| L | BTS ID |
| M | 합계 |

### 헤더 구조
- 행 1: 우상단 Priority/개수 요약 헤더 (col 18~19)
- 행 2: 제목 (A2:M2 병합, Arial 18pt Bold, 색상 #002060) + 우상단 P1 COUNTIF
- 행 3~4: 컬럼 헤더
  - A~H, M: 행3-4 병합 단일 헤더
  - I3:L3 병합 → "Test Result"
  - 행4 서브헤더: Android / iOS / Date / BTS ID
- 행 5부터: 데이터

### 스타일
- 폰트: **Arial** (제목 18pt Bold / 헤더·데이터 9pt)
- 헤더 셀 배경: `#BFBFBF`, Bold
- 행 1 요약 헤더 배경: `#D9D9D9`
- 데이터 셀: 흰색 배경
- 행 고정: freeze_panes = A5
- 자동 필터: A4:L4
- 데이터 행 높이: 30pt

### 컬럼 너비
A=13.75, B=20.25, C=22.13, D=22.75, E=24.0, F=22.13, G=8.25, H=32.38, I=12.63, J=10.0, K=10.0, L=13.63, M=43.5

### xlsx 생성 방법
Python `openpyxl` 라이브러리 + `C:/Users/rkd98/TC_Automation/generate_tc.py` 스크립트를 활용한다.
`tc_data` 리스트에 `(1D, 2D, 3D, 4D, 5D, 6D, Priority, Expected Result)` 튜플을 추가하면 된다.

---

## 명령어 트리거

| 트리거 | 액션 |
|--------|------|
| "TC 작성해줘" | Figma MCP로 현재 화면 읽고 TC 생성 |
| "TC xlsx로 뽑아줘" | xlsx 파일로 출력 |
| "커버리지 확인해줘" | 체크리스트 기준으로 누락 TC 보고 |
| "P1만 뽑아줘" | 우선순위 P1 TC만 필터링하여 출력 |

---

## 주의사항
- **Figma는 반드시 node-id 단위로 개별 프레임을 호출한다** (전체 섹션 호출 금지)
- node-id가 URL에 없으면 사용자에게 요청한다
- Figma 프레임명을 화면명으로 그대로 사용한다
- 기대결과는 반드시 사용자 관점으로 작성한다 ("~가 노출된다", "~으로 이동한다")
- 동일 컴포넌트라도 상태(정상/예외/경계값)별로 TC를 분리한다
