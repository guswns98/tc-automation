# Quantus QA Automation

Figma 디자인을 분석하여 QA 테스트 케이스(TC)를 자동 생성하고 xlsx 파일로 저장하는 자동화 도구입니다.

## 주요 기능

- **Figma 디자인 읽기**: Figma MCP를 통해 개별 프레임 단위로 디자인 데이터를 추출
- **TC 자동 생성**: 화면 구성요소(버튼, 탭, 입력값, 상태 등)를 분석하여 테스트 케이스 작성
- **xlsx 파일 출력**: `배포 전 체크리스트` 양식 기준의 xlsx 파일 자동 저장
- **다중 화면 병렬 처리**: 여러 Figma URL 입력 시 동시 처리 후 결과 병합

## 구조

```
TC_Automation/
├── generate_tc.py   # TC 데이터 정의 및 xlsx 생성 스크립트
├── output/          # 생성된 xlsx 파일 저장 위치
└── CLAUDE.md        # Claude 동작 규칙 및 TC 작성 기준
```

## 사용법

1. Figma 프레임 URL (node-id 포함)을 Claude에게 전달
2. `TC 작성해줘` 명령 입력
3. `output/` 폴더에 `QA_[화면명]_TC.xlsx` 파일 자동 생성

### 명령어

| 명령어 | 설명 |
|--------|------|
| `TC 작성해줘` | Figma 화면 읽고 TC 생성 + xlsx 저장 |
| `TC xlsx로 뽑아줘` | xlsx 파일로 출력 |
| `커버리지 확인해줘` | 누락 TC 체크리스트 보고 |
| `P1만 뽑아줘` | 우선순위 P1 TC만 필터링 |

## TC 컬럼 구조

| 열 | 항목 |
|----|------|
| A~F | 1~6 Depth (테스트 계층) |
| G | Priority (P1 / P2 / P3) |
| H | Expected Result (기대결과) |
| I~O | Test Result (이슈 트래킹, Android/iOS, Date, BTS ID) |

## 우선순위 기준

- **P1**: 화면 진입, 핵심 기능 동작, 구매 플로우, 에러 상태 처리
- **P2**: UI 노출 검증, 배지/텍스트 정확성, 스크롤
- **P3**: 안내 문구 세부 텍스트, 엣지 케이스

## 요구사항

- Python 3.x
- `openpyxl` 라이브러리
- Figma MCP 서버 연결
