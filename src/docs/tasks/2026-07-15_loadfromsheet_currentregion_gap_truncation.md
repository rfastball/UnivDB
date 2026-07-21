# 이슈: clsTable.LoadFromSheet - 다중 셀 Range 입력 시 CurrentRegion 절단을 그대로 신뢰

작성일: 2026-07-15
발견 경로: 소비 프로젝트 `지출 리팩토링`에서 "특이건 시트 진행상황이 미처리건
대시보드에 반영 안 된다"는 사용자 제보 조사 중 확인. 실측 진단 매크로로 재현.
상태: 미해결 (리포트만 - 수정은 다음 UnivDB 세션에서 결정, 소비측은 우회 적용 예정)
심각도: 높음 (silent data loss - 오류 없이 실데이터 대부분이 조용히 누락됨)

## 1. 증상

`UnivDB.RegisterTable "X", ws.Range("A1").CurrentRegion` 처럼 호출부가 CurrentRegion을
미리 평가해 다중 셀 Range로 넘기면, 실데이터 중간에 완전히 빈 행이 하나라도 있을 때
그 행에서 로드가 끊긴다. `LoadFromSheet`는 이 경우 재탐색을 하지 않고 넘어온 Range를
그대로 신뢰한다.

소비측 실제 사례 (`지출 리팩토링` Q_Logic_Worklist.bas:25):
```vba
UnivDB.RegisterTable "SPEC_CASES", Sheets("특이건").Range("A1").CurrentRegion
```
`특이건` 시트는 실데이터가 179행까지 이어지지만 15행이 완전 공백행이라
`CurrentRegion`이 1~14행만 잡는다. 15행 이후 실데이터(예: 17·18행의 계약/납품 키)가
SPEC_CASES 테이블에 아예 실리지 않아, 이를 키로 삼는
`Q_Logic_Worklist.ApplyComplexBusinessLogic`의 특이건 -> 진행상황 매핑이 조용히
미스된다. 사용자에게는 "특이건 메모를 넣었는데 미처리건 대시보드에 반영이 안 된다"로
보인다.

진단(임시 매크로, 실워크북 실측):
```
EndUpLastRow=179; CurrentRegionLastRow=14; MissedByCR=165; FullBlankRowsInSpan=161
```

## 2. 확정된 메커니즘

`clsTable.LoadFromSheet` (clsTable.cls:63-97):

```vba
ElseIf TypeOf vSource Is Range Then
    If vSource.Cells.Count = 1 Then
        ' Case B-1: 단일 셀 -> GetRealUsedRange(Find 기반, 빈 행에 안전)로 재탐색
        Set rngTarget = GetRealUsedRange(vSource.Worksheet, vSource)
    Else
        ' Case B-2: 다중 셀 Range -> 그대로 신뢰, 재탐색 없음
        Set rngTarget = vSource
    End If
```

`GetRealUsedRange`(101-146행)는 `.Find(What:="*", ..., SearchDirection:=xlPrevious)`로
진짜 마지막 행/열을 찾아 중간 공백행에 흔들리지 않는다. 하지만 이 안전 경로는 "셀
1개짜리 Range 또는 Worksheet 객체"를 넘겼을 때만 탄다. 호출부가 편의상
`.Range("A1").CurrentRegion`을 직접 붙여 다중 셀 Range를 만들어 넘기면, 이미 잘린
결과를 그대로 신뢰하는 Case B-2로 빠진다 - 안전한 탐색 로직이 라이브러리 안에
존재하는데도 호출 방식 때문에 우회된다.

## 3. 함정 포인트 (왜 눈에 안 띄나)

- 데이터가 조밀한 시트(중간 공백행 없음)에서는 CurrentRegion과 GetRealUsedRange의
  결과가 동일해서 절대 드러나지 않는다. 소비 프로젝트의 RAW/매핑/처리결과/고지서
  시트는 현재 이 조건이라 무증상.
- "실데이터가 성기고, 사람이 수기로 관리하는 시트"(자유서식 메모 시트 등)에서만
  발현된다 - 바로 이런 시트가 중간에 빈 행이 생기기 쉽다. 함정과 발현 조건이 상관돼
  있어 리뷰에서 우연히 마주치기 전엔 잡히지 않는다.
- 호출부 코드(`Sheets(...).Range("A1").CurrentRegion`)는 관용적으로 "전체 데이터
  읽기"로 읽혀서 리뷰어가 위험 신호로 인지하기 어렵다. 오히려 `LoadFromSheet`에 안전
  경로가 이미 존재한다는 사실이 이 함정을 반직관적으로 만든다.

## 4. 처방 후보 (수정 보류 - 다음 세션 논의용)

- **A. 호출부 컨벤션 문서화**: `RegisterTable`/`LoadFromSheet`에 넘기는 Range는 항상
  단일 셀(`Range("A1")`)로 넘기라는 컨벤션을 Reference에 명기. 다중 셀 Range 직접
  전달은 "의도적으로 좁힌 서브셋을 확신할 때만" 쓰라고 경고. 최소 변경이지만 소비처가
  문서를 안 읽으면 여전히 무장.
- **B. Case B-2에서도 안전 재탐색**: 다중 셀 Range가 넘어와도 그 시작 셀에서
  `GetRealUsedRange`를 한 번 더 태워, 실제 End(xlUp)/Find 결과가 넘어온 Range보다 크면
  확장하거나 최소한 로그/경고. 다만 "의도적으로 좁힌 범위"(A1:E100 서브셋 로드)라는
  합법적 유스케이스와 구분이 어려워 오탐 위험이 있다.
- **C. 안전한 전체범위 탐색을 Public API로 노출**: `GetRealUsedRange`에 대응하는
  공개 함수(예: `UnivDB.GetSafeUsedRange(ws)`)를 제공해, 소비 프로젝트가
  `ws.Range("A1").CurrentRegion` 대신 명시적으로 안전 경로를 고르도록 유도.

우선순위 판단은 보류. [[2026-07-14_dumptosheet_no_clear_footgun]]과 같은 "CurrentRegion이
조용히 데이터를 자르거나 되살린다" 계열 반복 패턴 - 이번이 두 번째 사례. 세 번째가
나오면 A/B/C 중 하나로 수렴을 검토할 것.

## 5. 소비측 우회 (제안, 미적용)

`지출 리팩토링` Q_Logic_Worklist.bas:25:
```vba
UnivDB.RegisterTable "SPEC_CASES", Sheets("특이건").Range("A1")   ' .CurrentRegion 제거
```
단일 셀로 넘기면 Case B-1의 `GetRealUsedRange`(Find 기반)를 타서 빈 행에 안전해진다.

같은 함수 내 RAW/MAP_BRANCH/MAP_UNIT/PROC_RESULTS/BILL_INFO도 동일 패턴
(`.Range("A1").CurrentRegion`)이라 구조적으로 같은 함정을 안고 있으나, 현재는 데이터가
조밀해 무증상 - 예방적 동일 수정 여부는 소비측 판단.
