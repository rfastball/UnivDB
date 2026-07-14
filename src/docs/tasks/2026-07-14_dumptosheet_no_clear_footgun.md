# 이슈: clsTable.DumpToSheet가 대상 영역을 클리어하지 않음 - 데이터 축소 시 유령행 잔존

작성일: 2026-07-14
발견 경로: 소비 프로젝트 `지출 리팩토링` 코드 리뷰. "시트 덤프 시 기존값을
깔끔히 날리는가 / 미처리건 갱신이 바르게 덤프되는가" 점검 중 확인.
상태: **해결 (라이브러리, 2026-07-14)** - 처방 B를 기본값 반전으로 채택(§6). 소비측
(`지출 리팩토링`)의 호출부 선클리어 우회는 새 UnivDB 벤더링 시 잉여가 됨(그쪽 몫).
심각도: 중 (silent data corruption 가능 - 오류 없이 옛 데이터가 되살아남)

## 1. 증상

`clsTable.DumpToSheet`로 같은 시작셀에 반복 덤프할 때, 이번 데이터가 지난번보다
**행 수가 적으면** 초과분(옛 데이터)이 시트에 그대로 남는다. 오류는 없다.

전형적 피해 경로 - 다운스트림이 그 시트를 `Range("A1").CurrentRegion`으로 읽는 경우:
유령행이 실데이터와 물리적으로 인접하므로 `CurrentRegion`이 이를 통째로 흡수하고,
이미 사라졌어야 할 레코드가 살아있는 데이터로 재유입된다.

소비측 실제 사례(`지출 리팩토링` Q_Logic_Sync):
- `송신목록_미처리`(raw 저장소)는 처리가 진행될수록 행이 단조 감소하는데
  `tUnproc.DumpToSheet` 로 덮어써 옛 미처리건이 잔존 → 하위 `Proc_Make_Skeleton_Report`가
  `Sheets("송신목록_미처리").Range("A1").CurrentRegion` 으로 읽어 처리완료건이 미처리건
  보고서에 되살아남.

## 2. 확정된 메커니즘

`clsTable.DumpToSheet` (clsTable.cls:530-559):

```vba
Set rngTarget = rngStart.Resize(m_lRowCount, m_lColCount)   ' 새 데이터 크기만큼만
' rngTarget.ClearContents ' 필요 시 주석 해제                ' ← 클리어가 주석 처리됨
...
rngTarget.Value2 = m_vData                                  ' 딱 그 사각형만 덮어씀
```

- 출력 범위를 **새 데이터 크기**(`m_lRowCount`×`m_lColCount`)로만 잡아 `Value2` 일괄 대입.
- 유일한 클리어 라인(541행)이 주석 처리되어 있어, 이전 덤프가 더 컸을 때 아래·오른쪽
  잔여 셀을 건드리지 않는다.
- 대비: `clsReportBuilder.BuildToSheet`(clsReportBuilder.cls:295-296)는 쓰기 전에
  `ws.Range(sAddr).CurrentRegion.ClearContents`로 옛 영역 전체를 지운다. 즉 라이브러리
  내부에서도 두 출력 경로의 클리어 계약이 **비대칭**이다.

## 3. 함정 포인트 (왜 눈에 안 띄나)

- 41만 건 Upsert처럼 행이 늘기만 하는(단조 증가) 사용에선 절대 안 드러난다.
  행이 줄어드는 순간에만 발현 → 리뷰·스모크에서 지나치기 쉽다.
- `주석 처리된 클리어를 살리는 것도 오답`이다. 541행을 살려도 `Resize(새크기)`
  사각형만 지우므로 **줄어든 경우의 잔여행은 여전히 안 지워진다.** 올바른 클리어는
  쓰기 전 옛 영역(`CurrentRegion` 또는 이전 크기) 기준이어야 한다.

## 4. 처방 후보

- **A. 계약 명시 (문서만, 최소 변경)**: DumpToSheet는 "새 크기 사각형만 덮어쓴다.
  영역 축소 시 선클리어는 호출자 책임"임을 Reference에 명기. BuildToSheet와의 비대칭을
  의도된 계약으로 고정. → 본 리포트와 함께 Reference §5 I/O에 경고 1줄 추가함.
- **B. 옵트인 클리어 인자**: `DumpToSheet(Range, [bAutoFit], [eTextGuard], [bClearRegion])`.
  `bClearRegion:=True`면 쓰기 전 `rngStart.CurrentRegion.ClearContents`. 기본 False로
  기존 호출 무영향. 소비측이 선클리어 2줄을 덜 쓰게 됨.
- **C. 기본 클리어로 전환 (파괴적)**: 비권장. `Resize`만 겨냥해 인접 보조열/수식열을
  안 건드리는 현 동작에 의존하는 소비처가 있을 수 있어 회귀 위험.

권장: **A(즉시) + B(다음 마이너)**. C는 지양.

## 5. 소비측 우회 (적용 완료, 2026-07-14)

`지출 리팩토링` Q_Logic_Sync.bas 호출부 2곳에 덤프 직전 선클리어 추가:

```vba
wsMaster.Range("A1").CurrentRegion.ClearContents   ' Proc_Sync_NewData
tMaster.DumpToSheet wsMaster.Range("A1"), True

wsUnproc.Range("A1").CurrentRegion.ClearContents   ' Proc_Extract_Unprocessed
tUnproc.DumpToSheet wsUnproc.Range("A1"), True
```

`ClearContents`(값만)로 충분: 빈 셀이 `CurrentRegion`을 끊어 유령행 재유입을 차단.
`BuildToSheet` 경유 출력(미처리건 뷰·처리결과·반납필요)은 이미 내부 클리어가 있어 무영향.

## 6. 라이브러리 해결 (2026-07-14)

**채택: 처방 B, 단 기본값을 True로 반전** (원안 B는 기본 False 옵트인이었음).

- `clsTable.DumpToSheet`에 `Optional bClearRegion As Boolean = True` 추가. 쓰기 전
  `rngStart.CurrentRegion.ClearContents`로 옛 extent 선클리어. 이 선클리어를
  `If m_lRowCount = 0 Then Exit Sub` **앞에** 배치해 빈 결과 덤프도 영역을 비운다.
- `clsReportBuilder.BuildToSheet`는 이미 자체 선클리어(옛 extent)·NumberFormat 리셋을
  하므로 유일 내부 호출을 `bClearRegion:=False`로 옵트아웃 → 기존 동작 바이트 동일 유지.
- 회귀 테스트 3종(`TestDumpLeadingZero.bas`): `Test_DumpToSheet_ShrinkClearsGhostRows`
  (유령행 제거+CurrentRegion 미흡수), `Test_DumpToSheet_ClearOptOut_PreservesGhost`
  (옵트아웃이 외과적 쓰기 보존), `Test_DumpToSheet_EmptyClearsRegion`(빈 덤프 클리어).
- Reference §5 계약 노트를 "클리어 안 함/호출자 책임" → "기본 클리어"로 재작성.

**근거 (왜 B-반전이고 A/C가 아닌가)**: 유령행 재유입은 오류 없이 옛 레코드를 살아있는
데이터로 되살리는 silent data corruption — 프로젝트가 스스로 정의한 최악의 실패 양식이다.
A(문서화)는 이 함정을 "정상 계약"으로 승격시켜 무장 해제된 채 방치한다. 원안 B(기본 False)도
문서 안 읽는 소비처엔 함정이 그대로 무장 — 이번 버그를 맞은 바로 그 집단. 반면 과다클리어
위험(인접 열)은 `BuildToSheet`가 이미 감수 중인 동일 위험이라 새 위험 클래스 도입이 아니다.
C(무인자 파괴적 전환)와 달리 기본 True + 옵트아웃은 외과적 쓰기 소비처에 명시적 탈출구를 남긴다.
