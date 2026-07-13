# 이슈: 메모리 해제(ClearDB/Terminate) 비용 - 키당 DbRowList 객체 폭발

작성일: 2026-07-13
발견 경로: 소비 프로젝트 `지출 리팩토링`의 로직 벤치(modBench)에서 버튼 Proc의
Finally 구간(`cleardb`/`teardown`)이 활성화 직후 2초대로 관측됨. 원인 추적 결과
비용원이 UnivDB 인덱스 페이로드 설계에 있음을 확인.
상태: **원인 확정 / 지배 성분 분리측정 완료(성분 A 압도적 지배) / 수정 미착수**
검증 방법: 코드 정독 + 소비측 벤치 실측 + `TestBenchDb.Test_BenchDb_Teardown` 분리측정
(2026-07-13, §5에 수치). 성분 A(인덱스 객체)가 성분 B(데이터 배열)를 ~30배 압도.

## 1. 증상 (소비측 실측)

`지출 리팩토링` modBench 하네스, 기본 백엔드(dbBackendDict), warmup 제외 중앙값(ms):

| Proc | 주요 연산 구간 | activate | **cleardb** | **teardown** |
|---|---|---|---|---|
| Proc_Find_Return_Required | rowloop 1203 | 0 | **2090** | - |
| Proc_Find_Missing_Bills | join 1691 | 0 | **934** | **1094** |

- `cleardb` = `UnivDB.ClearDB` (등록 테이블 일괄 해제)
- `teardown` = 결과 clsTable들 `Set ... = Nothing` (Missing 한정 별도 계측)
- `activate`(`.Activate`+`.Select`)는 0ms - 시트 전환 자체는 비용원이 아님.
  체감 지연의 실체는 activate "이후" Finally에서 도는 이 해제 연쇄다.

## 2. 확정된 메커니즘

`ClearDB` 본체는 두 줄뿐이다 (UnivDB.bas:142-147):

```vba
Public Sub ClearDB()
    If Not m_dicTables Is Nothing Then
        m_dicTables.RemoveAll
        Set m_dicTables = Nothing
    End If
End Sub
```

즉 비용은 이 코드가 아니라, 마지막 참조가 끊길 때 연쇄되는 **소멸자 캐스케이드**다.
소멸 대상은 두 성분:

### 성분 A - 키당 DbRowList 객체 (인덱스 페이로드)

`DictHashIndex.AddRow`는 새 키마다 객체를 하나 생성한다 (DictHashIndex.cls:36-45):

```vba
If mDict.exists(sKey) Then
    Set rows = mDict(sKey)
Else
    Set rows = New DbRowList      ' <- 키마다 COM 객체 1개
    rows.BindRaw mOwnerPtr, mGeneration
    mDict.Add sKey, rows
End If
rows.Append rowId
```

고유키성 컬럼(고지번호, DW 계약납품통합번호 등 행마다 거의 유일)에 인덱스가 서면
**고유키 수 ~= 행 수만큼 DbRowList 인스턴스**가 생긴다. VBA 클래스 인스턴스 소멸은
개당 오버헤드(참조계수 해제 + vtable 정리)가 있어, 수만 개를 하나씩 파괴하는 것이
해제 시간의 유력 지배원이다. DbRowList 자체 페이로드는 `Long()` 배열이라 가볍다
(DbRowList.cls:12) - 비싼 건 배열 내용이 아니라 "객체 개수"다.

### 성분 B - m_vData 대형 문자열 배열

등록된 각 테이블은 시트 `Value2`를 2D Variant로 통째 보유한다 (clsTable.cls:22, 86).
코드성 컬럼이 많아 셀 대부분이 BSTR이고, 배열 해제 = 셀마다 BSTR 해제(O(행x열)).
`Class_Terminate`에서 `m_vData = Empty`로 방출 (clsTable.cls:40-44). 로드한 만큼
되돌리는 비용이라 상당 부분 불가피.

## 3. 핵심 정정 - 인덱스 백엔드 전환은 해법이 아니다

"느리니 dbBackendCore로 바꾸자"는 이 문제에 무효다. `CoreHashIndex`도 키마다
동일하게 객체를 만든다 (CoreHashIndex.cls:48-53):

```vba
If isNew Then
    Set rows = New DbRowList      ' <- Dict 백엔드와 동일
    rows.BindRaw mOwnerPtr, mGeneration
    Set mRows(slot) = rows
End If
mRows(slot).Append rowId
```

두 백엔드는 키->슬롯 **조회 엔진만** 다르다(Scripting.Dictionary vs Core_clsFastHash).
값 페이로드(키당 DbRowList 객체)는 동일하므로 **소멸시킬 객체 수가 같고 해제 비용도
같다.** 성분 A는 백엔드 선택과 무관한 공통 설계다.

## 4. 처방 후보 (효과/리스크순)

| 방향 | 내용 | 효과 | 리스크 |
|---|---|---|---|
| **A. 인덱스 페이로드 평탄화 (CSR)** | 키당 DbRowList 객체 대신 공유 Long 배열 1개 + 키별 (offset, count). 소멸이 O(객체수) -> O(1) 배열 free | 성분 A 사실상 제거 | 인덱스 백엔드 내부 재작성. 단 DbRowList는 인덱스 단독 소유(SelectRows/GetUniqueKeys는 복사본 반환 - clsTable.cls:363-437)라 외부 계약 영향은 국소적 |
| **B. 조기 컬럼 투영** | 큰 테이블을 조회만 할 때 필요한 컬럼만 물질화 | 성분 B 축소 | 소비측 Proc별 SelectColumns 배치 |
| **C. 동기 해제 회피** | 클릭 경로 Finally의 ClearDB 생략, 다음 InitDB 재등록 시 지연 해제 | 체감만 이동(근본 아님) | 메모리 증가, 비용 재등장 |

## 5. 분리측정 결과 - 성분 A 압도적 지배 (2026-07-13 실측)

`TestBenchDb.Test_BenchDb_Teardown` 추가. 방법: 신선 테이블(20k행)의 마지막 참조를 끊는
`Set T = Nothing` 동기 소멸 구간을 QPC로 감싸되, **인덱스 유무만 바꿔 차분**으로 성분 A를
분리(warmup 제외 3회 median, ms):

| backend | data_only (B) | with_uniq_idx (A+B) | with_dup_idx (B+미미A) | **idxA_uniq (A 분리치)** |
|---|---|---|---|---|
| dict | 2 | 62 | 3 | **60** |
| core | 2 | 59 | 2 | **57** |

- `data_only`=인덱스 없음(성분 B), `with_uniq_idx`=고유키 20k 인덱스(DbRowList 20k), `with_dup_idx`=고중복 5키.

판정:
- **성분 A가 지배 (~30배)**: 인덱스 소멸 순비용 ≈ 60ms vs 데이터 배열 해제 ≈ 2ms.
  성분 B(m_vData BSTR 배열)는 사실상 무시 가능.
- **A는 행 수가 아니라 고유키 수에 스케일**: 같은 20k행이라도 고중복(5키)이면 `with_dup`≈`data_only`
  (3 vs 2). 고유키 20k에서만 60ms 폭발 → 성분 A = "키당 DbRowList 객체 수" 확정.
- **백엔드 무관(§3 실증)**: dict idxA=60 ≈ core idxA=57. CoreHashIndex도 동일 페이로드라
  백엔드 전환은 해법이 아님이 실측으로 재확인.

결론: **처방 A(CSR 인덱스 페이로드 평탄화)가 결정타.** 성분 B(처방 B/C 대상)는 지렛대가 없다.
(소비측 cleardb 2090ms는 이 20k 케이스의 60ms를 선형 외삽한 규모 - 다수 고유키성 컬럼×테이블에
인덱스가 서면서 DbRowList 총량이 수십만 개로 불어난 결과. 스케일 법칙은 위에서 실증됨.)

다음 스텝(이 리포트 범위 밖): §4 처방 A(CSR) 착수 여부 결정.

## 6. 참조 위치

- UnivDB.bas:142-147 (ClearDB 본체)
- clsTable.cls:22, 24 (m_vData / m_dicHashIndices), :40-44 (Class_Terminate)
- DictHashIndex.cls:29-45 (AddRow, 키당 New DbRowList), :21 (mDict)
- CoreHashIndex.cls:35-54 (AddRow, 동일 객체 생성), :22 (mRows() As DbRowList)
- DbRowList.cls:12-13 (Long 배열 페이로드)
