# 이슈: 메모리 해제(ClearDB/Terminate) 비용 - 키당 DbRowList 객체 폭발

작성일: 2026-07-13
발견 경로: 소비 프로젝트 `지출 리팩토링`의 로직 벤치(modBench)에서 버튼 Proc의
Finally 구간(`cleardb`/`teardown`)이 활성화 직후 2초대로 관측됨. 원인 추적 결과
비용원이 UnivDB 인덱스 페이로드 설계에 있음을 확인.
상태: **원인 확정 / 지배 성분 분리측정 미착수 / 수정 미착수** (코드 정독 판독)
검증 방법: 코드 정독 + 소비측 벤치 실측치. 소멸자별 분리 계측은 미실시("측정 미검증" 표기).

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

## 5. 미해결 - 지배 성분 분리측정이 선행

성분 A(인덱스 객체)와 B(문자열 배열) 중 어느 쪽이 지배적인지에 따라 방향이 갈린다:
- A 지배 -> **처방 A(CSR)가 결정타**
- B 지배 -> A는 헛수고, B/C로 선회

`TestBenchDb`에 teardown 계측이 없으므로, 대형 테이블을 등록/인덱싱한 뒤
(1) 인덱스만 해제 (2) 데이터 배열만 해제를 각각 시간계측하는 마이크로벤치를 심어
성분을 쪼개는 것이 다음 스텝이다. 이 리포트는 그 측정 전까지의 원인 판독 고정본.

## 6. 참조 위치

- UnivDB.bas:142-147 (ClearDB 본체)
- clsTable.cls:22, 24 (m_vData / m_dicHashIndices), :40-44 (Class_Terminate)
- DictHashIndex.cls:29-45 (AddRow, 키당 New DbRowList), :21 (mDict)
- CoreHashIndex.cls:35-54 (AddRow, 동일 객체 생성), :22 (mRows() As DbRowList)
- DbRowList.cls:12-13 (Long 배열 페이로드)
