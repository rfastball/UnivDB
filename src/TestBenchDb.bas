Attribute VB_Name = "TestBenchDb"
' @desc DB 레이어 성능 baseline 측정 벤치마크 담당
' @depends UnivDB, modEngine, modIndexKey, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' Phase 7 준비 벤치마크 스위트 (2026-07-05_next_direction.md 7.4 목록).
' 목적: 현재 백엔드(DictHashIndex + 밀봉된 컬럼 reader 경로)의 baseline을 jsonl에 보존.
' Phase 7에서 CoreHashIndex / 핫 컬럼 캐시 도입 시 같은 케이스를 재실행해 비교한다.
' 규약: warm-up 1회 + 3회 측정 median, LogShim.LogEvent "bench_db_*"로 기록.

' QPC는 시간 측정 전용이며 배포 제외 모듈(Test_*)에만 선언 (Test_CoreBench 패턴)
#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
#End If

Private Const BENCH_ROWS As Long = 20000
Private Const PLAN_MASTER_ROWS As Long = 27401
Private Const PLAN_MASTER_COLS As Long = 34
Private Const PLAN_DW_ROWS As Long = 91000
Private Const PLAN_DW_COLS As Long = 27

Private mFreq As Currency
Private mT0 As Currency
Private mT20k As clsTable ' 공유 벤치 테이블 (지연 생성, 테스트 간 재사용)
Private mPlanMaster As Variant ' 마스터형 원본 배열 (지연 생성, 테스트 간 재사용)

Private Sub TicStart()
    If mFreq = 0 Then QueryPerformanceFrequency mFreq
    QueryPerformanceCounter mT0
End Sub

Private Function TicMs() As Double
    Dim t1 As Currency
    QueryPerformanceCounter t1
    TicMs = (t1 - mT0) / mFreq * 1000#
End Function

Private Function BackendName(ByVal b As Long) As String
    If b = dbBackendCore Then BackendName = "core" Else BackendName = "dict"
End Function

Private Function Median3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    Dim mx As Double, mn As Double
    mx = a: If b > mx Then mx = b
    If c > mx Then mx = c
    mn = a: If b < mn Then mn = b
    If c < mn Then mn = c
    Median3 = a + b + c - mx - mn
End Function

' ==============================================================================
' 1. 인덱스 빌드 (저중복 = 고유 키 20k / 고중복 = 키 5종)
' ==============================================================================

Public Sub Test_BenchDb_IndexBuild()
    Dim T As clsTable
    Dim idx As DbHashIndex
    Dim msLow(0 To 2) As Double, msHigh(0 To 2) As Double
    Dim R As Long
    Dim b As Long

    For b = dbBackendDict To dbBackendCore
        ' warm-up (지연 초기화 배출)
        Set T = BuildBenchTable(BENCH_ROWS)
        T.IndexBackend = b
        Set idx = T.GetHashIndex("ID")

        For R = 0 To 2
            Set T = BuildBenchTable(BENCH_ROWS) ' 캐시 없는 신선한 테이블에서 빌드 측정
            T.IndexBackend = b
            TicStart
            Set idx = T.GetHashIndex("ID")
            msLow(R) = TicMs
            LogShim.AssertEq "bench_idx_low_keys_" & BackendName(b) & "_" & R, CLng(BENCH_ROWS), idx.CountKeys

            Set T = BuildBenchTable(BENCH_ROWS)
            T.IndexBackend = b
            TicStart
            Set idx = T.GetHashIndex("Dept")
            msHigh(R) = TicMs
            LogShim.AssertEq "bench_idx_high_keys_" & BackendName(b) & "_" & R, 5, idx.CountKeys
        Next R

        LogShim.LogEvent "bench_db_index_build", _
            Array("backend", "rows", "lowdup_ms", "highdup_ms"), _
            Array(BackendName(b), BENCH_ROWS, CLng(Median3(msLow(0), msLow(1), msLow(2))), _
                  CLng(Median3(msHigh(0), msHigh(1), msHigh(2))))
    Next b
End Sub

' ==============================================================================
' 2. Where 동등 (hit / miss / 고중복 / 저중복) + 다중 Where 교집합
' ==============================================================================

Public Sub Test_BenchDb_WhereEquality()
    Const ITER_LIGHT As Long = 200
    Const ITER_HEAVY As Long = 20
    Const ITER_MULTI As Long = 50
    Dim T As clsTable
    Dim q As clsQuery
    Dim i As Long, R As Long
    Dim msHit(0 To 2) As Double, msMiss(0 To 2) As Double
    Dim msDup(0 To 2) As Double, msMulti(0 To 2) As Double

    Dim b As Long

    Set T = GetBench20k()

    For b = dbBackendDict To dbBackendCore
        T.IndexBackend = b ' 백엔드 전환 (캐시 실효 후 재빌드)
        ' 인덱스 프리워밍 (빌드 비용은 1번 벤치가 담당)
        T.GetHashIndex "ID"
        T.GetHashIndex "Dept"
        T.GetHashIndex "Score"

        ' warm-up
        Set q = UnivDB.Query(T).Where("ID", BENCH_ROWS \ 2)

        For R = 0 To 2
            TicStart
            For i = 1 To ITER_LIGHT
                Set q = UnivDB.Query(T).Where("ID", BENCH_ROWS \ 2) ' hit (저중복 = 고유 키)
            Next i
            msHit(R) = TicMs

            TicStart
            For i = 1 To ITER_LIGHT
                Set q = UnivDB.Query(T).Where("ID", -1) ' miss
            Next i
            msMiss(R) = TicMs

            TicStart
            For i = 1 To ITER_HEAVY
                Set q = UnivDB.Query(T).Where("Dept", "D3") ' 고중복 (~4000행 적재)
            Next i
            msDup(R) = TicMs

            TicStart
            For i = 1 To ITER_MULTI
                Set q = UnivDB.Query(T).Where("Dept", "D3").Where("Score", 503) ' 교집합
            Next i
            msMulti(R) = TicMs
        Next R

        LogShim.AssertEq "bench_where_hit_count_" & BackendName(b), 1, T.SelectRows("ID", BENCH_ROWS \ 2).Count
        LogShim.AssertEq "bench_where_dup_count_" & BackendName(b), 4000, T.SelectRows("Dept", "D3").Count

        LogShim.LogEvent "bench_db_where", _
            Array("backend", "rows", "iters_light", "iters_heavy", "iters_multi", _
                  "hit_ms", "miss_ms", "highdup_ms", "multi_ms"), _
            Array(BackendName(b), BENCH_ROWS, ITER_LIGHT, ITER_HEAVY, ITER_MULTI, _
                  CLng(Median3(msHit(0), msHit(1), msHit(2))), _
                  CLng(Median3(msMiss(0), msMiss(1), msMiss(2))), _
                  CLng(Median3(msDup(0), msDup(1), msDup(2))), _
                  CLng(Median3(msMulti(0), msMulti(1), msMulti(2))))
    Next b

    T.IndexBackend = dbBackendDict ' 공유 테이블 기본 백엔드 복원
End Sub

' ==============================================================================
' 3. Filter 순회 (수치 > / 문자 LIKE) - 밀봉된 컬럼 reader 경로
' ==============================================================================

Public Sub Test_BenchDb_FilterScan()
    Dim T As clsTable
    Dim tRes As clsTable
    Dim R As Long
    Dim msNum(0 To 2) As Double, msLike(0 To 2) As Double

    Set T = GetBench20k()

    ' warm-up
    Set tRes = UnivDB.Query(T).Filter("Score", ">", 500).SelectTable()

    For R = 0 To 2
        TicStart
        Set tRes = UnivDB.Query(T).Filter("Score", ">", 500).SelectTable()
        msNum(R) = TicMs

        TicStart
        Set tRes = UnivDB.Query(T).Filter("Tag", "LIKE", "*7*").SelectTable()
        msLike(R) = TicMs
    Next R

    ' Score = i Mod 1000, 501..999가 각 20회 = 9980행 (+헤더)
    Set tRes = UnivDB.Query(T).Filter("Score", ">", 500).SelectTable()
    LogShim.AssertEq "bench_filter_num_rows", CLng(9981), tRes.RowCount
    Set tRes = UnivDB.Query(T).Filter("Tag", "LIKE", "*7*").SelectTable()
    LogShim.AssertEq "bench_filter_like_nonempty", True, tRes.RowCount > 1

    LogShim.LogEvent "bench_db_filter", _
        Array("rows", "numeric_ms", "like_ms"), _
        Array(BENCH_ROWS, CLng(Median3(msNum(0), msNum(1), msNum(2))), _
              CLng(Median3(msLike(0), msLike(1), msLike(2))))
End Sub

' ==============================================================================
' 4. Join (1:1 / 1:N / N:N / no-match-heavy left)
' ==============================================================================

Public Sub Test_BenchDb_Join()
    Dim tL11 As clsTable, tR11 As clsTable
    Dim tL1N As clsTable, tR1N As clsTable
    Dim tLNN As clsTable, tRNN As clsTable
    Dim tLNo As clsTable, tRNo As clsTable
    Dim tRes As clsTable
    Dim R As Long
    Dim b As Long
    Dim ms11(0 To 2) As Double, ms1N(0 To 2) As Double
    Dim msNN(0 To 2) As Double, msNoMatch(0 To 2) As Double

    Set tL11 = BuildJoinTable(5000, 5000, 0): Set tR11 = BuildJoinTable(5000, 5000, 0) ' 1:1 고유 키
    Set tL1N = BuildJoinTable(1000, 1000, 0): Set tR1N = BuildJoinTable(5000, 1000, 0) ' 1:N
    Set tLNN = BuildJoinTable(5000, 500, 0): Set tRNN = BuildJoinTable(5000, 500, 0)   ' N:N 키당 100쌍
    Set tLNo = BuildJoinTable(5000, 5000, 0): Set tRNo = BuildJoinTable(5000, 5000, 100000) ' 도메인 분리

    For b = dbBackendDict To dbBackendCore
        tL11.IndexBackend = b: tR11.IndexBackend = b
        tL1N.IndexBackend = b: tR1N.IndexBackend = b
        tLNN.IndexBackend = b: tRNN.IndexBackend = b
        tLNo.IndexBackend = b: tRNo.IndexBackend = b

        ' 1:1 - warm-up이 인덱스 빌드 흡수, 측정은 조인 본체
        Set tRes = JoinTable(tL11, tR11, "K", "K", jtInner, "R_")
        For R = 0 To 2
            TicStart
            Set tRes = JoinTable(tL11, tR11, "K", "K", jtInner, "R_")
            ms11(R) = TicMs
        Next R
        LogShim.AssertEq "bench_join_11_rows_" & BackendName(b), CLng(5001), tRes.RowCount

        ' 1:N
        Set tRes = JoinTable(tL1N, tR1N, "K", "K", jtInner, "R_")
        For R = 0 To 2
            TicStart
            Set tRes = JoinTable(tL1N, tR1N, "K", "K", jtInner, "R_")
            ms1N(R) = TicMs
        Next R
        LogShim.AssertEq "bench_join_1n_rows_" & BackendName(b), CLng(5001), tRes.RowCount

        ' N:N (결과 50k)
        Set tRes = JoinTable(tLNN, tRNN, "K", "K", jtInner, "R_")
        For R = 0 To 2
            TicStart
            Set tRes = JoinTable(tLNN, tRNN, "K", "K", jtInner, "R_")
            msNN(R) = TicMs
        Next R
        LogShim.AssertEq "bench_join_nn_rows_" & BackendName(b), CLng(50001), tRes.RowCount

        ' no-match-heavy left join
        Set tRes = JoinTable(tLNo, tRNo, "K", "K", jtLeft, "R_")
        For R = 0 To 2
            TicStart
            Set tRes = JoinTable(tLNo, tRNo, "K", "K", jtLeft, "R_")
            msNoMatch(R) = TicMs
        Next R
        LogShim.AssertEq "bench_join_nomatch_rows_" & BackendName(b), CLng(5001), tRes.RowCount

        LogShim.LogEvent "bench_db_join", _
            Array("backend", "left_rows", "right_rows", "j11_ms", "j1n_ms", "jnn_ms", "nomatch_left_ms"), _
            Array(BackendName(b), 5000, 5000, CLng(Median3(ms11(0), ms11(1), ms11(2))), _
                  CLng(Median3(ms1N(0), ms1N(1), ms1N(2))), _
                  CLng(Median3(msNN(0), msNN(1), msNN(2))), _
                  CLng(Median3(msNoMatch(0), msNoMatch(1), msNoMatch(2))))
    Next b
End Sub

' ==============================================================================
' 4b. Join wide-table (20k행 x 30컬럼 1:1) - 헤더 구간 RawArray 반복 복사 비용 관측
' ==============================================================================

Public Sub Test_BenchDb_JoinWide()
    Const WIDE_ROWS As Long = 20000
    Const WIDE_COLS As Long = 30
    Dim tL As clsTable, tR As clsTable
    Dim tRes As clsTable
    Dim R As Long
    Dim msInner(0 To 2) As Double

    Set tL = BuildWideJoinTable(WIDE_ROWS, WIDE_COLS)
    Set tR = BuildWideJoinTable(WIDE_ROWS, WIDE_COLS)

    ' warm-up (인덱스 빌드 흡수, 측정은 조인 본체 + 물질화)
    Set tRes = JoinTable(tL, tR, "K", "K", jtInner, "R_")

    For R = 0 To 2
        TicStart
        Set tRes = JoinTable(tL, tR, "K", "K", jtInner, "R_")
        msInner(R) = TicMs
    Next R

    LogShim.AssertEq "bench_join_wide_rows", CLng(WIDE_ROWS + 1), tRes.RowCount
    LogShim.AssertEq "bench_join_wide_cols", CLng(WIDE_COLS * 2), tRes.ColCount

    LogShim.LogEvent "bench_db_join_wide", _
        Array("rows", "cols", "inner_ms"), _
        Array(WIDE_ROWS, WIDE_COLS, CLng(Median3(msInner(0), msInner(1), msInner(2))))
End Sub

' ==============================================================================
' 5. GroupBy (짧은 키 5종 / 긴 키 180자 pad 50종)
' ==============================================================================

Public Sub Test_BenchDb_GroupBy()
    Dim T As clsTable
    Dim tRes As clsTable
    Dim R As Long
    Dim b As Long
    Dim msShort(0 To 2) As Double, msLong(0 To 2) As Double

    Set T = GetBench20k()

    ' R2 이후 단일 컬럼 GroupBy는 해시 인덱스를 소비하므로 백엔드별 비교 (인덱스 프리워밍,
    ' 빌드 비용은 1번 벤치가 담당 - 여기 수치는 커서 순회 + 집계만)
    For b = dbBackendDict To dbBackendCore
        T.IndexBackend = b
        T.GetHashIndex "Dept"
        T.GetHashIndex "LongKey"

        ' warm-up
        Set tRes = T.GroupBy("Dept", "Score:SUM")

        For R = 0 To 2
            TicStart
            Set tRes = T.GroupBy("Dept", "Score:SUM")
            msShort(R) = TicMs

            TicStart
            Set tRes = T.GroupBy("LongKey", "Score:SUM")
            msLong(R) = TicMs
        Next R

        Set tRes = T.GroupBy("Dept", "Score:SUM")
        LogShim.AssertEq "bench_groupby_short_rows_" & BackendName(b), CLng(6), tRes.RowCount
        Set tRes = T.GroupBy("LongKey", "Score:SUM")
        LogShim.AssertEq "bench_groupby_long_rows_" & BackendName(b), CLng(51), tRes.RowCount

        LogShim.LogEvent "bench_db_groupby", _
            Array("backend", "rows", "shortkey_ms", "longkey_ms"), _
            Array(BackendName(b), BENCH_ROWS, CLng(Median3(msShort(0), msShort(1), msShort(2))), _
                  CLng(Median3(msLong(0), msLong(1), msLong(2))))
    Next b

    T.IndexBackend = dbBackendDict ' 공유 벤치 테이블 기본 백엔드 복원
End Sub

' ==============================================================================
' 6. 컬럼 접근 (R4: hot column cache 지렛대 측정)
' copy_ms = 호출당 CopyColumn1D 순수 비용 (캐시가 절약 가능한 상한)
' filter_ms / groupby_warm_ms = 그 복사를 포함한 소비자 전체 비용
' 캐시 지렛대 = copy 비중. CoreHashIndex의 join ~19% 개선과 비교해 도입 여부 결정
' ==============================================================================

Public Sub Test_BenchDb_ColumnAccess()
    Const ITER As Long = 50
    Dim T As clsTable
    Dim rdr As DbColumnReader
    Dim q As clsQuery
    Dim tRes As clsTable
    Dim i As Long, R As Long
    Dim msCopy(0 To 2) As Double, msFilter(0 To 2) As Double, msGroup(0 To 2) As Double

    Set T = GetBench20k()
    T.IndexBackend = dbBackendDict
    T.GetHashIndex "Dept" ' GroupBy warm 경로용 인덱스 프리워밍

    ' warm-up
    Set rdr = T.GetColumnReader("Score")
    Set q = UnivDB.Query(T).Filter("Score", ">", 500)
    Set tRes = T.GroupBy("Dept", "Score:SUM")

    For R = 0 To 2
        TicStart
        For i = 1 To ITER
            Set rdr = T.GetColumnReader("Score") ' 호출당 CopyColumn1D 1회
        Next i
        msCopy(R) = TicMs

        TicStart
        For i = 1 To ITER
            Set q = UnivDB.Query(T).Filter("Score", ">", 500) ' 복사 + 전 행 스캔
        Next i
        msFilter(R) = TicMs

        TicStart
        For i = 1 To ITER
            Set tRes = T.GroupBy("Dept", "Score:SUM") ' warm 인덱스: agg 컬럼 복사 + 커서 집계
        Next i
        msGroup(R) = TicMs
    Next R

    LogShim.AssertEq "bench_colaccess_reader_alive", True, Not (rdr Is Nothing)
    LogShim.AssertEq "bench_colaccess_group_rows", CLng(6), tRes.RowCount

    LogShim.LogEvent "bench_db_column_access", _
        Array("rows", "iter", "copy_ms", "filter_ms", "groupby_warm_ms"), _
        Array(BENCH_ROWS, ITER, CLng(Median3(msCopy(0), msCopy(1), msCopy(2))), _
              CLng(Median3(msFilter(0), msFilter(1), msFilter(2))), _
              CLng(Median3(msGroup(0), msGroup(1), msGroup(2))))
End Sub

' ==============================================================================
' 7. Teardown 분리측정 (성분 A: 키당 DbRowList 객체 / 성분 B: m_vData BSTR 배열)
' 마지막 참조를 끊는 Set T = Nothing 구간(동기 소멸 캐스케이드)을 QPC로 감싸되,
' 인덱스 유무만 바꿔 차분(with_uniq - data_only)으로 성분 A를 분리한다.
' 참조: docs/tasks/2026-07-13_teardown_cost_index_payload.md §5
' 주의: 인덱스 반환값을 잡으면 사슬이 안 끊겨 성분 A가 측정에서 누락된다 -> 반환 폐기.
'       공유 캐시 GetBench20k 금지 -> 매번 신선 테이블이 유일 참조가 되게 한다.
' ==============================================================================

Public Sub Test_BenchDb_Teardown()
    Dim T As clsTable
    Dim idx As DbHashIndex
    Dim R As Long
    Dim b As Long
    Dim msData(0 To 2) As Double, msUniq(0 To 2) As Double, msDup(0 To 2) As Double
    Dim medData As Double, medUniq As Double, medDup As Double

    For b = dbBackendDict To dbBackendCore
        ' warm-up: 지연 초기화 배출 + 인덱스 페이로드 전제 단언 (고유키 20k / 고중복 5종)
        Set T = BuildBenchTable(BENCH_ROWS): Set T = Nothing
        Set T = BuildBenchTable(BENCH_ROWS): T.IndexBackend = b
        Set idx = T.GetHashIndex("ID")
        LogShim.AssertEq "bench_teardown_uniq_keys_" & BackendName(b), CLng(BENCH_ROWS), idx.CountKeys
        Set idx = Nothing: Set T = Nothing
        Set T = BuildBenchTable(BENCH_ROWS): T.IndexBackend = b
        Set idx = T.GetHashIndex("Dept")
        LogShim.AssertEq "bench_teardown_dup_keys_" & BackendName(b), 5, idx.CountKeys
        Set idx = Nothing: Set T = Nothing

        For R = 0 To 2
            ' 성분 B: 인덱스 없는 테이블 소멸 (m_vData BSTR 배열 해제)
            Set T = BuildBenchTable(BENCH_ROWS)
            TicStart
            Set T = Nothing
            msData(R) = TicMs

            ' 성분 A + B: 고유키 인덱스(DbRowList 20k) + 데이터 소멸. 반환 폐기 = 사슬 단독 소유
            Set T = BuildBenchTable(BENCH_ROWS)
            T.IndexBackend = b
            T.GetHashIndex "ID"
            TicStart
            Set T = Nothing
            msUniq(R) = TicMs

            ' 성분 B + 미미한 A: 고중복 인덱스(DbRowList 5) + 데이터 소멸
            Set T = BuildBenchTable(BENCH_ROWS)
            T.IndexBackend = b
            T.GetHashIndex "Dept"
            TicStart
            Set T = Nothing
            msDup(R) = TicMs
        Next R

        medData = Median3(msData(0), msData(1), msData(2))
        medUniq = Median3(msUniq(0), msUniq(1), msUniq(2))
        medDup = Median3(msDup(0), msDup(1), msDup(2))

        ' idxA_uniq_ms = 고유키 인덱스 소멸 순비용 (성분 A 분리치)
        LogShim.LogEvent "bench_db_teardown", _
            Array("backend", "rows", "data_only_ms", "with_uniq_idx_ms", "with_dup_idx_ms", "idxA_uniq_ms"), _
            Array(BackendName(b), BENCH_ROWS, CLng(medData), CLng(medUniq), CLng(medDup), _
                  CLng(medUniq - medData))
    Next b
End Sub

' ==============================================================================
' 데이터 생성 헬퍼
' ==============================================================================

' @desc 10만 행 Update와 동일 열 숫자 OR의 실행 비용을 기록함
Public Sub Test_BenchDb_QueryWritesAndNumericOr()
    Const NROWS As Long = 100000
    Dim T As clsTable, source As clsTable, result As clsTable, q As clsQuery
    Dim arms(0 To 99) As Variant, r As Long, i As Long, changed As Long
    Dim msUpdate(0 To 2) As Double, msOr(0 To 2) As Double, elapsed As Double

    Set source = BuildJoinTable(NROWS, 1000, 0)
    For i = 0 To 99: arms(i) = Array("K", "=", i + 1): Next i
    For r = -1 To 2
        Set T = BuildJoinTable(NROWS, 2, 0)
        T.GetHashIndex "K"
        TicStart
        changed = UnivDB.Query(T).Where("K", 1).SetVal("K", 3).Update()
        elapsed = TicMs
        If r >= 0 Then msUpdate(r) = elapsed
        LogShim.AssertEq "bench_batch_update_count_" & r, NROWS \ 2, changed
        LogShim.AssertEq "bench_batch_update_index_" & r, NROWS \ 2, T.SelectRows("K", 3).Count

        TicStart
        Set q = UnivDB.Query(source).FilterAny(arms)
        elapsed = TicMs
        If r >= 0 Then msOr(r) = elapsed
        Set result = q.ToTable()
        LogShim.AssertEq "bench_numeric_or_count_" & r, 10001, result.RowCount
        LogShim.AssertEq "bench_numeric_or_first_" & r, 1, result.Value(2, "Val")
        LogShim.AssertEq "bench_numeric_or_last_" & r, 99100, result.Value(result.RowCount, "Val")
    Next r
    LogShim.LogEvent "bench_db_query_writes_or", _
        Array("rows", "arms", "update_warm_ms", "numeric_or_ms"), _
        Array(NROWS, 100, Median3(msUpdate(0), msUpdate(1), msUpdate(2)), _
              Median3(msOr(0), msOr(1), msOr(2)))
End Sub

' @desc 일괄 추가와 복수 컬럼 삭제의 비용을 기록함
Public Sub Test_BenchDb_BulkAppendAndDrop()
    Dim T As clsTable, batch As clsTable, r As Long, i As Long, drops As String
    Dim msAppend(0 To 2) As Double, msDrop(0 To 2) As Double, elapsed As Double
    Set batch = BuildJoinTable(10000, 10000, 50000)
    For i = 10 To 29
        If Len(drops) > 0 Then drops = drops & ","
        drops = drops & "Val" & i
    Next i
    For r = -1 To 2
        Set T = BuildJoinTable(50000, 50000, 0)
        TicStart
        T.AppendFrom batch
        elapsed = TicMs
        If r >= 0 Then msAppend(r) = elapsed
        LogShim.AssertEq "bench_append_rows_" & r, 60001, T.RowCount
        LogShim.AssertEq "bench_append_first_" & r, 50001, T.Value(50002, "K")
        LogShim.AssertEq "bench_append_last_" & r, 60000, T.Value(60001, "K")
        Set T = BuildWideJoinTable(20000, 30)
        TicStart
        T.RemoveColumns drops
        elapsed = TicMs
        If r >= 0 Then msDrop(r) = elapsed
        LogShim.AssertEq "bench_drop_cols_" & r, 10, T.ColCount
        LogShim.AssertEq "bench_drop_value_" & r, 200010, T.Value(20001, "Val9")
    Next r
    LogShim.LogEvent "bench_db_bulk_append_drop", Array("append_ms", "drop_ms"), _
        Array(Median3(msAppend(0), msAppend(1), msAppend(2)), Median3(msDrop(0), msDrop(1), msDrop(2)))
End Sub

' ==============================================================================
' 8. 2026-10-31 성능 계획 벤치 (docs/tasks/2026-10-31_perf_implementation_plan.md)
' 데이터 형태: 마스터형 27,401 x 34 / DW형 91,000 x 27, 셀 2/3 문자열 (MakePlanData)
' ==============================================================================

' @desc 마스터형 테이블의 인덱스 warm 소량 Select 물질화 비용을 기록함
Public Sub Test_BenchDb_PlanSelectSmall()
    Dim T As clsTable, tRes As clsTable, r As Long, el As Double, ms(0 To 2) As Double

    Set T = PlanTable(PlanMaster())
    T.GetHashIndex "G"
    For r = -1 To 2
        TicStart
        Set tRes = UnivDB.Query(T).Where("G", "G7").ToTable()
        el = TicMs
        If r >= 0 Then ms(r) = el
        LogShim.AssertEq "bench_plan_select_rows_" & r, CLng(549), tRes.RowCount
        LogShim.AssertEq "bench_plan_select_last_" & r, "R25TA00027357", tRes.Value(549, "K")
    Next r
    LogShim.LogEvent "bench_db_select_small", Array("rows", "cols", "result_rows", "ms"), _
        Array(PLAN_MASTER_ROWS, PLAN_MASTER_COLS, 548, Median3(ms(0), ms(1), ms(2)))
End Sub

' @desc 마스터형 x DW형 inner/left anti 조인 비용을 기록함 (인덱스 warm)
Public Sub Test_BenchDb_PlanJoin()
    Dim tL As clsTable, tR As clsTable, tRes As clsTable, r As Long, el As Double
    Dim msInner(0 To 2) As Double, msAnti(0 To 2) As Double

    Set tL = PlanTable(PlanMaster())
    Set tR = PlanTable(MakePlanData(PLAN_DW_ROWS, PLAN_DW_COLS, PLAN_MASTER_ROWS - 3000))
    tL.GetHashIndex "K"
    tR.GetHashIndex "K"
    For r = -1 To 2
        TicStart
        Set tRes = UnivDB.Join(tL, tR, "K", jtInner, "R_")
        el = TicMs
        If r >= 0 Then msInner(r) = el
        LogShim.AssertEq "bench_plan_join_inner_rows_" & r, CLng(3001), tRes.RowCount
        LogShim.AssertEq "bench_plan_join_inner_cols_" & r, CLng(PLAN_MASTER_COLS + PLAN_DW_COLS), tRes.ColCount
        Set tRes = Nothing

        TicStart
        Set tRes = UnivDB.Join(tL, tR, "K", jtLeftAnti)
        el = TicMs
        If r >= 0 Then msAnti(r) = el
        LogShim.AssertEq "bench_plan_join_anti_rows_" & r, CLng(PLAN_MASTER_ROWS - 3000 + 1), tRes.RowCount
        Set tRes = Nothing
    Next r
    LogShim.LogEvent "bench_db_join_large", _
        Array("left_rows", "right_rows", "match_keys", "inner_ms", "left_anti_ms"), _
        Array(PLAN_MASTER_ROWS, PLAN_DW_ROWS, 3000, Median3(msInner(0), msInner(1), msInner(2)), _
              Median3(msAnti(0), msAnti(1), msAnti(2)))
End Sub

' @desc 마스터형 전 행 갱신 UpsertFrom 비용을 기록함 (대상 인덱스 warm)
Public Sub Test_BenchDb_PlanUpsert()
    Dim tgt As clsTable, src As clsTable, r As Long, el As Double, ms(0 To 2) As Double

    Set src = PlanTable(PlanMaster())
    For r = -1 To 2
        Set tgt = PlanTable(PlanMaster())
        tgt.GetHashIndex "K"
        TicStart
        tgt.UpsertFrom src, "K"
        el = TicMs
        If r >= 0 Then ms(r) = el
        LogShim.AssertEq "bench_plan_upsert_rows_" & r, CLng(PLAN_MASTER_ROWS + 1), tgt.RowCount
        Set tgt = Nothing
    Next r
    LogShim.LogEvent "bench_db_upsert", Array("rows", "cols", "all_update_ms"), _
        Array(PLAN_MASTER_ROWS, PLAN_MASTER_COLS, Median3(ms(0), ms(1), ms(2)))
End Sub

' @desc 마스터형 테이블에 Query.Insert 1행 추가 비용을 기록함
Public Sub Test_BenchDb_PlanQueryInsert()
    Dim T As clsTable, r As Long, el As Double, ms(0 To 2) As Double

    Set T = PlanTable(PlanMaster())
    For r = -1 To 2
        TicStart
        UnivDB.Query(T).SetVal("K", "NEW" & r).SetVal("C3", 1.5).Insert
        el = TicMs
        If r >= 0 Then ms(r) = el
        LogShim.AssertEq "bench_plan_insert_rows_" & r, CLng(PLAN_MASTER_ROWS + 3 + r), T.RowCount
        LogShim.AssertEq "bench_plan_insert_key_" & r, "NEW" & r, T.Value(T.RowCount, "K")
    Next r
    LogShim.LogEvent "bench_db_query_insert", Array("rows", "cols", "one_row_ms"), _
        Array(PLAN_MASTER_ROWS, PLAN_MASTER_COLS, Median3(ms(0), ms(1), ms(2)))
End Sub

' @desc DW형 91k 행 고유 키/50그룹 인덱스 빌드 비용을 기록함
Public Sub Test_BenchDb_PlanIndexBuild()
    Dim v As Variant, T As clsTable, idx As DbHashIndex, r As Long, el As Double
    Dim msUniq(0 To 2) As Double, msGrp(0 To 2) As Double

    v = MakePlanData(PLAN_DW_ROWS, 2, 0)
    For r = -1 To 2
        Set T = PlanTable(v)
        TicStart
        Set idx = T.GetHashIndex("K")
        el = TicMs
        If r >= 0 Then msUniq(r) = el
        LogShim.AssertEq "bench_plan_idx_uniq_keys_" & r, PLAN_DW_ROWS, idx.CountKeys
        Set idx = Nothing: Set T = Nothing

        Set T = PlanTable(v)
        TicStart
        Set idx = T.GetHashIndex("G")
        el = TicMs
        If r >= 0 Then msGrp(r) = el
        LogShim.AssertEq "bench_plan_idx_grp_keys_" & r, 50, idx.CountKeys
        Set idx = Nothing: Set T = Nothing
    Next r
    LogShim.LogEvent "bench_db_index_build_large", Array("rows", "unique_ms", "group50_ms"), _
        Array(PLAN_DW_ROWS, Median3(msUniq(0), msUniq(1), msUniq(2)), Median3(msGrp(0), msGrp(1), msGrp(2)))
End Sub

' @desc 마스터형 시트/배열 적재(LoadFromSheet/LoadFromArray, Trim 기본/생략) 비용을 기록함
Public Sub Test_BenchDb_PlanLoad()
    Dim ws As Worksheet, T As clsTable, v As Variant, r As Long, el As Double
    Dim msSheet(0 To 2) As Double, msArray(0 To 2) As Double
    Dim msSheetNo(0 To 2) As Double, msArrayNo(0 To 2) As Double

    v = PlanMaster()
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Range(ws.Cells(1, 1), ws.Cells(PLAN_MASTER_ROWS + 1, PLAN_MASTER_COLS)).Value2 = v
    For r = -1 To 2
        Set T = New clsTable
        TicStart
        T.LoadFromSheet ws
        el = TicMs
        If r >= 0 Then msSheet(r) = el
        LogShim.AssertEq "bench_plan_load_sheet_rows_" & r, CLng(PLAN_MASTER_ROWS + 1), T.RowCount
        LogShim.AssertEq "bench_plan_load_sheet_cols_" & r, CLng(PLAN_MASTER_COLS), T.ColCount
        Set T = Nothing

        Set T = New clsTable
        TicStart
        T.LoadFromArray v, True
        el = TicMs
        If r >= 0 Then msArray(r) = el
        LogShim.AssertEq "bench_plan_load_array_rows_" & r, CLng(PLAN_MASTER_ROWS + 1), T.RowCount
        Set T = Nothing

        Set T = New clsTable
        TicStart
        T.LoadFromSheet ws, False
        el = TicMs
        If r >= 0 Then msSheetNo(r) = el
        LogShim.AssertEq "bench_plan_load_sheet_notrim_rows_" & r, CLng(PLAN_MASTER_ROWS + 1), T.RowCount
        Set T = Nothing

        Set T = New clsTable
        TicStart
        T.LoadFromArray v, True, False
        el = TicMs
        If r >= 0 Then msArrayNo(r) = el
        LogShim.AssertEq "bench_plan_load_array_notrim_rows_" & r, CLng(PLAN_MASTER_ROWS + 1), T.RowCount
        Set T = Nothing
    Next r
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    LogShim.LogEvent "bench_db_load", _
        Array("rows", "cols", "sheet_ms", "array_ms", "sheet_notrim_ms", "array_notrim_ms"), _
        Array(PLAN_MASTER_ROWS, PLAN_MASTER_COLS, Median3(msSheet(0), msSheet(1), msSheet(2)), _
              Median3(msArray(0), msArray(1), msArray(2)), Median3(msSheetNo(0), msSheetNo(1), msSheetNo(2)), _
              Median3(msArrayNo(0), msArrayNo(1), msArrayNo(2)))
End Sub

Private Function PlanMaster() As Variant
    If IsEmpty(mPlanMaster) Then mPlanMaster = MakePlanData(PLAN_MASTER_ROWS, PLAN_MASTER_COLS, 0)
    PlanMaster = mPlanMaster
End Function

Private Function PlanTable(ByRef v As Variant) As clsTable
    Dim T As New clsTable
    T.LoadFromArray v, True
    Set PlanTable = T
End Function

' 1열 K(고유 키 R25TA########, 오프셋), 2열 G(50그룹), 나머지는 숫자/영문/한글 혼합
Private Function MakePlanData(ByVal nRows As Long, ByVal nCols As Long, ByVal keyOffset As Long) As Variant
    Dim v As Variant, i As Long, j As Long, hangul As String
    hangul = ChrW(&HAC00) & ChrW(&HB098) & ChrW(&HC0C1) & ChrW(&HC0AC)
    ReDim v(1 To nRows + 1, 1 To nCols)
    v(1, 1) = "K": v(1, 2) = "G"
    For j = 3 To nCols: v(1, j) = "C" & j: Next j
    For i = 2 To nRows + 1
        v(i, 1) = "R25TA" & Format$(i - 1 + keyOffset, "00000000")
        v(i, 2) = "G" & ((i - 1) Mod 50)
        For j = 3 To nCols
            Select Case j Mod 3
                Case 0: v(i, j) = CDbl(i) * 1.5 + j
                Case 1: v(i, j) = "V" & Format$(i, "00000000")
                Case 2: v(i, j) = hangul & (i Mod 977)
            End Select
        Next j
    Next i
    MakePlanData = v
End Function

Private Function GetBench20k() As clsTable
    If mT20k Is Nothing Then Set mT20k = BuildBenchTable(BENCH_ROWS)
    Set GetBench20k = mT20k
End Function

' 컬럼: ID(고유) / Dept(5종) / Score(0..999) / Tag(문자) / LongKey(180자 pad 50종)
Private Function BuildBenchTable(ByVal lRows As Long) As clsTable
    Dim vData() As Variant
    Dim T As New clsTable
    Dim i As Long
    Dim pad As String

    pad = String$(180, "p")
    ReDim vData(1 To lRows + 1, 1 To 5)
    vData(1, 1) = "ID": vData(1, 2) = "Dept": vData(1, 3) = "Score"
    vData(1, 4) = "Tag": vData(1, 5) = "LongKey"
    For i = 1 To lRows
        vData(i + 1, 1) = i
        vData(i + 1, 2) = "D" & (i Mod 5)
        vData(i + 1, 3) = i Mod 1000
        vData(i + 1, 4) = "TAG" & (i Mod 97)
        vData(i + 1, 5) = pad & (i Mod 50)
    Next i

    T.LoadFromArray vData
    Set BuildBenchTable = T
End Function

' K = i (고유 키 1:1), Val1..Val(nCols-1) = 수치 채움
Private Function BuildWideJoinTable(ByVal nRows As Long, ByVal nCols As Long) As clsTable
    Dim vData() As Variant
    Dim T As New clsTable
    Dim i As Long, c As Long

    ReDim vData(1 To nRows + 1, 1 To nCols)
    vData(1, 1) = "K"
    For c = 2 To nCols
        vData(1, c) = "Val" & (c - 1)
    Next c
    For i = 1 To nRows
        vData(i + 1, 1) = i
        For c = 2 To nCols
            vData(i + 1, c) = i * 10 + c
        Next c
    Next i

    T.LoadFromArray vData
    Set BuildWideJoinTable = T
End Function

' K = lOffset + ((i-1) Mod nDistinct) + 1 (nDistinct = nRows면 고유 키), Val = i
Private Function BuildJoinTable(ByVal nRows As Long, ByVal nDistinct As Long, _
                                ByVal lOffset As Long) As clsTable
    Dim vData() As Variant
    Dim T As New clsTable
    Dim i As Long

    ReDim vData(1 To nRows + 1, 1 To 2)
    vData(1, 1) = "K": vData(1, 2) = "Val"
    For i = 1 To nRows
        vData(i + 1, 1) = lOffset + ((i - 1) Mod nDistinct) + 1
        vData(i + 1, 2) = i
    Next i

    T.LoadFromArray vData
    Set BuildJoinTable = T
End Function
