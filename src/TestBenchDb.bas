Attribute VB_Name = "TestBenchDb"
' @desc DB 레이어 성능 baseline 측정 벤치마크 담당
' @depends UnivDB, modEngine, modIndexKey, LogShim
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

Private mFreq As Currency
Private mT0 As Currency
Private mT20k As clsTable ' 공유 벤치 테이블 (지연 생성, 테스트 간 재사용)

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
