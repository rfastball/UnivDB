Attribute VB_Name = "ZPerfProbe2"
' @desc 내부 전체 복사와 인덱스 고정비 실측 담당 (임시)
' @depends UnivDB, modEngine, modIndexKey, LogShim
Option Explicit

' 임시 측정 모듈 (2026-10-31 개선 후보 검증). 측정 후 삭제한다.

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
#End If

Private Const MASTER_ROWS As Long = 27401
Private Const MASTER_COLS As Long = 34

Private mFreq As Currency
Private mT0 As Currency
Private mMs(0 To 2) As Double

Private Sub TicStart()
    If mFreq = 0 Then QueryPerformanceFrequency mFreq
    QueryPerformanceCounter mT0
End Sub

Private Function TicMs() As Double
    Dim t1 As Currency
    QueryPerformanceCounter t1
    TicMs = (t1 - mT0) / mFreq * 1000#
End Function

Private Sub Keep(ByVal r As Long, ByVal el As Double)
    If r >= 0 Then mMs(r) = el
End Sub

Private Function Med() As Double
    Dim mx As Double, mn As Double
    mx = mMs(0): If mMs(1) > mx Then mx = mMs(1)
    If mMs(2) > mx Then mx = mMs(2)
    mn = mMs(0): If mMs(1) < mn Then mn = mMs(1)
    If mMs(2) < mn Then mn = mMs(2)
    Med = mMs(0) + mMs(1) + mMs(2) - mx - mn
End Function

Private Sub LogCase(ByVal sCase As String, ByVal sPath As String, ByVal rows As Long)
    LogShim.LogEvent "probe2", Array("case", "path", "rows", "ms"), Array(sCase, sPath, rows, Med())
End Sub

' 1열 K(고유 키, 오프셋), 2열 G(50그룹), 나머지는 숫자/영문/한글 혼합
Private Function MakeData(ByVal nRows As Long, ByVal nCols As Long, ByVal keyOffset As Long) As Variant
    Dim v As Variant, i As Long, j As Long
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
                Case 2: v(i, j) = ChrW(&HAC00) & ChrW(&HB098) & ChrW(&HC0C1) & ChrW(&HC0AC) & (i Mod 977)
            End Select
        Next j
    Next i
    MakeData = v
End Function

Private Function MakeTable(ByRef v As Variant) As clsTable
    Dim tbl As clsTable
    Set tbl = New clsTable
    tbl.LoadFromArray v, True
    Set MakeTable = tbl
End Function

' @desc SelectTable/AdoptArray 물질화의 전체 복사 비중을 기록함
Public Sub Probe_Materialize()
    Dim vSrc As Variant, tbl As clsTable, res As clsTable, v As Variant, arr As Variant
    Dim r As Long, el As Double, nRes As Long

    vSrc = MakeData(MASTER_ROWS, MASTER_COLS, 0)
    Set tbl = MakeTable(vSrc)
    tbl.GetHashIndex "G"

    For r = -1 To 2
        TicStart
        v = tbl.RawArray
        v = Empty
        Keep r, TicMs
    Next r
    LogCase "materialize", "rawarray_copy_release", MASTER_ROWS

    For r = -1 To 2
        TicStart
        Set res = UnivDB.Query(tbl).Where("G", "G7").ToTable()
        Keep r, TicMs
        nRes = res.RowCount - 1
        Set res = Nothing
    Next r
    LogCase "materialize", "select_where_small_total", nRes

    For r = -1 To 2
        TicStart
        Set res = UnivDB.Query(tbl).ToTable()
        Keep r, TicMs
        nRes = res.RowCount - 1
        Set res = Nothing
    Next r
    LogCase "materialize", "select_all_total", nRes

    For r = -1 To 2
        arr = vSrc
        Set res = New clsTable
        TicStart
        res.AdoptArray arr, True
        Keep r, TicMs
        Set res = Nothing
    Next r
    LogCase "materialize", "adopt_copy_full", MASTER_ROWS
End Sub

' @desc 조인의 양쪽 RawArray 복사 비중을 기록함
Public Sub Probe_Join()
    Dim tL As clsTable, tR As clsTable, res As clsTable, v As Variant
    Dim r As Long, nRes As Long

    Set tL = MakeTable(MakeData(MASTER_ROWS, MASTER_COLS, 0))
    Set tR = MakeTable(MakeData(91000, 27, MASTER_ROWS - 3000))
    tL.GetHashIndex "K"
    tR.GetHashIndex "K"

    For r = -1 To 2
        TicStart
        Set res = UnivDB.Join(tL, tR, "K", jtInner, "R_")
        Keep r, TicMs
        nRes = res.RowCount - 1
        Set res = Nothing
    Next r
    LogCase "join", "inner_3k_match_total", nRes

    For r = -1 To 2
        TicStart
        Set res = UnivDB.Join(tL, tR, "K", jtLeftAnti)
        Keep r, TicMs
        nRes = res.RowCount - 1
        Set res = Nothing
    Next r
    LogCase "join", "left_anti_total", nRes

    For r = -1 To 2
        TicStart
        v = tL.RawArray
        v = Empty
        Keep r, TicMs
    Next r
    LogCase "join", "rawarray_left_27k_x34", MASTER_ROWS

    For r = -1 To 2
        TicStart
        v = tR.RawArray
        v = Empty
        Keep r, TicMs
    Next r
    LogCase "join", "rawarray_right_91k_x27", 91000
End Sub

' 공개 경로를 거치지 않는 일괄 빌드 원형: 2D 배열 직접 읽기 + 인라인 정규화 + CSR
Private Function BulkBuild(ByRef v As Variant, ByVal col As Long, ByVal singleLookup As Boolean) As Long
    Dim d As Object, p As DbCsrPayload, i As Long, s As String, keyId As Long, c0 As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set p = New DbCsrPayload
    For i = 2 To UBound(v, 1)
        If VarType(v(i, col)) = vbString Then
            s = v(i, col)
        ElseIf IsError(v(i, col)) Then
            GoTo NextRow
        Else
            s = modIndexKey.NormalizeIndexKey(v(i, col))
        End If
        If singleLookup Then
            c0 = d.Count
            keyId = d(s)
            If d.Count > c0 Then
                keyId = c0 + 1
                d(s) = keyId
            End If
        Else
            If d.exists(s) Then
                keyId = d(s)
            Else
                keyId = d.Count + 1
                d.Add s, keyId
            End If
        End If
        p.AddEntry keyId, i
NextRow:
    Next i
    BulkBuild = d.Count
End Function

Private Sub RunIndexOverhead(ByVal n As Long)
    Dim vSrc As Variant, tbl As clsTable, idx As DbHashIndex, r As Long, col As Long, nk As Long
    Dim names As Variant
    names = Array("", "unique_K", "dup50_G")
    vSrc = MakeData(n, 2, 0)
    For col = 1 To 2
        For r = -1 To 2
            Set tbl = MakeTable(vSrc)
            TicStart
            Set idx = tbl.GetHashIndexByIndex(col)
            Keep r, TicMs
            nk = idx.CountKeys
            Set idx = Nothing: Set tbl = Nothing
        Next r
        LogCase "index_" & names(col), "real_dict", n

        For r = -1 To 2
            TicStart
            nk = BulkBuild(vSrc, col, False)
            Keep r, TicMs
        Next r
        LogCase "index_" & names(col), "bulk_direct_exists_add", n

        For r = -1 To 2
            TicStart
            nk = BulkBuild(vSrc, col, True)
            Keep r, TicMs
        Next r
        LogCase "index_" & names(col), "bulk_direct_single_lookup", n
    Next col
End Sub

' @desc 인덱스 빌드 공개 경로의 고정비와 일괄 빌드 원형을 기록함
Public Sub Probe_IndexOverhead()
    RunIndexOverhead MASTER_ROWS
    RunIndexOverhead 91000
End Sub

' @desc UpsertFrom 중복 선검증과 스탬프 배열 대안의 비용을 기록함
Public Sub Probe_Upsert()
    Dim vSrc As Variant, tgt As clsTable, src As clsTable, idx As DbHashIndex
    Dim vS As Variant, r As Long, i As Long, row As Long, nDup As Long
    Dim seenKeys As Object, nk As String, stamp() As Long, tgtRow() As Long

    vSrc = MakeData(MASTER_ROWS, MASTER_COLS, 0)
    Set src = MakeTable(vSrc)
    vS = src.RawArray

    For r = -1 To 2
        TicStart
        Set seenKeys = CreateObject("Scripting.Dictionary")
        seenKeys.CompareMode = vbTextCompare
        For i = 2 To UBound(vS, 1)
            If Not IsError(vS(i, 1)) Then
                nk = modIndexKey.NormalizeIndexKey(vS(i, 1))
                If seenKeys.exists(nk) Then nDup = nDup + 1
                seenKeys.Add nk, i
            End If
        Next i
        Set seenKeys = Nothing
        Keep r, TicMs
    Next r
    LogCase "upsert", "seenkeys_pass_current", MASTER_ROWS

    Set tgt = MakeTable(vSrc)
    Set idx = tgt.GetHashIndex("K")
    For r = -1 To 2
        TicStart
        For i = 2 To UBound(vS, 1)
            idx.TryGetFirstRowId vS(i, 1), row
        Next i
        Keep r, TicMs
    Next r
    LogCase "upsert", "lookup_pass_only", MASTER_ROWS

    For r = -1 To 2
        TicStart
        ReDim stamp(1 To tgt.RowCount)
        ReDim tgtRow(2 To UBound(vS, 1))
        For i = 2 To UBound(vS, 1)
            If idx.TryGetFirstRowId(vS(i, 1), row) Then
                If stamp(row) <> 0 Then nDup = nDup + 1 Else stamp(row) = i
                tgtRow(i) = row
            End If
        Next i
        Keep r, TicMs
    Next r
    LogCase "upsert", "lookup_plus_stamp", MASTER_ROWS
    LogShim.AssertEq "probe2_upsert_no_dup", 0, nDup

    For r = -1 To 2
        Set tgt = MakeTable(vSrc)
        tgt.GetHashIndex "K"
        TicStart
        tgt.UpsertFrom src, "K"
        Keep r, TicMs
        Set tgt = Nothing
    Next r
    LogCase "upsert", "upsertfrom_total_all_update", MASTER_ROWS
End Sub

' @desc Query.Insert 1행 비용을 기록함
Public Sub Probe_QueryInsert()
    Dim tbl As clsTable, r As Long
    Set tbl = MakeTable(MakeData(MASTER_ROWS, MASTER_COLS, 0))
    For r = -1 To 2
        TicStart
        UnivDB.Query(tbl).SetVal("K", "NEW" & r).Insert
        Keep r, TicMs
    Next r
    LogCase "insert", "query_insert_one_row", MASTER_ROWS
End Sub

' @desc 조인 내부 단계(커서 생성, 키 탐침 순회) 비용을 기록함
Public Sub Probe_JoinParts()
    Dim tL As clsTable, tR As clsTable, idxL As DbHashIndex, idxR As DbHashIndex
    Dim cur As DbIndexCursor, r As Long, cnt As Long, nHit As Long

    Set tL = MakeTable(MakeData(MASTER_ROWS, MASTER_COLS, 0))
    Set tR = MakeTable(MakeData(91000, 27, MASTER_ROWS - 3000))
    Set idxL = tL.GetHashIndex("K")
    Set idxR = tR.GetHashIndex("K")

    For r = -1 To 2
        TicStart
        Set cur = idxR.Cursor
        Keep r, TicMs
        Set cur = Nothing
    Next r
    LogCase "join_parts", "cursor_create_right_91k", 91000

    For r = -1 To 2
        Set cur = idxR.Cursor
        nHit = 0
        TicStart
        Do While cur.MoveNext
            If idxL.TryGetRowCount(cur.Key, cnt) Then nHit = nHit + 1
        Loop
        Keep r, TicMs
    Next r
    LogCase "join_parts", "probe_pass_right_cursor_x_left", 91000
    LogShim.AssertEq "probe2_join_hits", 3000, nHit

    For r = -1 To 2
        Set cur = idxL.Cursor
        nHit = 0
        TicStart
        Do While cur.MoveNext
            If idxR.TryGetRowCount(cur.Key, cnt) Then nHit = nHit + 1
        Loop
        Keep r, TicMs
    Next r
    LogCase "join_parts", "probe_pass_left_cursor_x_right", MASTER_ROWS

    For r = -1 To 2
        Set cur = idxR.Cursor
        TicStart
        Do While cur.MoveNext
        Loop
        Keep r, TicMs
    Next r
    LogCase "join_parts", "cursor_walk_only_91k", 91000
End Sub

' @desc 개선 후보 측정 전체를 순서대로 실행함
Public Sub Probe_All()
    ZPerfProbe2.Probe_Materialize
    ZPerfProbe2.Probe_Join
    ZPerfProbe2.Probe_IndexOverhead
    ZPerfProbe2.Probe_Upsert
    ZPerfProbe2.Probe_QueryInsert
End Sub
