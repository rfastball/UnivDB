Attribute VB_Name = "ZPerfProbe"
' @desc 성능 브리프 검증용 키 접기와 Trim 비용 실측 담당
' @depends modIndexKey, LogShim
Option Explicit

' 임시 측정 모듈 (2026-10-31 브리프 검증). 측정 후 삭제한다.

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
Private Declare PtrSafe Function LCMapStringW Lib "kernel32" (ByVal locale As Long, ByVal flags As Long, ByVal src As LongPtr, ByVal srcLen As Long, ByVal dest As LongPtr, ByVal destLen As Long) As Long
Private Declare PtrSafe Function GetThreadLocale Lib "kernel32" () As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
Private Declare Function LCMapStringW Lib "kernel32" (ByVal locale As Long, ByVal flags As Long, ByVal src As Long, ByVal srcLen As Long, ByVal dest As Long, ByVal destLen As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
#End If

Private Const LCMAP_UPPERCASE As Long = &H200
Private Const LCMAP_HALFWIDTH As Long = &H400000

Private Const MODE_TEXT As Long = 0
Private Const MODE_BINARY_RAW As Long = 1
Private Const MODE_UCASE As Long = 2
Private Const MODE_LCMAP As Long = 3
Private Const MODE_STRCONV As Long = 4
Private Const MODE_LCMAP_INPLACE As Long = 5
Private Const MODE_HYBRID_BYTES As Long = 6
Private Const MODE_HYBRID_LIKE As Long = 7

Private mWidthPattern As String

Private mFreq As Currency
Private mT0 As Currency
Private mBuf As String
Private mLocale As Long

Private Sub TicStart()
    If mFreq = 0 Then QueryPerformanceFrequency mFreq
    QueryPerformanceCounter mT0
End Sub

Private Function TicMs() As Double
    Dim t1 As Currency
    QueryPerformanceCounter t1
    TicMs = (t1 - mT0) / mFreq * 1000#
End Function

Private Function Median3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    Dim mx As Double, mn As Double
    mx = a: If b > mx Then mx = b
    If c > mx Then mx = c
    mn = a: If b < mn Then mn = b
    If c < mn Then mn = c
    Median3 = a + b + c - mx - mn
End Function

' 폭/가나 접기가 의미를 갖는 영역: U+3000-30FF, U+3131-318E, U+FF00-FFFF
Private Function InWidthRange(ByVal cu As Long) As Boolean
    InWidthRange = (cu >= &H3000& And cu <= &H30FF&) Or (cu >= &H3131& And cu <= &H318E&) Or (cu >= &HFF00&)
End Function

Private Function NeedsFoldBytes(ByRef s As String) As Boolean
    Dim b() As Byte, i As Long, hi As Long
    If Len(s) = 0 Then Exit Function
    b = s
    For i = 0 To UBound(b) Step 2
        hi = b(i + 1)
        If hi = &H30 Or hi = &H31 Or hi = &HFF Then
            If InWidthRange(hi * 256& + b(i)) Then NeedsFoldBytes = True: Exit Function
        End If
    Next i
End Function

Private Function NeedsFoldLike(ByRef s As String) As Boolean
    If Len(mWidthPattern) = 0 Then
        mWidthPattern = "*[" & ChrW(&H3000) & "-" & ChrW(&H30FF) & ChrW(&H3131) & "-" & ChrW(&H318E) & _
                        ChrW(&HFF00) & "-" & ChrW(&HFFFF) & "]*"
    End If
    NeedsFoldLike = (s Like mWidthPattern)
End Function

Private Function FoldLc(ByRef s As String) As String
    Dim n As Long
    If Len(s) = 0 Then Exit Function
    If Len(mBuf) < Len(s) * 2 Then mBuf = String$(Len(s) * 2 + 64, vbNullChar)
    n = LCMapStringW(mLocale, LCMAP_UPPERCASE Or LCMAP_HALFWIDTH, StrPtr(s), Len(s), StrPtr(mBuf), Len(mBuf))
    If n = 0 Then Err.Raise 9999, "ZPerfProbe.FoldLc", "LCMapStringW failed"
    FoldLc = Left$(mBuf, n)
End Function

Private Function FoldKey(ByRef s As String, ByVal mode As Long) As String
    Select Case mode
        Case MODE_UCASE: FoldKey = UCase$(s)
        Case MODE_LCMAP: FoldKey = FoldLc(s)
        Case MODE_STRCONV: FoldKey = StrConv(s, vbUpperCase + vbNarrow)
        Case Else: FoldKey = s
    End Select
End Function

Private Function ModeName(ByVal mode As Long) As String
    ModeName = Choose(mode + 1, "dict_text", "dict_binary_raw", "ucase_binary", "lcmap_binary", "strconv_binary", "lcmap_inplace_binary", _
                      "hybrid_bytes_binary", "hybrid_like_binary")
End Function

' DictHashIndex.AddRow 경로 복제: NormalizeIndexKey -> (접기) -> exists/Add -> CSR AddEntry
Private Function TimeInlineBuild(ByRef vKeys As Variant, ByVal n As Long, ByVal mode As Long, ByRef outKeys As Long) As Double
    Dim d As Object, p As DbCsrPayload, i As Long, keyId As Long, fl As Long
    Dim s As String, k As String, orig() As String

    Set d = CreateObject("Scripting.Dictionary")
    If mode = MODE_TEXT Then d.CompareMode = vbTextCompare
    Set p = New DbCsrPayload
    ReDim orig(1 To 1024)

    TicStart
    For i = 1 To n
        s = modIndexKey.NormalizeIndexKey(vKeys(i))
        Select Case mode
            Case MODE_TEXT, MODE_BINARY_RAW: k = s
            Case MODE_UCASE: k = UCase$(s)
            Case MODE_LCMAP_INPLACE
                k = s
                If Len(k) > 0 Then
                    fl = LCMapStringW(mLocale, LCMAP_UPPERCASE Or LCMAP_HALFWIDTH, StrPtr(s), Len(s), StrPtr(k), Len(k))
                    If fl = 0 Then
                        k = FoldLc(s)
                    ElseIf fl < Len(k) Then
                        k = Left$(k, fl)
                    End If
                End If
            Case MODE_HYBRID_BYTES
                If NeedsFoldBytes(s) Then k = FoldLc(s) Else k = UCase$(s)
            Case MODE_HYBRID_LIKE
                If NeedsFoldLike(s) Then k = FoldLc(s) Else k = UCase$(s)
            Case Else: k = FoldKey(s, mode)
        End Select
        If d.exists(k) Then
            keyId = d(k)
        Else
            keyId = d.Count + 1
            d.Add k, keyId
            If mode <> MODE_TEXT Then
                If keyId > UBound(orig) Then ReDim Preserve orig(1 To UBound(orig) * 2)
                orig(keyId) = s
            End If
        End If
        p.AddEntry keyId, i
    Next i
    TimeInlineBuild = TicMs
    outKeys = d.Count
End Function

' 공개 경로: clsTable.GetHashIndex (현행 Dict/Core)
Private Function TimeRealBuild(ByRef vTbl As Variant, ByVal backend As Long, ByRef outKeys As Long) As Double
    Dim T As clsTable, idx As DbHashIndex
    Set T = New clsTable
    T.LoadFromArray vTbl, True
    T.IndexBackend = backend
    TicStart
    Set idx = T.GetHashIndex("K")
    TimeRealBuild = TicMs
    outKeys = idx.CountKeys
End Function

Private Sub MakeKeys(ByVal n As Long, ByVal shape As String, ByRef vKeys As Variant, ByRef vTbl As Variant)
    Dim i As Long, s As String
    ReDim vKeys(1 To n)
    ReDim vTbl(1 To n + 1, 1 To 2)
    vTbl(1, 1) = "K": vTbl(1, 2) = "Val"
    For i = 1 To n
        If shape = "ascii" Then
            s = "R25TA" & Format$(i, "00000000")
        Else
            s = ChrW(&HAC00) & ChrW(&HB098) & ChrW(&HC0C1) & ChrW(&HC0AC) & Format$(i, "000000")
        End If
        vKeys(i) = s
        vTbl(i + 1, 1) = s
        vTbl(i + 1, 2) = i
    Next i
End Sub

Private Sub RunIndexCase(ByVal n As Long, ByVal shape As String)
    Dim vKeys As Variant, vTbl As Variant, mode As Long, r As Long, nk As Long
    Dim ms(0 To 2) As Double, t As Double, b As Long

    MakeKeys n, shape, vKeys, vTbl
    For mode = MODE_TEXT To MODE_STRCONV
        For r = -1 To 2
            t = TimeInlineBuild(vKeys, n, mode, nk)
            If r >= 0 Then ms(r) = t
        Next r
        LogShim.AssertEq "probe_keys_" & ModeName(mode) & "_" & n, n, nk
        LogShim.LogEvent "probe_index_build", Array("shape", "keys", "path", "ms"), _
            Array(shape, n, ModeName(mode), Median3(ms(0), ms(1), ms(2)))
    Next mode
    For b = dbBackendDict To dbBackendCore
        For r = -1 To 2
            t = TimeRealBuild(vTbl, b, nk)
            If r >= 0 Then ms(r) = t
        Next r
        LogShim.AssertEq "probe_real_keys_" & b & "_" & n, n, nk
        LogShim.LogEvent "probe_index_build", Array("shape", "keys", "path", "ms"), _
            Array(shape, n, IIf(b = dbBackendCore, "real_core", "real_dict"), Median3(ms(0), ms(1), ms(2)))
    Next b
End Sub

' @desc 인덱스 빌드 경로별 비용을 27k/91k/182k 키에서 기록함
Public Sub Probe_IndexBuild()
    mLocale = GetThreadLocale()
    RunIndexCase 27000, "ascii"
    RunIndexCase 91000, "ascii"
    RunIndexCase 182000, "ascii"
    RunIndexCase 91000, "hangul"
End Sub

Private Function DictTextSame(ByVal a As String, ByVal b As String) As Boolean
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d.Add a, 1
    DictTextSame = d.exists(b)
End Function

Private Sub LogPair(ByVal label As String, ByVal a As String, ByVal b As String)
    LogShim.LogEvent "probe_key_semantics", _
        Array("pair", "dict_text", "ucase", "lcmap", "strconv", "strcomp_text"), _
        Array(label, DictTextSame(a, b), (UCase$(a) = UCase$(b)), (FoldLc(a) = FoldLc(b)), _
              (StrConv(a, vbUpperCase + vbNarrow) = StrConv(b, vbUpperCase + vbNarrow)), _
              (StrComp(a, b, vbTextCompare) = 0))
End Sub

Private Sub RunFoldCase(ByVal n As Long, ByVal shape As String)
    Dim vKeys As Variant, vTbl As Variant, modes As Variant, mode As Variant, r As Long, nk As Long
    Dim ms(0 To 2) As Double, t As Double

    MakeKeys n, shape, vKeys, vTbl
    modes = Array(MODE_TEXT, MODE_UCASE, MODE_LCMAP_INPLACE, MODE_HYBRID_BYTES, MODE_HYBRID_LIKE)
    For Each mode In modes
        For r = -1 To 2
            t = TimeInlineBuild(vKeys, n, CLng(mode), nk)
            If r >= 0 Then ms(r) = t
        Next r
        LogShim.AssertEq "probe_fold_keys_" & ModeName(CLng(mode)) & "_" & n, n, nk
        LogShim.LogEvent "probe_index_fold", Array("shape", "keys", "path", "ms"), _
            Array(shape, n, ModeName(CLng(mode)), Median3(ms(0), ms(1), ms(2)))
    Next mode
End Sub

' @desc BMP 전 코드포인트에서 하이브리드 접기와 LCMapStringW 접기 일치를 기록함
Public Sub Probe_FoldSweep()
    Dim c As Long, s As String, hy As String, lc As String, nMis As Long, nDet As Long, sample As String
    Dim nDictMis As Long, dictSample As String, d As Object, base As String
    mLocale = GetThreadLocale()
    For c = 1 To &HFFFF&
        If c < &HD800& Or c > &HDFFF& Then
            s = ChrW(c)
            lc = FoldLc(s)
            If NeedsFoldBytes(s) Then hy = lc Else hy = UCase$(s)
            If hy <> lc Then
                nMis = nMis + 1
                If nMis <= 30 Then sample = sample & Hex$(c) & " "
            End If
            If NeedsFoldBytes(s) <> NeedsFoldLike(s) Then nDet = nDet + 1
        End If
    Next c
    LogShim.LogEvent "probe_fold_sweep", Array("hybrid_vs_lcmap_mismatch", "sample", "bytes_vs_like_detect_mismatch"), _
        Array(nMis, sample, nDet)

    ' 단일 문자 동치류 양방향 대조: Dictionary text 모드 vs 하이브리드 접기 + binary
    Dim dt As Object, db As Object, nSplit As Long, splitSample As String, nMerge As Long, mergeSample As String
    Set dt = CreateObject("Scripting.Dictionary")
    dt.CompareMode = vbTextCompare
    Set db = CreateObject("Scripting.Dictionary")
    For c = 1 To &HFFFF&
        If c < &HD800& Or c > &HDFFF& Then
            s = ChrW(c)
            If NeedsFoldBytes(s) Then hy = FoldLc(s) Else hy = UCase$(s)
            If dt.exists(s) Then
                base = dt(s)
                If Not db.exists(hy) Then
                    nSplit = nSplit + 1
                    If nSplit <= 2000 Then splitSample = splitSample & Hex$(c) & "~" & Hex$(AscW(base)) & " "
                ElseIf db(hy) <> base And dt.exists(db(hy)) Then
                    If dt(db(hy)) <> base Then
                        nSplit = nSplit + 1
                        If nSplit <= 2000 Then splitSample = splitSample & Hex$(c) & "~" & Hex$(AscW(base)) & " "
                    End If
                End If
            Else
                dt.Add s, s
            End If
            If db.exists(hy) Then
                If Not dt.exists(db(hy)) Then
                ElseIf dt(db(hy)) <> dt(s) Then
                    nMerge = nMerge + 1
                    If nMerge <= 30 Then mergeSample = mergeSample & Hex$(c) & "~" & Hex$(AscW(db(hy))) & " "
                End If
            Else
                db.Add hy, s
            End If
        End If
    Next c
    LogShim.LogEvent "probe_fold_vs_dict", _
        Array("dict_merges_fold_splits", "split_sample", "fold_merges_dict_splits", "merge_sample"), _
        Array(nSplit, splitSample, nMerge, mergeSample)
End Sub

' @desc 제자리 LCMapStringW 접기 비용을 27k/91k/182k 키에서 기록함
Public Sub Probe_IndexFold()
    mLocale = GetThreadLocale()
    RunFoldCase 27000, "ascii"
    RunFoldCase 91000, "ascii"
    RunFoldCase 182000, "ascii"
    RunFoldCase 91000, "hangul"
End Sub

' @desc 키 접기 방식별 동치 판정을 Dictionary text 모드와 대조해 기록함
Public Sub Probe_KeySemantics()
    mLocale = GetThreadLocale()
    LogPair "ascii_case", "abc", "ABC"
    LogPair "fullwidth_latin", "ABC", ChrW(&HFF21) & ChrW(&HFF22) & ChrW(&HFF23)
    LogPair "fullwidth_digit", "R25TA123", "R25TA" & ChrW(&HFF11) & ChrW(&HFF12) & ChrW(&HFF13)
    LogPair "fullwidth_lower", "abc", ChrW(&HFF41) & ChrW(&HFF42) & ChrW(&HFF43)
    LogPair "halfwidth_hangul", ChrW(&H3131), ChrW(&HFFA1)
    LogPair "fullwidth_space", "a b", "a" & ChrW(&H3000) & "b"
    LogPair "kata_hira", ChrW(&H30A2), ChrW(&H3042)
    LogPair "halfwidth_kata_voiced", ChrW(&H30AC), ChrW(&HFF76) & ChrW(&HFF9E)
    LogPair "sharp_s", "strasse", "stra" & ChrW(&HDF) & "e"
    LogPair "final_sigma", ChrW(&H3C3), ChrW(&H3C2)
    LogPair "e_combining", ChrW(&HE9), "e" & ChrW(&H301)
    LogPair "accented_upper", ChrW(&HE9), ChrW(&HC9)
    LogPair "accent_vs_plain", "e", ChrW(&HE9)
    LogPair "nbsp", "a b", "a" & ChrW(&HA0) & "b"
    LogPair "hyphen", "ab", "a-b"
    LogPair "cyrillic", ChrW(&H434), ChrW(&H414)
    LogPair "roman_numeral", ChrW(&H2170), ChrW(&H2160)
    LogPair "circled", ChrW(&H24D0), ChrW(&H24B6)
    LogPair "hangul_same", ChrW(&HAC00), ChrW(&HAC00)
End Sub

Private Sub MakeWide(ByVal nRows As Long, ByVal nCols As Long, ByRef v As Variant)
    Dim i As Long, j As Long, s As String
    ReDim v(1 To nRows + 1, 1 To nCols)
    For j = 1 To nCols: v(1, j) = "C" & j: Next j
    For i = 2 To nRows + 1
        For j = 1 To nCols
            Select Case j Mod 3
                Case 0: v(i, j) = CDbl(i) * 1.5 + j
                Case 1: s = "R25TA" & Format$(i, "00000000")
                Case 2: s = ChrW(&HAC00) & ChrW(&HB098) & ChrW(&HC0C1) & ChrW(&HC0AC) & (i Mod 977)
            End Select
            If j Mod 3 <> 0 Then
                If (i + j) Mod 200 = 0 Then s = " " & s & " "
                v(i, j) = s
            End If
        Next j
    Next i
End Sub

Private Sub TrimCurrent(ByRef m As Variant)
    Dim i As Long, j As Long
    For i = LBound(m, 1) To UBound(m, 1)
        For j = LBound(m, 2) To UBound(m, 2)
            If VarType(m(i, j)) = vbString Then
                m(i, j) = Trim$(m(i, j))
            End If
        Next j
    Next i
End Sub

Private Sub TrimEdgeLR(ByRef m As Variant)
    Dim i As Long, j As Long
    For i = LBound(m, 1) To UBound(m, 1)
        For j = LBound(m, 2) To UBound(m, 2)
            If VarType(m(i, j)) = vbString Then
                If Left$(m(i, j), 1) = " " Then
                    m(i, j) = Trim$(m(i, j))
                ElseIf Right$(m(i, j), 1) = " " Then
                    m(i, j) = Trim$(m(i, j))
                End If
            End If
        Next j
    Next i
End Sub

Private Sub TrimEdgeAsc(ByRef m As Variant)
    Dim i As Long, j As Long
    For i = LBound(m, 1) To UBound(m, 1)
        For j = LBound(m, 2) To UBound(m, 2)
            If VarType(m(i, j)) = vbString Then
                If Len(m(i, j)) > 0 Then
                    If AscW(m(i, j)) = 32 Then
                        m(i, j) = Trim$(m(i, j))
                    ElseIf AscW(Right$(m(i, j), 1)) = 32 Then
                        m(i, j) = Trim$(m(i, j))
                    End If
                End If
            End If
        Next j
    Next i
End Sub

Private Function CountDiff(ByRef a As Variant, ByRef b As Variant) As Long
    Dim i As Long, j As Long
    For i = LBound(a, 1) To UBound(a, 1)
        For j = LBound(a, 2) To UBound(a, 2)
            If VarType(a(i, j)) <> VarType(b(i, j)) Then
                CountDiff = CountDiff + 1
            ElseIf a(i, j) <> b(i, j) Then
                CountDiff = CountDiff + 1
            End If
        Next j
    Next i
End Function

' @desc 27k x 34 적재의 복사와 Trim 루프 변형 비용을 기록함
Public Sub Probe_TrimPass()
    Const NROWS As Long = 27401
    Const NCOLS As Long = 34
    Dim vSrc As Variant, m As Variant, ref As Variant, r As Long, mode As Long
    Dim ms(0 To 2) As Double, t As Double, tbl As clsTable

    MakeWide NROWS, NCOLS, vSrc
    ref = vSrc
    TrimCurrent ref

    For r = -1 To 2
        TicStart
        m = vSrc
        t = TicMs
        If r >= 0 Then ms(r) = t
    Next r
    LogShim.LogEvent "probe_trim_pass", Array("rows", "cols", "path", "ms", "diff"), _
        Array(NROWS, NCOLS, "array_copy", Median3(ms(0), ms(1), ms(2)), 0)

    For mode = 0 To 2
        For r = -1 To 2
            m = vSrc
            TicStart
            Select Case mode
                Case 0: TrimCurrent m
                Case 1: TrimEdgeLR m
                Case 2: TrimEdgeAsc m
            End Select
            t = TicMs
            If r >= 0 Then ms(r) = t
        Next r
        LogShim.LogEvent "probe_trim_pass", Array("rows", "cols", "path", "ms", "diff"), _
            Array(NROWS, NCOLS, Choose(mode + 1, "trim_current", "trim_edge_leftright", "trim_edge_ascw"), _
                  Median3(ms(0), ms(1), ms(2)), CountDiff(ref, m))
    Next mode

    For r = -1 To 2
        Set tbl = New clsTable
        TicStart
        tbl.LoadFromArray vSrc, True
        t = TicMs
        If r >= 0 Then ms(r) = t
        Set tbl = Nothing
    Next r
    LogShim.LogEvent "probe_trim_pass", Array("rows", "cols", "path", "ms", "diff"), _
        Array(NROWS, NCOLS, "loadfromarray_current", Median3(ms(0), ms(1), ms(2)), 0)
End Sub

' @desc 브리프 검증 측정 3종을 순서대로 실행함
Public Sub Probe_All()
    ZPerfProbe.Probe_KeySemantics
    ZPerfProbe.Probe_TrimPass
    ZPerfProbe.Probe_IndexBuild
End Sub
