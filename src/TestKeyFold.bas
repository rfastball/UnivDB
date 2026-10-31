Attribute VB_Name = "TestKeyFold"
' @desc Dict 백엔드 하이브리드 키 접기 의미론과 원본 표기 보존 단언 담당
' @depends UnivDB, modIndexKey, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' 2026-10-31 계획 단계 6 / 검증 문서 2.2절 표. Dict 백엔드(기본)만 - Core는 TestHashIndex 담당.

' @desc 대소문자/전각/반각 자모/반각 가나 동치와 e/e-acute, sharp-s/ss, 가나 종류 구분을 인덱스와 조인에서 단언함
Public Sub Test_KeyFold_DictEquivalence()
    Dim same As Variant, diff As Variant, pair As Variant, n As Long

    same = Array( _
        Array("abc", "ABC"), _
        Array(ChrW(&HE9), ChrW(&HC9)), _
        Array(ChrW(&HFF21) & ChrW(&HFF22) & ChrW(&HFF23), "ABC"), _
        Array(ChrW(&HFF11) & ChrW(&HFF12), "12"), _
        Array(ChrW(&HFF41) & ChrW(&HFF42), "AB"), _
        Array("a" & ChrW(&H3000) & "b", "a b"), _
        Array(ChrW(&H3131), ChrW(&HFFA1)), _
        Array(ChrW(&H30AC), ChrW(&HFF76) & ChrW(&HFF9E)), _
        Array("x" & ChrW(&H30AC) & "y", "X" & ChrW(&HFF76) & ChrW(&HFF9E) & "Y"))
    diff = Array( _
        Array("e", ChrW(&HE9)), _
        Array(ChrW(&HDF), "ss"), _
        Array(ChrW(&H30AB), ChrW(&H304B)), _
        Array("a", "b"), _
        Array("1", "1.0"))

    For Each pair In same
        n = n + 1
        AssertPair "fold_same_" & n, pair(0), pair(1), True
    Next pair
    n = 0
    For Each pair In diff
        n = n + 1
        AssertPair "fold_diff_" & n, pair(0), pair(1), False
    Next pair
End Sub

' @desc 커서 키/GroupBy/GetUniqueKeys 표기가 첫 등장 원본(접기 전)으로 보존됨을 단언함
Public Sub Test_KeyFold_FirstSpellingPreserved()
    Dim v(1 To 5, 1 To 3) As Variant, T As clsTable, res As clsTable, keys As Variant
    v(1, 1) = "K": v(1, 2) = "K2": v(1, 3) = "N"
    v(2, 1) = "abc": v(2, 2) = ChrW(&H30AC): v(2, 3) = 1
    v(3, 1) = "ABC": v(3, 2) = ChrW(&HFF76) & ChrW(&HFF9E): v(3, 3) = 2
    v(4, 1) = ChrW(&HFF41) & "bc": v(4, 2) = ChrW(&H30AC): v(4, 3) = 4
    v(5, 1) = "zz": v(5, 2) = "q": v(5, 3) = 8
    Set T = New clsTable
    T.LoadFromArray v

    keys = T.GetUniqueKeys("K")
    LogShim.AssertEq "fold_unique_count", 2, UBound(keys) + 1
    LogShim.AssertEq "fold_unique_first", "abc", keys(0)

    Set res = T.GroupBy("K", "N:SUM")
    LogShim.AssertEq "fold_group_rows", 3, res.RowCount
    LogShim.AssertEq "fold_group_key", "abc", res.Value(2, "K")
    LogShim.AssertEq "fold_group_sum", 7, res.Value(2, "N_SUM")

    Set res = T.GroupBy("K,K2", "N:SUM")
    LogShim.AssertEq "fold_group2_rows", 3, res.RowCount
    LogShim.AssertEq "fold_group2_k", "abc", res.Value(2, "K")
    LogShim.AssertEq "fold_group2_k2", ChrW(&H30AC), res.Value(2, "K2")
    LogShim.AssertEq "fold_group2_sum", 7, res.Value(2, "N_SUM")

    LogShim.AssertEq "fold_where", 3, UnivDB.Query(T).Where("K", "Abc").ToTable().RowCount - 1
End Sub

' @desc 폭/가나 영역 판별 경계와 빈 문자열, LCMapStringW 경로의 대문자화를 단언함
Public Sub Test_KeyFold_FoldIndexKeyDirect()
    Dim s As String
    s = "": LogShim.AssertEq "fold_direct_empty", "", modIndexKey.FoldIndexKey(s)
    s = "aBc1": LogShim.AssertEq "fold_direct_ascii", "ABC1", modIndexKey.FoldIndexKey(s)
    s = ChrW(&HAC00) & "a": LogShim.AssertEq "fold_direct_hangul", ChrW(&HAC00) & "A", modIndexKey.FoldIndexKey(s)
    s = ChrW(&HFF41) & "a": LogShim.AssertEq "fold_direct_fullwidth", "AA", modIndexKey.FoldIndexKey(s)
    s = ChrW(&H3130) & "a": LogShim.AssertEq "fold_direct_below_jamo", ChrW(&H3130) & "A", modIndexKey.FoldIndexKey(s)
    s = ChrW(&H318F) & "a": LogShim.AssertEq "fold_direct_above_jamo", ChrW(&H318F) & "A", modIndexKey.FoldIndexKey(s)
    s = ChrW(&H2018) & "a": LogShim.AssertEq "fold_direct_curly_quote_kept", ChrW(&H2018) & "A", modIndexKey.FoldIndexKey(s)
    s = String$(300, ChrW(&H30AC))
    LogShim.AssertEq "fold_direct_long_len", 600, Len(modIndexKey.FoldIndexKey(s))
End Sub

' 두 값을 같은 열에 두면 키 1개/2개, 한쪽만 둔 테이블에서 다른 쪽 조회/조인 일치 여부를 단언한다
Private Sub AssertPair(ByVal label As String, ByVal a As String, ByVal b As String, ByVal expectSame As Boolean)
    Dim v(1 To 3, 1 To 1) As Variant, va(1 To 2, 1 To 1) As Variant, vb(1 To 2, 1 To 1) As Variant
    Dim T As clsTable, tA As clsTable, tB As clsTable, cnt As Long

    v(1, 1) = "K": v(2, 1) = a: v(3, 1) = b
    Set T = New clsTable
    T.LoadFromArray v
    LogShim.AssertEq label & "_keys", IIf(expectSame, 1, 2), T.GetHashIndex("K").CountKeys

    va(1, 1) = "K": va(2, 1) = a
    vb(1, 1) = "K": vb(2, 1) = b
    Set tA = New clsTable: tA.LoadFromArray va
    Set tB = New clsTable: tB.LoadFromArray vb
    LogShim.AssertEq label & "_lookup", expectSame, tA.GetHashIndex("K").TryGetRowCount(b, cnt)
    LogShim.AssertEq label & "_join", IIf(expectSame, 2, 1), modEngine.JoinTable(tA, tB, "K", "K", jtInner).RowCount
End Sub
