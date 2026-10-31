Attribute VB_Name = "TestLoadTrim"
' @desc 적재 Trim 계약과 bTrim:=False 옵션의 공백 경계 단언 담당
' @depends UnivDB, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' @desc 기본 적재는 문자열 셀을 Trim하고 bTrim:=False는 셀 공백을 유지하되 헤더 Trim은 유지함을 단언함
Public Sub Test_LoadTrim_Option()
    Dim T As clsTable

    Set T = New clsTable
    T.LoadFromArray TrimData()
    LogShim.AssertEq "trim_default_cell", "a", T.Value(2, "K")
    LogShim.AssertEq "trim_default_number", 5, T.Value(2, "N")

    Set T = New clsTable
    T.LoadFromArray TrimData(), True, False
    LogShim.AssertEq "trim_off_cell", " a ", T.Value(2, "K")
    LogShim.AssertEq "trim_off_header_trimmed", "K", T.Value(1, "K")
    LogShim.AssertEq "trim_off_has_col", True, T.HasColumn("K")

    UnivDB.InitDB
    UnivDB.RegisterTable "TrimOff_T", TrimData(), False
    LogShim.AssertEq "trim_off_register", " a ", UnivDB.GetTable("TrimOff_T").Value(2, "K")
    UnivDB.RegisterTable "TrimOn_T", TrimData()
    LogShim.AssertEq "trim_on_register", "a", UnivDB.GetTable("TrimOn_T").Value(2, "K")
    UnivDB.ClearDB
End Sub

' @desc bTrim:=False에서 Filter "IS"는 셀을 Trim해 일치하고 Where/Join은 원문 키라 불일치함을 단언함
Public Sub Test_LoadTrim_OffSurfacesDiverge()
    Dim T As clsTable, tR As clsTable, v(1 To 2, 1 To 1) As Variant

    Set T = New clsTable
    T.LoadFromArray TrimData(), True, False
    LogShim.AssertEq "trim_off_filter_is", 1, UnivDB.Query(T).Filter("K", "IS", "a").ToTable().RowCount - 1
    LogShim.AssertEq "trim_off_where", 0, UnivDB.Query(T).Where("K", "a").ToTable().RowCount - 1
    LogShim.AssertEq "trim_off_where_raw", 1, UnivDB.Query(T).Where("K", " a ").ToTable().RowCount - 1
    v(1, 1) = "K": v(2, 1) = "a"
    Set tR = New clsTable
    tR.LoadFromArray v
    LogShim.AssertEq "trim_off_join", 1, modEngine.JoinTable(T, tR, "K", "K", jtInner).RowCount
End Sub

' @desc bTrim과 무관하게 중복 헤더를 9999로 거부하고 기존 상태를 보존함을 단언함
Public Sub Test_LoadTrim_DuplicateHeaderStillRejected()
    Dim T As clsTable, v(1 To 2, 1 To 2) As Variant, n As Long
    v(1, 1) = "K": v(1, 2) = " k ": v(2, 1) = 1: v(2, 2) = 2
    Set T = New clsTable
    T.LoadFromArray TrimData()
    On Error Resume Next
    T.LoadFromArray v, True, False
    n = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "trim_off_dup_errno", 9999, n
    LogShim.AssertEq "trim_off_dup_state_kept", "a", T.Value(2, "K")
End Sub

Private Function TrimData() As Variant
    Dim v(1 To 2, 1 To 2) As Variant
    v(1, 1) = " K ": v(1, 2) = "N"
    v(2, 1) = " a ": v(2, 2) = 5
    TrimData = v
End Function
