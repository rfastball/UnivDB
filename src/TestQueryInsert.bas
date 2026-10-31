Attribute VB_Name = "TestQueryInsert"
' @desc Query.Insert 1행 추가 경로의 값, Trim, fail-fast, 세대 단언 담당
' @depends UnivDB, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' @desc 새 행 매핑/Trim, 미지정 열 Empty, 인덱스 재빌드, 세대 증가를 단언함
Public Sub Test_QueryInsert_AppendsTrimmedRow()
    Dim T As clsTable, gen As Long, n As Long

    Set T = NewInsertTable()
    T.GetHashIndex "ID"
    gen = T.Generation
    n = UnivDB.Query(T).SetVal(" name ", "  Zed  ").SetVal("ID", 9).Insert()
    LogShim.AssertEq "insert_count", 1, n
    LogShim.AssertEq "insert_rows", 4, T.RowCount
    LogShim.AssertEq "insert_trimmed", "Zed", T.Value(4, "Name")
    LogShim.AssertEq "insert_number", 9, T.Value(4, "ID")
    LogShim.AssertEq "insert_unset_empty", True, IsEmpty(T.Value(4, "Score"))
    LogShim.AssertEq "insert_gen", True, T.Generation > gen
    LogShim.AssertEq "insert_index", 4, T.SelectRows("ID", 9)(1)
    LogShim.AssertEq "insert_header_kept", "Name", T.Value(1, "Name")
    LogShim.AssertEq "insert_old_row", "b", T.Value(3, "Name")
End Sub

' @desc 미존재 열 Insert가 테이블을 바꾸지 않고 9999로 실패함을 단언함
Public Sub Test_QueryInsert_UnknownColumnFailsFast()
    Dim T As clsTable, gen As Long, n As Long

    Set T = NewInsertTable()
    gen = T.Generation
    On Error Resume Next
    UnivDB.Query(T).SetVal("ID", 5).SetVal("Nope", 1).Insert
    n = Err.Number
    On Error GoTo 0
    LogShim.AssertEq "insert_unknown_errno", 9999, n
    LogShim.AssertEq "insert_unknown_rows", 3, T.RowCount
    LogShim.AssertEq "insert_unknown_gen", gen, T.Generation
End Sub

' ID/Name/Score: (1,a,10) (2,b,20)
Private Function NewInsertTable() As clsTable
    Dim v(1 To 3, 1 To 3) As Variant, T As clsTable
    v(1, 1) = "ID": v(1, 2) = "Name": v(1, 3) = "Score"
    v(2, 1) = 1: v(2, 2) = "a": v(2, 3) = 10
    v(3, 1) = 2: v(3, 2) = "b": v(3, 3) = 20
    Set T = New clsTable
    T.LoadFromArray v
    Set NewInsertTable = T
End Function
