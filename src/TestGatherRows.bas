Attribute VB_Name = "TestGatherRows"
' @desc 내부 행 수집(GatherRows/KeepRows)과 그 소비 경로 단언 담당
' @depends UnivDB, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' @desc GatherRows의 rowId 0 건너뛰기, 열 재배치, 헤더 행, 대상 오프셋을 단언함
Public Sub Test_GatherRows_Basics()
    Dim t As clsTable, dest As Variant, ids() As Long, cols() As Long

    Set t = NewGatherTable()
    ReDim dest(1 To 4, 1 To 4)
    dest(2, 2) = "keep"

    ReDim ids(0 To 2): ids(0) = 4: ids(1) = 0: ids(2) = 2
    ReDim cols(1 To 2): cols(1) = 3: cols(2) = 1
    t.GatherRows ids, 3, cols, dest, 1, 2
    LogShim.AssertEq "gather_r0_c3", 300, dest(1, 2)
    LogShim.AssertEq "gather_r0_c1", "c", dest(1, 3)
    LogShim.AssertEq "gather_zero_untouched", "keep", dest(2, 2)
    LogShim.AssertEq "gather_zero_untouched_empty", True, IsEmpty(dest(2, 3))
    LogShim.AssertEq "gather_r2_c3", 100, dest(3, 2)
    LogShim.AssertEq "gather_r2_c1", "a", dest(3, 3)
    LogShim.AssertEq "gather_col1_untouched", True, IsEmpty(dest(1, 1))

    ReDim ids(0 To 0): ids(0) = 1
    ReDim cols(0 To 2): cols(0) = 2: cols(1) = 3: cols(2) = 1
    t.GatherRows ids, 1, cols, dest, 4, 1
    LogShim.AssertEq "gather_header_1", "NAME", dest(4, 1)
    LogShim.AssertEq "gather_header_2", "VAL", dest(4, 2)
    LogShim.AssertEq "gather_header_3", "ID", dest(4, 3)

    ReDim ids(0 To 0): ids(0) = 2
    t.GatherRows ids, 0, cols, dest, 4, 1
    LogShim.AssertEq "gather_n0_noop", "NAME", dest(4, 1)
End Sub

' @desc KeepRows의 헤더 보존, 순서, 세대 증가, 전 행 삭제를 단언함
Public Sub Test_GatherRows_KeepRows()
    Dim t As clsTable, ids() As Long, gen As Long

    Set t = NewGatherTable()
    t.GetHashIndex "ID"
    gen = t.Generation
    ReDim ids(0 To 1): ids(0) = 4: ids(1) = 2
    t.KeepRows ids, 2
    LogShim.AssertEq "keep_rows", 3, t.RowCount
    LogShim.AssertEq "keep_header", "ID", t.Value(1, "ID")
    LogShim.AssertEq "keep_first", "c", t.Value(2, "ID")
    LogShim.AssertEq "keep_second", "a", t.Value(3, "ID")
    LogShim.AssertEq "keep_gen_bumped", True, t.Generation > gen
    LogShim.AssertEq "keep_index_rebuilt", 1, t.SelectRows("ID", "a").Count
    LogShim.AssertEq "keep_index_row", 3, t.SelectRows("ID", "a")(1)

    t.KeepRows ids, 0
    LogShim.AssertEq "keep_none_rows", 1, t.RowCount
    LogShim.AssertEq "keep_none_header", "VAL", t.Value(1, "VAL")
End Sub

' @desc Query Select/Top/Delete와 AppendFrom 자기 추가가 수집 경로에서 값과 순서를 보존함을 단언함
Public Sub Test_GatherRows_Consumers()
    Dim t As clsTable, res As clsTable, n As Long

    Set t = NewGatherTable()
    Set res = UnivDB.Query(t).Where("Name", "y").SelectColumns("Val, ID").ToTable()
    LogShim.AssertEq "consumer_proj_rows", 3, res.RowCount
    LogShim.AssertEq "consumer_proj_hdr", "Val", res.Value(1, "Val")
    LogShim.AssertEq "consumer_proj_c1", 200, res.Value(2, "Val")
    LogShim.AssertEq "consumer_proj_c2", "c", res.Value(3, "ID")

    Set res = UnivDB.Query(t).OrderBy("Val", True).Top(2).ToTable()
    LogShim.AssertEq "consumer_top_rows", 3, res.RowCount
    LogShim.AssertEq "consumer_top_hdr", "ID", res.Value(1, "ID")
    LogShim.AssertEq "consumer_top_first", "c", res.Value(2, "ID")
    LogShim.AssertEq "consumer_top_second", "b", res.Value(3, "ID")

    Set res = UnivDB.Query(t).Where("ID", "zz").ToTable()
    LogShim.AssertEq "consumer_empty_rows", 1, res.RowCount
    LogShim.AssertEq "consumer_empty_hdr", "NAME", res.Value(1, "Name")

    n = UnivDB.Query(t).Where("Name", "y").Delete()
    LogShim.AssertEq "consumer_delete_count", 2, n
    LogShim.AssertEq "consumer_delete_rows", 2, t.RowCount
    LogShim.AssertEq "consumer_delete_left", "a", t.Value(2, "ID")
    LogShim.AssertEq "consumer_delete_hdr", "VAL", t.Value(1, "Val")

    t.AppendFrom t
    LogShim.AssertEq "consumer_self_append_rows", 3, t.RowCount
    LogShim.AssertEq "consumer_self_append_last", 100, t.Value(3, "Val")

    n = UnivDB.Query(t).Delete()
    LogShim.AssertEq "consumer_delete_all_count", 2, n
    LogShim.AssertEq "consumer_delete_all_rows", 1, t.RowCount
End Sub

' ID/NAME/VAL: (a,x,100) (b,y,200) (c,y,300)
Private Function NewGatherTable() As clsTable
    Dim v(1 To 4, 1 To 3) As Variant, t As clsTable
    v(1, 1) = "ID": v(1, 2) = "NAME": v(1, 3) = "VAL"
    v(2, 1) = "a": v(2, 2) = "x": v(2, 3) = 100
    v(3, 1) = "b": v(3, 2) = "y": v(3, 3) = 200
    v(4, 1) = "c": v(4, 2) = "y": v(4, 3) = 300
    Set t = New clsTable
    t.LoadFromArray v
    Set NewGatherTable = t
End Function
