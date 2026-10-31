Attribute VB_Name = "TestUpsertFrom"
' @desc UpsertFrom 열 매핑, 혼합 갱신/삽입, 자기 병합, 중복 오류 단언 담당
' @depends LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' @desc 소스 열 순서가 다르고 미매핑 열이 섞인 갱신/삽입 혼합 병합 결과를 단언함
Public Sub Test_UpsertFrom_MixedMapping()
    Dim T As clsTable, src As clsTable, gen As Long

    Set T = NewUpsertTarget()
    Set src = NewUpsertSource()
    gen = T.Generation
    T.UpsertFrom src, "ID"
    LogShim.AssertEq "upsert_mix_rows", 4, T.RowCount
    LogShim.AssertEq "upsert_mix_gen", True, T.Generation > gen
    LogShim.AssertEq "upsert_mix_r2_a", "na1", T.Value(2, "A")
    LogShim.AssertEq "upsert_mix_r2_b", "nb1", T.Value(2, "B")
    LogShim.AssertEq "upsert_mix_r2_id_written", "1", T.Value(2, "ID")
    LogShim.AssertEq "upsert_mix_r2_extra", "x1", T.Value(2, "Extra")
    LogShim.AssertEq "upsert_mix_r3_a", "na2", T.Value(3, "A")
    LogShim.AssertEq "upsert_mix_r3_extra", "x2", T.Value(3, "Extra")
    LogShim.AssertEq "upsert_mix_r4_id", 3, T.Value(4, "ID")
    LogShim.AssertEq "upsert_mix_r4_b", "nb3", T.Value(4, "B")
    LogShim.AssertEq "upsert_mix_r4_extra_empty", True, IsEmpty(T.Value(4, "Extra"))
    LogShim.AssertEq "upsert_mix_index", 4, T.SelectRows("ID", 3)(1)
    LogShim.AssertEq "upsert_mix_src_untouched", "nb2", src.Value(2, "B")
End Sub

' @desc 자기 자신과의 병합이 값과 행 수를 보존함을 단언함
Public Sub Test_UpsertFrom_Self()
    Dim T As clsTable
    Set T = NewUpsertTarget()
    T.UpsertFrom T, "ID"
    LogShim.AssertEq "upsert_self_rows", 3, T.RowCount
    LogShim.AssertEq "upsert_self_a", "a2", T.Value(3, "A")
    LogShim.AssertEq "upsert_self_extra", "x1", T.Value(2, "Extra")
End Sub

' @desc 중복 소스 키 오류 메시지의 키와 두 소스 행 번호를 단언함
Public Sub Test_UpsertFrom_DuplicateMessage()
    Dim T As clsTable, src As clsTable, v(1 To 5, 1 To 2) As Variant, desc As String, n As Long
    v(1, 1) = "ID": v(1, 2) = "A"
    v(2, 1) = "new": v(3, 1) = 2: v(4, 1) = "NEW": v(5, 1) = "2"
    Set src = New clsTable
    src.LoadFromArray v
    Set T = NewUpsertTarget()
    On Error Resume Next
    T.UpsertFrom src, "ID"
    n = Err.Number: desc = Err.Description
    On Error GoTo 0
    LogShim.AssertEq "upsert_dup_new_errno", 9999, n
    LogShim.AssertEq "upsert_dup_new_msg", "Duplicate source key [NEW] in column [ID] at rows 2 and 4", desc
    LogShim.AssertEq "upsert_dup_new_untouched", "a2", T.Value(3, "A")

    v(4, 1) = "other"
    src.LoadFromArray v
    On Error Resume Next
    T.UpsertFrom src, "ID"
    n = Err.Number: desc = Err.Description
    On Error GoTo 0
    LogShim.AssertEq "upsert_dup_existing_errno", 9999, n
    LogShim.AssertEq "upsert_dup_existing_msg", "Duplicate source key [2] in column [ID] at rows 3 and 5", desc
    LogShim.AssertEq "upsert_dup_existing_rows", 3, T.RowCount
End Sub

' ID/A/B/Extra: (1,a1,b1,x1) (2,a2,b2,x2)
Private Function NewUpsertTarget() As clsTable
    Dim v(1 To 3, 1 To 4) As Variant, T As clsTable
    v(1, 1) = "ID": v(1, 2) = "A": v(1, 3) = "B": v(1, 4) = "Extra"
    v(2, 1) = 1: v(2, 2) = "a1": v(2, 3) = "b1": v(2, 4) = "x1"
    v(3, 1) = 2: v(3, 2) = "a2": v(3, 3) = "b2": v(3, 4) = "x2"
    Set T = New clsTable
    T.LoadFromArray v
    Set NewUpsertTarget = T
End Function

' B/ID/A/Ghost (열 순서 다름, Ghost는 대상에 없음): 2 갱신, 3 삽입, "1" 갱신(정규화 키 일치)
Private Function NewUpsertSource() As clsTable
    Dim v(1 To 4, 1 To 4) As Variant, T As clsTable
    v(1, 1) = "B": v(1, 2) = "id": v(1, 3) = "A": v(1, 4) = "Ghost"
    v(2, 1) = "nb2": v(2, 2) = 2: v(2, 3) = "na2": v(2, 4) = "g2"
    v(3, 1) = "nb3": v(3, 2) = 3: v(3, 3) = "na3": v(3, 4) = "g3"
    v(4, 1) = "nb1": v(4, 2) = "1": v(4, 3) = "na1": v(4, 4) = "g1"
    Set T = New clsTable
    T.LoadFromArray v
    Set NewUpsertSource = T
End Function
