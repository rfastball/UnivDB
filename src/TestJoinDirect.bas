Attribute VB_Name = "TestJoinDirect"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

Public Sub Test_JoinDirect_AllJoinTypes()
    Dim tLeft As clsTable
    Dim tRight As clsTable
    Dim tResult As clsTable

    Set tLeft = NewJoinDirectLeft()
    Set tRight = NewJoinDirectRight()

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtInner)
    LogShim.AssertEq "join_direct_inner_rows", 5, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtLeft)
    LogShim.AssertEq "join_direct_left_rows", 7, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtRight)
    LogShim.AssertEq "join_direct_right_rows", 6, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtFull)
    LogShim.AssertEq "join_direct_full_rows", 8, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtLeftAnti)
    LogShim.AssertEq "join_direct_left_anti_rows", 3, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtRightAnti)
    LogShim.AssertEq "join_direct_right_anti_rows", 2, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtLeftSemi)
    LogShim.AssertEq "join_direct_left_semi_rows", 3, tResult.RowCount

    Set tResult = UnivDB.Join(tLeft, tRight, "Dept", jtRightSemi)
    LogShim.AssertEq "join_direct_right_semi_rows", 3, tResult.RowCount
End Sub

Public Sub Test_JoinDirect_DifferentKeyColumnNames()
    Dim tLeft As clsTable
    Dim tRight As clsTable
    Dim tResult As clsTable

    Set tLeft = NewJoinDirectLeft()
    Set tRight = NewJoinDirectRightDifferentKey()

    Set tResult = UnivDB.Join(tLeft, tRight, leftKey:="Dept", rightKey:="DeptKey", jType:=jtInner)
    LogShim.AssertEq "join_direct_different_keys_rows", 5, tResult.RowCount
End Sub

Private Function NewJoinDirectLeft() As clsTable
    Dim vData(1 To 5, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "ID": vData(1, 2) = "Dept"
    vData(2, 1) = 1: vData(2, 2) = "A"
    vData(3, 1) = 2: vData(3, 2) = "A"
    vData(4, 1) = 3: vData(4, 2) = "B"
    vData(5, 1) = 4: vData(5, 2) = "X"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewJoinDirectLeft = T
End Function

Private Function NewJoinDirectRight() As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "Dept": vData(1, 2) = "Label"
    vData(2, 1) = "A": vData(2, 2) = "Alpha1"
    vData(3, 1) = "A": vData(3, 2) = "Alpha2"
    vData(4, 1) = "C": vData(4, 2) = "Gamma"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewJoinDirectRight = T
End Function

Private Function NewJoinDirectRightDifferentKey() As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant
    Dim T As clsTable

    vData(1, 1) = "DeptKey": vData(1, 2) = "Label"
    vData(2, 1) = "A": vData(2, 2) = "Alpha1"
    vData(3, 1) = "A": vData(3, 2) = "Alpha2"
    vData(4, 1) = "C": vData(4, 2) = "Gamma"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData
    Set NewJoinDirectRightDifferentKey = T
End Function
