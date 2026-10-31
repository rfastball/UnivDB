Attribute VB_Name = "TestJoinOrder"
' @desc 조인 유형별 결과 행 순서 고정 회귀 단언 담당
' @depends modEngine, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' 2026-10-31 성능 계획 단계 0: Join 결과 행 순서를 현행 코드 기준으로 고정한다.
' 순서 계약 = 공통 키(오른쪽 커서 순, 키 내 왼쪽 x 오른쪽 중첩) -> 왼쪽 전용(왼쪽 커서 순)
' -> 오른쪽 전용(오른쪽 커서 순). 1:N, N:N, 대소문자 변이 키, 양쪽 전용 키를 섞는다.

Private Const HDR As String = "K|LV|K_R|RV"

' @desc inner/left/right/full 조인 결과 전체 행 순서를 양 백엔드에서 단언함
Public Sub Test_JoinOrder_OuterFamily()
    Dim tL As clsTable, tR As clsTable, b As Long, sfx As String
    Dim common As String

    common = "b|L2|b|R1;b|L2|B|R4;b|L2|b|R7;b|L6|b|R1;b|L6|B|R4;b|L6|b|R7;" & _
             "a|L1|a|R2;a|L1|a|R5;a|L3|a|R2;a|L3|a|R5"

    Set tL = NewOrderLeft(): Set tR = NewOrderRight()
    For b = dbBackendDict To dbBackendCore
        tL.IndexBackend = b: tR.IndexBackend = b
        sfx = "_" & b
        LogShim.AssertEq "join_order_inner" & sfx, HDR & ";" & common, _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtInner))
        LogShim.AssertEq "join_order_left" & sfx, HDR & ";" & common & ";c|L4|~|~;x|L5|~|~;y|L7|~|~", _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtLeft))
        LogShim.AssertEq "join_order_right" & sfx, HDR & ";" & common & ";~|~|z|R3;~|~|w|R6", _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtRight))
        LogShim.AssertEq "join_order_full" & sfx, _
            HDR & ";" & common & ";c|L4|~|~;x|L5|~|~;y|L7|~|~;~|~|z|R3;~|~|w|R6", _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtFull))
    Next b
End Sub

' @desc semi/anti 조인 결과 전체 행 순서와 접두 헤더를 양 백엔드에서 단언함
Public Sub Test_JoinOrder_SemiAnti()
    Dim tL As clsTable, tR As clsTable, b As Long, sfx As String

    Set tL = NewOrderLeft(): Set tR = NewOrderRight()
    For b = dbBackendDict To dbBackendCore
        tL.IndexBackend = b: tR.IndexBackend = b
        sfx = "_" & b
        LogShim.AssertEq "join_order_left_semi" & sfx, "K|LV;b|L2;b|L6;a|L1;a|L3", _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtLeftSemi))
        LogShim.AssertEq "join_order_right_semi" & sfx, "R_K|R_RV;b|R1;B|R4;b|R7;a|R2;a|R5", _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtRightSemi, "R_"))
        LogShim.AssertEq "join_order_left_anti" & sfx, "K|LV;c|L4;x|L5;y|L7", _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtLeftAnti))
        LogShim.AssertEq "join_order_right_anti" & sfx, "K|RV;z|R3;w|R6", _
            RowsSig(modEngine.JoinTable(tL, tR, "K", "K", jtRightAnti))
    Next b
End Sub

' @desc 일치 키가 없는 조인과 빈 테이블 조인의 결과 형태를 단언함
Public Sub Test_JoinOrder_EmptySides()
    Dim tL As clsTable, tR As clsTable, tE As clsTable
    Dim v(1 To 1, 1 To 2) As Variant

    Set tL = NewOrderLeft(): Set tR = NewOrderRight()
    v(1, 1) = "K": v(1, 2) = "EV"
    Set tE = New clsTable
    tE.LoadFromArray v

    LogShim.AssertEq "join_order_empty_inner", "K|LV|K_R|EV", RowsSig(modEngine.JoinTable(tL, tE, "K", "K", jtInner))
    LogShim.AssertEq "join_order_empty_left", _
        "K|LV|K_R|EV;a|L1|~|~;a|L3|~|~;b|L2|~|~;b|L6|~|~;c|L4|~|~;x|L5|~|~;y|L7|~|~", _
        RowsSig(modEngine.JoinTable(tL, tE, "K", "K", jtLeft))
    LogShim.AssertEq "join_order_empty_right", _
        "K|EV|K_R|RV;~|~|b|R1;~|~|B|R4;~|~|b|R7;~|~|a|R2;~|~|a|R5;~|~|z|R3;~|~|w|R6", _
        RowsSig(modEngine.JoinTable(tE, tR, "K", "K", jtRight))
    LogShim.AssertEq "join_order_empty_both", "K|EV|K_R|EV_R", RowsSig(modEngine.JoinTable(tE, tE, "K", "K", jtFull))
    LogShim.AssertEq "join_order_empty_left_anti", "K|LV;a|L1;a|L3;b|L2;b|L6;c|L4;x|L5;y|L7", _
        RowsSig(modEngine.JoinTable(tL, tE, "K", "K", jtLeftAnti))
End Sub

' 왼쪽: a,b 중복(1:N, N:N 구성), c/x/y 왼쪽 전용
Private Function NewOrderLeft() As clsTable
    Dim v(1 To 8, 1 To 2) As Variant, t As clsTable
    v(1, 1) = "K": v(1, 2) = "LV"
    v(2, 1) = "a": v(2, 2) = "L1"
    v(3, 1) = "b": v(3, 2) = "L2"
    v(4, 1) = "a": v(4, 2) = "L3"
    v(5, 1) = "c": v(5, 2) = "L4"
    v(6, 1) = "x": v(6, 2) = "L5"
    v(7, 1) = "b": v(7, 2) = "L6"
    v(8, 1) = "y": v(8, 2) = "L7"
    Set t = New clsTable
    t.LoadFromArray v
    Set NewOrderLeft = t
End Function

' 오른쪽: b 3행(대소문자 변이 B 포함), a 2행, z/w 오른쪽 전용
Private Function NewOrderRight() As clsTable
    Dim v(1 To 8, 1 To 2) As Variant, t As clsTable
    v(1, 1) = "K": v(1, 2) = "RV"
    v(2, 1) = "b": v(2, 2) = "R1"
    v(3, 1) = "a": v(3, 2) = "R2"
    v(4, 1) = "z": v(4, 2) = "R3"
    v(5, 1) = "B": v(5, 2) = "R4"
    v(6, 1) = "a": v(6, 2) = "R5"
    v(7, 1) = "w": v(7, 2) = "R6"
    v(8, 1) = "b": v(8, 2) = "R7"
    Set t = New clsTable
    t.LoadFromArray v
    Set NewOrderRight = t
End Function

' 헤더 포함 전체 행을 "셀|셀;셀|셀" 서명으로 (Empty = ~)
Private Function RowsSig(ByVal t As clsTable) As String
    Dim v As Variant, r As Long, c As Long, s As String
    v = t.RawArray
    For r = 1 To t.RowCount
        If r > 1 Then s = s & ";"
        For c = 1 To t.ColCount
            If c > 1 Then s = s & "|"
            If IsEmpty(v(r, c)) Then s = s & "~" Else s = s & CStr(v(r, c))
        Next c
    Next r
    RowsSig = s
End Function
