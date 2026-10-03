VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Core_clsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc 대량 데이터용 대소문자 무시 집합 엔진
' @depends Core_Enums, Core_clsFastHash
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

Private mEngine As Core_clsFastHash

Private Sub Class_Initialize()
    Set mEngine = New Core_clsFastHash
End Sub

Private Sub Class_Terminate()
    Set mEngine = Nothing
End Sub

' @desc 내부 엔진을 비워 재사용 가능한 상태로 전환
Public Sub Clear()
    mEngine.Clear
End Sub

' @desc 빈 값이 아닌 단일 값을 집합에 추가
Public Sub Add(ByVal vItem As Variant)
    If Not IsEmpty(vItem) And Not IsNull(vItem) And Not IsError(vItem) Then
        mEngine.AddKey vItem
    End If
End Sub

' @desc 단일 값을 집합에서 제거
Public Sub Remove(ByVal vItem As Variant)
    If Not IsNull(vItem) And Not IsError(vItem) Then
        mEngine.RemoveKey vItem
    End If
End Sub

' @desc 단일 값의 집합 포함 여부 반환
Public Function Contains(ByVal vItem As Variant) As Boolean
    If IsNull(vItem) Or IsError(vItem) Then Exit Function
    Contains = mEngine.ExistsKey(vItem)
End Function

Public Property Get Count() As Long
    Count = mEngine.Count
End Property

' @desc 배열 또는 단일 값을 집합에 고속 적재
Public Sub LoadFromArray(ByVal vArr As Variant)
    If Not IsArray(vArr) Then
        Me.Add vArr
        Exit Sub
    End If
    mEngine.LoadKeys vArr
End Sub

' @desc Single-pass로 교집합과 양방향 차집합을 분류해 반환
Public Function Compare(ByVal otherSet As Core_clsSet, _
                        Optional ByVal mode As CoreSetRetMode = Core_Enums.srAllSplit) As Variant
    Dim setCommon As Core_clsSet, setDiffOther As Core_clsSet, setDiffMe As Core_clsSet
    Dim engCommon As Core_clsFastHash, engDiffOther As Core_clsFastHash, engDiffMe As Core_clsFastHash

    If mode = Core_Enums.srIntersect Or mode = Core_Enums.srAllSplit Then
        Set setCommon = New Core_clsSet
        Set engCommon = setCommon.InternalEngine
    End If
    If mode = Core_Enums.srDiffOther Or mode = Core_Enums.srAllSplit Then
        Set setDiffOther = New Core_clsSet
        Set engDiffOther = setDiffOther.InternalEngine
    End If
    If mode = Core_Enums.srDiffMe Or mode = Core_Enums.srAllSplit Then
        Set setDiffMe = New Core_clsSet
        Set engDiffMe = setDiffMe.InternalEngine
    End If

    ' 분류 전체를 엔진 내부 단일 패스로 위임, 결과는 zero-copy view로 바인딩됨
    mEngine.ClassifySplit otherSet.InternalEngine, engCommon, engDiffOther, engDiffMe

    Select Case mode
        Case Core_Enums.srIntersect: Set Compare = setCommon
        Case Core_Enums.srDiffOther: Set Compare = setDiffOther
        Case Core_Enums.srDiffMe:    Set Compare = setDiffMe
        Case Core_Enums.srAllSplit:  Compare = Array(setCommon, setDiffMe, setDiffOther)
    End Select
End Function

' @desc 교집합을 새 집합으로 반환
Public Function Intersection(ByVal otherSet As Core_clsSet) As Core_clsSet
    Set Intersection = Compare(otherSet, Core_Enums.srIntersect)
End Function

' @desc 내 차집합을 새 집합으로 반환
Public Function Difference(ByVal otherSet As Core_clsSet) As Core_clsSet
    Set Difference = Compare(otherSet, Core_Enums.srDiffMe)
End Function

' @desc 상대 차집합을 새 집합으로 반환
Public Function DifferenceInput(ByVal otherSet As Core_clsSet) As Core_clsSet
    Set DifferenceInput = Compare(otherSet, Core_Enums.srDiffOther)
End Function

' @desc 합집합을 새 집합으로 반환
Public Function Union(ByVal otherSet As Core_clsSet) As Core_clsSet
    Dim newSet As New Core_clsSet
    Dim targetEng As Core_clsFastHash

    Set targetEng = newSet.InternalEngine
    targetEng.EnsureCapacity Me.Count + otherSet.Count
    targetEng.MergeFrom mEngine
    targetEng.MergeFrom otherSet.InternalEngine
    Set Union = newSet
End Function

' @desc 집합 키를 1차원 배열로 반환
Public Function ToArray() As Variant
    ToArray = mEngine.KeysArray
End Function

Public Property Get InternalEngine() As Core_clsFastHash
    Set InternalEngine = mEngine
End Property

Public Property Get KeysArray() As Variant
    KeysArray = mEngine.KeysArray
End Property
