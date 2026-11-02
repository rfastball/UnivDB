VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Core_clsLongList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc 동적 확장 Long 리스트 프리미티브 제공 담당
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' Core 공개 프리미티브 (설계 결정 3항: 저수준 컨테이너는 Core 책임).
' 최소 표면: Append / Count / ItemAt / ToArray / Clear / EnsureCapacity.
' 1 기반 ItemAt, ToArray는 0 기반 Long 배열 (빈 리스트면 Array()).

Private Const INITIAL_CAPACITY As Long = 16

Private mItems() As Long
Private mCount As Long
Private mCapacity As Long

Private Sub Class_Initialize()
    Clear
End Sub

' @desc 모든 항목을 비워 초기 크기로 재설정함
Public Sub Clear()
    mCapacity = INITIAL_CAPACITY
    ReDim mItems(1 To mCapacity)
    mCount = 0
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

' @desc 값 받아 리스트 끝에 추가함
Public Sub Append(ByVal v As Long)
    If mCount = mCapacity Then
        mCapacity = mCapacity * 2
        ReDim Preserve mItems(1 To mCapacity)
    End If
    mCount = mCount + 1
    mItems(mCount) = v
End Sub

' @desc 1 기반 인덱스 받아 값 반환, 범위 밖 시 9 에러
Public Function ItemAt(ByVal idx As Long) As Long
    If idx < 1 Or idx > mCount Then
        Err.Raise 9, "Core_clsLongList.ItemAt", "Index out of range"
    End If
    ItemAt = mItems(idx)
End Function

' @desc 값들을 0 기반 Long 배열로 반환, 빈 리스트 시 Array() 반환
Public Function ToArray() As Variant
    Dim arr() As Long
    Dim i As Long

    If mCount = 0 Then
        ToArray = Array()
        Exit Function
    End If
    ReDim arr(0 To mCount - 1)
    For i = 1 To mCount
        arr(i - 1) = mItems(i)
    Next i
    ToArray = arr
End Function

' @desc 예상 항목 수 받아 저장소를 선확장함
Public Sub EnsureCapacity(ByVal N As Long)
    Dim newCap As Long

    If N <= mCapacity Then Exit Sub
    newCap = mCapacity
    Do While newCap < N
        newCap = newCap * 2
    Loop
    mCapacity = newCap
    ReDim Preserve mItems(1 To mCapacity)
End Sub
