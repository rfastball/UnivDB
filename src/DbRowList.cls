VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mRows() As Long
Private mCount As Long

' 세대 계약 결속 (ObjPtr 스냅샷 - 객체 참조 아님, 순환 참조 방지)
Private mOwnerPtr As LongPtr
Private mGeneration As Long

' @desc 테이블을 받아 현재 세대에 결속함
Public Sub BindTo(ByVal tbl As clsTable)
    mOwnerPtr = ObjPtr(tbl)
    mGeneration = tbl.Generation
End Sub

Friend Sub BindRaw(ByVal ownerPtr As LongPtr, ByVal gen As Long)
    mOwnerPtr = ownerPtr
    mGeneration = gen
End Sub

' @desc 테이블을 받아 결속 세대와 일치 검증함, stale 시 9998 에러
Public Sub EnsureCurrent(ByVal tbl As clsTable)
    modDbContract.AssertCurrentGeneration mOwnerPtr, mGeneration, tbl, "DbRowList.EnsureCurrent"
End Sub

Friend Property Get OwnerPtr() As LongPtr
    OwnerPtr = mOwnerPtr
End Property

Friend Property Get BoundGeneration() As Long
    BoundGeneration = mGeneration
End Property

Public Sub Append(ByVal rowId As Long)
    If mCount = 0 Then
        ReDim mRows(1 To 8)
    ElseIf mCount = UBound(mRows) Then
        ReDim Preserve mRows(1 To UBound(mRows) * 2)
    End If
    
    mCount = mCount + 1
    mRows(mCount) = rowId
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function RowAt(ByVal index As Long) As Long
    If index < 1 Or index > mCount Then
        Err.Raise 9, "DbRowList.RowAt", "Row index out of range"
    End If
    
    RowAt = mRows(index)
End Function

Public Function ToArray() As Variant
    Dim result() As Long
    Dim i As Long
    
    If mCount = 0 Then
        ToArray = Array()
        Exit Function
    End If
    
    ReDim result(1 To mCount)
    For i = 1 To mCount
        result(i) = mRows(i)
    Next i
    
    ToArray = result
End Function

Public Sub Clear()
    Erase mRows
    mCount = 0
End Sub

' 공유 Long 배열의 (offset, count) 구간을 정확 크기로 적재한다 (DbCsrPayload 복사 경로)
Friend Sub LoadFromLongSpan(ByRef rowIds() As Long, ByVal offset As Long, ByVal count As Long)
    Dim i As Long

    Clear
    If count <= 0 Then Exit Sub

    ReDim mRows(1 To count)
    For i = 1 To count
        mRows(i) = rowIds(offset + i - 1)
    Next i
    mCount = count
End Sub

Public Sub LoadFromCollection(ByVal rows As Collection)
    Dim rowId As Variant
    
    Clear
    If rows Is Nothing Then Exit Sub
    
    For Each rowId In rows
        Append CLng(rowId)
    Next rowId
End Sub

Public Function ToCollection() As Collection
    Dim rows As New Collection
    Dim i As Long
    
    For i = 1 To mCount
        rows.Add mRows(i)
    Next i
    
    Set ToCollection = rows
End Function
