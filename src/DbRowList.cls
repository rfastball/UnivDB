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
