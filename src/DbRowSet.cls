VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRowSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mFlags() As Boolean
Private mMaxRow As Long
Private mCount As Long

Public Sub Clear()
    Erase mFlags
    mMaxRow = 0
    mCount = 0
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Sub FromAllRows(ByVal tableRowCount As Long, Optional ByVal firstDataRow As Long = 2)
    Dim rowId As Long
    
    Clear
    If tableRowCount < firstDataRow Then Exit Sub
    
    EnsureCapacity tableRowCount
    For rowId = firstDataRow To tableRowCount
        AddRow rowId
    Next rowId
End Sub

Public Sub LoadFromRowList(ByVal rows As DbRowList)
    Dim i As Long
    
    Clear
    If rows Is Nothing Then Exit Sub
    
    For i = 1 To rows.Count
        AddRow rows.RowAt(i)
    Next i
End Sub

Public Sub AddRow(ByVal rowId As Long)
    If rowId < 1 Then Exit Sub
    EnsureCapacity rowId
    
    If Not mFlags(rowId) Then
        mFlags(rowId) = True
        mCount = mCount + 1
    End If
End Sub

Public Function ContainsRow(ByVal rowId As Long) As Boolean
    If rowId < 1 Or rowId > mMaxRow Then Exit Function
    ContainsRow = mFlags(rowId)
End Function

Public Sub IntersectWithRowList(ByVal rows As DbRowList)
    Dim keep() As Boolean
    Dim i As Long
    Dim rowId As Long
    Dim newCount As Long
    
    If mCount = 0 Then Exit Sub
    If rows Is Nothing Or rows.Count = 0 Then
        Clear
        Exit Sub
    End If
    
    ReDim keep(1 To mMaxRow)
    For i = 1 To rows.Count
        rowId = rows.RowAt(i)
        If rowId >= 1 And rowId <= mMaxRow Then
            If mFlags(rowId) And Not keep(rowId) Then
                keep(rowId) = True
                newCount = newCount + 1
            End If
        End If
    Next i
    
    mFlags = keep
    mCount = newCount
End Sub

Public Function ToArray() As Variant
    Dim result() As Long
    Dim rowId As Long
    Dim pos As Long
    
    If mCount = 0 Then
        ToArray = Array()
        Exit Function
    End If
    
    ReDim result(0 To mCount - 1)
    For rowId = 1 To mMaxRow
        If mFlags(rowId) Then
            result(pos) = rowId
            pos = pos + 1
        End If
    Next rowId
    
    ToArray = result
End Function

Private Sub EnsureCapacity(ByVal rowId As Long)
    Dim oldMax As Long
    
    If rowId <= mMaxRow Then Exit Sub
    
    oldMax = mMaxRow
    mMaxRow = rowId
    
    If oldMax = 0 Then
        ReDim mFlags(1 To mMaxRow)
    Else
        ReDim Preserve mFlags(1 To mMaxRow)
    End If
End Sub
