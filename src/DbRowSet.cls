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
    modDbContract.AssertCurrentGeneration mOwnerPtr, mGeneration, tbl, "DbRowSet.EnsureCurrent"
End Sub

Friend Property Get OwnerPtr() As LongPtr
    OwnerPtr = mOwnerPtr
End Property

Friend Property Get BoundGeneration() As Long
    BoundGeneration = mGeneration
End Property

' 결합 대상 RowList의 결속을 승계하거나, 서로 다른 결속이면 즉시 실패
Private Sub AdoptOrVerifyBinding(ByVal rows As DbRowList)
    If rows Is Nothing Then Exit Sub
    If rows.OwnerPtr = 0 Then Exit Sub

    If mOwnerPtr = 0 Then
        BindRaw rows.OwnerPtr, rows.BoundGeneration
    ElseIf mOwnerPtr <> rows.OwnerPtr Or mGeneration <> rows.BoundGeneration Then
        Err.Raise modDbContract.DB_ERR_STALE_GENERATION, "DbRowSet", _
            "Derived structures bound to different table generations"
    End If
End Sub

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

    AdoptOrVerifyBinding rows
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

    AdoptOrVerifyBinding rows
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
