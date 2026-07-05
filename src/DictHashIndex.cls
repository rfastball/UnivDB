VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictHashIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DbHashIndex

Private mDict As Object

' 세대 계약 결속 (ObjPtr 스냅샷 - 객체 참조 아님, 순환 참조 방지)
Private mOwnerPtr As LongPtr
Private mGeneration As Long

Private Sub Class_Initialize()
    Set mDict = CreateObject("Scripting.Dictionary")
    mDict.CompareMode = vbTextCompare
End Sub

Private Sub Class_Terminate()
    Set mDict = Nothing
End Sub

Public Sub AddRow(ByVal key As Variant, ByVal rowId As Long)
    Dim sKey As String
    Dim rows As DbRowList

    If IsError(key) Then Exit Sub
    sKey = NormalizeKey(key)

    If mDict.exists(sKey) Then
        Set rows = mDict(sKey)
    Else
        Set rows = New DbRowList
        rows.BindRaw mOwnerPtr, mGeneration
        mDict.Add sKey, rows
    End If

    rows.Append rowId
End Sub

' @desc 테이블을 받아 현재 세대에 결속함
Public Sub BindTo(ByVal tbl As clsTable)
    mOwnerPtr = ObjPtr(tbl)
    mGeneration = tbl.Generation
End Sub

' @desc 테이블을 받아 결속 세대와 일치 검증함, stale 시 9998 에러
Public Sub EnsureCurrent(ByVal tbl As clsTable)
    modDbContract.AssertCurrentGeneration mOwnerPtr, mGeneration, tbl, "DictHashIndex.EnsureCurrent"
End Sub

Public Function TryGetRows(ByVal key As Variant, ByRef outRows As DbRowList) As Boolean
    Dim sKey As String
    
    If IsError(key) Then
        Set outRows = Nothing
        TryGetRows = False
        Exit Function
    End If
    
    sKey = NormalizeKey(key)
    If mDict.exists(sKey) Then
        Set outRows = mDict(sKey)
        TryGetRows = True
    Else
        Set outRows = Nothing
        TryGetRows = False
    End If
End Function

Public Property Get CountKeys() As Long
    CountKeys = mDict.Count
End Property

Public Sub Clear()
    mDict.RemoveAll
End Sub

Public Function Cursor() As DbIndexCursor
    Dim cur As New DbIndexCursor
    Set Cursor = cur.Init(mDict)
End Function

Private Function NormalizeKey(ByVal key As Variant) As String
    NormalizeKey = modIndexKey.NormalizeIndexKey(key)
End Function

' ==============================================================================
' DbHashIndex 인터페이스 위임 (기존 Public 표면 유지)
' ==============================================================================

Private Sub DbHashIndex_AddRow(ByVal key As Variant, ByVal rowId As Long)
    AddRow key, rowId
End Sub

Private Function DbHashIndex_TryGetRows(ByVal key As Variant, ByRef outRows As DbRowList) As Boolean
    DbHashIndex_TryGetRows = TryGetRows(key, outRows)
End Function

Private Property Get DbHashIndex_CountKeys() As Long
    DbHashIndex_CountKeys = CountKeys
End Property

Private Sub DbHashIndex_Clear()
    Clear
End Sub

Private Function DbHashIndex_Cursor() As DbIndexCursor
    Set DbHashIndex_Cursor = Cursor()
End Function

Private Sub DbHashIndex_EnsureCurrent(ByVal tbl As clsTable)
    EnsureCurrent tbl
End Sub
