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

Private mDict As Object

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
        mDict.Add sKey, rows
    End If
    
    rows.Append rowId
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

Public Function KeysArray() As Variant
    If mDict.Count = 0 Then
        KeysArray = Array()
    Else
        KeysArray = mDict.Keys
    End If
End Function

Public Function Cursor() As DbIndexCursor
    Dim cur As New DbIndexCursor
    Set Cursor = cur.Init(mDict)
End Function

Private Function NormalizeKey(ByVal key As Variant) As String
    If IsEmpty(key) Then
        NormalizeKey = ""
    Else
        NormalizeKey = CStr(key)
    End If
End Function
