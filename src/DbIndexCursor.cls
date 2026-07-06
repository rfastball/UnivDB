VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbIndexCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDict As Object
Private mKeys As Variant
Private mRows() As DbRowList
Private mArrayBacked As Boolean
Private mPos As Long
Private mCount As Long

Public Function Init(ByVal dict As Object) As DbIndexCursor
    Set mDict = dict
    mArrayBacked = False
    Erase mRows
    mCount = dict.Count
    If mCount = 0 Then
        mKeys = Array()
    Else
        mKeys = dict.Keys
    End If
    mPos = -1
    Set Init = Me
End Function

Public Function InitArrays(ByVal keys As Variant, ByRef rows() As DbRowList, ByVal count As Long) As DbIndexCursor
    Dim i As Long

    Set mDict = Nothing
    mArrayBacked = True
    mCount = count
    If mCount = 0 Then
        mKeys = Array()
        Erase mRows
    Else
        mKeys = keys
        ReDim mRows(0 To mCount - 1)
        For i = 0 To mCount - 1
            Set mRows(i) = rows(i + 1)
        Next i
    End If
    mPos = -1
    Set InitArrays = Me
End Function

Public Function MoveNext() As Boolean
    If mPos + 1 >= mCount Then Exit Function
    mPos = mPos + 1
    MoveNext = True
End Function

Public Property Get Key() As Variant
    If mPos < 0 Or mPos >= mCount Then
        Err.Raise 9, "DbIndexCursor.Key", "Cursor is not positioned on an entry"
    End If
    Key = mKeys(mPos)
End Property

Public Property Get Rows() As DbRowList
    If mPos < 0 Or mPos >= mCount Then
        Err.Raise 9, "DbIndexCursor.Rows", "Cursor is not positioned on an entry"
    End If
    If mArrayBacked Then
        Set Rows = mRows(mPos)
    Else
        Set Rows = mDict(mKeys(mPos))
    End If
End Property
