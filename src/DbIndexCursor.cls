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

' CSR 모드 (키 순서 = keyId 1..mCount, keys 배열은 0기반 - keyId = mPos + 1)
Private mPayload As DbCsrPayload
Private mCsrBacked As Boolean
Private mOwnerPtr As LongPtr
Private mGeneration As Long

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

' CSR 페이로드 기반 초기화. keys는 keyId 순서(0기반), payload keyId = 위치 + 1.
' ownerPtr/gen은 Rows 복사본에 전파할 세대 결속.
Friend Function InitCsr(ByVal keys As Variant, ByVal payload As DbCsrPayload, _
                        ByVal count As Long, ByVal ownerPtr As LongPtr, _
                        ByVal gen As Long) As DbIndexCursor
    Set mDict = Nothing
    mArrayBacked = False
    Erase mRows
    mCsrBacked = True
    Set mPayload = payload
    mCount = count
    If mCount = 0 Then
        mKeys = Array()
    Else
        mKeys = keys
    End If
    mOwnerPtr = ownerPtr
    mGeneration = gen
    mPos = -1
    Set InitCsr = Me
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
    If mCsrBacked Then
        ' CSR 모드: 세대 결속을 탑재한 fresh 복사본 (핫 경로는 RowCount/CopyRows 사용)
        Set Rows = mPayload.MakeRowList(mPos + 1, mOwnerPtr, mGeneration)
    ElseIf mArrayBacked Then
        Set Rows = mRows(mPos)
    Else
        Set Rows = mDict(mKeys(mPos))
    End If
End Property

' 현재 키의 행 수 (복사 없음, CSR 모드 전용)
Public Property Get RowCount() As Long
    If mPos < 0 Or mPos >= mCount Then
        Err.Raise 9, "DbIndexCursor.RowCount", "Cursor is not positioned on an entry"
    End If
    RowCount = mPayload.CountOf(mPos + 1)
End Property

' 현재 키의 rowId span을 호출자 재사용 버퍼에 복사 (CSR 모드 전용)
Public Sub CopyRows(ByRef outRowIds() As Long, ByRef outCount As Long)
    If mPos < 0 Or mPos >= mCount Then
        Err.Raise 9, "DbIndexCursor.CopyRows", "Cursor is not positioned on an entry"
    End If
    mPayload.CopySpan mPos + 1, outRowIds, outCount
End Sub
