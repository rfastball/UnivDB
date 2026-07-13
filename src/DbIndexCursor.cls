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

' CSR 페이로드 기반 단일 모드 (구 dict/array 모드는 CSR 전환으로 폐기).
' 키 순서 = keyId 1..mCount, keys 배열은 0기반 - keyId = mPos + 1.
Private mKeys As Variant
Private mPayload As DbCsrPayload
Private mPos As Long
Private mCount As Long
Private mOwnerPtr As LongPtr
Private mGeneration As Long

' CSR 페이로드 기반 초기화. ownerPtr/gen은 Rows 복사본에 전파할 세대 결속.
Friend Function InitCsr(ByVal keys As Variant, ByVal payload As DbCsrPayload, _
                        ByVal count As Long, ByVal ownerPtr As LongPtr, _
                        ByVal gen As Long) As DbIndexCursor
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
    ' 세대 결속을 탑재한 fresh 복사본 (핫 경로는 RowCount/CopyRows 사용)
    Set Rows = mPayload.MakeRowList(mPos + 1, mOwnerPtr, mGeneration)
End Property

' 현재 키의 행 수 (복사 없음)
Public Property Get RowCount() As Long
    If mPos < 0 Or mPos >= mCount Then
        Err.Raise 9, "DbIndexCursor.RowCount", "Cursor is not positioned on an entry"
    End If
    RowCount = mPayload.CountOf(mPos + 1)
End Property

' 현재 키의 rowId span을 호출자 재사용 버퍼에 복사
Public Sub CopyRows(ByRef outRowIds() As Long, ByRef outCount As Long)
    If mPos < 0 Or mPos >= mCount Then
        Err.Raise 9, "DbIndexCursor.CopyRows", "Cursor is not positioned on an entry"
    End If
    mPayload.CopySpan mPos + 1, outRowIds, outCount
End Sub
