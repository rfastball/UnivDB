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

' 구조: mDict가 정규화 키 -> keyId(1기반 밀집 Long)를 결정하고,
' rowId 페이로드는 mPayload(CSR 공유 배열)가 보관한다. 키당 객체 0개 -
' 소멸이 키 수와 무관 (이슈 2026-07-13_teardown_cost_index_payload 처방 A).
Private mDict As Object
Private mPayload As DbCsrPayload

' 세대 계약 결속 (ObjPtr 스냅샷 - 객체 참조 아님, 순환 참조 방지)
Private mOwnerPtr As LongPtr
Private mGeneration As Long

Private Sub Class_Initialize()
    Set mDict = CreateObject("Scripting.Dictionary")
    mDict.CompareMode = vbTextCompare
    Set mPayload = New DbCsrPayload
End Sub

Private Sub Class_Terminate()
    Set mDict = Nothing
    Set mPayload = Nothing
End Sub

Public Sub AddRow(ByVal key As Variant, ByVal rowId As Long)
    Dim sKey As String
    Dim keyId As Long

    If IsError(key) Then Exit Sub
    sKey = NormalizeKey(key)

    If mDict.exists(sKey) Then
        keyId = mDict(sKey)
    Else
        keyId = mDict.Count + 1
        mDict.Add sKey, keyId
    End If

    mPayload.AddEntry keyId, rowId
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

' 반환 리스트는 세대 결속을 탑재한 fresh 복사본 (인덱스 내부 상태와 분리)
Public Function TryGetRows(ByVal key As Variant, ByRef outRows As DbRowList) As Boolean
    Dim sKey As String

    If IsError(key) Then
        Set outRows = Nothing
        TryGetRows = False
        Exit Function
    End If

    sKey = NormalizeKey(key)
    If mDict.exists(sKey) Then
        Set outRows = mPayload.MakeRowList(mDict(sKey), mOwnerPtr, mGeneration)
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
    mPayload.Clear
End Sub

Public Function Cursor() As DbIndexCursor
    Dim cur As New DbIndexCursor
    Dim vKeys As Variant

    ' keyId는 mDict.Count+1 순차 부여이므로 dict.Keys 순서(삽입 순서) = keyId 1..Count
    If mDict.Count = 0 Then
        vKeys = Array()
    Else
        vKeys = mDict.Keys
    End If
    Set Cursor = cur.InitCsr(vKeys, mPayload, mDict.Count, mOwnerPtr, mGeneration)
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
