VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoreHashIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc FastHash 엔진 기반 실험 해시 인덱스 백엔드 담당
' @depends DbHashIndex, modDbContract, modIndexKey, Core_clsFastHash, Core_modMemCore
Option Explicit

Implements DbHashIndex

' 구조: Core_clsFastHash가 정규화 키와 slot(1 기반)을 결정하고, rowId 페이로드는
' mPayload(CSR 공유 배열)가 slot = keyId로 보관한다. 키당 객체 0개 - 소멸이
' 키 수와 무관 (이슈 2026-07-13_teardown_cost_index_payload 처방 A).
' build-only 전제: AddRow/Clear만 사용 (FastHash RemoveKey는 swap-last로
' slot을 재배치하므로 인덱스 용도에서 금지 - 재구축은 Clear 후 재적재).
' CSR 페이로드는 slot이 build 동안 1..Count 밀집·삽입 순서 안정임에 의존한다.

Private mEngine As Core_clsFastHash
Private mPayload As DbCsrPayload

' 세대 계약 결속 (ObjPtr 스냅샷 - 객체 참조 아님, 순환 참조 방지)
Private mOwnerPtr As LongPtr
Private mGeneration As Long

Private Sub Class_Initialize()
    Set mEngine = New Core_clsFastHash
    Set mPayload = New DbCsrPayload
End Sub

' @desc 키와 rowID를 받아 인덱스에 적재함
Public Sub AddRow(ByVal key As Variant, ByVal rowId As Long)
    Dim vKey As Variant
    Dim sKey As String
    Dim slot As Long
    Dim isNew As Boolean

    If IsError(key) Then Exit Sub
    sKey = modIndexKey.NormalizeIndexKey(key)
    vKey = sKey
    slot = mEngine.EnsureKeyH(vKey, Core_HashKeyText(sKey), isNew)

    mPayload.AddEntry slot, rowId
End Sub

' @desc 테이블을 받아 현재 세대에 결속함
Public Sub BindTo(ByVal tbl As clsTable)
    mOwnerPtr = ObjPtr(tbl)
    mGeneration = tbl.Generation
End Sub

' @desc 테이블을 받아 결속 세대와 일치 검증함, stale 시 9998 에러
Public Sub EnsureCurrent(ByVal tbl As clsTable)
    modDbContract.AssertCurrentGeneration mOwnerPtr, mGeneration, tbl, "CoreHashIndex.EnsureCurrent"
End Sub

' @desc 키 받아 RowList 반환 여부 반환, 미존재/Error 키 시 False
' 반환 리스트는 세대 결속을 탑재한 fresh 복사본 (인덱스 내부 상태와 분리)
Public Function TryGetRows(ByVal key As Variant, ByRef outRows As DbRowList) As Boolean
    Dim vKey As Variant
    Dim slot As Long

    If IsError(key) Then
        Set outRows = Nothing
        Exit Function
    End If

    vKey = modIndexKey.NormalizeIndexKey(key)
    slot = mEngine.FindIndex(vKey)
    If slot > 0 Then
        Set outRows = mPayload.MakeRowList(slot, mOwnerPtr, mGeneration)
        TryGetRows = True
    Else
        Set outRows = Nothing
    End If
End Function

' @desc 키의 rowId span을 버퍼에 복사 여부 반환, 미존재/Error 키 시 False에 0건
Public Function TryGetRowIdSpan(ByVal key As Variant, ByRef outRowIds() As Long, ByRef outCount As Long) As Boolean
    Dim vKey As Variant
    Dim slot As Long

    outCount = 0
    If IsError(key) Then Exit Function

    vKey = modIndexKey.NormalizeIndexKey(key)
    slot = mEngine.FindIndex(vKey)
    If slot > 0 Then
        mPayload.CopySpan slot, outRowIds, outCount
        TryGetRowIdSpan = True
    End If
End Function

' @desc 키의 행 수만 반환 여부 반환(복사 없음), 미존재/Error 키 시 False에 0건
Public Function TryGetRowCount(ByVal key As Variant, ByRef outCount As Long) As Boolean
    Dim vKey As Variant
    Dim slot As Long

    outCount = 0
    If IsError(key) Then Exit Function

    vKey = modIndexKey.NormalizeIndexKey(key)
    slot = mEngine.FindIndex(vKey)
    If slot > 0 Then
        outCount = mPayload.CountOf(slot)
        TryGetRowCount = True
    End If
End Function

' @desc 키의 첫 rowId 반환 여부 반환, 미존재/Error 키 시 False
Public Function TryGetFirstRowId(ByVal key As Variant, ByRef outRowId As Long) As Boolean
    Dim vKey As Variant
    Dim slot As Long

    If IsError(key) Then Exit Function

    vKey = modIndexKey.NormalizeIndexKey(key)
    slot = mEngine.FindIndex(vKey)
    If slot > 0 Then
        outRowId = mPayload.FirstRowId(slot)
        TryGetFirstRowId = True
    End If
End Function

Public Property Get CountKeys() As Long
    CountKeys = mEngine.Count
End Property

Public Sub Clear()
    mEngine.Clear
    mPayload.Clear
End Sub

' @desc 키-RowList 순회용 커서 반환 (CSR 경로 - 삽입 순서 보존)
Public Function Cursor() As DbIndexCursor
    Dim vKeys As Variant
    Dim cur As New DbIndexCursor

    vKeys = mEngine.KeysArray
    Set Cursor = cur.InitCsr(vKeys, mPayload, mEngine.Count, mOwnerPtr, mGeneration)
End Function

Friend Property Get OwnerPtr() As LongPtr
    OwnerPtr = mOwnerPtr
End Property

Friend Property Get BoundGeneration() As Long
    BoundGeneration = mGeneration
End Property

' ==============================================================================
' DbHashIndex 인터페이스 위임 (기존 Public 표면 유지)
' ==============================================================================

Private Sub DbHashIndex_AddRow(ByVal key As Variant, ByVal rowId As Long)
    AddRow key, rowId
End Sub

Private Function DbHashIndex_TryGetRows(ByVal key As Variant, ByRef outRows As DbRowList) As Boolean
    DbHashIndex_TryGetRows = TryGetRows(key, outRows)
End Function

Private Function DbHashIndex_TryGetRowIdSpan(ByVal key As Variant, ByRef outRowIds() As Long, ByRef outCount As Long) As Boolean
    DbHashIndex_TryGetRowIdSpan = TryGetRowIdSpan(key, outRowIds, outCount)
End Function

Private Function DbHashIndex_TryGetRowCount(ByVal key As Variant, ByRef outCount As Long) As Boolean
    DbHashIndex_TryGetRowCount = TryGetRowCount(key, outCount)
End Function

Private Function DbHashIndex_TryGetFirstRowId(ByVal key As Variant, ByRef outRowId As Long) As Boolean
    DbHashIndex_TryGetFirstRowId = TryGetFirstRowId(key, outRowId)
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
