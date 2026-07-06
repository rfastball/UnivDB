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

' 구조: Core_clsFastHash가 정규화 키와 slot을 결정하고,
' mRows 병렬 배열이 slot(1 기반) -> DbRowList payload를 연결한다.
' build-only 전제: AddRow/Clear만 사용 (FastHash RemoveKey는 swap-last로
' slot을 재배치하므로 인덱스 용도에서 금지 - 재구축은 Clear 후 재적재).

Private mEngine As Core_clsFastHash
Private mRows() As DbRowList
Private mRowsCap As Long

' 세대 계약 결속 (ObjPtr 스냅샷 - 객체 참조 아님, 순환 참조 방지)
Private mOwnerPtr As LongPtr
Private mGeneration As Long

Private Sub Class_Initialize()
    Set mEngine = New Core_clsFastHash
    mRowsCap = 0
End Sub

' @desc 키와 rowID를 받아 인덱스에 적재함
Public Sub AddRow(ByVal key As Variant, ByVal rowId As Long)
    Dim vKey As Variant
    Dim sKey As String
    Dim slot As Long
    Dim isNew As Boolean
    Dim rows As DbRowList

    If IsError(key) Then Exit Sub
    sKey = modIndexKey.NormalizeIndexKey(key)
    vKey = sKey
    slot = mEngine.EnsureKeyH(vKey, Core_HashKeyText(sKey), isNew)

    If slot > mRowsCap Then GrowRows slot
    If isNew Then
        Set rows = New DbRowList
        rows.BindRaw mOwnerPtr, mGeneration
        Set mRows(slot) = rows
    End If
    mRows(slot).Append rowId
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
        Set outRows = mRows(slot)
        TryGetRows = True
    Else
        Set outRows = Nothing
    End If
End Function

Public Property Get CountKeys() As Long
    CountKeys = mEngine.Count
End Property

Public Sub Clear()
    mEngine.Clear
    Erase mRows
    mRowsCap = 0
End Sub

' @desc 키-RowList 순회용 커서 반환 (경계 물질화 - 삽입 순서 보존)
Public Function Cursor() As DbIndexCursor
    Dim d As Object
    Dim vKeys As Variant
    Dim i As Long
    Dim cur As New DbIndexCursor

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    vKeys = mEngine.KeysArray
    For i = 0 To mEngine.Count - 1
        d.Add vKeys(i), mRows(i + 1) ' build-only라 slot = 삽입 순서 (밀집)
    Next i
    Set Cursor = cur.Init(d)
End Function

Friend Property Get OwnerPtr() As LongPtr
    OwnerPtr = mOwnerPtr
End Property

Friend Property Get BoundGeneration() As Long
    BoundGeneration = mGeneration
End Property

Private Sub GrowRows(ByVal lNeed As Long)
    Dim newCap As Long

    newCap = mRowsCap
    If newCap = 0 Then newCap = 16
    Do While newCap < lNeed
        newCap = newCap * 2
    Loop
    If mRowsCap = 0 Then
        ReDim mRows(1 To newCap)
    Else
        ReDim Preserve mRows(1 To newCap)
    End If
    mRowsCap = newCap
End Sub

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
