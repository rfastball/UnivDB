VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictHashIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc Scripting.Dictionary 기반 기본 해시 인덱스 백엔드 담당
' @depends DbHashIndex, DbCsrPayload, modDbContract, modIndexKey
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

Implements DbHashIndex

' 구조: mDict가 접기 키(FoldIndexKey) -> keyId(1기반 밀집 Long)를 결정하고,
' rowId 페이로드는 mPayload(CSR 공유 배열)가 보관한다. 키당 객체 0개 -
' 소멸이 키 수와 무관 (이슈 2026-07-13_teardown_cost_index_payload 처방 A).
' 비교는 binary (대소문자/폭 동치는 접기가 담당 - 2026-10-31 하이브리드 접기).
' 원본 표기(첫 등장 정규화 키)는 mKeyText(keyId - 1)에 보존 - Cursor.Key/GroupBy 렌더 표기.
Private mDict As Object
Private mPayload As DbCsrPayload
Private mKeyText() As String
Private mKeyCap As Long

' 세대 계약 결속 (ObjPtr 스냅샷 - 객체 참조 아님, 순환 참조 방지)
Private mOwnerPtr As LongPtr
Private mGeneration As Long
Private mInvalidated As Boolean

Private Sub Class_Initialize()
    Set mDict = CreateObject("Scripting.Dictionary") ' binary 비교 (기본)
    Set mPayload = New DbCsrPayload
End Sub

Private Sub Class_Terminate()
    Set mDict = Nothing
    Set mPayload = Nothing
End Sub

Public Sub AddRow(ByVal key As Variant, ByVal rowId As Long)
    RequireValid
    Dim sKey As String

    If IsError(key) Then Exit Sub
    sKey = modIndexKey.NormalizeIndexKey(key)
    AddNormalized sKey, rowId
End Sub

' 일괄 빌드: data(firstRow..lastRow, col)을 직접 읽는다 (열 사본, 행당 ByVal 복사와 인터페이스 디스패치 없음).
' 문자열 셀은 그대로, 그 외는 NormalizeIndexKey, Error 셀 skip - AddRow 계약과 동치.
Friend Sub BuildFromArray(ByRef data As Variant, ByVal col As Long, ByVal firstRow As Long, ByVal lastRow As Long)
    RequireValid
    Dim i As Long, sKey As String

    For i = firstRow To lastRow
        If VarType(data(i, col)) = vbString Then
            sKey = data(i, col)
            AddNormalized sKey, i
        ElseIf Not IsError(data(i, col)) Then
            sKey = modIndexKey.NormalizeIndexKey(data(i, col))
            AddNormalized sKey, i
        End If
    Next i
End Sub

' 정규화 키 1건 적재. 새 키는 Item 읽기(자동 추가) + 쓰기 - exists + Add의 이중 조회 회피.
Private Sub AddNormalized(ByRef sKey As String, ByVal rowId As Long)
    Dim f As String, c0 As Long, keyId As Long

    f = modIndexKey.FoldIndexKey(sKey)
    c0 = mDict.Count
    keyId = mDict(f)
    If mDict.Count > c0 Then
        keyId = c0 + 1
        mDict(f) = keyId
        If keyId > mKeyCap Then GrowKeyText keyId
        mKeyText(keyId - 1) = sKey
    End If
    mPayload.AddEntry keyId, rowId
End Sub

' 키 받아 keyId 반환 (미존재 시 0, 조회 전용 - 자동 추가 없음). Error 키는 호출부가 거른다.
Private Function FindKeyId(ByVal key As Variant) As Long
    Dim sKey As String
    sKey = modIndexKey.NormalizeIndexKey(key)
    sKey = modIndexKey.FoldIndexKey(sKey)
    If mDict.exists(sKey) Then FindKeyId = mDict(sKey)
End Function

Private Sub GrowKeyText(ByVal lNeed As Long)
    If mKeyCap = 0 Then mKeyCap = 64
    Do While mKeyCap < lNeed
        mKeyCap = mKeyCap * 2
    Loop
    ReDim Preserve mKeyText(0 To mKeyCap - 1)
End Sub

' @desc 테이블을 받아 현재 세대에 결속함
Public Sub BindTo(ByVal tbl As clsTable)
    RequireValid
    mOwnerPtr = ObjPtr(tbl)
    mGeneration = tbl.Generation
End Sub

' @desc 테이블을 받아 결속 세대와 일치 검증함, stale 시 9998 에러
Public Sub EnsureCurrent(ByVal tbl As clsTable)
    RequireValid
    modDbContract.AssertCurrentGeneration mOwnerPtr, mGeneration, tbl, "DictHashIndex.EnsureCurrent"
End Sub

' 반환 리스트는 세대 결속을 탑재한 fresh 복사본 (인덱스 내부 상태와 분리)
Public Function TryGetRows(ByVal key As Variant, ByRef outRows As DbRowList) As Boolean
    RequireValid
    Dim keyId As Long

    Set outRows = Nothing
    If IsError(key) Then Exit Function

    keyId = FindKeyId(key)
    If keyId > 0 Then
        Set outRows = mPayload.MakeRowList(keyId, mOwnerPtr, mGeneration)
        TryGetRows = True
    End If
End Function

' @desc 키의 rowId span을 버퍼에 복사 여부 반환, 미존재/Error 키 시 False에 0건
Public Function TryGetRowIdSpan(ByVal key As Variant, ByRef outRowIds() As Long, ByRef outCount As Long) As Boolean
    RequireValid
    Dim keyId As Long

    outCount = 0
    If IsError(key) Then Exit Function

    keyId = FindKeyId(key)
    If keyId > 0 Then
        mPayload.CopySpan keyId, outRowIds, outCount
        TryGetRowIdSpan = True
    End If
End Function

' @desc 키의 행 수만 반환 여부 반환(복사 없음), 미존재/Error 키 시 False에 0건
Public Function TryGetRowCount(ByVal key As Variant, ByRef outCount As Long) As Boolean
    RequireValid
    Dim keyId As Long

    outCount = 0
    If IsError(key) Then Exit Function

    keyId = FindKeyId(key)
    If keyId > 0 Then
        outCount = mPayload.CountOf(keyId)
        TryGetRowCount = True
    End If
End Function

' @desc 키의 첫 rowId 반환 여부 반환, 미존재/Error 키 시 False
Public Function TryGetFirstRowId(ByVal key As Variant, ByRef outRowId As Long) As Boolean
    RequireValid
    Dim keyId As Long

    If IsError(key) Then Exit Function

    keyId = FindKeyId(key)
    If keyId > 0 Then
        outRowId = mPayload.FirstRowId(keyId)
        TryGetFirstRowId = True
    End If
End Function

Public Property Get CountKeys() As Long
    RequireValid
    CountKeys = mDict.Count
End Property

Public Sub Clear()
    RequireValid
    mDict.RemoveAll
    mPayload.Clear
    Erase mKeyText
    mKeyCap = 0
End Sub

Public Function Cursor() As DbIndexCursor
    RequireValid
    Dim cur As New DbIndexCursor

    ' 키 순서 = keyId 1..Count = 첫 등장 순서, 표기는 첫 등장 원본 (접기 전)
    Set Cursor = cur.InitCsr(mKeyText, mPayload, mDict.Count, mOwnerPtr, mGeneration)
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

Private Sub DbHashIndex_Invalidate()
    mInvalidated = True
End Sub

Private Sub RequireValid()
    If mInvalidated Then
        Err.Raise modDbContract.DB_ERR_STALE_GENERATION, "DbHashIndex", "Index invalidated: reacquire from table"
    End If
End Sub
