VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCsrPayload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc 인덱스 rowId 페이로드를 CSR 평탄 배열로 보관 담당
' @depends DbRowList
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' 키당 DbRowList 객체(수만 개 개별 소멸이 teardown 비용원)를 공유 Long 배열로 대체하는
' CSR 페이로드. 백엔드(DictHashIndex/CoreHashIndex)는 키->keyId(1기반 밀집) 매핑만 담당하고
' rowId 보관/조회는 전부 여기로 위임한다. (이슈 2026-07-13_teardown_cost_index_payload)
'
' 수명 규약:
' - AddEntry 스트리밍 누적(빌드 버퍼) -> 첫 읽기 접근자에서 lazy finalize(CSR 산출)
' - AddEntry는 dirty 마킹만 하므로 읽기 후 재적재도 안전(다음 읽기에서 재finalize)
' - 빌드 버퍼는 finalize 후에도 유지(재finalize 원본, 행당 Long 2개라 가벼움)
' - 키 내 rowId 순서 = AddEntry 호출 순서 (안정 counting-sort scatter가 보장)
' - 소멸 = Long 배열 Erase 몇 회 (키 수와 무관하게 O(1) 등급)

' [빌드 버퍼]
Private mEntryKeyId() As Long
Private mEntryRowId() As Long
Private mEntryCount As Long
Private mEntryCap As Long
Private mCounts() As Long   ' keyId -> 건수
Private mKeyCount As Long
Private mKeyCap As Long

' [finalize 산출 - CSR]
Private mOffsets() As Long  ' (1 To mKeyCount+1), CountOf(k) = mOffsets(k+1) - mOffsets(k)
Private mRowIds() As Long   ' (1 To mEntryCount) 공유 rowId 풀
Private mFinalized As Boolean

' keyId(1기반 밀집)와 rowId를 빌드 버퍼에 누적한다
Friend Sub AddEntry(ByVal keyId As Long, ByVal rowId As Long)
    If mEntryCount = mEntryCap Then GrowEntries
    mEntryCount = mEntryCount + 1
    mEntryKeyId(mEntryCount) = keyId
    mEntryRowId(mEntryCount) = rowId

    If keyId > mKeyCap Then GrowCounts keyId
    mCounts(keyId) = mCounts(keyId) + 1
    If keyId > mKeyCount Then mKeyCount = keyId

    mFinalized = False
End Sub

Friend Sub Clear()
    Erase mEntryKeyId
    Erase mEntryRowId
    mEntryCount = 0
    mEntryCap = 0
    Erase mCounts
    mKeyCount = 0
    mKeyCap = 0
    Erase mOffsets
    Erase mRowIds
    mFinalized = False
End Sub

Friend Property Get KeyCount() As Long
    KeyCount = mKeyCount
End Property

Friend Property Get EntryCount() As Long
    EntryCount = mEntryCount
End Property

' keyId의 행 수를 반환한다 (복사 없음)
Friend Function CountOf(ByVal keyId As Long) As Long
    EnsureFinalized
    CountOf = mOffsets(keyId + 1) - mOffsets(keyId)
End Function

' keyId의 첫 rowId를 반환한다 (UpsertFrom 등 첫 매치 전용 경로)
Friend Function FirstRowId(ByVal keyId As Long) As Long
    EnsureFinalized
    FirstRowId = mRowIds(mOffsets(keyId))
End Function

' keyId의 rowId span을 호출자 버퍼 buf(1..outCount)에 복사한다.
' buf는 호출자 소유 재사용 버퍼 - 용량 부족 시 배가 ReDim, 내용은 호출마다 덮어씀.
Friend Sub CopySpan(ByVal keyId As Long, ByRef buf() As Long, ByRef outCount As Long)
    Dim base As Long, i As Long, cap As Long

    EnsureFinalized
    outCount = mOffsets(keyId + 1) - mOffsets(keyId)
    If outCount = 0 Then Exit Sub

    On Error Resume Next
    cap = UBound(buf) ' 미할당이면 0 유지
    On Error GoTo 0
    If cap < outCount Then
        If cap = 0 Then cap = 16
        Do While cap < outCount
            cap = cap * 2
        Loop
        ReDim buf(1 To cap)
    End If

    base = mOffsets(keyId) - 1
    For i = 1 To outCount
        buf(i) = mRowIds(base + i)
    Next i
End Sub

' keyId의 rowId들을 새 DbRowList 복사본으로 반환한다 (세대 결속 탑재, 호환 경로 전용)
Friend Function MakeRowList(ByVal keyId As Long, ByVal ownerPtr As LongPtr, ByVal gen As Long) As DbRowList
    Dim lst As New DbRowList

    EnsureFinalized
    lst.BindRaw ownerPtr, gen
    lst.LoadFromLongSpan mRowIds, mOffsets(keyId), mOffsets(keyId + 1) - mOffsets(keyId)
    Set MakeRowList = lst
End Function

' === 내부 구현 ===

' prefix-sum + 안정 scatter로 CSR을 산출한다. 엔트리를 e 오름차순(=삽입 순서)으로
' 각 키의 다음 빈 칸에 쓰므로 키 내 상대 순서가 AddEntry 순서와 일치한다.
Private Sub EnsureFinalized()
    Dim k As Long, e As Long
    Dim nextFree() As Long

    If mFinalized Then Exit Sub

    ReDim mOffsets(1 To mKeyCount + 1)
    mOffsets(1) = 1
    For k = 1 To mKeyCount
        mOffsets(k + 1) = mOffsets(k) + mCounts(k)
    Next k

    If mEntryCount > 0 Then
        ReDim mRowIds(1 To mEntryCount)
        ReDim nextFree(1 To mKeyCount)
        For k = 1 To mKeyCount
            nextFree(k) = mOffsets(k)
        Next k
        For e = 1 To mEntryCount
            k = mEntryKeyId(e)
            mRowIds(nextFree(k)) = mEntryRowId(e)
            nextFree(k) = nextFree(k) + 1
        Next e
    Else
        Erase mRowIds
    End If

    mFinalized = True
End Sub

Private Sub GrowEntries()
    If mEntryCap = 0 Then
        mEntryCap = 64
        ReDim mEntryKeyId(1 To mEntryCap)
        ReDim mEntryRowId(1 To mEntryCap)
    Else
        mEntryCap = mEntryCap * 2
        ReDim Preserve mEntryKeyId(1 To mEntryCap)
        ReDim Preserve mEntryRowId(1 To mEntryCap)
    End If
End Sub

Private Sub GrowCounts(ByVal lNeed As Long)
    Dim newCap As Long

    newCap = mKeyCap
    If newCap = 0 Then newCap = 16
    Do While newCap < lNeed
        newCap = newCap * 2
    Loop
    If mKeyCap = 0 Then
        ReDim mCounts(1 To newCap)
    Else
        ReDim Preserve mCounts(1 To newCap)
    End If
    mKeyCap = newCap
End Sub
