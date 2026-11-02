VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Core_clsFastHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc 병렬 배열 기반 고속 해시 집합 엔진
' @depends Core_modMemCore
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' 구조: mBuckets(해시 And mMask) = 체인 헤드 엔트리 인덱스(1 기반, 0=빈).
' 엔트리 1..mCount는 항상 밀집. 삭제는 체인 unlink + swap-last로 밀집 유지.
' 키 등가: 문자열은 StrComp vbTextCompare, 숫자류(Boolean/Date 포함)는 CDbl 등가
' (0단계 특성화 테스트로 확정한 Scripting.Dictionary 의미론).
'
' slot API: FindIndexH/EnsureKeyH가 반환하는 엔트리 인덱스는 1 기반이며
' KeyAt/HashAt과 같은 좌표계다. AddKey/AddKeyH/EnsureKeyH만 수행하는 동안 기존
' slot은 유지된다. RemoveKey/Clear는 swap-last/초기화로 외부 slot-bound payload를
' 무효화하므로, payload 소유자는 자체 세대 계약으로 방어해야 한다.
'
' view 모드: Compare 결과셋은 키를 복사하지 않고 부모 엔진 참조 + 인덱스만 보유.
' 부모의 인덱스 재배치(RemoveKey/Clear)는 MutStamp로 감지해 이후 접근 시 명시 오류.
' 부모 AddKey는 기존 인덱스를 보존하므로 view가 계속 유효함. view에 쓰기가 오면
' 그 시점 자동 물질화(자체 저장소로 복사 전환)한다.

Private Const INITIAL_BUCKETS As Long = 16
Private Const INITIAL_CAPACITY As Long = 16
Private Const ERR_VIEW_STALE As Long = vbObjectError + 513

Private mBuckets() As Long
Private mNextIdx() As Long
Private mKeys() As Variant
Private mHashes() As Long
Private mFlags() As Boolean
Private mCount As Long
Private mCapacity As Long
Private mMask As Long
Private mMutStamp As Long

Private mViewParent As Core_clsFastHash
Private mViewIdx() As Long
Private mViewMember() As Boolean
Private mViewCount As Long
Private mViewStamp As Long

Private Sub Class_Initialize()
    Clear
End Sub

' @desc 모든 엔트리를 비워 초기 크기로 재설정함
Public Sub Clear()
    Set mViewParent = Nothing
    Erase mViewIdx
    Erase mViewMember
    mViewCount = 0
    mMutStamp = mMutStamp + 1
    mMask = INITIAL_BUCKETS - 1
    mCapacity = INITIAL_CAPACITY
    ReDim mBuckets(0 To mMask)
    ReDim mNextIdx(1 To mCapacity)
    ReDim mKeys(1 To mCapacity)
    ReDim mHashes(1 To mCapacity)
    ReDim mFlags(1 To mCapacity)
    mCount = 0
End Sub

Public Property Get Count() As Long
    If mViewParent Is Nothing Then
        Count = mCount
    Else
        EnsureReadable
        Count = mViewCount
    End If
End Property

Public Property Get MutStamp() As Long
    MutStamp = mMutStamp
End Property

' view이면서 부모가 구조 변경된 상태인지 반환 (비발생 진단 프로브)
Public Property Get IsStale() As Boolean
    If mViewParent Is Nothing Then Exit Property
    IsStale = (mViewParent.MutStamp <> mViewStamp)
End Property

' @desc view면 부모에서 자체 저장소로 물질화해 분리함
Public Sub EnsureOwner()
    MaterializeIfView
End Sub

' @desc 부모 엔진과 인덱스 배열(0 기반, n개)로 view를 구성함
Public Sub BindView(ByVal parentEng As Core_clsFastHash, ByRef idxArr() As Long, ByVal N As Long)
    Dim i As Long
    Dim pc As Long

    Clear
    Set mViewParent = parentEng
    mViewStamp = parentEng.MutStamp
    mViewCount = N
    If N > 0 Then
        ReDim mViewIdx(0 To N - 1)
        For i = 0 To N - 1
            mViewIdx(i) = idxArr(i)
        Next i
    End If
    pc = parentEng.Count
    If pc > 0 Then
        ReDim mViewMember(1 To pc)
        For i = 0 To N - 1
            mViewMember(mViewIdx(i)) = True
        Next i
    End If
End Sub

' @desc 예상 엔트리 수 받아 배열과 버킷을 선확장함
Public Sub EnsureCapacity(ByVal N As Long)
    Dim newCap As Long
    Dim needBuckets As Long

    MaterializeIfView
    If N <= 0 Then Exit Sub
    newCap = mCapacity
    Do While newCap < N
        newCap = newCap * 2
    Loop
    If newCap > mCapacity Then
        mCapacity = newCap
        ReDim Preserve mNextIdx(1 To mCapacity)
        ReDim Preserve mKeys(1 To mCapacity)
        ReDim Preserve mHashes(1 To mCapacity)
        ReDim Preserve mFlags(1 To mCapacity)
    End If

    needBuckets = NextPow2(N + N \ 3)
    If needBuckets > mMask + 1 Then
        mMask = needBuckets - 1
        RebuildBuckets
    End If
End Sub

' @desc 키 받아 신규 삽입 시 True 반환, 중복/무효 키 시 False
Public Function AddKey(ByRef vKey As Variant) As Boolean
    If Not IsValidKey(vKey) Then Exit Function
    MaterializeIfView
    AddKey = AddKeyH(vKey, HashOf(vKey))
End Function

' @desc 사전 계산 해시로 유효 키 삽입, 신규 시 True (엔진 간 이동 전용)
Public Function AddKeyH(ByRef vKey As Variant, ByVal h As Long) As Boolean
    Dim isNew As Boolean

    MaterializeIfView
    EnsureKeyH vKey, h, isNew
    AddKeyH = isNew
End Function

' @desc 사전 계산 해시로 유효 키 slot 인덱스 반환, 부재 시 0 (1 기반 slot 좌표)
Public Function FindIndexH(ByRef vKey As Variant, ByVal h As Long) As Long
    Dim prevIdx As Long

    MaterializeIfView
    FindIndexH = FindSlot(vKey, h, prevIdx)
End Function

' @desc 사전 계산 해시로 유효 키 slot을 보장하고 1 기반 slot 반환, 신규 여부를 outIsNew에 기록
Public Function EnsureKeyH(ByRef vKey As Variant, ByVal h As Long, ByRef outIsNew As Boolean) As Long
    Dim prevIdx As Long
    Dim idx As Long
    Dim B As Long

    MaterializeIfView
    outIsNew = False
    idx = FindSlot(vKey, h, prevIdx)
    If idx > 0 Then
        EnsureKeyH = idx
        Exit Function
    End If

    If mCount = mCapacity Then GrowEntries
    If (mCount + 1) * 4 > (mMask + 1) * 3 Then GrowBuckets

    mCount = mCount + 1
    mKeys(mCount) = vKey
    mHashes(mCount) = h
    mFlags(mCount) = False
    B = h And mMask
    mNextIdx(mCount) = mBuckets(B)
    mBuckets(B) = mCount
    outIsNew = True
    EnsureKeyH = mCount
End Function

' @desc 키 존재 여부 반환, 무효 키 시 False
Public Function ExistsKey(ByRef vKey As Variant) As Boolean
    Dim prevIdx As Long
    Dim pidx As Long

    If Not IsValidKey(vKey) Then Exit Function
    If mViewParent Is Nothing Then
        ExistsKey = (FindSlot(vKey, HashOf(vKey), prevIdx) > 0)
    Else
        EnsureReadable
        If mViewCount = 0 Then Exit Function
        pidx = mViewParent.FindIndex(vKey)
        If pidx > 0 Then
            If pidx <= UBound(mViewMember) Then ExistsKey = mViewMember(pidx)
        End If
    End If
End Function

' @desc 키의 엔트리 인덱스 반환, 부재/무효 키 시 0
Public Function FindIndex(ByRef vKey As Variant) As Long
    Dim prevIdx As Long

    If Not IsValidKey(vKey) Then Exit Function
    MaterializeIfView
    FindIndex = FindSlot(vKey, HashOf(vKey), prevIdx)
End Function

' @desc 키 존재 시 방문 플래그 표시 후 True 반환, 부재 시 False
Public Function ExistsAndMark(ByRef vKey As Variant) As Boolean
    If Not IsValidKey(vKey) Then Exit Function
    MaterializeIfView
    ExistsAndMark = ExistsAndMarkH(vKey, HashOf(vKey))
End Function

' @desc 사전 계산 해시로 존재 시 플래그 표시 후 True (엔진 간 이동 전용)
Public Function ExistsAndMarkH(ByRef vKey As Variant, ByVal h As Long) As Boolean
    Dim idx As Long
    Dim prevIdx As Long

    MaterializeIfView
    idx = FindSlot(vKey, h, prevIdx)
    If idx > 0 Then
        mFlags(idx) = True
        ExistsAndMarkH = True
    End If
End Function

' @desc 키 받아 제거 성공 시 True 반환, 부재/무효 키 시 False
Public Function RemoveKey(ByRef vKey As Variant) As Boolean
    Dim h As Long
    Dim cur As Long
    Dim prevIdx As Long
    Dim b2 As Long
    Dim walk As Long

    If Not IsValidKey(vKey) Then Exit Function
    MaterializeIfView
    h = HashOf(vKey)
    cur = FindSlot(vKey, h, prevIdx)
    If cur = 0 Then Exit Function

    If prevIdx = 0 Then
        mBuckets(h And mMask) = mNextIdx(cur)
    Else
        mNextIdx(prevIdx) = mNextIdx(cur)
    End If

    If cur < mCount Then
        mKeys(cur) = mKeys(mCount)
        mHashes(cur) = mHashes(mCount)
        mFlags(cur) = mFlags(mCount)
        mNextIdx(cur) = mNextIdx(mCount)
        b2 = mHashes(cur) And mMask
        If mBuckets(b2) = mCount Then
            mBuckets(b2) = cur
        Else
            walk = mBuckets(b2)
            Do While mNextIdx(walk) <> mCount
                walk = mNextIdx(walk)
            Loop
            mNextIdx(walk) = cur
        End If
    End If

    mKeys(mCount) = Empty
    mCount = mCount - 1
    ' swap-last로 인덱스가 재배치되므로 이 엔진을 부모로 삼은 view들을 무효화
    mMutStamp = mMutStamp + 1
    RemoveKey = True
End Function

' @desc 키들을 0 기반 배열로 반환, 빈 엔진 시 Array() 반환
Public Function KeysArray() As Variant
    Dim arr() As Variant
    Dim i As Long

    If Not mViewParent Is Nothing Then
        EnsureReadable
        If mViewCount = 0 Then
            KeysArray = Array()
            Exit Function
        End If
        ReDim arr(0 To mViewCount - 1)
        For i = 0 To mViewCount - 1
            arr(i) = mViewParent.KeyAt(mViewIdx(i))
        Next i
        KeysArray = arr
        Exit Function
    End If

    If mCount = 0 Then
        KeysArray = Array()
        Exit Function
    End If
    ReDim arr(0 To mCount - 1)
    For i = 1 To mCount
        arr(i - 1) = mKeys(i)
    Next i
    KeysArray = arr
End Function

' @desc 배열 받아 유효 키만 선확장 후 일괄 추가함 (핫패스 인라인 루프)
Public Sub LoadKeys(ByRef vArr As Variant)
    Dim lb As Long
    Dim ub As Long
    Dim i As Long
    Dim h As Long
    Dim idx As Long
    Dim B As Long
    Dim s As String
    Dim vKey As Variant

    If Not IsArray(vArr) Then Exit Sub
    On Error GoTo EmptyArr
    lb = LBound(vArr)
    ub = UBound(vArr)
    On Error GoTo 0
    If ub < lb Then Exit Sub

    MaterializeIfView
    EnsureCapacity mCount + (ub - lb + 1)
    ' AddKey/FindSlot 호출 중첩을 제거한 인라인 루프 (측정: 디스패치 ~0.1us/호출)
    For i = lb To ub
        vKey = vArr(i)
        If IsValidKey(vKey) Then
            If VarType(vKey) = vbString Then
                s = vKey
                h = Core_HashKeyText(s)
            Else
                h = Core_HashKeyNumber(CDbl(vKey))
            End If
            idx = mBuckets(h And mMask)
            Do While idx > 0
                If mHashes(idx) = h Then
                    If KeysEqual(mKeys(idx), vKey) Then Exit Do
                End If
                idx = mNextIdx(idx)
            Loop
            If idx = 0 Then
                If mCount = mCapacity Then GrowEntries
                If (mCount + 1) * 4 > (mMask + 1) * 3 Then GrowBuckets
                mCount = mCount + 1
                mKeys(mCount) = vKey
                mHashes(mCount) = h
                mFlags(mCount) = False
                B = h And mMask
                mNextIdx(mCount) = mBuckets(B)
                mBuckets(B) = mCount
            End If
        End If
    Next i
    Exit Sub
EmptyArr:
End Sub

' @desc 키/해시/개수를 0 기반 배열로 일괄 반출함
Public Sub ExportEntries(ByRef outKeys() As Variant, ByRef outHashes() As Long, ByRef outN As Long)
    Dim i As Long
    Dim p As Long

    If Not mViewParent Is Nothing Then
        EnsureReadable
        outN = mViewCount
        If mViewCount = 0 Then Exit Sub
        ReDim outKeys(0 To mViewCount - 1)
        ReDim outHashes(0 To mViewCount - 1)
        For i = 0 To mViewCount - 1
            p = mViewIdx(i)
            outKeys(i) = mViewParent.KeyAt(p)
            outHashes(i) = mViewParent.HashAt(p)
        Next i
        Exit Sub
    End If

    outN = mCount
    If mCount = 0 Then Exit Sub
    ReDim outKeys(0 To mCount - 1)
    ReDim outHashes(0 To mCount - 1)
    For i = 1 To mCount
        outKeys(i - 1) = mKeys(i)
        outHashes(i - 1) = mHashes(i)
    Next i
End Sub

' @desc 상대 엔진 키를 단일 패스 분류해 결과를 view로 바인딩함
Public Sub ClassifySplit(ByVal otherEng As Core_clsFastHash, _
                         ByVal engCommon As Core_clsFastHash, _
                         ByVal engDiffOther As Core_clsFastHash, _
                         ByVal engDiffMe As Core_clsFastHash)
    Dim ks() As Variant
    Dim hs() As Long
    Dim N As Long
    Dim i As Long
    Dim h As Long
    Dim idx As Long
    Dim comIdx() As Long
    Dim comN As Long
    Dim doIdx() As Long
    Dim doN As Long
    Dim dmIdx() As Long
    Dim dmN As Long

    ' view 체인 방지: 결과 view의 부모(Me, otherEng)는 항상 owner로 고정
    MaterializeIfView
    otherEng.EnsureOwner
    ResetFlags
    otherEng.ExportEntries ks, hs, N

    ' 결과 삽입(BSTR 복사 + 프로브) 대신 부모 인덱스만 누적 (집합 키는 유일하므로 dedupe 불필요)
    If N > 0 Then
        ReDim comIdx(0 To N - 1)
        ReDim doIdx(0 To N - 1)
    End If
    For i = 0 To N - 1
        h = hs(i)
        idx = mBuckets(h And mMask)
        Do While idx > 0
            If mHashes(idx) = h Then
                If KeysEqual(mKeys(idx), ks(i)) Then Exit Do
            End If
            idx = mNextIdx(idx)
        Loop
        If idx > 0 Then
            mFlags(idx) = True
            comIdx(comN) = i + 1
            comN = comN + 1
        Else
            doIdx(doN) = i + 1
            doN = doN + 1
        End If
    Next i

    If Not engDiffMe Is Nothing Then
        If mCount > 0 Then ReDim dmIdx(0 To mCount - 1)
        For i = 1 To mCount
            If Not mFlags(i) Then
                dmIdx(dmN) = i
                dmN = dmN + 1
            End If
        Next i
        engDiffMe.BindView Me, dmIdx, dmN
    End If
    If Not engCommon Is Nothing Then engCommon.BindView otherEng, comIdx, comN
    If Not engDiffOther Is Nothing Then engDiffOther.BindView otherEng, doIdx, doN
End Sub

' @desc 모든 방문 플래그를 False로 초기화함
Public Sub ResetFlags()
    Dim i As Long
    MaterializeIfView
    For i = 1 To mCount
        mFlags(i) = False
    Next i
End Sub

' @desc 1 기반 인덱스 받아 키 반환
Public Function KeyAt(ByVal idx As Long) As Variant
    If mViewParent Is Nothing Then
        KeyAt = mKeys(idx)
    Else
        EnsureReadable
        KeyAt = mViewParent.KeyAt(mViewIdx(idx - 1))
    End If
End Function

' @desc 1 기반 인덱스 받아 방문 플래그 반환
Public Function FlagAt(ByVal idx As Long) As Boolean
    MaterializeIfView
    FlagAt = mFlags(idx)
End Function

' @desc 1 기반 인덱스 받아 캐시된 해시 반환
Public Function HashAt(ByVal idx As Long) As Long
    If mViewParent Is Nothing Then
        HashAt = mHashes(idx)
    Else
        EnsureReadable
        HashAt = mViewParent.HashAt(mViewIdx(idx - 1))
    End If
End Function

' @desc 다른 엔진의 키들을 캐시 해시 재사용으로 일괄 병합함
Public Sub MergeFrom(ByVal other As Core_clsFastHash)
    Dim i As Long
    Dim N As Long

    MaterializeIfView
    N = other.Count
    If N = 0 Then Exit Sub
    EnsureCapacity mCount + N
    For i = 1 To N
        AddKeyH other.KeyAt(i), other.HashAt(i)
    Next i
End Sub

Private Sub EnsureReadable()
    If mViewParent Is Nothing Then Exit Sub
    If mViewParent.MutStamp <> mViewStamp Then
        Err.Raise ERR_VIEW_STALE, "Core_clsFastHash", _
                  "view 무효: 원본 집합이 변경되어 이 결과를 더 이상 신뢰할 수 없음"
    End If
End Sub

Private Sub MaterializeIfView()
    Dim p As Core_clsFastHash
    Dim N As Long
    Dim i As Long
    Dim srcIdx As Long
    Dim B As Long

    If mViewParent Is Nothing Then Exit Sub
    EnsureReadable
    Set p = mViewParent
    N = mViewCount

    ' owner 상태로 초기화 후 부모에서 복사 (키 유일 보장이라 프로브 없이 직접 삽입)
    Set mViewParent = Nothing
    mMask = INITIAL_BUCKETS - 1
    mCapacity = INITIAL_CAPACITY
    ReDim mBuckets(0 To mMask)
    ReDim mNextIdx(1 To mCapacity)
    ReDim mKeys(1 To mCapacity)
    ReDim mHashes(1 To mCapacity)
    ReDim mFlags(1 To mCapacity)
    mCount = 0
    EnsureCapacity N
    For i = 0 To N - 1
        srcIdx = mViewIdx(i)
        mCount = mCount + 1
        mKeys(mCount) = p.KeyAt(srcIdx)
        mHashes(mCount) = p.HashAt(srcIdx)
        mFlags(mCount) = False
        B = mHashes(mCount) And mMask
        mNextIdx(mCount) = mBuckets(B)
        mBuckets(B) = mCount
    Next i
    Erase mViewIdx
    Erase mViewMember
    mViewCount = 0
End Sub

Private Function IsValidKey(ByRef vKey As Variant) As Boolean
    If IsObject(vKey) Then Exit Function
    If IsArray(vKey) Then Exit Function
    If IsEmpty(vKey) Then Exit Function
    If IsNull(vKey) Then Exit Function
    If IsError(vKey) Then Exit Function
    If VarType(vKey) = vbString Then
        IsValidKey = True
    Else
        IsValidKey = IsNumeric(vKey) Or VarType(vKey) = vbDate Or VarType(vKey) = vbBoolean
    End If
End Function

Private Function HashOf(ByRef vKey As Variant) As Long
    Dim s As String
    If VarType(vKey) = vbString Then
        s = vKey
        HashOf = Core_HashKeyText(s)
    Else
        HashOf = Core_HashKeyNumber(CDbl(vKey))
    End If
End Function

Private Function FindSlot(ByRef vKey As Variant, ByVal h As Long, ByRef prevIdx As Long) As Long
    Dim idx As Long

    prevIdx = 0
    idx = mBuckets(h And mMask)
    Do While idx > 0
        If mHashes(idx) = h Then
            If KeysEqual(mKeys(idx), vKey) Then
                FindSlot = idx
                Exit Function
            End If
        End If
        prevIdx = idx
        idx = mNextIdx(idx)
    Loop
End Function

Private Function KeysEqual(ByRef a As Variant, ByRef B As Variant) As Boolean
    Dim aIsStr As Boolean
    Dim bIsStr As Boolean

    aIsStr = (VarType(a) = vbString)
    bIsStr = (VarType(B) = vbString)
    If aIsStr And bIsStr Then
        KeysEqual = (StrComp(a, B, vbTextCompare) = 0)
    ElseIf (Not aIsStr) And (Not bIsStr) Then
        KeysEqual = (CDbl(a) = CDbl(B))
    End If
End Function

Private Sub GrowEntries()
    mCapacity = mCapacity * 2
    ReDim Preserve mNextIdx(1 To mCapacity)
    ReDim Preserve mKeys(1 To mCapacity)
    ReDim Preserve mHashes(1 To mCapacity)
    ReDim Preserve mFlags(1 To mCapacity)
End Sub

Private Sub GrowBuckets()
    mMask = (mMask + 1) * 2 - 1
    RebuildBuckets
End Sub

Private Sub RebuildBuckets()
    Dim i As Long
    Dim B As Long

    ReDim mBuckets(0 To mMask)
    For i = 1 To mCount
        B = mHashes(i) And mMask
        mNextIdx(i) = mBuckets(B)
        mBuckets(B) = i
    Next i
End Sub

Private Function NextPow2(ByVal N As Long) As Long
    Dim p As Long
    p = INITIAL_BUCKETS
    Do While p < N
        p = p * 2
    Loop
    NextPow2 = p
End Function
