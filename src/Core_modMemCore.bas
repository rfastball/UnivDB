Attribute VB_Name = "Core_modMemCore"
' @desc 메모리 API 선언과 안전 래퍼 제공 담당
' @depends Core_Enums
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' 프로젝트에서 유일하게 메모리 API를 선언/실행하는 통제 모듈.
' 모든 진입점은 0 포인터와 0 이하 길이를 무동작으로 거른다 (상시 안전핀).

' 해시 시드와 24비트 Long 믹서: h <= &HFFFFFF 이면 h*33 <= 553M < 2^31 이라
' 오버플로(에러 6) 없이 순수 Long 연산만 사용. 결과는 항상 [0, 2^24) 비음수.
Private Const HASH_SEED_L As Long = 5381

' LSet UDT 비트캐스트용: Double의 8바이트를 DLL 호출 없이 Long 두 개로 읽는다
Private Type TDblBits
    d As Double
End Type
Private Type TLngPair
    lo As Long
    hi As Long
End Type

#If VBA7 Then

Private Declare PtrSafe Function LCMapStringW Lib "kernel32" (ByVal locale As Long, ByVal flags As Long, ByVal src As LongPtr, ByVal srcLen As Long, ByVal dest As LongPtr, ByVal destLen As Long) As Long
Private Declare PtrSafe Function GetThreadLocale Lib "kernel32" () As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal byteLen As LongPtr)

' @desc 대상/원본 포인터와 길이 받아 복사함, 0 포인터나 0 이하 길이 시 무동작
Public Sub Core_SafeMemCopy(ByVal destPtr As LongPtr, ByVal srcPtr As LongPtr, ByVal byteLen As Long)
    If destPtr = 0 Then Exit Sub
    If srcPtr = 0 Then Exit Sub
    If byteLen <= 0 Then Exit Sub
    RtlMoveMemory destPtr, srcPtr, byteLen
End Sub

' @desc 문자열 받아 BSTR 주소와 바이트 길이 반환, 빈 문자열 시 False
Public Function Core_GetBSTRInfo(ByRef s As String, ByRef pStr As LongPtr, ByRef byteLen As Long) As Boolean
    pStr = StrPtr(s)
    byteLen = LenB(s)
    If pStr = 0 Then byteLen = 0
    Core_GetBSTRInfo = (byteLen > 0)
End Function

' @desc 문자열 받아 BSTR 길이 프리픽스 반환, 빈 문자열 시 0 (진단 전용)
Public Function Core_ReadBSTRPrefix(ByRef s As String) As Long
    Dim p As LongPtr
    Dim n As Long
    p = StrPtr(s)
    If p = 0 Then Exit Function
    If LenB(s) = 0 Then Exit Function
    Core_SafeMemCopy VarPtr(n), p - 4, 4&
    Core_ReadBSTRPrefix = n
End Function

' @desc 문자열 받아 Long 배열에 1회 복사 후 Long 개수 반환, 빈 문자열 시 0
Public Function Core_BlitStringToLongs(ByRef s As String, ByRef outArr() As Long) As Long
    Dim pStr As LongPtr
    Dim byteLen As Long
    Dim nLongs As Long
    If Not Core_GetBSTRInfo(s, pStr, byteLen) Then Exit Function
    nLongs = (byteLen + 3) \ 4
    ReDim outArr(0 To nLongs - 1)
    Core_SafeMemCopy VarPtr(outArr(0)), pStr, byteLen
    Core_BlitStringToLongs = nLongs
End Function

' @desc Double 받아 하위/상위 4바이트를 Long 두 개로 복사함
Public Sub Core_BlitDoubleToLongs(ByVal d As Double, ByRef lo As Long, ByRef hi As Long)
    Core_SafeMemCopy VarPtr(lo), VarPtr(d), 4&
    Core_SafeMemCopy VarPtr(hi), VarPtr(d) + 4, 4&
End Sub

#Else

Private Declare Function LCMapStringW Lib "kernel32" (ByVal locale As Long, ByVal flags As Long, ByVal src As Long, ByVal srcLen As Long, ByVal dest As Long, ByVal destLen As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSrc As Long, ByVal byteLen As Long)

' @desc 대상/원본 포인터와 길이 받아 복사함, 0 포인터나 0 이하 길이 시 무동작
Public Sub Core_SafeMemCopy(ByVal destPtr As Long, ByVal srcPtr As Long, ByVal byteLen As Long)
    If destPtr = 0 Then Exit Sub
    If srcPtr = 0 Then Exit Sub
    If byteLen <= 0 Then Exit Sub
    RtlMoveMemory destPtr, srcPtr, byteLen
End Sub

' @desc 문자열 받아 BSTR 주소와 바이트 길이 반환, 빈 문자열 시 False
Public Function Core_GetBSTRInfo(ByRef s As String, ByRef pStr As Long, ByRef byteLen As Long) As Boolean
    pStr = StrPtr(s)
    byteLen = LenB(s)
    If pStr = 0 Then byteLen = 0
    Core_GetBSTRInfo = (byteLen > 0)
End Function

' @desc 문자열 받아 BSTR 길이 프리픽스 반환, 빈 문자열 시 0 (진단 전용)
Public Function Core_ReadBSTRPrefix(ByRef s As String) As Long
    Dim p As Long
    Dim n As Long
    p = StrPtr(s)
    If p = 0 Then Exit Function
    If LenB(s) = 0 Then Exit Function
    Core_SafeMemCopy VarPtr(n), p - 4, 4&
    Core_ReadBSTRPrefix = n
End Function

' @desc 문자열 받아 Long 배열에 1회 복사 후 Long 개수 반환, 빈 문자열 시 0
Public Function Core_BlitStringToLongs(ByRef s As String, ByRef outArr() As Long) As Long
    Dim pStr As Long
    Dim byteLen As Long
    Dim nLongs As Long
    If Not Core_GetBSTRInfo(s, pStr, byteLen) Then Exit Function
    nLongs = (byteLen + 3) \ 4
    ReDim outArr(0 To nLongs - 1)
    Core_SafeMemCopy VarPtr(outArr(0)), pStr, byteLen
    Core_BlitStringToLongs = nLongs
End Function

' @desc Double 받아 하위/상위 4바이트를 Long 두 개로 복사함
Public Sub Core_BlitDoubleToLongs(ByVal d As Double, ByRef lo As Long, ByRef hi As Long)
    Core_SafeMemCopy VarPtr(lo), VarPtr(d), 4&
    Core_SafeMemCopy VarPtr(hi), VarPtr(d) + 4, 4&
End Sub

#End If

' ----- 해시 함수 (인스턴스 무관 순수 함수 - 엔진 간 키 이동 시 일관성 보장) -----
' StrComp(vbTextCompare)의 동등 키는 같은 해시여야 한다.
' 문자별 UCase/길이 해시는 합성 문자, 전각 문자, 무시되는 문자에서 이 불변식을 깬다.
' Windows sort key는 대소문자/문자 폭/가나 종류를 무시한다. 추가 충돌은 KeysEqual이 판별한다.

' @desc 키 문자열 받아 로캘 텍스트 동등성에 맞는 해시 반환, API 실패 시 9999
Public Function Core_HashKeyText(ByRef s As String) As Long
    Const SORT_FLAGS As Long = &H30401 ' SORTKEY | IGNORECASE | IGNOREKANATYPE | IGNOREWIDTH
    Dim bytes() As Byte, source As String
    Dim n As Long, i As Long, h As Long, locale As Long

    source = s
    If Len(source) = 0 Then source = vbNullChar
    locale = GetThreadLocale()
    n = LCMapStringW(locale, SORT_FLAGS, StrPtr(source), Len(source), 0, 0)
    If n = 0 Then Err.Raise 9999, "Core_HashKeyText", "Cannot size text sort key"
    ReDim bytes(0 To n - 1)
    If LCMapStringW(locale, SORT_FLAGS, StrPtr(source), Len(source), VarPtr(bytes(0)), n) = 0 Then
        Err.Raise 9999, "Core_HashKeyText", "Cannot create text sort key"
    End If
    h = HASH_SEED_L
    For i = 0 To n - 1
        h = ((h * 33) Xor bytes(i)) And &HFFFFFF
    Next i
    Core_HashKeyText = h
End Function
' @desc 숫자 키 받아 해시 반환, -0은 +0으로 정규화 후 계산
Public Function Core_HashKeyNumber(ByVal d As Double) As Long
    Dim dbits As TDblBits
    Dim pair As TLngPair
    Dim h As Long

    If d = 0 Then d = 0
    dbits.d = d
    LSet pair = dbits
    h = ((HASH_SEED_L * 33) Xor (pair.lo And &HFFFF&)) And &HFFFFFF
    h = ((h * 33) Xor ((pair.lo \ &H10000) And &HFFFF&)) And &HFFFFFF
    h = ((h * 33) Xor (pair.hi And &HFFFF&)) And &HFFFFFF
    Core_HashKeyNumber = ((h * 33) Xor ((pair.hi \ &H10000) And &HFFFF&)) And &HFFFFFF
End Function
