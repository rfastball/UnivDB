Attribute VB_Name = "modIndexKey"
' @desc 해시 인덱스 백엔드 공통 키 정규화 담당
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' 백엔드 독립 키 계약 (모든 DbHashIndex 구현이 공유):
' - Empty -> "" (빈 문자열 키)
' - Date -> canonical ISO: "yyyy-mm-dd" (시간 성분 0) / "yyyy-mm-dd hh:nn:ss" - locale 무관 (D3)
' - 그 외 CStr 문자열화
' - IsError 값은 정규화 대상 아님 - 호출부(AddRow/TryGetRows)가 사전에 skip
' - 비교는 대소문자/폭 무시. Dict 백엔드: FoldIndexKey 접기 + binary 비교 (2026-10-31 하이브리드 접기),
'   Core 백엔드: vbTextCompare 등가. 두 백엔드는 미할당/서식 제어 문자와 반각 가나 탁음 결합에서만 다르다

' 인덱스 백엔드 선택자 (clsTable.IndexBackend) - 자동 선택 없음, 명시적 옵션 (Phase 8 정책)
Public Enum DbIndexBackend
    dbBackendDict = 0 ' Scripting.Dictionary 기반 (기본)
    dbBackendCore = 1 ' Core_clsFastHash 기반 (실험)
End Enum

' 키 접기 (FoldIndexKey) - UnivFilter 벤더링 파일이라 Core_*에 의존하지 않고 여기서 직접 선언
#If VBA7 Then
Private Declare PtrSafe Function LCMapStringW Lib "kernel32" (ByVal locale As Long, ByVal flags As Long, ByVal src As LongPtr, ByVal srcLen As Long, ByVal dest As LongPtr, ByVal destLen As Long) As Long
Private Declare PtrSafe Function GetThreadLocale Lib "kernel32" () As Long
#Else
Private Declare Function LCMapStringW Lib "kernel32" (ByVal locale As Long, ByVal flags As Long, ByVal src As Long, ByVal srcLen As Long, ByVal dest As Long, ByVal destLen As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
#End If

Private Const LCMAP_UPPERCASE As Long = &H200&
Private Const LCMAP_HALFWIDTH As Long = &H400000
Private mFoldBuf As String ' LCMapStringW 출력 버퍼 (재사용)

' @desc 인덱스 키 값 받아 정규화된 문자열 키 반환, Empty 시 빈 문자열
Public Function NormalizeIndexKey(ByVal key As Variant) As String
    If IsEmpty(key) Then
        NormalizeIndexKey = ""
    ElseIf VarType(key) = vbDate Then
        ' 날짜 canonical 키 (D3): CStr(date)는 locale 의존이라 세션 밖 키 재현성이 없다.
        ' 시간 성분 0이면 날짜만 - GroupBy 렌더 표기의 00:00:00 노이즈 방지 (단사성 유지).
        If CDbl(key) = Int(CDbl(key)) Then
            NormalizeIndexKey = Format$(key, "yyyy-mm-dd")
        Else
            NormalizeIndexKey = Format$(key, "yyyy-mm-dd hh:nn:ss")
        End If
    Else
        NormalizeIndexKey = CStr(key)
    End If
End Function

' 복합 키 인코딩 (R3): 길이 접두 "len:value" 조각을 "|"로 연결한다.
' 길이가 파싱 권위이므로 값에 "|"나 ":"가 있어도 충돌 불가 (이스케이프 불필요).
' 접두 숫자와 구분 문자는 대소문자가 없어 vbTextCompare 아래서도 단사성이 유지된다.
' Empty는 NormalizeIndexKey 계약대로 "0:" 조각이 되고, Error 값은 호출부가 행 skip.

' @desc 키 파트 받아 길이 접두 인코딩 조각 반환, 복합 키 조립용
Public Function EncodeCompositeKeyPart(ByVal v As Variant) As String
    Dim s As String
    s = NormalizeIndexKey(v)
    EncodeCompositeKeyPart = CStr(Len(s)) & ":" & s
End Function

' 정규화 키의 저장/비교용 접기 (Dict 백엔드, GroupBy 복합 키, Upsert 삽입 키 공통).
' 폭/가나 영역(U+3000-30FF, U+3131-318E, U+FF00-FFFF) 코드 유닛이 있을 때만
' LCMapStringW(UPPERCASE|HALFWIDTH) - 전각 라틴/숫자/공백, 반각 한글 자모, 반각 가나를 접는다 -
' 아니면 UCase$. 판별은 UTF-16 상위 바이트 검사. 원본 표기 보존은 호출부 책임.
' @desc 정규화 키 받아 대소문자/폭 접기 키 반환, 빈 문자열 시 빈 문자열
Public Function FoldIndexKey(ByRef s As String) As String
    Dim b() As Byte, i As Long, n As Long

    If Len(s) = 0 Then Exit Function
    b = s
    For i = 1 To UBound(b) Step 2
        Select Case b(i)
            Case &H30, &HFF: GoTo NeedsWidthFold
            Case &H31: If b(i - 1) >= &H31 And b(i - 1) <= &H8E Then GoTo NeedsWidthFold
        End Select
    Next i
    FoldIndexKey = UCase$(s)
    Exit Function

NeedsWidthFold:
    n = Len(s) * 2 + 16 ' 반각화는 길이를 늘릴 수 있다 (탁음 가나 1자 -> 반각 2자)
    If Len(mFoldBuf) < n Then mFoldBuf = String$(n, vbNullChar)
    n = LCMapStringW(GetThreadLocale(), LCMAP_UPPERCASE Or LCMAP_HALFWIDTH, StrPtr(s), Len(s), StrPtr(mFoldBuf), Len(mFoldBuf))
    If n = 0 Then Err.Raise 9999, "modIndexKey.FoldIndexKey", "LCMapStringW failed"
    FoldIndexKey = Left$(mFoldBuf, n)
End Function
