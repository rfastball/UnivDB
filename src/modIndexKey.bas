Attribute VB_Name = "modIndexKey"
' @desc 해시 인덱스 백엔드 공통 키 정규화 담당
Option Explicit

' 백엔드 독립 키 계약 (모든 DbHashIndex 구현이 공유):
' - Empty -> "" (빈 문자열 키)
' - Date -> canonical ISO: "yyyy-mm-dd" (시간 성분 0) / "yyyy-mm-dd hh:nn:ss" - locale 무관 (D3)
' - 그 외 CStr 문자열화
' - IsError 값은 정규화 대상 아님 - 호출부(AddRow/TryGetRows)가 사전에 skip
' - 비교는 대소문자 무시 (vbTextCompare) - 백엔드의 저장소 설정이 보장

' 인덱스 백엔드 선택자 (clsTable.IndexBackend) - 자동 선택 없음, 명시적 옵션 (Phase 8 정책)
Public Enum DbIndexBackend
    dbBackendDict = 0 ' Scripting.Dictionary 기반 (기본)
    dbBackendCore = 1 ' Core_clsFastHash 기반 (실험)
End Enum

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
