Attribute VB_Name = "modIndexKey"
' @desc 해시 인덱스 백엔드 공통 키 정규화 담당
Option Explicit

' 백엔드 독립 키 계약 (모든 DbHashIndex 구현이 공유):
' - Empty -> "" (빈 문자열 키)
' - 그 외 CStr 문자열화
' - IsError 값은 정규화 대상 아님 - 호출부(AddRow/TryGetRows)가 사전에 skip
' - 비교는 대소문자 무시 (vbTextCompare) - 백엔드의 저장소 설정이 보장

' @desc 인덱스 키 값 받아 정규화된 문자열 키 반환, Empty 시 빈 문자열
Public Function NormalizeIndexKey(ByVal key As Variant) As String
    If IsEmpty(key) Then
        NormalizeIndexKey = ""
    Else
        NormalizeIndexKey = CStr(key)
    End If
End Function
