Attribute VB_Name = "Core_modNorm"
' @desc Variant 값을 안전한 문자열 키로 정규화
Option Explicit

' @desc Variant 값을 오류 없이 표시 문자열로 변환해 반환
Public Function Core_SafeText(ByVal vValue As Variant) As String
    If IsError(vValue) Then
        Core_SafeText = "#ERROR"
    ElseIf IsNull(vValue) Then
        Core_SafeText = vbNullString
    ElseIf IsEmpty(vValue) Then
        Core_SafeText = vbNullString
    Else
        Core_SafeText = CStr(vValue)
    End If
End Function

' @desc Variant 값을 대소문자 무시 비교용 키 문자열로 변환해 반환
Public Function Core_SafeKey(ByVal vValue As Variant) As String
    Core_SafeKey = UCase$(Trim$(Core_SafeText(vValue)))
End Function
