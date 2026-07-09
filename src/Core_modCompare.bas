Attribute VB_Name = "Core_modCompare"
' @desc Core 값 비교 reference 함수 제공
' @depends Core_modNorm
Option Explicit

' @desc 두 Variant 값을 Core 기본 정책으로 비교해 -1/0/1 반환
Public Function Core_CompareValues(ByVal vLeft As Variant, ByVal vRight As Variant, _
                                   Optional ByVal compareMode As VbCompareMethod = vbTextCompare) As Long
    If Core_IsErrorLike(vLeft) Or Core_IsErrorLike(vRight) Then
        Core_CompareValues = Core_CompareText(Core_SafeText(vLeft), Core_SafeText(vRight), compareMode)
    ElseIf Core_IsDateValue(vLeft) And Core_IsDateValue(vRight) Then
        Core_CompareValues = Core_CompareDouble(CDbl(vLeft), CDbl(vRight))
    ElseIf Core_IsNumericValue(vLeft) And Core_IsNumericValue(vRight) Then
        Core_CompareValues = Core_CompareDouble(CDbl(vLeft), CDbl(vRight))
    Else
        Core_CompareValues = Core_CompareText(Core_SafeText(vLeft), Core_SafeText(vRight), compareMode)
    End If
End Function

Private Function Core_CompareDouble(ByVal dLeft As Double, ByVal dRight As Double) As Long
    If dLeft < dRight Then
        Core_CompareDouble = -1
    ElseIf dLeft > dRight Then
        Core_CompareDouble = 1
    Else
        Core_CompareDouble = 0
    End If
End Function

Private Function Core_CompareText(ByVal sLeft As String, ByVal sRight As String, _
                                  ByVal compareMode As VbCompareMethod) As Long
    Core_CompareText = Sgn(StrComp(sLeft, sRight, compareMode))
End Function

Private Function Core_IsErrorLike(ByVal vValue As Variant) As Boolean
    Core_IsErrorLike = IsError(vValue)
End Function

Private Function Core_IsDateValue(ByVal vValue As Variant) As Boolean
    Core_IsDateValue = (VarType(vValue) = vbDate)
End Function

Private Function Core_IsNumericValue(ByVal vValue As Variant) As Boolean
    Select Case VarType(vValue)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            Core_IsNumericValue = True
    End Select
End Function
