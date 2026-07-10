Attribute VB_Name = "modOperators"
' @desc 연산자/집계함수 문자열 화이트리스트 검증 담당
Option Explicit

' 검증 계약 (audit 2026-07-10, V1/V6): 연산자/집계함수 문자열을 받는 모든 표면은
' 이 모듈의 단일 정규화 지점을 경유한다. 미지 문자열은 즉시 Err.Raise 9999 -
' 조용한 0행/무동작/빈 컬럼(관용 모드)을 금지한다. 컬럼명 오타의 기존
' Err.Raise 9999와 같은 코드로 오류 처리 대칭(V2)을 맞춘다.

' @desc Filter 연산자 받아 정규화 문자열 반환, 미지 연산자 시 Err.Raise 9999
Public Function NormalizeFilterOp(ByVal sOp As String) As String
    Dim s As String
    s = UCase(Trim$(sOp))
    Select Case s
        Case "=", "<>", ">", "<", ">=", "<=", "IS", "IS NOT", "LIKE", "NOT LIKE"
            NormalizeFilterOp = s
        Case Else
            Err.Raise 9999, "modOperators.NormalizeFilterOp", _
                "Unknown filter operator: '" & sOp & "' (allowed: = <> > < >= <= IS [NOT] / [NOT] LIKE)"
    End Select
End Function

' @desc 집계함수명 받아 정규화 문자열 반환, 미지 함수 시 Err.Raise 9999
Public Function NormalizeAggFunc(ByVal sFunc As String) As String
    Dim s As String
    s = UCase(Trim$(sFunc))
    Select Case s
        Case "SUM", "AVG", "COUNT", "MIN", "MAX"
            NormalizeAggFunc = s
        Case Else
            Err.Raise 9999, "modOperators.NormalizeAggFunc", _
                "Unknown aggregate function: '" & sFunc & "' (allowed: SUM AVG COUNT MIN MAX)"
    End Select
End Function

' @desc 산술/결합 연산자 받아 정규화 문자열 반환, 미지 연산자 시 Err.Raise 9999
Public Function NormalizeComputeOp(ByVal sOp As String) As String
    Dim s As String
    s = Trim$(sOp)
    Select Case s
        Case "+", "-", "*", "/", "&"
            NormalizeComputeOp = s
        Case Else
            Err.Raise 9999, "modOperators.NormalizeComputeOp", _
                "Unknown compute operator: '" & sOp & "' (allowed: + - * / &)"
    End Select
End Function

' @desc 비교 연산자 받아 정규화 문자열 반환, 미지 연산자 시 Err.Raise 9999
Public Function NormalizeCompareOp(ByVal sOp As String) As String
    Dim s As String
    s = Trim$(sOp)
    Select Case s
        Case "=", "<>", ">", ">=", "<", "<="
            NormalizeCompareOp = s
        Case Else
            Err.Raise 9999, "modOperators.NormalizeCompareOp", _
                "Unknown compare operator: '" & sOp & "' (allowed: = <> > >= < <=)"
    End Select
End Function
