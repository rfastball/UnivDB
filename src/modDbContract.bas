Attribute VB_Name = "modDbContract"
' @desc 파생 구조의 테이블-세대 결속 계약 검증 담당
' @depends clsTable
Option Explicit

' 세대 계약 위반(stale 파생 구조 사용) 전용 에러 번호 (일반 에러 9999와 구분)
Public Const DB_ERR_STALE_GENERATION As Long = 9998

' @desc 결속 스냅샷과 테이블 받아 세대 일치 검증함, 불일치 시 9998 에러
Public Sub AssertCurrentGeneration(ByVal ownerPtr As LongPtr, ByVal boundGen As Long, _
                                   ByVal tbl As clsTable, ByVal sSource As String)
    ' 미결속(ownerPtr=0)은 계약 밖 사용 - 검사 통과
    If ownerPtr = 0 Then Exit Sub

    If tbl Is Nothing Then
        Err.Raise DB_ERR_STALE_GENERATION, sSource, _
            "Stale derived structure: bound table reference missing"
    End If

    If ObjPtr(tbl) <> ownerPtr Or tbl.Generation <> boundGen Then
        Err.Raise DB_ERR_STALE_GENERATION, sSource, _
            "Stale derived structure: table generation changed"
    End If
End Sub
