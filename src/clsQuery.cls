VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' ==============================================================================
' Class: clsQuery
' 설명: clsTable에 대한 조건 검색을 최적화된 순서(Eq -> Range)로 수행
' 전략: 인덱스를 사용하는 동등조건(Where)을 먼저 수행하여 스캔 범위를 줄임
' ==============================================================================

Private m_SourceTable As clsTable
Private m_setCandidates As DbRowSet ' 현재 살아남은 후보 행 번호(RowID) 집합
Private m_bIsFirstCondition As Boolean
Private m_dicUpdates As Object       ' 업데이트할 컬럼과 값 저장 (Set 절)

' [Projection & Sort & Limit]
Private m_sTargetColumns As String   ' 출력할 컬럼 목록 (쉼표 구분)
Private m_sSortColumn As String      ' 정렬할 컬럼명
Private m_bSortDesc As Boolean       ' 내림차순 여부
Private m_lLimit As Long             ' 반환할 최대 행 개수 (0이면 제한 없음)

Private Sub Class_Initialize()
    ' 기본 상태 설정
    m_bIsFirstCondition = True
    Set m_dicUpdates = CreateObject("Scripting.Dictionary")
    
    ' 기능별 기본값 초기화
    m_sTargetColumns = "*"           ' 기본: 전체 컬럼
    m_sSortColumn = ""               ' 기본: 정렬 없음
    m_bSortDesc = False              ' 기본: 오름차순
    m_lLimit = 0                     ' 기본: 개수 제한 없음
End Sub

Private Sub Class_Terminate()
    Set m_SourceTable = Nothing
    Set m_setCandidates = Nothing
    Set m_dicUpdates = Nothing
End Sub

' ------------------------------------------------------------------------------
' [Step 0] 쿼리 대상 테이블 설정
' ------------------------------------------------------------------------------
Public Function From(ByVal tTable As clsTable) As clsQuery
    Set m_SourceTable = tTable
    Set m_setCandidates = Nothing
    m_bIsFirstCondition = True
    m_dicUpdates.RemoveAll
    Set From = Me
End Function

' ------------------------------------------------------------------------------
' [Step 1] 동등 조건 (Equality) - 인덱스 활용 (Fast)
' 설명: 가장 먼저 호출하여 후보군을 대폭 줄이는 것을 권장
' ------------------------------------------------------------------------------
Public Function Where(ByVal sColName As String, ByVal vValue As Variant) As clsQuery
    Dim idx As DbHashIndex
    Dim rows As DbRowList

    ' 0. 세대 계약: 이전 조건의 후보군이 stale이면 즉시 실패
    EnsureCandidatesCurrent

    ' 1. 내부 HashIndex에서 행 번호 추출 (O(1))
    RequireColumn sColName, "clsQuery.Where"
    Set idx = m_SourceTable.GetHashIndex(sColName)

    ' 2. DbRowList를 후보 RowSet에 반영
    If Not idx Is Nothing Then
        If idx.TryGetRows(vValue, rows) Then
            If m_bIsFirstCondition Then
                Set m_setCandidates = New DbRowSet
                m_setCandidates.BindTo m_SourceTable
                m_setCandidates.LoadFromRowList rows
                m_bIsFirstCondition = False
            Else
                m_setCandidates.IntersectWithRowList rows
            End If
        Else
            Set m_setCandidates = New DbRowSet
            m_setCandidates.BindTo m_SourceTable
            m_bIsFirstCondition = False
        End If
    ElseIf m_bIsFirstCondition Then
        Set m_setCandidates = New DbRowSet
        m_setCandidates.BindTo m_SourceTable
        m_bIsFirstCondition = False
    Else
        m_setCandidates.Clear
    End If

    Set Where = Me
End Function

' ==============================================================================
' [Scanning Filters] - 동작 확인된 AndFilter의 유전자를 그대로 이식
' ==============================================================================

' 1. 기본 필터 (v2.5 JIT Architecture)
' 설명: 연산자(sOp)에 따라 수치/문자 경로를 분리하고, 모든 전처리를 루프 외부로 격리 (Hoisting).
'       루프 내부의 함수 호출 및 메모리 할당을 제거하여 성능을 극대화.
'       스캔 코어는 EvaluateArm(단일 조건 공용 헬퍼)에 있다 - FilterAny와 공유.
Public Function Filter(ByVal sColName As String, ByVal sOp As String, ByVal vValue As Variant) As clsQuery
    Dim vCandidates As Variant

    ' 연산자 화이트리스트 검증 (audit F1, V1/V2): 미지 연산자는 후보 유무와 무관하게
    ' 즉시 raise - 컬럼명 오타(아래 9999)와 오류 처리 대칭. 정규화도 이 단일 지점에서.
    sOp = modOperators.NormalizeFilterOp(sOp)
    ValidateFilterArm sColName, sOp, vValue, "clsQuery.Filter"

    EnsureCandidatesCurrent
    If m_bIsFirstCondition Then InitializeAllRows
    If m_setCandidates.Count = 0 Then Set Filter = Me: Exit Function

    vCandidates = m_setCandidates.ToArray
    Set m_setCandidates = EvaluateArm(sColName, sOp, vValue, vCandidates, "clsQuery.Filter")
    m_bIsFirstCondition = False
    Set Filter = Me
End Function

' 단일 조건 스캔 코어 (Filter/FilterAny 공용): 컬럼 존재/수치 비교값 검증과
' 수치/문자 dual-path 루프를 한 곳에 보유. 전제: sOp는 NormalizeFilterOp 정규화 완료.
' vCandidates(후보 rowId 0-based 배열)는 ByRef로 받아 복사를 회피한다 (읽기 전용).
Private Function EvaluateArm(ByVal sColName As String, ByVal sOp As String, _
                             ByVal vValue As Variant, ByRef vCandidates As Variant, _
                             ByVal sSource As String) As DbRowSet
    Dim setSurvivors As New DbRowSet
    Dim i As Long

    ' 1. [Hoisting] 루프 외부에서 모든 전처리 완료
    Dim rdr As DbColumnReader
    Dim vCol As Variant

    ' 1-1. 컬럼 접근 표면 획득 (Phase 6.8 밀봉 - 레이아웃 비의존)
    Set rdr = m_SourceTable.GetColumnReader(sColName)
    If rdr Is Nothing Then Err.Raise 9999, sSource, "Column not found: " & sColName

    setSurvivors.BindTo m_SourceTable

    ' 1-2. 컬럼 1D 배열 로컬 변수에 할당 (루프 내부는 직접 인덱싱 유지)
    vCol = rdr.Values

    ' 2. [Dual Path] 연산자 종류에 따라 실행 경로 분기 (sOp는 호출자에서 정규화 완료)
    Select Case sOp
        Case "=", "<>", ">", "<", ">=", "<="
            ' ==========================================================
            ' [Path A] 수치 연산 경로 (Numeric Path)
            ' ==========================================================
            Dim dTarget As Double
            Dim vCell As Variant

            dTarget = CDbl(vValue)
            
            ' 최적화된 루프
            For i = LBound(vCandidates) To UBound(vCandidates)
                vCell = vCol(vCandidates(i))
                ' Empty=NULL 계약(D1): IsNumeric(Empty)=True라 0으로 참여하던 것을 차단.
                ' Error/Date는 이미 IsNumeric=False로 제외 - 인덱스 계약과 정합.
                If Not IsEmpty(vCell) And IsNumeric(vCell) Then
                    Select Case sOp
                        Case "=": If CDbl(vCell) = dTarget Then setSurvivors.AddRow vCandidates(i)
                        Case "<>": If CDbl(vCell) <> dTarget Then setSurvivors.AddRow vCandidates(i)
                        Case ">": If CDbl(vCell) > dTarget Then setSurvivors.AddRow vCandidates(i)
                        Case "<": If CDbl(vCell) < dTarget Then setSurvivors.AddRow vCandidates(i)
                        Case ">=": If CDbl(vCell) >= dTarget Then setSurvivors.AddRow vCandidates(i)
                        Case "<=": If CDbl(vCell) <= dTarget Then setSurvivors.AddRow vCandidates(i)
                    End Select
                End If
            Next i
            
        Case "IS", "IS NOT", "LIKE", "NOT LIKE"
            ' ==========================================================
            ' [Path B] 문자열 연산 경로 (String Path)
            ' ==========================================================
            Dim sTarget As String
            sTarget = FastNorm(vValue) ' FastNorm으로 비교값 사전 정규화
            
            ' Error=skip 계약(D1): Error 셀은 어느 연산에서도 결과에 참여하지 않는다.
            ' 부정 연산(IS NOT/NOT LIKE)에서도 "not equal"로 걸리지 않고 skip - raw CStr(Error)의
            ' locale 문자열 은밀 참여를 차단. (Empty는 CStr(Empty)="" 로 IS "" 매칭 유지)
            Select Case sOp
                Case "IS" ' StrComp 기반의 대소문자 무시 비교 (=)
                    For i = LBound(vCandidates) To UBound(vCandidates)
                        vCell = vCol(vCandidates(i))
                        If Not IsError(vCell) Then
                            If StrComp(FastNorm(vCell), sTarget, vbTextCompare) = 0 Then setSurvivors.AddRow vCandidates(i)
                        End If
                    Next i
                Case "IS NOT" ' StrComp 기반의 대소문자 무시 비교 (<>)
                    For i = LBound(vCandidates) To UBound(vCandidates)
                        vCell = vCol(vCandidates(i))
                        If Not IsError(vCell) Then
                            If StrComp(FastNorm(vCell), sTarget, vbTextCompare) <> 0 Then setSurvivors.AddRow vCandidates(i)
                        End If
                    Next i
                Case "LIKE" ' VBA Like 연산자
                    sTarget = UCase(sTarget) ' Like는 vbTextCompare를 지원 안하므로 UCase 필요
                    For i = LBound(vCandidates) To UBound(vCandidates)
                        vCell = vCol(vCandidates(i))
                        If Not IsError(vCell) Then
                            If UCase(FastNorm(vCell)) Like sTarget Then setSurvivors.AddRow vCandidates(i)
                        End If
                    Next i
                Case "NOT LIKE"
                     sTarget = UCase(sTarget)
                     For i = LBound(vCandidates) To UBound(vCandidates)
                        vCell = vCol(vCandidates(i))
                        If Not IsError(vCell) Then
                            If Not (UCase(FastNorm(vCell)) Like sTarget) Then setSurvivors.AddRow vCandidates(i)
                        End If
                    Next i
            End Select
    End Select

    Set EvaluateArm = setSurvivors
End Function

' ==============================================================================
' [Step 1.5] FilterAny - OR 그룹 (교차 컬럼 OR, 표현력 확장 C-1)
' 의미론: 한 호출 = 하나의 OR 그룹. 그룹 안 arm들은 OR, 그룹과 체인의 나머지는 AND.
'   결과 = (arm1 OR arm2 OR ...) AND 이전 체인. FilterAny 여러 번 호출은 그룹 간 AND.
' arm 형식: Array(컬럼, 연산자, 값) 3요소. 연산자는 Filter와 같은 화이트리스트.
'   개별 arm 인자 외에 arm 배열 1개(동적 목록, KeepAny 위임)도 허용 - 재귀 평탄화.
' 검증: 전체 arm을 평가 전에 문장 수준 검증 (audit F3 교훈 - 부분 축소 없이 9999).
'   Filter와 동일하게 후보 0행 early-out보다 검증이 앞선다.
' 성능: arm당 컬럼 사본 1회 (Filter N회 체인과 동일 비용 프로파일).
' ==============================================================================

' @desc 조건 arm들 받아 하나라도 맞는 행 유지 (OR 그룹, 체인 나머지와는 AND)
Public Function FilterAny(ParamArray vArms() As Variant) As clsQuery
    Dim colArms As New Collection
    Dim vItem As Variant, vArm As Variant
    Dim k As Long, lb As Long
    Dim lArmCount As Long
    Dim sCols() As String, sOps() As String, vVals() As Variant
    Dim vCandidates As Variant
    Dim setUnion As DbRowSet
    Dim setArm As DbRowSet

    ' 1. arm 평탄화 (개별 Array 인자 / arm 배열 혼용 허용)
    If UBound(vArms) < LBound(vArms) Then
        Err.Raise 9999, "clsQuery.FilterAny", "FilterAny requires at least one Array(column, op, value) arm"
    End If
    For Each vItem In vArms
        FlattenArms vItem, colArms
    Next vItem
    lArmCount = colArms.Count
    If lArmCount = 0 Then
        Err.Raise 9999, "clsQuery.FilterAny", "FilterAny requires at least one Array(column, op, value) arm"
    End If

    ' 2. 문장 수준 검증 - 전체 arm을 평가 앞에서 (미지 연산자/컬럼 오타/비숫자 비교값)
    ReDim sCols(1 To lArmCount): ReDim sOps(1 To lArmCount): ReDim vVals(1 To lArmCount)
    k = 0
    For Each vArm In colArms
        k = k + 1
        ValidateArmShape vArm
        lb = LBound(vArm)
        If IsObject(vArm(lb)) Or IsArray(vArm(lb)) _
           Or IsObject(vArm(lb + 1)) Or IsArray(vArm(lb + 1)) _
           Or IsObject(vArm(lb + 2)) Or IsArray(vArm(lb + 2)) Then
            Err.Raise 9999, "clsQuery.FilterAny", "Malformed arm: elements must be scalar (column, op, value)"
        End If
        sCols(k) = Trim(CStr(vArm(lb)))
        sOps(k) = modOperators.NormalizeFilterOp(CStr(vArm(lb + 1)))
        vVals(k) = vArm(lb + 2)

        ValidateFilterArm sCols(k), sOps(k), vVals(k), "clsQuery.FilterAny"
    Next vArm

    ' 3. 실행: 같은 후보 스냅샷에 arm별 독립 스캔 -> 합집합 (중복 행은 비트맵이 무계수)
    EnsureCandidatesCurrent
    If m_bIsFirstCondition Then InitializeAllRows
    If m_setCandidates.Count = 0 Then Set FilterAny = Me: Exit Function

    vCandidates = m_setCandidates.ToArray
    For k = 1 To lArmCount
        Set setArm = EvaluateArm(sCols(k), sOps(k), vVals(k), vCandidates, "clsQuery.FilterAny")
        If setUnion Is Nothing Then
            Set setUnion = setArm
        Else
            setUnion.UnionWith setArm
        End If
    Next k

    Set m_setCandidates = setUnion
    m_bIsFirstCondition = False
    Set FilterAny = Me
End Function

' arm 인자 재귀 평탄화: 첫 요소가 배열이면 arm 목록으로 보고 내부로 재귀
' (유효 arm의 첫 요소는 컬럼명 스칼라라 모호성 없음)
Private Sub FlattenArms(ByVal vItem As Variant, ByVal colArms As Collection)
    Dim vSub As Variant
    If IsArmList(vItem) Then
        For Each vSub In vItem
            FlattenArms vSub, colArms
        Next vSub
    Else
        colArms.Add vItem
    End If
End Sub

' arm 목록 판별: 배열이고 첫 요소가 배열이면 True (빈 배열/판독 불가는 False)
Private Function IsArmList(ByVal vItem As Variant) As Boolean
    Dim vFirst As Variant
    If Not IsArray(vItem) Then Exit Function
    On Error Resume Next
    vFirst = vItem(LBound(vItem))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    IsArmList = IsArray(vFirst)
End Function

' arm 형태 검증: 3요소 1D 배열이 아니면 9999 (2D 배열/비배열 포함)
Private Sub ValidateArmShape(ByVal vArm As Variant)
    Dim lProbe As Long
    If Not IsArray(vArm) Then
        Err.Raise 9999, "clsQuery.FilterAny", "Each arm must be Array(column, op, value)"
    End If
    On Error Resume Next
    lProbe = UBound(vArm, 2)
    If Err.Number = 0 Then
        On Error GoTo 0
        Err.Raise 9999, "clsQuery.FilterAny", "Each arm must be a 1D Array(column, op, value)"
    End If
    Err.Clear
    On Error GoTo 0
    If UBound(vArm) - LBound(vArm) <> 2 Then
        Err.Raise 9999, "clsQuery.FilterAny", "Each arm must have exactly 3 elements (column, op, value)"
    End If
End Sub

' ==============================================================================
' [Step 2.7] Join - 조인을 쿼리 체인에 합류 (audit F17)
' 의미론: 물질화 장벽(materialization barrier). Join 시점까지 스테이징된
' 조건/투영/정렬/Top을 ToTable()로 실행해 좌측 테이블을 확정한 뒤, 조인 결과로
' 재소싱해 체인을 초기 상태에서 계속한다 (지연 엔진이 단일 소스 전제라 지연 조인 불가).
' ==============================================================================

' @desc 우측 테이블과 조인 후 결과로 재소싱해 체인 유지, 물질화 장벽 (스테이징 전부 실행)
Public Function Join(ByVal vRight As Variant, ByVal sLeftKey As String, _
                     Optional ByVal jType As joinType = jtInner, _
                     Optional ByVal sRightPrefix As String = "", _
                     Optional ByVal sRightKey As String = "") As clsQuery
    Dim tJoined As clsTable
    ' vRight는 테이블명(String)/clsTable 혼용 - UnivDB.Join의 ResolveTable이 해석.
    ' 기본 jType은 파사드와 정합(jtInner). ReportBuilder.JoinTable의 jtLeft 기본과 다름 주의.
    Set tJoined = UnivDB.Join(Me.ToTable(), vRight, sLeftKey, jType, sRightPrefix, sRightKey)
    Me.From tJoined            ' 후보/업데이트 리셋 (From 계약)
    ResetPipelineState         ' 투영/정렬/Top 기본값 복원 (From이 안 하는 부분)
    Set Join = Me
End Function

' 물질화 장벽 통과 후 파이프라인 상태를 Class_Initialize 기본값으로 복원
Private Sub ResetPipelineState()
    m_sTargetColumns = "*"
    m_sSortColumn = ""
    m_bSortDesc = False
    m_lLimit = 0
End Sub

' ==============================================================================
' [Step 2.5] Projection 설정 (SQL의 SELECT 절)
' 설명: 결과로 출력할 컬럼과 그 순서를 지정합니다. (기본값: "*")===============
' 설명: 출력할 컬럼 목록을 유연하게 입력받아 내부 문자열로 합칩니다.
' 사용법: .SelectColumns("Col1", "Col2") 또는 .SelectColumns("Col1, Col2")
' ------------------------------------------------------------------------------
Public Function SelectColumns(ParamArray vCols() As Variant) As clsQuery
    Dim vItem As Variant, vSubItem As Variant
    Dim sFinal As String, sCol As String
    
    ' 1. 인자가 없는 경우 기본값 "*" 설정
    If UBound(vCols) < 0 Then
        m_sTargetColumns = "*"
        Set SelectColumns = Me: Exit Function
    End If
    
    sFinal = ""
    
    ' 2. 모든 인자를 순회하며 결합
    For Each vItem In vCols
        ' 핵심: 단일 항목이든 배열이든 동일한 로직으로 처리하기 위해 내부 함수처럼 분리
        Call Internal_ProcessItem(vItem, sFinal)
    Next vItem
    
    ' 3. 최종 결과 할당
    m_sTargetColumns = IIf(Len(sFinal) = 0, "*", sFinal)
    Set SelectColumns = Me
End Function

' [Internal] 타입 오류를 방지하며 아이템을 분석하여 sFinal에 추가
Private Sub Internal_ProcessItem(ByVal vInput As Variant, ByRef sFinal As String)
    Dim vSub As Variant, arrTemp As Variant, sCol As String
    
    ' A. 배열인 경우 (SelectFinal에서 넘겨준 경우 등)
    If IsArray(vInput) Then
        For Each vSub In vInput
            Call Internal_ProcessItem(vSub, sFinal) ' 재귀적 처리를 통해 중첩 배열 해결
        Next vSub
        
    ' B. 단일 값인 경우 (에러, Null, 객체 방어)
    Else
        ' 1. 기초 방어 (Error, Null, Empty 체크)
        If IsError(vInput) Or IsNull(vInput) Or IsEmpty(vInput) Then Exit Sub
        If IsObject(vInput) Then Exit Sub
        
        ' 2. 콤마 포함 여부 확인 및 분리
        If InStr(CStr(vInput), ",") > 0 Then
            arrTemp = Split(CStr(vInput), ",")
            For Each vSub In arrTemp
                sCol = Trim(CStr(vSub))
                If Len(sCol) > 0 Then
                    If Len(sFinal) > 0 Then sFinal = sFinal & ","
                    sFinal = sFinal & sCol
                End If
            Next vSub
        Else
            ' 3. 최종 문자열 변환 및 추가
            sCol = Trim(CStr(vInput))
            If Len(sCol) > 0 Then
                If Len(sFinal) > 0 Then sFinal = sFinal & ","
                sFinal = sFinal & sCol
            End If
        End If
    End If
End Sub
' ==============================================================================
' [Step 2.8] 정렬 설정 (SQL의 ORDER BY)
' 설명: SelectTable() 실행 시 반환될 결과 테이블을 정렬하도록 예약합니다.
' 활용: 기존 clsTable의 QuickSort 엔진을 재사용합니다.
' ==============================================================================
Public Function OrderBy(ByVal sColName As String, Optional ByVal bDesc As Boolean = False) As clsQuery
    RequireColumn sColName, "clsQuery.OrderBy"
    m_sSortColumn = Trim$(sColName)
    m_bSortDesc = bDesc
    Set OrderBy = Me
End Function
' ==============================================================================
' [Step 2.9] 결과 제한 (SQL의 TOP / LIMIT)
' 설명: 반환할 최대 행 개수를 설정합니다.
' ==============================================================================
Public Function Top(ByVal lCount As Long) As clsQuery
    If lCount < 0 Then lCount = 0
    m_lLimit = lCount
    Set Top = Me
End Function

' ------------------------------------------------------------------------------
' [Step 2.6] Update 값 설정 (SQL SET)
' 설명: 업데이트할 컬럼과 값을 설정합니다. (체이닝 지원)
' ------------------------------------------------------------------------------
Public Function SetValue(ByVal sColName As String, ByVal vValue As Variant) As clsQuery
    If m_dicUpdates.exists(sColName) Then
        m_dicUpdates(sColName) = vValue
    Else
        m_dicUpdates.Add sColName, vValue
    End If
    Set SetValue = Me
End Function

' @desc 업데이트 값 스테이징 후 자기 자신 반환, canonical (SetValue 동의어)
Public Function SetVal(ByVal sColName As String, ByVal vValue As Variant) As clsQuery
    ' 값 스테이징 어휘 수렴 (audit F11): clsTableBuilder.SetVal과 같은 개념 - 이름 공유
    Set SetVal = Me.SetValue(sColName, vValue)
End Function

' ------------------------------------------------------------------------------
' [Step 4] Update 실행 (SQL UPDATE)
' 설명: 조건에 맞는 행들의 컬럼 값을 일괄 변경합니다.
' 반환: 변경된 행의 개수
' ------------------------------------------------------------------------------
Public Function Update() As Long
    Dim vRowID As Variant, vCol As Variant
    Dim lCount As Long

    EnsureCandidatesCurrent
    If m_setCandidates Is Nothing Then InitializeAllRows
    For Each vCol In m_dicUpdates
        RequireColumn CStr(vCol), "clsQuery.Update"
    Next vCol
    If m_dicUpdates.Count = 0 Or m_setCandidates.Count = 0 Then Update = 0: Exit Function
    
    For Each vRowID In m_setCandidates.ToArray
        For Each vCol In m_dicUpdates
            m_SourceTable.Value(CLng(vRowID), CStr(vCol)) = m_dicUpdates(vCol)
        Next vCol
        lCount = lCount + 1
    Next vRowID
    
    Update = lCount
End Function

' ------------------------------------------------------------------------------
' [Step 4.1] Delete 실행 (SQL DELETE)
' 설명: 조건에 맞는 행을 테이블에서 삭제합니다.
' 반환: 삭제된 행의 개수
' ------------------------------------------------------------------------------
Public Function Delete() As Long
    Dim vSource As Variant
    Dim vNewData As Variant
    Dim i As Long, j As Long, R As Long
    Dim lTotalRows As Long, lTotalCols As Long
    Dim lDeleteCount As Long
    Dim lKeepCount As Long

    EnsureCandidatesCurrent
    If m_setCandidates Is Nothing Then InitializeAllRows
    lDeleteCount = m_setCandidates.Count
    If lDeleteCount = 0 Then Delete = 0: Exit Function

    ' 1. 데이터 재구성 (삭제 대상은 후보 RowSet의 ContainsRow로 직접 판정)
    vSource = m_SourceTable.RawArray
    lTotalRows = UBound(vSource, 1)
    lTotalCols = UBound(vSource, 2)
    lKeepCount = lTotalRows - lDeleteCount

    ReDim vNewData(1 To lKeepCount, 1 To lTotalCols)

    R = 1
    For j = 1 To lTotalCols: vNewData(1, j) = vSource(1, j): Next j ' 헤더

    For i = 2 To lTotalRows
        If Not m_setCandidates.ContainsRow(i) Then
            R = R + 1
            For j = 1 To lTotalCols: vNewData(R, j) = vSource(i, j): Next j
        End If
    Next i
    
    ' 2. 테이블 갱신 (기존 셀 재배치라 이미 정규화됨 - 입양 경로)
    m_SourceTable.AdoptArray vNewData, True
    Set m_setCandidates = Nothing: m_bIsFirstCondition = True
    
    Delete = lDeleteCount
End Function

' ------------------------------------------------------------------------------
' [Step 4.2] Insert 실행 (SQL INSERT)
' 설명: SetValue로 설정된 값으로 새로운 행을 추가합니다.
' 반환: 추가된 행의 개수 (성공 시 1)
' ------------------------------------------------------------------------------
Public Function Insert() As Long
    Dim vSource As Variant
    Dim vNewData As Variant
    Dim lTotalRows As Long, lTotalCols As Long
    Dim i As Long, j As Long
    Dim dicHeaders As Object
    Dim vKey As Variant
    Dim sCol As String
    Dim lColIdx As Long
    
    If m_dicUpdates.Count = 0 Then Insert = 0: Exit Function
    
    ' 1. 원본 데이터 및 헤더 확보
    vSource = m_SourceTable.RawArray
    lTotalRows = UBound(vSource, 1)
    lTotalCols = UBound(vSource, 2)
    Set dicHeaders = m_SourceTable.Headers
    
    ' 2. 배열 확장 (행 추가) 및 데이터 복사
    ReDim vNewData(1 To lTotalRows + 1, 1 To lTotalCols)
    For i = 1 To lTotalRows
        For j = 1 To lTotalCols: vNewData(i, j) = vSource(i, j): Next j
    Next i
    
    ' 3. 새 데이터 매핑
    For Each vKey In m_dicUpdates.Keys
        sCol = UCase(Trim(CStr(vKey)))
        If dicHeaders.exists(sCol) Then
            lColIdx = dicHeaders(sCol)
            vNewData(lTotalRows + 1, lColIdx) = m_dicUpdates(vKey)
        Else
            Err.Raise 9999, "clsQuery", "Insert Error: Column not found [" & sCol & "]"
        End If
    Next vKey
    
    ' 4. 테이블 갱신
    m_SourceTable.LoadFromArray vNewData, True
    m_dicUpdates.RemoveAll
    
    Insert = 1
End Function

' ------------------------------------------------------------------------------
' [Step 3] 최종 실행 및 결과 반환 (Projection + Top/Limit + OrderBy 통합)
' ------------------------------------------------------------------------------
' @desc 쿼리 결과를 clsTable로 물질화해 반환, canonical 종결자 (audit F10)
Public Function ToTable() As clsTable
    ' 종결자 어휘 수렴: LINQ To* 규약 정합 (survey 축 D). SelectTable은 호환 유지.
    Set ToTable = SelectTable()
End Function

Public Function SelectTable() As clsTable
    Dim tResult As New clsTable
    Dim vResultData As Variant
    Dim vRowIDs As Variant
    Dim vRawSource As Variant
    
    Dim arrColNames() As String
    Dim arrSourceIndices() As Long ' Target Index -> Source Index 매핑
    Dim lTargetCols As Long
    Dim lResultRows As Long
    
    Dim i As Long, j As Long, R As Long
    Dim sCol As String
    
    ' 1. [Candidate Check] 후보군 확인 (stale 후보군은 세대 계약으로 즉시 실패)
    EnsureCandidatesCurrent
    If m_setCandidates Is Nothing Then
        InitializeAllRows
    End If
    lResultRows = m_setCandidates.Count
    
    ' 후보 rowID 정렬 -> Top -> projection. 출력에 정렬 열이 없어도 의미를 보존한다.
    vRowIDs = m_setCandidates.ToArray
    If Len(m_sSortColumn) > 0 Then m_SourceTable.SortRowIds vRowIDs, m_sSortColumn, m_bSortDesc

    ' 출력 크기는 정렬 여부와 무관하게 Top으로 제한한다.
    If m_lLimit > 0 Then
        If lResultRows > m_lLimit Then lResultRows = m_lLimit
    End If
    
    ' 2. [Column Parsing] 프로젝션 대상 컬럼 분석
    Dim dicHeaders As Object
    Set dicHeaders = m_SourceTable.Headers
    
    If m_sTargetColumns = "*" Or m_sTargetColumns = "" Then
        ' [Case A] 전체 컬럼 선택
        lTargetCols = m_SourceTable.ColCount
        ReDim arrSourceIndices(1 To lTargetCols)
        For j = 1 To lTargetCols: arrSourceIndices(j) = j: Next j
    Else
        ' [Case B] 특정 컬럼 선택
        arrColNames = Split(m_sTargetColumns, ",")
        lTargetCols = UBound(arrColNames) + 1
        ReDim arrSourceIndices(1 To lTargetCols)
        
        For j = 0 To UBound(arrColNames)
            sCol = UCase(Trim(arrColNames(j)))
            If dicHeaders.exists(sCol) Then
                arrSourceIndices(j + 1) = dicHeaders(sCol)
            Else
                Err.Raise 9999, , "SelectColumns Error: Column not found [" & sCol & "]"
            End If
        Next j
    End If
    
    ' 3. [Allocation] 결과 배열 할당 (Limit이 적용된 lResultRows 사용)
    ReDim vResultData(1 To lResultRows + 1, 1 To lTargetCols)
    vRawSource = m_SourceTable.RawArray
    
    ' 4. [Header Copy] 헤더 작성
    If m_sTargetColumns = "*" Then
        For j = 1 To lTargetCols: vResultData(1, j) = vRawSource(1, j): Next j
    Else
        For j = 0 To UBound(arrColNames): vResultData(1, j + 1) = Trim(arrColNames(j)): Next j
    End If
    
    ' 5. [Data Fetch] 데이터 고속 복사
    If lResultRows > 0 Then
        
        ' [핵심] lResultRows만큼만 루프 수행 (정렬 없는 경우 Limit까지만 돎)
        ' vRowIDs는 0-based array이므로 0부터 시작
        For i = 0 To lResultRows - 1
            R = vRowIDs(i)
            For j = 1 To lTargetCols
                vResultData(i + 2, j) = vRawSource(R, arrSourceIndices(j))
            Next j
        Next i
    End If
    
    ' 결과 로드 (원본 셀 프로젝션이라 이미 정규화됨 - 입양 경로)
    tResult.AdoptArray vResultData, True
    
    Set SelectTable = tResult
End Function

' ==============================================================================
' [Internal] 헬퍼 함수
' ==============================================================================

' 전체 행을 후보군으로 초기화 (Where 없이 Filter 먼저 호출 시)
Private Sub InitializeAllRows()
    Set m_setCandidates = New DbRowSet
    m_setCandidates.BindTo m_SourceTable
    m_setCandidates.FromAllRows m_SourceTable.RowCount, 2
    m_bIsFirstCondition = False
End Sub

' 세대 계약: 보유 중인 후보 RowSet이 소스 테이블의 현재 세대와 일치하는지 검증
Private Sub EnsureCandidatesCurrent()
    If m_setCandidates Is Nothing Then Exit Sub
    m_setCandidates.EnsureCurrent m_SourceTable
End Sub


' 쿼리 결과를 VBE의 직접 실행 창에 출력 (디버깅용)
Public Sub DebugPrint()
    Dim T As clsTable
    Set T = Me.SelectTable()
    
    If T.RowCount = 0 Then
        Debug.Print "[Query Result] No rows returned."
        Exit Sub
    End If
    
    Dim vData As Variant
    vData = T.RawArray
    
    Dim i As Long, j As Long
    Dim sLine As String
    
    ' 헤더 출력
    sLine = ""
    For j = 1 To UBound(vData, 2)
        sLine = sLine & vData(1, j) & vbTab
    Next j
    Debug.Print sLine
    Debug.Print String(Len(sLine) + (UBound(vData, 2) * 4), "-") ' 구분선
    
    ' 데이터 출력
    For i = 2 To UBound(vData, 1)
        sLine = ""
        For j = 1 To UBound(vData, 2)
            sLine = sLine & vData(i, j) & vbTab
        Next j
        Debug.Print sLine
    Next i
End Sub


Private Sub RequireColumn(ByVal name As String, ByVal source As String)
    If Not m_SourceTable.HasColumn(name) Then
        Err.Raise 9999, source, "Column not found: " & name
    End If
End Sub

Private Sub ValidateFilterArm(ByVal column As String, ByVal op As String, ByVal value As Variant, ByVal source As String)
    RequireColumn column, source
    If IsObject(value) Or IsArray(value) Then
        Err.Raise 9999, source, "Filter compare value must be scalar"
    End If
    Select Case op
        Case "=", "<>", ">", "<", ">=", "<="
            If Not IsNumeric(value) Then
                Err.Raise 9999, source, "Numeric operator requires a numeric compare value"
            End If
    End Select
End Sub
