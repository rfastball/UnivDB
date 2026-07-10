VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 컬럼명 -> 배열 인덱스(1-based) 매핑
Private m_dicHeaders As Object
' ID -> 1차원 배열(Variant Array) 매핑
Private m_dicRows As Object

' 현재 작업 중인 행의 ID 포인터
Private m_sCursor As String
' 현재 작업 중인 행의 1차원 배열 데이터 (캐시)
Private m_vCursor As Variant

' 빌더의 기준이 되는 Key 컬럼명
Private m_sKeyColumn As String

Private Sub Class_Initialize()
    Set m_dicHeaders = CreateObject("Scripting.Dictionary")
    Set m_dicRows = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_dicHeaders = Nothing
    Set m_dicRows = Nothing
End Sub

' ---------------------------------------------------------
' [설정] 키 컬럼 및 초기화
' ---------------------------------------------------------
Public Property Let KeyColumn(ByVal sKey As String)
    m_sKeyColumn = UCase(Trim(sKey))
    If Not m_dicHeaders.exists(m_sKeyColumn) Then
        m_dicHeaders.Add m_sKeyColumn, m_dicHeaders.Count + 1
    End If
End Property

Public Sub InitFromTable(ByVal T As clsTable)
    Dim vKey As Variant
    
    m_dicHeaders.RemoveAll
    ' 원본 테이블의 Headers 딕셔너리를 안전하게 복제
    For Each vKey In T.Headers.Keys
        m_dicHeaders.Add UCase(Trim(CStr(vKey))), T.Headers(vKey)
    Next vKey
    
    ' 키 컬럼이 누락된 경우 강제 추가
    If Len(m_sKeyColumn) > 0 And Not m_dicHeaders.exists(m_sKeyColumn) Then
        m_dicHeaders.Add m_sKeyColumn, m_dicHeaders.Count + 1
    End If
End Sub

' ---------------------------------------------------------
' [내부 엔진] 커서 관리 및 동적 스키마 확장
' ---------------------------------------------------------
Private Sub p_CommitCursor()
    If Len(m_sCursor) > 0 Then
        m_dicRows(m_sCursor) = m_vCursor
    End If
End Sub

Private Sub p_ExpandSchema(ByVal sNewCol As String)
    Dim lNewIndex As Long
    lNewIndex = m_dicHeaders.Count + 1
    m_dicHeaders.Add sNewCol, lNewIndex
    
    ' 현재 커서 확장
    If IsArray(m_vCursor) Then
        ReDim Preserve m_vCursor(1 To lNewIndex)
    Else
        ReDim m_vCursor(1 To lNewIndex)
    End If
    
    ' 기존 행들 일괄 확장 (VBA 1D Array ReDim은 빠름)
    Dim vKey As Variant, vTemp As Variant
    For Each vKey In m_dicRows.Keys
        If vKey <> m_sCursor Then
            vTemp = m_dicRows(vKey)
            If IsArray(vTemp) Then
                ReDim Preserve vTemp(1 To lNewIndex)
                m_dicRows(vKey) = vTemp
            End If
        End If
    Next vKey
End Sub

' ---------------------------------------------------------
' [조작] 레코드 단위 추출 및 이동
' ---------------------------------------------------------
Public Function Row(ByVal ID As String) As clsTableBuilder
    p_CommitCursor
    m_sCursor = CStr(ID)
    
    If m_dicRows.exists(m_sCursor) Then
        m_vCursor = m_dicRows(m_sCursor)
    Else
        If m_dicHeaders.Count > 0 Then
            ReDim m_vCursor(1 To m_dicHeaders.Count)
        Else
            m_vCursor = Array()
        End If
        m_dicRows.Add m_sCursor, m_vCursor
    End If
    
    ' 새 행이든 기존 행이든 Key 값은 확실히 박아둠
    If Len(m_sKeyColumn) > 0 Then Me.Val m_sKeyColumn, m_sCursor
    Set Row = Me
End Function

Public Function ExtractRow(ByVal SourceTable As clsTable, ByVal ID As String) As clsTableBuilder
    Dim colRows As Collection
    Dim dicSrcHeaders As Object
    Dim vRaw As Variant
    Dim lRow As Long
    Dim vKey As Variant
    
    ' 1. 타겟 Row로 커서 이동 (없으면 생성됨)
    Me.Row ID
    
    ' 2. 원본에서 고속으로 RowID 추출 (O(1))
    Set colRows = SourceTable.SelectRows(m_sKeyColumn, ID)
    
    ' 3. 데이터가 존재하면 고속 복사
    If colRows.Count > 0 Then
        lRow = colRows(1) ' 중복 시 첫 번째 행 기준
        vRaw = SourceTable.RawArray
        Set dicSrcHeaders = SourceTable.Headers
        
        For Each vKey In dicSrcHeaders.Keys
            ' 현재 빌더의 Val(Setter)를 이용해 안전하게 복사 (동적 스키마 자동 대응)
            Me.Val CStr(vKey), vRaw(lRow, dicSrcHeaders(vKey))
        Next vKey
    End If
    
    Set ExtractRow = Me
End Function

' ---------------------------------------------------------
' [조작] 값 읽기 / 쓰기 (동적 프로퍼티)
' ---------------------------------------------------------
Public Function Val(ByVal ColName As String, Optional ByVal Value As Variant) As Variant
    Dim sKey As String
    Dim lIndex As Long
    
    sKey = UCase(Trim(ColName))
    
    If IsMissing(Value) Then
        ' --- [GETTER] ---
        If m_dicHeaders.exists(sKey) Then
            lIndex = m_dicHeaders(sKey)
            On Error Resume Next ' 초기화되지 않은 배열 방어
            Val = m_vCursor(lIndex)
            If Err.Number <> 0 Then Val = Empty
            On Error GoTo 0
        Else
            Val = Empty
        End If
    Else
        ' --- [SETTER] ---
        If Len(m_sCursor) = 0 Then Err.Raise 9999, "clsTableBuilder", ".Row(ID)를 먼저 호출하세요."
        
        If Not m_dicHeaders.exists(sKey) Then
            p_ExpandSchema sKey
        End If
        
        m_vCursor(m_dicHeaders(sKey)) = Value
        Set Val = Me ' 체이닝 반환
    End If
End Function

Public Function PatchIf(ByVal ColName As String, ByVal Operator As String, ByVal CondVal As Variant, ByVal TargetCol As String, ByVal NewVal As Variant) As clsTableBuilder
    Dim vCurrent As Variant
    Dim bMatched As Boolean
    
    vCurrent = Me.Val(ColName)

    ' 미지 연산자는 즉시 raise (audit F4, V1) - On Error 진입 전에 검증
    Operator = modOperators.NormalizeCompareOp(Operator)

    ' 단순 비교 연산 - 행 데이터 수준 비교 오류(타입 불일치)는 무매칭으로 관용 (계약)
    On Error Resume Next
    Select Case Operator
        Case "=":  bMatched = (vCurrent = CondVal)
        Case "<>": bMatched = (vCurrent <> CondVal)
        Case ">":  bMatched = (vCurrent > CondVal)
        Case ">=": bMatched = (vCurrent >= CondVal)
        Case "<":  bMatched = (vCurrent < CondVal)
        Case "<=": bMatched = (vCurrent <= CondVal)
        Case Else: bMatched = False
    End Select
    On Error GoTo 0
    
    If bMatched Then Me.Val TargetCol, NewVal
    Set PatchIf = Me
End Function

' ---------------------------------------------------------
' [빌드] 데이터 병합 및 객체 반환
' ---------------------------------------------------------
Public Function ToTable() As clsTable
    p_CommitCursor
    
    Dim tNew As New clsTable
    If m_dicRows.Count = 0 Then
        Set ToTable = tNew
        Exit Function
    End If
    
    Dim lRows As Long, lCols As Long
    Dim vResult As Variant
    Dim vKey As Variant, vRow As Variant
    Dim i As Long, j As Long
    
    lRows = m_dicRows.Count
    lCols = m_dicHeaders.Count
    ReDim vResult(1 To lRows + 1, 1 To lCols)
    
    ' 1. 헤더 입력
    For Each vKey In m_dicHeaders.Keys
        vResult(1, m_dicHeaders(vKey)) = vKey
    Next vKey
    
    ' 2. 데이터 입력
    i = 2
    For Each vKey In m_dicRows.Keys
        vRow = m_dicRows(vKey)
        If IsArray(vRow) Then
            For j = 1 To UBound(vRow)
                If j <= lCols Then vResult(i, j) = vRow(j)
            Next j
        End If
        i = i + 1
    Next vKey
    
    ' 3. 반환 및 초기화
    tNew.LoadFromArray vResult
    Set ToTable = tNew
    Clear
End Function

Public Sub Clear()
    m_dicHeaders.RemoveAll
    m_dicRows.RemoveAll
    m_sCursor = ""
    Erase m_vCursor
    
    If Len(m_sKeyColumn) > 0 Then
        m_dicHeaders.Add m_sKeyColumn, 1
    End If
End Sub
