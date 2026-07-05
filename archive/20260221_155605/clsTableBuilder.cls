VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Public Property Let KeyColumn(ByVal sKey As String)
    m_sKeyColumn = sKey
    ' 키 컬럼도 헤더에 포함시킵니다.
    If Not m_dicHeaders.exists(sKey) Then
        m_dicHeaders.Add sKey, m_dicHeaders.Count + 1
    End If
End Property

Public Sub InitFromTable(ByVal T As clsTable)
    Dim vHeader As Variant
    Dim i As Long
    
    ' 원본 테이블의 헤더 정보를 그대로 복제
    m_dicHeaders.RemoveAll
    vHeader = T.header
    For i = 1 To UBound(vHeader)
        m_dicHeaders.Add vHeader(i), i
    Next i
    
    ' 키 컬럼이 헤더에 없는 경우 추가
    If Not m_dicHeaders.exists(m_sKeyColumn) Then
        m_dicHeaders.Add m_sKeyColumn, m_dicHeaders.Count + 1
    End If
End Sub


' 현재 작업중인 커서(m_vCursor)의 변경사항을 m_dicRows에 최종 반영
Private Sub p_CommitCursor()
    If Len(m_sCursor) > 0 Then
        m_dicRows(m_sCursor) = m_vCursor
    End If
End Sub

' 작업할 행을 ID로 지정. 없으면 신규 생성
Public Function Row(ByVal ID As String) As clsTableBuilder
    ' 이전 커서가 있었다면 변경사항을 저장
    p_CommitCursor
    
    m_sCursor = ID
    
    If m_dicRows.exists(ID) Then
        ' 기존 행 로드
        m_vCursor = m_dicRows(ID)
    Else
        ' 신규 행 생성
        ' ReDim Preserve는 비어있을 때 사용 불가. 분기 처리.
        If m_dicHeaders.Count > 0 Then
            ReDim m_vCursor(1 To m_dicHeaders.Count)
        Else
            ' 헤더가 아예 없는 초기 상태.
            ' 일단 빈 배열로 초기화. Val에서 확장됨.
            m_vCursor = Array()
        End If
        m_dicRows.Add ID, m_vCursor
    End If
    
    ' KeyColumn에 ID 자동 입력
    Me.Val m_sKeyColumn, ID
    
    Set Row = Me
End Function

' 원본 테이블에서 특정 ID의 행을 찾아 빌더로 복사
Public Function ExtractRow(SourceTable As clsTable, ByVal ID As String) As clsTableBuilder
    Dim T As clsTable
    Set T = SourceTable.Query.SelectRows(m_sKeyColumn, "=", ID).ToTable
    
    ' Select 결과가 있으면 첫번째 행을 복사
    If T.RowCount > 0 Then
        Dim vRow As Variant
        vRow = T.GetRow(1) ' 첫번째 행 전체를 가져옴
        
        Dim i As Long
        Dim sCol As String
        Dim vHeader As Variant
        vHeader = T.header
        
        ' Row 메서드를 호출하여 커서 초기화
        Me.Row ID
        
        ' 헤더를 순회하며 데이터 복사
        For i = 1 To UBound(vHeader)
            sCol = vHeader(i)
            Me.Val sCol, vRow(i)
        Next i
    Else
        ' Select 결과가 없으면 그냥 해당 ID로 커서만 이동
        Me.Row ID
    End If
    
    Set ExtractRow = Me
End Function


' 새 컬럼 추가 시 모든 행의 배열 크기 확장
Private Sub p_ExpandSchema(ByVal NewColumnName As String)
    ' 1. 헤더 딕셔너리 확장
    Dim NewIndex As Long
    NewIndex = m_dicHeaders.Count + 1
    m_dicHeaders.Add NewColumnName, NewIndex
    
    ' 2. 현재 커서 확장
    ReDim Preserve m_vCursor(1 To NewIndex)
    
    ' 3. 저장된 모든 행(m_dicRows)의 배열 크기 확장
    Dim vKey As Variant
    Dim vTemp As Variant
    For Each vKey In m_dicRows.Keys
        ' 현재 커서가 아닌 행들만 확장 (커서는 이미 위에서 확장됨)
        If vKey <> m_sCursor Then
            vTemp = m_dicRows(vKey)
            ReDim Preserve vTemp(1 To NewIndex)
            m_dicRows(vKey) = vTemp
        End If
    Next vKey
End Sub

' [Getter/Setter] 값 읽기/쓰기. 컬럼 없으면 자동 생성(동적 스키마).
Public Function Val(ByVal ColName As String, Optional ByVal Value As Variant) As Variant
    If IsMissing(Value) Then
        ' --- Getter ---
        If m_dicHeaders.exists(ColName) Then
            Dim Index As Long: Index = m_dicHeaders(ColName)
            ' 현재 커서가 비어있는 경우 에러 방지
            If IsArray(m_vCursor) And UBound(m_vCursor) >= Index Then
                Val = m_vCursor(Index)
            Else
                Val = Empty ' 커서에 아직 데이터가 없거나, 배열이 작으면 Empty 반환
            End If
        Else
            Val = Empty ' 존재하지 않는 컬럼 조회는 Empty 반환
        End If
    Else
        ' --- Setter ---
        If Len(m_sCursor) = 0 Then Err.Raise 9999, "clsTableBuilder", ".Row(ID)를 먼저 호출하여 작업 행을 지정해야 합니다."
        
        ' 1. 컬럼이 없으면 스키마 확장
        If Not m_dicHeaders.exists(ColName) Then
            p_ExpandSchema ColName
        End If
        
        ' 2. 값 쓰기
        m_vCursor(m_dicHeaders(ColName)) = Value
        
        ' 3. 메서드 체이닝을 위해 Me 객체 반환
        Set Val = Me
    End If
End Function

' 현재 행에 대한 조건부 값 변경
Public Function PatchIf(ByVal ColName As String, ByVal Operator As String, ByVal CondVal As Variant, ByVal TargetCol As String, ByVal NewVal As Variant) As clsTableBuilder
    Dim CurrentVal As Variant
    CurrentVal = Me.Val(ColName)
    
    Dim Condition As Boolean
    Select Case Operator
        Case "=":   Condition = (CurrentVal = CondVal)
        Case "<>":  Condition = (CurrentVal <> CondVal)
        Case ">":   Condition = (CurrentVal > CondVal)
        Case ">=":  Condition = (CurrentVal >= CondVal)
        Case "<":   Condition = (CurrentVal < CondVal)
        Case "<=":  Condition = (CurrentVal <= CondVal)
        Case Else:  Condition = False
    End Select
    
    If Condition Then
        Me.Val(TargetCol) = NewVal
    End If
    
    Set PatchIf = Me
End Function


' 빌더에 쌓인 데이터를 새로운 clsTable 객체로 변환하여 반환
Public Function ToTable() As clsTable
    ' 1. 최종 커서 변경사항 반영
    p_CommitCursor
    
    ' 데이터가 없는 경우 빈 테이블 반환
    If m_dicRows.Count = 0 Then
        Set ToTable = UnivDB.CreateTable
        Exit Function
    End If
    
    ' 2. 헤더와 데이터 배열 준비
    Dim vHeaders As Variant: vHeaders = m_dicHeaders.Keys
    ReDim vData(1 To m_dicRows.Count, 1 To m_dicHeaders.Count)
    
    Dim i As Long: i = 1
    Dim vKey As Variant
    Dim vRow As Variant
    Dim j As Long
    
    ' 3. m_dicRows의 1차원 배열들을 2차원 배열로 조립
    For Each vKey In m_dicRows.Keys
        vRow = m_dicRows(vKey)
        ' 행 데이터가 배열이고, 크기가 맞는 경우에만 복사
        If IsArray(vRow) And UBound(vRow) = m_dicHeaders.Count Then
            For j = 1 To m_dicHeaders.Count
                vData(i, j) = vRow(j)
            Next j
        End If
        i = i + 1
    Next vKey
    
    ' 4. 최종 데이터(헤더+데이터) 배열 생성
    Dim vResult As Variant
    ReDim vResult(1 To UBound(vData, 1) + 1, 1 To UBound(vData, 2))
    
    ' 헤더 복사 (vHeaders는 0-based, vResult는 1-based)
    For j = 0 To UBound(vHeaders)
        vResult(1, j + 1) = vHeaders(j)
    Next j
    
    ' 데이터 복사
    For i = 1 To UBound(vData, 1)
        For j = 1 To UBound(vData, 2)
            vResult(i + 1, j) = vData(i, j)
        Next j
    Next i
    
    ' 5. 새 clsTable 객체 생성 및 데이터 로드
    Dim T As New clsTable
    T.LoadFromArray vResult
    Set ToTable = T
    
    ' 6. 내부 메모리 정리
    Clear
End Function

' 빌더 내부 상태 초기화
Public Sub Clear()
    m_dicHeaders.RemoveAll
    m_dicRows.RemoveAll
    m_sCursor = ""
    Erase m_vCursor
    
    ' KeyColumn은 유지해야 다음에 재사용 가능
    ' 따라서 m_sKeyColumn은 초기화하지 않음.
    ' 단, 헤더에는 다시 추가해줘야 함.
    If Len(m_sKeyColumn) > 0 Then
        m_dicHeaders.Add m_sKeyColumn, 1
    End If
End Sub


Private Sub Class_Initialize()
    Set m_dicHeaders = CreateObject("Scripting.Dictionary")
    Set m_dicRows = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_dicHeaders = Nothing
    Set m_dicRows = Nothing
End Sub
