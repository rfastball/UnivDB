VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==============================================================================
' Class: clsReportBuilder
' 설명: UnivDB 기반의 리포트 생성 전용 빌더 (복잡한 로직 단순화)
' ==============================================================================
Option Explicit

Private m_tWork As clsTable      ' 작업 중인 테이블
Private m_dicFormats As Object   ' 서식 (NumberFormat)
Private m_dicWidths As Object    ' 너비 (ColumnWidth)
Private m_dicAligns As Object    ' 정렬 (HorizontalAlignment)
Private m_dicHeaders As Object   ' 헤더 별칭 (Header Alias)
Private m_bIsSafeCopy As Boolean ' [NEW] 원본 보호를 위한 복사본 여부

Private Sub Class_Initialize()
    Set m_dicFormats = CreateObject("Scripting.Dictionary")
    Set m_dicWidths = CreateObject("Scripting.Dictionary")
    Set m_dicAligns = CreateObject("Scripting.Dictionary")
    Set m_dicHeaders = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_tWork = Nothing
    Set m_dicFormats = Nothing
    Set m_dicWidths = Nothing
    Set m_dicAligns = Nothing
    Set m_dicHeaders = Nothing
End Sub

' [1] 초기화 (데이터 로드)
Public Function Init(ByVal tSource As clsTable) As clsReportBuilder
    Set m_tWork = tSource
    m_bIsSafeCopy = False ' 초기화 시에는 원본을 참조함
    Set Init = Me
End Function

' [Internal] 원본 데이터 보호를 위한 복사본 생성 (Copy-on-Write)
Private Sub EnsureSafeCopy()
    If m_bIsSafeCopy Then Exit Sub
    
    Dim tClone As New clsTable
    ' 원본 배열을 복사하여 새로운 테이블 생성
    tClone.LoadFromArray m_tWork.RawArray, True
    Set m_tWork = tClone
    m_bIsSafeCopy = True
End Sub

' [2] 컬럼 추가 (쉼표 구분 문자열 지원)
Public Function AddColumns(ByVal sColList As String) As clsReportBuilder
    EnsureSafeCopy ' 데이터 구조 변경 전 복사
    Dim vCol As Variant
    For Each vCol In Split(sColList, ",")
        m_tWork.AddColumn Trim(vCol)
    Next vCol
    Set AddColumns = Me
End Function

' [NEW] 컬럼 삭제 (쉼표 구분 문자열 지원)
Public Function RemoveColumns(ByVal sColList As String) As clsReportBuilder
    EnsureSafeCopy ' 데이터 구조 변경 전 복사
    Dim vCol As Variant
    For Each vCol In Split(sColList, ",")
        ' m_tWork는 Builder 내부의 clsTable 객체입니다.
        m_tWork.RemoveColumn Trim(vCol)
    Next vCol
    Set RemoveColumns = Me
End Function

' [3] 키 생성 (Create Key)
Public Function CreateKey(ByVal sNewCol As String, ByVal sSrcCol As String, ByVal lLen As Long) As clsReportBuilder
    EnsureSafeCopy ' 데이터 값 변경 전 복사
    Dim i As Long
    m_tWork.AddColumn sNewCol
    For i = 2 To m_tWork.RowCount
        m_tWork.Value(i, sNewCol) = Left(m_tWork.Value(i, sSrcCol), lLen)
    Next i
    Set CreateKey = Me
End Function

' ==============================================================================
' [4] 조인 (JoinTable) - 리포트 빌더용 메서드 체이닝 버전
' Update: sRightPrefix 매개변수 추가 (헤더 충돌 방지)
' ==============================================================================
Public Function JoinTable(ByVal tRef As clsTable, _
                          ByVal sKey As String, _
                          Optional ByVal jType As joinType = jtLeft, _
                          Optional ByVal sRightPrefix As String = "", _
                          Optional ByVal sRightKey As String = "") As clsReportBuilder
                          
    ' UnivDB.Join을 호출하여 테이블 병합 수행
    ' [UPDATE] 추가된 jType과 sRightPrefix를 파사드(UnivDB)로 전달
    Set m_tWork = UnivDB.Join(m_tWork, tRef, sKey, jType, sRightPrefix, sRightKey)
    m_bIsSafeCopy = True ' Join 결과는 항상 새로운 테이블임
    
    ' 메서드 체이닝을 위해 자기 자신(Me)을 반환
    Set JoinTable = Me
End Function

' [5] 매핑 값 채우기 (Map Value)
Public Function FillFromMap(ByVal sTargetCol As String, ByVal sSrcCol As String, Optional ByVal vDefault As Variant = "") As clsReportBuilder
    EnsureSafeCopy ' 데이터 값 변경 전 복사
    Dim i As Long
    Dim vVal As Variant
    
    For i = 2 To m_tWork.RowCount
        vVal = m_tWork.Value(i, sSrcCol)
        If IsEmpty(vVal) Then vVal = vDefault
        m_tWork.Value(i, sTargetCol) = vVal
    Next i
    Set FillFromMap = Me
End Function

' [6] 날짜 변환 (Convert Dates)
Public Function ConvertDates(ParamArray vCols() As Variant) As clsReportBuilder
    EnsureSafeCopy ' 데이터 값 변경 전 복사
    Dim i As Long, j As Long
    Dim sCol As String
    
    If UBound(vCols) < 0 Then
        Set ConvertDates = Me
        Exit Function
    End If
    
    For i = 2 To m_tWork.RowCount
        For j = LBound(vCols) To UBound(vCols)
            sCol = CStr(vCols(j))
            ' Internal_ConvertToDate 헬퍼 함수 사용
            m_tWork.Value(i, sCol) = Internal_ConvertToDate(m_tWork.Value(i, sCol))
        Next j
    Next i
    
    Set ConvertDates = Me
End Function

' ==============================================================================
' [NEW] 데이터 필터링 및 수정 (clsQuery 기능 연동)
' ==============================================================================

' 조건에 맞는 행만 남기기 (Keep)
Public Function Filter(ByVal sColName As String, ByVal sOp As String, ByVal vValue As Variant) As clsReportBuilder
    ' UnivDB.Query를 통해 필터링된 결과 테이블로 교체
    Set m_tWork = UnivDB.Query(m_tWork).Filter(sColName, sOp, vValue).SelectTable()
    m_bIsSafeCopy = True ' SelectTable 결과는 항상 새로운 테이블임
    Set Filter = Me
End Function

' @desc 조건에 맞는 행만 즉시 유지, Filter 동의어 (즉시 실행 어휘 - audit F7)
Public Function Keep(ByVal sColName As String, ByVal sOp As String, ByVal vValue As Variant) As clsReportBuilder
    ' clsQuery.Filter(지연)와의 동명이의 완화: Keep/Exclude 대칭쌍이 canonical 어휘.
    Set Keep = Me.Filter(sColName, sOp, vValue)
End Function

' @desc arm 중 하나라도 맞는 행 즉시 유지 (교차 컬럼 OR, FilterAny 위임)
Public Function KeepAny(ParamArray vArms() As Variant) As clsReportBuilder
    ' ParamArray는 재전달 시 단일 배열 요소가 된다 - FilterAny의 재귀 평탄화가 해석.
    ' 문장 수준 오류는 m_tWork 교체 전에 raise (변형 없는 실패).
    Set m_tWork = UnivDB.Query(m_tWork).FilterAny(vArms).SelectTable()
    m_bIsSafeCopy = True ' SelectTable 결과는 항상 새로운 테이블임
    Set KeepAny = Me
End Function

' 조건에 맞는 행 삭제하기 (Remove)
Public Function Exclude(ByVal sColName As String, ByVal sOp As String, ByVal vValue As Variant) As clsReportBuilder
    EnsureSafeCopy ' [중요] 원본 삭제 방지를 위해 복사본 생성 후 수행
    ' UnivDB.Query.Delete를 통해 원본 테이블에서 즉시 삭제
    UnivDB.Query(m_tWork).Filter(sColName, sOp, vValue).Delete
    Set Exclude = Me
End Function

' 조건에 맞는 행의 값 일괄 변경
Public Function UpdateValue(ByVal sTargetCol As String, ByVal vNewVal As Variant, _
                            ByVal sCondCol As String, ByVal sOp As String, ByVal vCondVal As Variant) As clsReportBuilder
    EnsureSafeCopy ' [중요] 원본 수정 방지를 위해 복사본 생성 후 수행
    ' 예: .UpdateValue("Status", "Expired", "Date", "<", Date)
    UnivDB.Query(m_tWork) _
          .Filter(sCondCol, sOp, vCondVal) _
          .SetValue(sTargetCol, vNewVal) _
          .Update
          
    Set UpdateValue = Me
End Function

' @desc 조건 매칭 행의 값 일괄 패치, canonical (UpdateValue 동의어, 빌더 PatchIf 인자 순서)
Public Function PatchIf(ByVal sCondCol As String, ByVal sOp As String, ByVal vCondVal As Variant, _
                        ByVal sTargetCol As String, ByVal vNewVal As Variant) As clsReportBuilder
    ' 조건부 패치 어휘 수렴 (audit F11): clsTableBuilder.PatchIf와 같은 개념/인자 순서.
    ' UpdateValue(target, new, cond, op, condVal)는 레거시 인자 순서로 유지.
    Set PatchIf = Me.UpdateValue(sTargetCol, vNewVal, sCondCol, sOp, vCondVal)
End Function

' ==============================================================================
' [Internal Helper] 날짜 변환 로직 (사용자 정의 로직 이식)
' ==============================================================================
Private Function Internal_ConvertToDate(vValue As Variant) As Variant
    Dim sVal As String
    sVal = Trim(CStr(vValue))
    
    ' 1. 빈 값 처리
    If sVal = "" Or sVal = "-" Then
        Internal_ConvertToDate = Empty
        Exit Function
    End If
    
    On Error Resume Next
    
    ' 2. YYYYMMDD (8자리 숫자)
    If Len(sVal) = 8 And IsNumeric(sVal) Then
        Internal_ConvertToDate = DateSerial(Left(sVal, 4), Mid(sVal, 5, 2), Right(sVal, 2))
        If Err.Number = 0 Then Exit Function
        Err.Clear
    End If
    
    ' 3. 구분자가 있는 날짜
    If IsDate(sVal) Then
        Internal_ConvertToDate = CDate(sVal)
    Else
        Internal_ConvertToDate = vValue
    End If
    
    On Error GoTo 0
End Function

' [7] 사용자 정의 연산용 테이블 반환
Public Property Get Table() As clsTable
    Set Table = m_tWork
End Property

' @desc 작업 테이블을 clsTable로 반환, canonical 종결자 (audit F10, Table 동의어)
Public Function ToTable() As clsTable
    Set ToTable = m_tWork
End Function

' [8] 최종 선택 (Select)
' ==============================================================================
' [신규] SelectColumns (Projection 설정) - ParamArray 적용
' 설명: 최종 리포트에 포함할 컬럼과 순서를 지정합니다.
' ------------------------------------------------------------------------------
Public Function SelectColumns(ParamArray vCols() As Variant) As clsReportBuilder
    ' 1. UnivDB.Query의 SelectColumns가 이미 ParamArray 및 배열 처리를
    '    완벽하게 지원하므로, vCols 배열을 그대로 전달합니다.
    ' 2. SelectTable()을 통해 추출된 결과로 작업 테이블을 교체합니다.
    Set m_tWork = UnivDB.Query(m_tWork).SelectColumns(vCols).SelectTable()
    
    ' 3. SelectTable 결과는 항상 새로운 테이블이므로 안전 복사 상태로 마킹합니다.
    m_bIsSafeCopy = True
    Set SelectColumns = Me
End Function

' ==============================================================================
' [Legacy/Deprecated] SelectFinal (하위 호환성 유지 - 신규 코드는 SelectColumns 사용)
' 설명: 이전 버전과의 호환을 위해 유지하며, 내부적으로 SelectColumns를 호출합니다.
'       audit F14: 같은 기능 두 이름 - canonical은 SelectColumns (제거 계획 없음).
' ------------------------------------------------------------------------------
Public Function SelectFinal(ParamArray vCols() As Variant) As clsReportBuilder
    ' ParamArray로 전달받은 vCols를 그대로 넘겨 처리합니다.
    Set SelectFinal = Me.SelectColumns(vCols)
End Function

' [NEW] 컬럼 속성 정의 (선언적 서식 설정)
' 설명: 컬럼의 너비, 표시형식, 정렬, 헤더명을 한 번에 설정합니다.
Public Function DefineColumn(ByVal sColName As String, _
                             Optional ByVal vWidth As Variant, _
                             Optional ByVal vFormat As Variant, _
                             Optional ByVal vAlign As Variant, _
                             Optional ByVal vHeaderAlias As Variant) As clsReportBuilder
    Dim sKey As String
    sKey = UCase(Trim(sColName))
    
    If Not IsMissing(vWidth) Then m_dicWidths(sKey) = vWidth
    If Not IsMissing(vFormat) Then m_dicFormats(sKey) = vFormat
    If Not IsMissing(vAlign) Then m_dicAligns(sKey) = vAlign
    If Not IsMissing(vHeaderAlias) Then m_dicHeaders(sKey) = vHeaderAlias
    
    Set DefineColumn = Me
End Function

' [9] 출력 (Build & Dump)
Public Sub BuildToSheet(ByVal ws As Worksheet, ByVal sAddr As String, Optional ByVal arrHeaders As Variant)
    Dim rngTable As Range
    Dim iCol As Long
    Dim sKey As String
    Dim lRows As Long

    ' 1. 기존 데이터 클리어
    ws.AutoFilterMode = False
    ws.Range(sAddr).CurrentRegion.ClearContents
    
    ' 2. 서식 고정 (텍스트)
    ws.Range(sAddr).Resize(m_tWork.RowCount, m_tWork.ColCount).NumberFormat = "@"
    
    ' 3. 덤프
    m_tWork.DumpToSheet ws.Range(sAddr)
    
    ' 4. 헤더 리네이밍
    If Not IsMissing(arrHeaders) Then
        If IsArray(arrHeaders) Then
            ws.Range(sAddr).Resize(1, UBound(arrHeaders) + 1).Value = arrHeaders
        End If
    End If
    
    ' 5. UI 정리
    With ws.Range(sAddr).Resize(1, m_tWork.ColCount)
        .VerticalAlignment = xlBottom
        .HorizontalAlignment = xlCenter
    End With
    
    If Not ws.AutoFilterMode Then ws.Range(sAddr).AutoFilter
    '오토핏 적용(주석처리)
    'ws.Range(sAddr).CurrentRegion.EntireColumn.AutoFit
    
    ' [NEW] 6. 사용자 정의 서식 적용 (DefineColumn 설정 반영)
    Set rngTable = ws.Range(sAddr).CurrentRegion
    lRows = rngTable.Rows.Count
    
    If lRows > 0 Then
    For iCol = 1 To rngTable.Columns.Count
        sKey = UCase(Trim(rngTable.Cells(1, iCol).Value))
        
        ' (1) 표시 형식 (데이터 영역만)
        If lRows > 1 And m_dicFormats.exists(sKey) Then
            rngTable.Columns(iCol).Offset(1, 0).Resize(lRows - 1).NumberFormat = m_dicFormats(sKey)
        End If
        
        ' (2) 정렬 (데이터 영역에만 적용하여 헤더의 중앙 정렬 유지)
        If lRows > 1 And m_dicAligns.exists(sKey) Then
            Dim targetRng As Range
            Set targetRng = rngTable.Columns(iCol).Offset(1, 0).Resize(lRows - 1)
            
            Select Case UCase(m_dicAligns(sKey))
                Case "L", "LEFT": targetRng.HorizontalAlignment = xlLeft
                Case "C", "CENTER": targetRng.HorizontalAlignment = xlCenter
                Case "R", "RIGHT": targetRng.HorizontalAlignment = xlRight
            End Select
        End If
            
            ' (3) 너비 (AutoFit 덮어쓰기)
            If m_dicWidths.exists(sKey) Then rngTable.Columns(iCol).ColumnWidth = m_dicWidths(sKey)
            
            ' (4) 헤더 별칭 (마지막에 적용)
            If m_dicHeaders.exists(sKey) Then rngTable.Cells(1, iCol).Value = m_dicHeaders(sKey)
        Next iCol
    End If
    
    Application.ScreenUpdating = True
End Sub


