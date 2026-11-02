VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbHashIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc 해시 인덱스 백엔드 교체용 인터페이스 표면 담당
Option Explicit

' 인터페이스 계약 (구현: DictHashIndex, 후속 후보: CoreHashIndex)
' - 키 정규화: modIndexKey.NormalizeIndexKey (Empty -> "", CStr, 대소문자 무시)
' - IsError 키는 저장/조회 모두 skip
' - 세대 계약: EnsureCurrent가 stale 사용을 9998로 차단

' @desc 키와 rowID를 받아 인덱스에 적재함
Public Sub AddRow(ByVal key As Variant, ByVal rowId As Long)
End Sub

' @desc 키 받아 RowList 반환 여부 반환, 미존재/Error 키 시 False
Public Function TryGetRows(ByVal key As Variant, ByRef outRows As DbRowList) As Boolean
End Function

' span 계열 (객체 생성 없는 핫 경로용): 버퍼는 호출자 소유 재사용,
' 부족 시 구현이 배가 ReDim, 내용은 호출마다 덮어씀 - 소비는 outCount 기준.

' @desc 키의 rowId span을 버퍼에 복사 여부 반환, 미존재/Error 키 시 False에 0건
Public Function TryGetRowIdSpan(ByVal key As Variant, ByRef outRowIds() As Long, ByRef outCount As Long) As Boolean
End Function

' @desc 키의 행 수만 반환 여부 반환(복사 없음), 미존재/Error 키 시 False에 0건
Public Function TryGetRowCount(ByVal key As Variant, ByRef outCount As Long) As Boolean
End Function

' @desc 키의 첫 rowId 반환 여부 반환, 미존재/Error 키 시 False
Public Function TryGetFirstRowId(ByVal key As Variant, ByRef outRowId As Long) As Boolean
End Function

' @desc 고유 키 개수 반환
Public Property Get CountKeys() As Long
End Property

' @desc 인덱스 내용을 비움
Public Sub Clear()
End Sub

' @desc 키-RowList 순회용 커서 반환
Public Function Cursor() As DbIndexCursor
End Function

' @desc 테이블을 받아 결속 세대와 일치 검증함, stale 시 9998 에러
Public Sub EnsureCurrent(ByVal tbl As clsTable)
End Sub
