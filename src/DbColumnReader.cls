VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbColumnReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @desc 단일 컬럼 값을 세대 결속 1D 표면으로 제공 담당
' @depends modDbContract
Option Explicit

' 컬럼 값 물질화 사본 (rowId 인덱스: 헤더행+1 .. RowCount, 데이터 행 없으면 Empty)
Private mValues As Variant
Private mFirstDataRow As Long
Private mLastDataRow As Long

' 세대 계약 결속 (ObjPtr 스냅샷 - 객체 참조 아님, 순환 참조 방지)
Private mOwnerPtr As LongPtr
Private mGeneration As Long

' clsTable 팩토리 전용 - 컬럼 사본과 결속 스냅샷 주입
' 일회용(생성-스캔-폐기) 전제. 값 변경은 세대를 올리지 않으므로
' 향후 clsTable이 reader를 캐시하려면 InvalidateIndexByColumnName에 축출 훅이 필요함
Friend Sub BindData(ByVal vColumn As Variant, ByVal ownerPtr As LongPtr, ByVal gen As Long)
    mValues = vColumn
    mOwnerPtr = ownerPtr
    mGeneration = gen

    If IsEmpty(mValues) Then
        mFirstDataRow = 1
        mLastDataRow = 0
    Else
        mFirstDataRow = LBound(mValues)
        mLastDataRow = UBound(mValues)
    End If
End Sub

' @desc rowId 받아 해당 행의 컬럼 값 반환, 범위 밖 시 첨자 에러
Public Function ValueAt(ByVal lRowId As Long) As Variant
    ValueAt = mValues(lRowId)
End Function

' @desc 컬럼 값 1D 배열 반환, 데이터 행 없으면 Empty 반환
Public Property Get Values() As Variant
    Values = mValues
End Property

Public Property Get FirstDataRow() As Long
    FirstDataRow = mFirstDataRow
End Property

Public Property Get LastDataRow() As Long
    LastDataRow = mLastDataRow
End Property

' @desc 테이블을 받아 결속 세대와 일치 검증함, stale 시 9998 에러
Public Sub EnsureCurrent(ByVal tbl As clsTable)
    modDbContract.AssertCurrentGeneration mOwnerPtr, mGeneration, tbl, "DbColumnReader.EnsureCurrent"
End Sub

Friend Property Get OwnerPtr() As LongPtr
    OwnerPtr = mOwnerPtr
End Property

Friend Property Get BoundGeneration() As Long
    BoundGeneration = mGeneration
End Property
