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
Private mColumnIndex As Long
Private mValueVersion As Long

' clsTable 팩토리 전용 - 컬럼 사본과 결속 스냅샷 주입
' 값 접근은 스냅샷 읽기. EnsureCurrent는 구조 세대와 해당 열의 값 버전을 함께 검증한다.
Friend Sub BindData(ByVal vColumn As Variant, ByVal ownerPtr As LongPtr, ByVal gen As Long, _
                    ByVal columnIndex As Long, ByVal valueVersion As Long)
    mValues = vColumn
    mOwnerPtr = ownerPtr
    mGeneration = gen
    mColumnIndex = columnIndex
    mValueVersion = valueVersion

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

' @desc 테이블의 구조 세대와 열 값 버전 검증함, 불일치 시 9998 에러
Public Sub EnsureCurrent(ByVal tbl As clsTable)
    modDbContract.AssertCurrentGeneration mOwnerPtr, mGeneration, tbl, "DbColumnReader.EnsureCurrent"
    If mOwnerPtr = 0 Then Exit Sub
    If tbl.ColumnValueVersion(mColumnIndex) <> mValueVersion Then
        Err.Raise modDbContract.DB_ERR_STALE_GENERATION, "DbColumnReader.EnsureCurrent", _
            "Stale column snapshot: column values changed"
    End If
End Sub

Friend Property Get OwnerPtr() As LongPtr
    OwnerPtr = mOwnerPtr
End Property

Friend Property Get BoundGeneration() As Long
    BoundGeneration = mGeneration
End Property
