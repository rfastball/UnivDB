VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ========================================================
' Class: clsSet
' 설명: 대용량 데이터 처리를 위한 고속 집합 연산 엔진
' 특징: Flagging 기법 및 Direct Access를 통한 극한의 최적화
' ========================================================

' [Enum] 반환 모드 정의
Public Enum SetRetMode
    srIntersect = 1 ' 교집합 (A ∩ B)
    srDiffMe = 2    ' 내 차집합 (A - B): 미사용 데이터
    srDiffOther = 3 ' 상대 차집합 (B - A): 데이터 없음(오류)
    srAllSplit = 4  ' [0]교집합, [1]내차집합, [2]상대차집합 (배열 반환)
End Enum

Private mDict As Object

Private Sub Class_Initialize()
    Set mDict = CreateObject("Scripting.Dictionary")
    mDict.CompareMode = vbTextCompare
End Sub

Private Sub Class_Terminate()
    Set mDict = Nothing
End Sub

' ========================================================
' [Basic] 기본 데이터 관리
' ========================================================

' 초기화 (재사용 시 필수)
Public Sub Clear()
    mDict.RemoveAll
End Sub

' 개별 추가
Public Sub Add(ByVal vItem As Variant)
    If Not IsEmpty(vItem) Then
        If Not mDict.exists(vItem) Then mDict(vItem) = Empty
    End If
End Sub

' 개별 삭제
Public Sub Remove(ByVal vItem As Variant)
    If mDict.exists(vItem) Then mDict.Remove vItem
End Sub

' 존재 여부 확인
Public Function Contains(ByVal vItem As Variant) As Boolean
    Contains = mDict.exists(vItem)
End Function

' 개수 확인
Public Property Get Count() As Long
    Count = mDict.Count
End Property

' [Loader] 1차원 배열 고속 적재
' 최적화: 함수 호출 오버헤드를 줄이기 위해 내부 로직 인라인 처리
Public Sub LoadFromArray(ByVal vArr As Variant)
    Dim i As Long
    Dim vVal As Variant
    
    If Not IsArray(vArr) Then Me.Add vArr: Exit Sub
    
    For i = LBound(vArr) To UBound(vArr)
        vVal = vArr(i)
        ' 빈 값 및 에러 값 방어
        If Not IsEmpty(vVal) And Not IsError(vVal) Then
            mDict(vVal) = Empty
        End If
    Next i
End Sub

' ========================================================
' [Core] 통합 비교 엔진 (Compare)
' 설명: Single-Pass 알고리즘으로 교집합/차집합을 동시 분류
' 최적화: Dictionary 직접 접근으로 함수 호출 비용 제거
' ========================================================
Public Function Compare(ByVal otherSet As clsSet, _
                        Optional ByVal mode As SetRetMode = srAllSplit) As Variant
    
    ' 결과 객체
    Dim setCommon As clsSet, setDiffOther As clsSet, setDiffMe As clsSet
    ' 직접 접근용 변수 (속도 향상 핵심)
    Dim dictCommon As Object, dictDiffOther As Object, dictDiffMe As Object
    
    Dim vKey As Variant, arrOther As Variant
    Dim i As Long
    
    ' 1. 상태 초기화: 기존 마킹(Flag) 제거 (O(N))
    For Each vKey In mDict.Keys
        mDict(vKey) = Empty
    Next vKey
    
    ' 2. 객체 준비 (Lazy Initialization)
    If mode = srIntersect Or mode = srAllSplit Then
        Set setCommon = New clsSet: Set dictCommon = setCommon.InternalDict
    End If
    If mode = srDiffOther Or mode = srAllSplit Then
        Set setDiffOther = New clsSet: Set dictDiffOther = setDiffOther.InternalDict
    End If
    If mode = srDiffMe Or mode = srAllSplit Then
        Set setDiffMe = New clsSet: Set dictDiffMe = setDiffMe.InternalDict
    End If

    ' 3. [Primary Loop] 입력값 순회 -> 교집합 & 상대 차집합 결정
    arrOther = otherSet.KeysArray
    If UBound(arrOther) >= LBound(arrOther) Then
        For i = LBound(arrOther) To UBound(arrOther)
            vKey = arrOther(i)
            
            If mDict.exists(vKey) Then
                ' [매칭 성공] 마킹(True) 및 교집합 등록
                mDict(vKey) = True
                If Not dictCommon Is Nothing Then dictCommon(vKey) = Empty
            Else
                ' [매칭 실패] 상대 차집합 등록
                If Not dictDiffOther Is Nothing Then dictDiffOther(vKey) = Empty
            End If
        Next i
    End If
    
    ' 4. [Secondary Loop] 내 장부 순회 -> 내 차집합(미사용) 결정
    If Not dictDiffMe Is Nothing Then
        For Each vKey In mDict.Keys
            ' 마킹되지 않은(Empty) 항목만 추출
            If IsEmpty(mDict(vKey)) Then dictDiffMe(vKey) = Empty
        Next vKey
    End If
    
    ' 5. 결과 반환
    Select Case mode
        Case srIntersect: Set Compare = setCommon
        Case srDiffOther: Set Compare = setDiffOther
        Case srDiffMe:    Set Compare = setDiffMe
        Case srAllSplit:  Compare = Array(setCommon, setDiffMe, setDiffOther)
    End Select
    
End Function

' ========================================================
' [Operations] 집합 연산 래퍼 & 유틸리티
' ========================================================

' 교집합 (A ∩ B)
Public Function Intersection(ByVal otherSet As clsSet) As clsSet
    Set Intersection = Compare(otherSet, srIntersect)
End Function

' 내 차집합 (A - B, 미사용 데이터)
Public Function Difference(ByVal otherSet As clsSet) As clsSet
    Set Difference = Compare(otherSet, srDiffMe)
End Function

' 상대 차집합 (B - A, 없는 데이터)
Public Function DifferenceInput(ByVal otherSet As clsSet) As clsSet
    Set DifferenceInput = Compare(otherSet, srDiffOther)
End Function

' 합집합 (A ∪ B) - Zero Overhead Version
Public Function Union(ByVal otherSet As clsSet) As clsSet
    Dim NewSet As New clsSet
    Dim TargetDict As Object
    Dim vArr As Variant, i As Long
    
    Set TargetDict = NewSet.InternalDict
    
    ' 내 데이터 주입 (검증 생략)
    If mDict.Count > 0 Then
        vArr = mDict.Keys
        For i = 0 To UBound(vArr)
            TargetDict(vArr(i)) = Empty
        Next i
    End If
    
    ' 상대 데이터 주입
    If otherSet.Count > 0 Then
        vArr = otherSet.KeysArray
        For i = 0 To UBound(vArr)
            TargetDict(vArr(i)) = Empty
        Next i
    End If
    
    Set Union = NewSet
End Function

' 배열로 반환
Public Function ToArray() As Variant
    If mDict.Count = 0 Then ToArray = Array() Else ToArray = mDict.Keys
End Function

' [Internal] 외부 모듈 직접 접근 허용 (최적화용)
Public Property Get InternalDict() As Object
    Set InternalDict = mDict
End Property

' [Internal] 키 배열 직접 접근 허용
Public Property Get KeysArray() As Variant
    If mDict.Count = 0 Then KeysArray = Array() Else KeysArray = mDict.Keys
End Property


