Attribute VB_Name = "TestNullTypeSemantics"
' @desc NULL/타입 의미론 현행 동작 회귀 고정 담당
' @depends UnivDB, LogShim
Option Explicit

' 비결정 D Phase D0 (docs/tasks/2026-07-08_null_type_semantics.md):
' 2026-07-07 조사 매트릭스의 각 칸을 실행 테스트로 고정한다.
' 여기의 기대값은 "현행 실동작"이지 "옳은 계약"이 아니다.
' D1(Empty NULL 계약)/D2(집계 정합)에서 기대값 변경은 별도 커밋으로 갱신된다.
' 전부 dict 경로만 사용 (IndexBackend 미지정 = 기본 DictHashIndex).

' ==============================================================================
' [1] Empty (빈 셀)
' ==============================================================================

' D1 계약 (Empty=NULL): Filter 숫자 경로에서 Empty는 불참 (0으로 참여하지 않음)
Public Sub Test_NullSem_EmptyNumericFilterExcluded()
    Dim T As clsTable
    Dim tR As clsTable
    Dim vData(1 To 4, 1 To 1) As Variant

    vData(1, 1) = "Score"
    vData(2, 1) = 10
    ' vData(3, 1) 미할당 = Empty
    vData(4, 1) = -5

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' Empty 불참: 10 < 100, -5 < 100 => 2행 (Empty 제외)
    Set tR = UnivDB.Query(T).Filter("Score", "<", 100).SelectTable()
    LogShim.AssertEq "nullsem_empty_numfilter_lt100", 2, tR.RowCount - 1

    ' Empty는 0으로 매칭되지 않음 => 0행
    Set tR = UnivDB.Query(T).Filter("Score", "=", 0).SelectTable()
    LogShim.AssertEq "nullsem_empty_numfilter_eq0", 0, tR.RowCount - 1
End Sub

' 매트릭스: Filter 문자 경로(IS)에서 Empty는 ""로 참여, 인덱스 키도 ""
Public Sub Test_NullSem_EmptyStringFilterAsBlank()
    Dim T As clsTable
    Dim tR As clsTable
    Dim vData(1 To 4, 1 To 1) As Variant

    vData(1, 1) = "Name"
    vData(2, 1) = "A"
    ' vData(3, 1) 미할당 = Empty
    vData(4, 1) = "B"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tR = UnivDB.Query(T).Filter("Name", "IS", "").SelectTable()
    LogShim.AssertEq "nullsem_empty_strfilter_is_blank", 1, tR.RowCount - 1

    ' 인덱스 키 경로도 동일: Empty -> "" 키
    LogShim.AssertEq "nullsem_empty_indexkey_blank", 1, T.SelectRows("Name", "").Count
End Sub

' D2 계약: Filter 문자 경로의 Trim 대칭화. 비교값(target)은 FastNorm으로 Trim되므로 셀 쪽도
' 대칭 정규화한다 - 셀 앞뒤 공백이 있어도 IS 매칭이 성립한다 (D2 이전엔 raw CStr라 불일치).
Public Sub Test_NullSem_StringFilterTrimSymmetry()
    Dim T As clsTable
    Dim tR As clsTable
    Dim vData(1 To 4, 1 To 1) As Variant

    vData(1, 1) = "Name"
    vData(2, 1) = "  A  "  ' 앞뒤 공백
    vData(3, 1) = "B"
    vData(4, 1) = "A"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' 셀 " A " 와 "A" 모두 IS "A" 에 매칭 => 2행
    Set tR = UnivDB.Query(T).Filter("Name", "IS", "A").SelectTable()
    LogShim.AssertEq "nullsem_strfilter_trim_is_a", 2, tR.RowCount - 1
End Sub

' D1 계약: 집계에서 Empty는 제외 (NULL처럼) - COUNT/SUM/AVG 모두 불참.
' 위험 지점 1 해소: 숫자 필터도 Empty 불참으로 통일돼 필터 행수와 COUNT가 정합한다.
Public Sub Test_NullSem_EmptyAggregateAndFilterConsistent()
    Dim T As clsTable
    Dim tG As clsTable
    Dim tR As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = "IT" ' vData(3, 2) 미할당 = Empty
    vData(4, 1) = "IT": vData(4, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tG = T.GroupBy("Dept", "Score:COUNT, Score:SUM, Score:AVG")
    LogShim.AssertEq "nullsem_empty_agg_count", 2, tG.Value(2, "Score_COUNT")
    LogShim.AssertEq "nullsem_empty_agg_sum", 40, tG.Value(2, "Score_SUM")
    LogShim.AssertEq "nullsem_empty_agg_avg", 20, tG.Value(2, "Score_AVG")

    ' 정합 고정: 필터도 2행(Empty 제외) = COUNT 2
    Set tR = UnivDB.Query(T).Filter("Score", "<", 100).SelectTable()
    LogShim.AssertEq "nullsem_empty_filter_vs_count", 2, tR.RowCount - 1
End Sub

' 매트릭스: OrderBy 숫자열에서 Empty는 0 취급 (CDbl(Empty)=0으로 숫자 비교)
Public Sub Test_NullSem_EmptyOrderByAsZero()
    Dim T As clsTable
    Dim vData(1 To 4, 1 To 1) As Variant

    vData(1, 1) = "N"
    vData(2, 1) = 5
    ' vData(3, 1) 미할당 = Empty
    vData(4, 1) = -3

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    T.OrderBy "N"
    ' 오름차순: -3, Empty(=0), 5
    LogShim.AssertEq "nullsem_empty_orderby_first", -3, T.Value(2, "N")
    LogShim.AssertEq "nullsem_empty_orderby_mid", "", T.Value(3, "N")
    LogShim.AssertEq "nullsem_empty_orderby_last", 5, T.Value(4, "N")
End Sub

' ==============================================================================
' [2] Error (#N/A 등)
' ==============================================================================

' 매트릭스: Filter 숫자 경로에서 Error는 제외 (IsNumeric(vbError)=False)
Public Sub Test_NullSem_ErrorNumericFilterExcluded()
    Dim T As clsTable
    Dim tR As clsTable
    Dim vData(1 To 4, 1 To 1) As Variant

    vData(1, 1) = "Score"
    vData(2, 1) = 10
    vData(3, 1) = CVErr(xlErrNA)
    vData(4, 1) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tR = UnivDB.Query(T).Filter("Score", ">", 0).SelectTable()
    LogShim.AssertEq "nullsem_err_numfilter_gt0", 2, tR.RowCount - 1
End Sub

' D1 계약 (Error=skip): Filter 문자 경로에서 Error 셀은 어느 연산에도 참여하지 않는다.
' (D0에서 raw CStr(vbError)가 크래시 대신 locale 문자열로 은밀히 참여함을 실증 -> D1에서 skip)
Public Sub Test_NullSem_ErrorStringFilterSkipped()
    Dim T As clsTable
    Dim tR As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 4, 1 To 1) As Variant

    vData(1, 1) = "Name"
    vData(2, 1) = "A"
    vData(3, 1) = CVErr(xlErrNA)
    vData(4, 1) = "B"

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    On Error Resume Next
    Set tR = UnivDB.Query(T).Filter("Name", "IS", "a").SelectTable()
    lErrNo = Err.Number
    On Error GoTo 0

    ' 크래시 없음, "a"엔 Error 행 불매칭 => 1행 (A)
    LogShim.AssertEq "nullsem_err_strfilter_nocrash", 0, lErrNo
    LogShim.AssertEq "nullsem_err_strfilter_is_a", 1, tR.RowCount - 1

    ' Error skip 고정: 부정 연산에서도 Error 행은 결과에 안 들어온다.
    ' "A"만 제외 => B 1행만 (Error 제외). D1 전에는 2행(B+Error)이었다.
    Set tR = UnivDB.Query(T).Filter("Name", "IS NOT", "A").SelectTable()
    LogShim.AssertEq "nullsem_err_strfilter_isnot_a", 1, tR.RowCount - 1
End Sub

' D1 계약 (Error=safe): OrderBy에서 Error 셀은 결정적으로 맨 뒤에 배치된다.
' (D0 현행도 우연히 "Error last"였고 D1이 이를 locale-무관/크래시-무관으로 보장 - assertion 불변)
Public Sub Test_NullSem_ErrorOrderBySortsLast()
    Dim T As clsTable
    Dim lErrNo As Long
    Dim vData(1 To 4, 1 To 1) As Variant

    vData(1, 1) = "N"
    vData(2, 1) = 10
    vData(3, 1) = CVErr(xlErrNA)
    vData(4, 1) = 5

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    On Error Resume Next
    T.OrderBy "N"
    lErrNo = Err.Number
    On Error GoTo 0

    ' 현행 고정: 크래시 없음, 행 보존(Error 미탈락), 숫자는 숫자끼리 정상 정렬
    LogShim.AssertEq "nullsem_err_orderby_nocrash", 0, lErrNo
    LogShim.AssertEq "nullsem_err_orderby_rowcount", 3, T.RowCount - 1
    LogShim.AssertEq "nullsem_err_orderby_r2", 5, T.Value(2, "N")
    LogShim.AssertEq "nullsem_err_orderby_r3", 10, T.Value(3, "N")
    ' Error 행은 텍스트로 맨 뒤 (locale-무관: 숫자가 문자보다 앞서 정렬)
    LogShim.AssertEq "nullsem_err_orderby_r4_iserr", True, IsError(T.Value(4, "N"))
    T.OrderBy "N", True
    LogShim.AssertEq "nullsem_err_desc_first", 10, T.Value(2, "N")
    LogShim.AssertEq "nullsem_err_desc_second", 5, T.Value(3, "N")
    LogShim.AssertEq "nullsem_err_desc_last", True, IsError(T.Value(4, "N"))
End Sub

' D2 계약 (Error=skip): 집계 컬럼의 Error 셀은 COUNT 포함 어느 집계에도 불참한다.
' COUNT=2(Error 제외), SUM=40, AVG=40/2=20(분모=숫자로 가산된 셀 수, 희석 없음)
Public Sub Test_NullSem_ErrorAggregateExcluded()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = "IT": vData(3, 2) = CVErr(xlErrNA)
    vData(4, 1) = "IT": vData(4, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    Set tG = T.GroupBy("Dept", "Score:COUNT, Score:SUM, Score:AVG")
    LogShim.AssertEq "nullsem_err_agg_count", 2, tG.Value(2, "Score_COUNT")
    LogShim.AssertEq "nullsem_err_agg_sum", 40, tG.Value(2, "Score_SUM")
    LogShim.AssertEq "nullsem_err_agg_avg", 20, tG.Value(2, "Score_AVG")
End Sub

' D2 계약 (Error=skip): GroupBy MIN/MAX는 Error 셀을 제외한다(raw Variant 비교 크래시 차단).
' D1(pin 4ef7ba7)에선 Error가 compare에 닿아 Type Mismatch(13) 크래시였다 - 이제 errno=0.
Public Sub Test_NullSem_MinMaxErrorSkipped()
    Dim T As clsTable
    Dim tG As clsTable
    Dim lErrNoMax As Long
    Dim lErrNoMin As Long
    Dim vMax As Variant
    Dim vMin As Variant
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = "IT": vData(3, 2) = CVErr(xlErrNA)
    vData(4, 1) = "IT": vData(4, 2) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    On Error Resume Next
    Set tG = T.GroupBy("Dept", "Score:MAX")
    lErrNoMax = Err.Number
    vMax = tG.Value(2, "Score_MAX")
    Err.Clear
    Set tG = T.GroupBy("Dept", "Score:MIN")
    lErrNoMin = Err.Number
    vMin = tG.Value(2, "Score_MIN")
    On Error GoTo 0

    ' 크래시 없음 + Error 제외한 실값으로 MIN/MAX 산출
    LogShim.AssertEq "nullsem_minmax_error_max_errno", 0, lErrNoMax
    LogShim.AssertEq "nullsem_minmax_error_min_errno", 0, lErrNoMin
    LogShim.AssertEq "nullsem_minmax_error_max_val", 30, vMax
    LogShim.AssertEq "nullsem_minmax_error_min_val", 10, vMin
End Sub

' ==============================================================================
' [3] 숫자 42 vs 문자 "42"
' ==============================================================================

' 매트릭스: 인덱스 키는 CStr 병합(같은 키), 숫자 필터는 CDbl로 둘 다 참여, SUM도 합산
Public Sub Test_NullSem_NumericStringMerge()
    Dim T As clsTable
    Dim tR As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Score"
    vData(2, 1) = "IT": vData(2, 2) = 42
    vData(3, 1) = "IT": vData(3, 2) = "42"
    vData(4, 1) = "IT": vData(4, 2) = 7

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' 인덱스 키: 42와 "42"는 같은 키
    LogShim.AssertEq "nullsem_numstr_indexkey_merge", 2, T.SelectRows("Score", 42).Count

    ' 숫자 필터: 둘 다 CDbl로 참여
    Set tR = UnivDB.Query(T).Filter("Score", "=", 42).SelectTable()
    LogShim.AssertEq "nullsem_numstr_numfilter_eq", 2, tR.RowCount - 1

    ' SUM: "42"도 IsNumeric이라 가산
    Set tG = T.GroupBy("Dept", "Score:SUM")
    LogShim.AssertEq "nullsem_numstr_agg_sum", 91, tG.Value(2, "Score_SUM")
End Sub

' 숫자와 문자열의 비교 순서는 입력 순열과 무관하다.
Public Sub Test_NullSem_MixedOrderByPermutation()
    Dim T As New clsTable
    Dim vData(1 To 4, 1 To 1) As Variant
    Dim vals As Variant, a As Long, b As Long, c As Long, d As Long
    vals = Array(2, 10, "15")
    vData(1, 1) = "N"
    For a = 0 To 2
        For b = 0 To 2
            If a <> b Then
                c = 3 - a - b
                vData(2, 1) = vals(a): vData(3, 1) = vals(b): vData(4, 1) = vals(c)
                For d = 0 To 1
                    T.LoadFromArray vData
                    T.OrderBy "N", CBool(d)
                    LogShim.AssertEq "sort_perm_first", vals(2 * d), T.Value(2, "N")
                    LogShim.AssertEq "sort_perm_middle", vals(1), T.Value(3, "N")
                    LogShim.AssertEq "sort_perm_last", vals(2 - 2 * d), T.Value(4, "N")
                Next d
            End If
        Next b
    Next a
End Sub

' 문자열부터 누적해도 단일/복합 그룹의 SUM/AVG는 순서 불변이다.
Public Sub Test_NullSem_AggregatePermutation()
    Dim T As New clsTable, result As clsTable
    Dim vData(1 To 4, 1 To 3) As Variant
    Dim vals As Variant, groupCols As Variant, a As Long, b As Long, c As Long, i As Long
    vals = Array("10", "20", 30)
    vData(1, 1) = "Dept": vData(1, 2) = "Tag": vData(1, 3) = "Amt"
    For i = 2 To 4
        vData(i, 1) = "IT": vData(i, 2) = "X"
    Next i
    For a = 0 To 2
        For b = 0 To 2
            If a <> b Then
                c = 3 - a - b
                vData(2, 3) = vals(a): vData(3, 3) = vals(b): vData(4, 3) = vals(c)
                T.LoadFromArray vData
                For Each groupCols In Array("Dept", "Dept,Tag")
                    Set result = T.GroupBy(CStr(groupCols), "Amt:SUM,Amt:AVG")
                    LogShim.AssertEq "agg_perm_sum", 60, result.Value(2, "Amt_SUM")
                    LogShim.AssertEq "agg_perm_avg", 20, result.Value(2, "Amt_AVG")
                    LogShim.AssertEq "agg_perm_numeric_type", vbDouble, VarType(result.Value(2, "Amt_SUM"))
                Next groupCols
            End If
        Next b
    Next a
End Sub

' ==============================================================================
' [4] 날짜
' ==============================================================================
' D2 계약: 숫자 필터에서 날짜는 제외(IsNumeric(Date)=False), 집계에선 SUM 불참·COUNT 포함이나
' AVG 분모는 숫자 셀 수라 희석되지 않는다 (D1 이전엔 AVG 희석의 세 번째 경로였음)
Public Sub Test_NullSem_DateNumericExcludedButCounted()
    Dim T As clsTable
    Dim tR As clsTable
    Dim tG As clsTable
    Dim vData(1 To 3, 1 To 2) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "Amount"
    vData(2, 1) = "IT": vData(2, 2) = 10
    vData(3, 1) = "IT": vData(3, 2) = DateSerial(2026, 7, 8)

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' 숫자 필터: 날짜 셀 제외
    Set tR = UnivDB.Query(T).Filter("Amount", ">", 0).SelectTable()
    LogShim.AssertEq "nullsem_date_numfilter_gt0", 1, tR.RowCount - 1

    ' 집계(D2): 날짜는 실값이라 COUNT엔 포함(2), SUM엔 불참(비숫자), AVG 분모=숫자 셀 수(1)
    ' => AVG = 10/1 = 10 (날짜에 의한 희석 없음)
    Set tG = T.GroupBy("Dept", "Amount:COUNT, Amount:SUM, Amount:AVG")
    LogShim.AssertEq "nullsem_date_agg_count", 2, tG.Value(2, "Amount_COUNT")
    LogShim.AssertEq "nullsem_date_agg_sum", 10, tG.Value(2, "Amount_SUM")
    LogShim.AssertEq "nullsem_date_agg_avg", 10, tG.Value(2, "Amount_AVG")
End Sub

' D3 계약 (Date=numeric): OrderBy에서 Date 두 값은 직렬값 수치 비교로 시간순 정렬된다.
' (D3 이전엔 CStr 텍스트 fallback - "9:00" vs "10:00", 연도 교차 등에서 locale 무관하게
'  텍스트 순서가 시간순과 갈렸다. 이 데이터셋은 텍스트 정렬이면 실패하도록 구성됨)
Public Sub Test_NullSem_DateOrderByChronological()
    Dim T As clsTable
    Dim vData(1 To 5, 1 To 1) As Variant

    vData(1, 1) = "D"
    vData(2, 1) = DateSerial(2026, 1, 9)
    vData(3, 1) = DateSerial(2025, 10, 1)
    vData(4, 1) = DateSerial(2025, 2, 3) + TimeSerial(10, 0, 0)
    vData(5, 1) = DateSerial(2025, 2, 3) + TimeSerial(9, 0, 0)

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    T.OrderBy "D"
    ' 시간순 오름차순: 02-03 09:00 < 02-03 10:00 < 10-01 < 2026-01-09
    LogShim.AssertEq "nullsem_date_orderby_r2", DateSerial(2025, 2, 3) + TimeSerial(9, 0, 0), T.Value(2, "D")
    LogShim.AssertEq "nullsem_date_orderby_r3", DateSerial(2025, 2, 3) + TimeSerial(10, 0, 0), T.Value(3, "D")
    LogShim.AssertEq "nullsem_date_orderby_r4", DateSerial(2025, 10, 1), T.Value(4, "D")
    LogShim.AssertEq "nullsem_date_orderby_r5", DateSerial(2026, 1, 9), T.Value(5, "D")
End Sub

' D3 계약 (Date key=ISO): 날짜 인덱스 키는 canonical ISO - locale 무관, 같은 Date 값으로
' 항상 재조회 가능 (roundtrip 불변식은 CStr 시절부터 유지)
Public Sub Test_NullSem_DateIndexKeySelfConsistent()
    Dim T As clsTable
    Dim vData(1 To 3, 1 To 1) As Variant

    vData(1, 1) = "D"
    vData(2, 1) = DateSerial(2026, 7, 8)
    vData(3, 1) = DateSerial(2026, 1, 1)

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    LogShim.AssertEq "nullsem_date_indexkey_roundtrip", 1, T.SelectRows("D", DateSerial(2026, 7, 8)).Count
End Sub

' D3 계약 (Date key=ISO canonical): 인덱스/GroupBy 날짜 키는 locale 무관 ISO 문자열.
' 시간 성분 0이면 "yyyy-mm-dd", 있으면 "yyyy-mm-dd hh:nn:ss". GroupBy 렌더 표기도 동일
' (단일 컬럼 경로=cur.Key, 복합 키 경로=NormalizeIndexKey 렌더 파트 - 양 경로 검증).
Public Sub Test_NullSem_DateIndexKeyIsoCanonical()
    Dim T As clsTable
    Dim tG As clsTable
    Dim vData(1 To 4, 1 To 3) As Variant

    vData(1, 1) = "Dept": vData(1, 2) = "D": vData(1, 3) = "Amount"
    vData(2, 1) = "IT": vData(2, 2) = DateSerial(2026, 7, 8): vData(2, 3) = 10
    vData(3, 1) = "IT": vData(3, 2) = DateSerial(2026, 7, 8): vData(3, 3) = 20
    vData(4, 1) = "IT": vData(4, 2) = DateSerial(2026, 7, 8) + TimeSerial(9, 5, 0): vData(4, 3) = 30

    Set T = UnivDB.CreateTable()
    T.LoadFromArray vData

    ' 단일 컬럼 경로 (인덱스 커서): 날짜-only와 datetime은 다른 키, 렌더는 ISO
    Set tG = T.GroupBy("D", "Amount:SUM")
    LogShim.AssertEq "nullsem_date_iso_groupby_rows", 2, tG.RowCount - 1
    LogShim.AssertEq "nullsem_date_iso_render_dateonly", "2026-07-08", tG.Value(2, "D")
    LogShim.AssertEq "nullsem_date_iso_render_datetime", "2026-07-08 09:05:00", tG.Value(3, "D")
    LogShim.AssertEq "nullsem_date_iso_sum_dateonly", 30, tG.Value(2, "Amount_SUM")

    ' 복합 키 경로 (EncodeCompositeKeyPart -> NormalizeIndexKey): 동일 canonical 경유
    Set tG = T.GroupBy("Dept, D", "Amount:SUM")
    LogShim.AssertEq "nullsem_date_iso_composite_rows", 2, tG.RowCount - 1
    LogShim.AssertEq "nullsem_date_iso_composite_render", "2026-07-08", tG.Value(2, "D")

    ' 인덱스 라운드트립: datetime 값도 같은 Date로 재조회 가능
    LogShim.AssertEq "nullsem_date_iso_roundtrip_dt", 1, _
        T.SelectRows("D", DateSerial(2026, 7, 8) + TimeSerial(9, 5, 0)).Count
End Sub
