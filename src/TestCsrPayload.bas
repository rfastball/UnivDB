Attribute VB_Name = "TestCsrPayload"
' @desc CSR 페이로드 계약(누적/finalize/재적재/브리지) 검증 담당
' @depends LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' DbCsrPayload 단독 유닛 스위트. keyId는 백엔드가 1기반 밀집으로 부여한다는 전제를
' 그대로 사용한다 (dict = Count+1 순차, core = FastHash slot).

Public Sub Test_CsrPayload_BasicSpans()
    Dim p As New DbCsrPayload
    Dim buf() As Long
    Dim n As Long

    ' 키 3종 교차 적재 - 키 내 순서 보존(안정 scatter) 검증
    p.AddEntry 1, 10
    p.AddEntry 2, 20
    p.AddEntry 1, 11
    p.AddEntry 3, 30
    p.AddEntry 2, 21
    p.AddEntry 1, 12

    LogShim.AssertEq "csr_keycount", 3, p.KeyCount
    LogShim.AssertEq "csr_entrycount", 6, p.EntryCount
    LogShim.AssertEq "csr_count_k1", 3, p.CountOf(1)
    LogShim.AssertEq "csr_count_k2", 2, p.CountOf(2)
    LogShim.AssertEq "csr_count_k3", 1, p.CountOf(3)
    LogShim.AssertEq "csr_first_k2", 20, p.FirstRowId(2)

    p.CopySpan 1, buf, n
    LogShim.AssertEq "csr_span_k1_n", 3, n
    LogShim.AssertEq "csr_span_k1_1", 10, buf(1)
    LogShim.AssertEq "csr_span_k1_2", 11, buf(2)
    LogShim.AssertEq "csr_span_k1_3", 12, buf(3)
End Sub

Public Sub Test_CsrPayload_DirtyRefinalize()
    Dim p As New DbCsrPayload
    Dim buf() As Long
    Dim n As Long

    p.AddEntry 1, 10
    LogShim.AssertEq "csr_dirty_before", 1, p.CountOf(1) ' 첫 읽기 = finalize 발동

    p.AddEntry 1, 11 ' 읽기 후 재적재 -> dirty
    LogShim.AssertEq "csr_dirty_after_count", 2, p.CountOf(1)
    p.CopySpan 1, buf, n
    LogShim.AssertEq "csr_dirty_order_1", 10, buf(1)
    LogShim.AssertEq "csr_dirty_order_2", 11, buf(2)
End Sub

Public Sub Test_CsrPayload_ClearAndBufferReuse()
    Dim p As New DbCsrPayload
    Dim q As New DbCsrPayload
    Dim buf() As Long
    Dim n As Long

    p.AddEntry 1, 10
    p.Clear
    LogShim.AssertEq "csr_clear_keycount", 0, p.KeyCount
    LogShim.AssertEq "csr_clear_entrycount", 0, p.EntryCount

    ' Clear 후 재적재 가능
    p.AddEntry 1, 99
    LogShim.AssertEq "csr_reload_first", 99, p.FirstRowId(1)

    ' 버퍼 재사용 계약: 큰 span 뒤 작은 span - 소비는 outCount 기준
    q.AddEntry 1, 1
    q.AddEntry 1, 2
    q.AddEntry 1, 3
    q.AddEntry 2, 7
    q.CopySpan 1, buf, n
    LogShim.AssertEq "csr_buf_big_n", 3, n
    q.CopySpan 2, buf, n
    LogShim.AssertEq "csr_buf_small_n", 1, n
    LogShim.AssertEq "csr_buf_small_v", 7, buf(1)
End Sub

Public Sub Test_CsrPayload_RowListBridge()
    Dim p As New DbCsrPayload
    Dim lst As DbRowList

    p.AddEntry 1, 5
    p.AddEntry 1, 8
    Set lst = p.MakeRowList(1, 0, 0) ' ownerPtr 0 = 미결속 (계약 검증 전용)
    LogShim.AssertEq "csr_rowlist_count", 2, lst.Count
    LogShim.AssertEq "csr_rowlist_first", 5, lst.RowAt(1)
    LogShim.AssertEq "csr_rowlist_second", 8, lst.RowAt(2)
End Sub
