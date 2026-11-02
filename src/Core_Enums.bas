Attribute VB_Name = "Core_Enums"
' @desc UnivCore 공통 Enum/상수 선언
Option Explicit

' 은퇴한 토글: 해시가 순수 VBA 단일 경로(byte 배열 + 폴딩 테이블)로 통일되어
' 더 이상 참조되지 않음. 외부 소비자 호환을 위해 선언만 유지.
Public Const CORE_USE_UNSAFE_MEM As Long = 1

Public Enum CoreSetRetMode
    srIntersect = 1
    srDiffMe = 2
    srDiffOther = 3
    srAllSplit = 4
End Enum
