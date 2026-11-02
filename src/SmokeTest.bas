Attribute VB_Name = "SmokeTest"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit
Public Sub Test_Smoke()
    LogShim.AssertEq "one_equals_one", 1, 1
End Sub
