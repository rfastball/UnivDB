Attribute VB_Name = "SmokeTest"
Option Explicit
Public Sub Test_Smoke()
    LogShim.AssertEq "one_equals_one", 1, 1
End Sub
