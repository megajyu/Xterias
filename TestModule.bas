Attribute VB_Name = "TestModule"
Option Explicit

Public Sub Test_ArCeil()
    Dim tRg As Range
    Set tRg = ArCeil(Selection)
    Debug.Print tRg.Address
End Sub
