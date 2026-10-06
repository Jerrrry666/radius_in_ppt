Attribute VB_Name = "RadiusNativeRibbon"
Option Explicit

Private ribbon As Object
Private radiusValue As Double
Private radiusUnit As String

Public Sub NativeOnLoad(ByVal ui As Object)
    Set ribbon = ui
    radiusValue = 0.3
    radiusUnit = "cm"
End Sub

Public Sub NativeGetValue(ByVal control As Object, ByRef returnedValue)
    returnedValue = Format$(radiusValue, "0.00")
End Sub

Public Sub NativeChangeValue(ByVal control As Object, ByVal text As String)
    On Error GoTo Failed
    radiusValue = RadiusNativeCore.ParseNumber(text)
    Exit Sub
Failed:
    ReportError Err.Description
    If Not ribbon Is Nothing Then ribbon.InvalidateControl "NativeRadiusValue"
End Sub

Public Sub NativeGetUnit(ByVal control As Object, ByRef returnedValue)
    If radiusUnit = "%" Then returnedValue = 1 Else returnedValue = 0
End Sub

Public Sub NativeChangeUnit(ByVal control As Object, ByVal selectedId As String, ByVal selectedIndex As Integer)
    If selectedIndex = 1 Then radiusUnit = "%" Else radiusUnit = "cm"
End Sub

Public Sub NativeApply(ByVal control As Object)
    On Error GoTo Failed
    RadiusNativeCore.ApplySelection radiusValue, radiusUnit
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativePreset(ByVal control As Object)
    On Error GoTo Failed
    radiusValue = RadiusNativeCore.ParseNumber(control.Tag)
    radiusUnit = "cm"
    RadiusNativeCore.ApplySelection radiusValue, radiusUnit
    If Not ribbon Is Nothing Then ribbon.Invalidate
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeRead(ByVal control As Object)
    Dim leaves As Collection
    On Error GoTo Failed
    Set leaves = RadiusNativeCore.SelectionLeaves()
    radiusValue = RadiusNativeCore.CurrentRadius(leaves(1))
    radiusUnit = "cm"
    If Not ribbon Is Nothing Then ribbon.Invalidate
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeProtect(ByVal control As Object)
    On Error GoTo Failed
    RadiusNativeCore.SetProtection (control.Tag = "on")
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Public Sub NativeSelfTest(ByVal control As Object)
    On Error GoTo Failed
    RadiusNativeCore.SelfTest
    MsgBox "Native radius core: 4 checks passed. Host controls and group transactions still require Mac validation.", vbInformation, "R-corner Native"
    Exit Sub
Failed:
    ReportError Err.Description
End Sub

Private Sub ReportError(ByVal message As String)
    Debug.Print "[NativeRibbon] " & message
    MsgBox message, vbExclamation, "R-corner Native"
End Sub
